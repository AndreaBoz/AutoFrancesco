"""Dalle fatture di acquisto dei fornitori diversi da FLOTT ai registri di accettazione
degli esterni e al loro registro giacenze.

Cambia il fornitore e cambia l'impaginazione della fattura: chi scrive il lotto in una
riga a parte ("Tipo dato: LOTTO"), chi lo mette sotto l'articolo ("lotto: 37 13/04/26"),
chi non lo scrive affatto. Quello che si fa dei dati, invece, e' lo stesso di FLOTT.
La scheda RIMANENZE del registro giacenze non viene mai toccata.
"""
from __future__ import annotations

import argparse
import datetime as dt
import re
import sys
from dataclasses import dataclass
from pathlib import Path

import openpyxl
import pymupdf

import acquisti
from acquisti import (GestoreGiacenze, Imballaggio, Registro, _apri, _nuovo_foglio,
                      _ordina, _suffissi, leggi_visti, risolvi, scrivi_visti, sottolinea)
from riepilogo import elenca, leggi_formati, plurale, salva_libro, FORMATI

RADICE = Path(__file__).resolve().parent

NOMI_REGISTRI = {
    "CAM": "EST-Controllo in accettazione MAT.PRIME e SEMILAVORATI.xlsx",
    "CAO": "EST-Controllo in accettazione OLIO.xlsx",
    "CAI": "EST-Controllo imballaggi primari.xlsx",
}
ETICHETTA = {"CAM": "pesce", "CAO": "olio", "CAI": "vasi"}

# Materiale di confezionamento e imballi di ritorno che non vanno nei registri.
IGNORA_IN_SILENZIO = ("VASSOIO", "VASSOI", "ETICHETTA", "ETICHETTE", "SIGILLO", "SIGILLI",
                      "SCATOLA", "SCATOLE", "PELLICOLA", "INTERFALDA", "INTERFALDE",
                      "BANCALE", "BANCALI", "PEDANA", "PEDANE", "CARTONE", "CARTONI")

# Le alici comprate fresche non portano scritta la zona di pesca: sono pescato locale.
ZONA_PESCATO_FRESCO = "37"
ZONE_GIACENZE = ("87", "37", "34", "27")
# il vaso di plastica da 40 ML sta nella scheda VASI 42, l'orcio da 314 ML in quella da 300
ALIAS_VASI = {40: 42, 314: 300}


# --------------------------------------------------------------------------- #
# lettura della fattura
# --------------------------------------------------------------------------- #
# Righe di servizio: riferimenti al DDT, codici articolo, righe di lotto, imballi resi,
# intestazioni di colonna e pie' di pagina del browser da cui si stampa il PDF.
RUMORE = re.compile(r"""^(?:
      DDT\s+\S+\s+del\s+[\d/.-]+
    | -{3,}
    | -\s*Riferimento\s+D\.?\s?d\.?\s?T\..*
    | da\s+DDT\s+nr\..*
    | \(Codice\s+Fornitore\)
    | Tipo\s+dato:.*
    | Rif\.\s*(?:testo|data|numero):.*
    | lotto:.*
    | \d{2}/\d{2}/\d{4}
    | \(.*\)
    | -\s*RESI:.*
    | Il\s+succitato.*|con\s+gli\s+alimenti.*|90/12/CEE.*|agg\.
    | Cod\.\s*articolo|Descrizione|Quantità|Prezzo\s+unitario|UM|Sconto\s+o|magg\.|%IVA
    | Prezzo\s+totale|Firefox|Chrome|Safari|https?://.*|\d+\s+di\s+\d+
    | \d{2}/\d{2}/\d{2},\s*\d{2}:\d{2}
    )$""", re.IGNORECASE | re.VERBOSE)
INIZIO_ARTICOLO = re.compile(r"^(?:DDT\s+\S+\s+del\s|-\s*Riferimento\s+D|da\s+DDT\s+nr\.)", re.I)
NUMERO = re.compile(r"^[\d.]+,\d+$")
UNITA = re.compile(r"^(?:PZ|NR|KG|LT|MT|CT|N)$", re.IGNORECASE)


@dataclass
class RigaFattura:
    descrizione: str
    quantita: float
    unita: str | None = None
    lotto: str | None = None
    scadenza: dt.date | None = None


@dataclass
class Fattura:
    numero: str            # senza la parte dopo la barra: 317/25 -> 317
    data: dt.date
    fornitore: str
    righe: list[RigaFattura]
    file: Path

    @property
    def base(self) -> str:
        return f"{self.numero}_{self.data:%d%m%y}"


CHIAVI = ("Regime fiscale:", "Indirizzo:", "Comune:", "Cap:", "Telefono:", "Email:",
          "Codice fiscale:", "Identificativo fiscale", "Nazione:", "PEC", "Riferimento")


def _fornitore(testo: str) -> str:
    """Il nome del cedente. Puo' occupare piu' righe, e sotto c'e' sempre una voce nota."""
    testa = testo.split("Cessionario", 1)[0]
    pezzi = testa.split("Denominazione:", 1)
    if len(pezzi) < 2:
        return ""
    nome = []
    for riga in pezzi[1].splitlines():
        riga = riga.strip()
        if not nome:
            nome.append(riga)
            continue
        if not riga or riga.startswith(CHIAVI):
            break
        nome.append(riga)
    return " ".join(x for x in nome if x).strip()


def _testa(testo: str) -> tuple[str, dt.date]:
    """Numero e data del documento: sono le due righe attaccate sotto la tipologia."""
    coda = testo.split("Codice destinatario", 1)[-1].splitlines()
    for i, riga in enumerate(coda):
        m = re.fullmatch(r"(\d{2}-\d{2}-\d{4})", riga.strip())
        if not m or i == 0:
            continue
        numero = coda[i - 1].strip()
        if not numero:
            continue
        return numero, dt.datetime.strptime(m.group(1), "%d-%m-%Y").date()
    raise ValueError("non trovo numero e data della fattura")


def leggi_fattura(percorso: Path) -> Fattura:
    testo = "\n".join(p.get_text() for p in pymupdf.open(percorso))
    numero, data = _testa(testo)

    corpo = testo.split("Prezzo totale", 1)[-1].split("RIEPILOGHI IVA", 1)[0]
    righe: list[RigaFattura] = []
    descrizione: list[str] = []
    lotto = scadenza = None
    prezzi = False              # sto scorrendo prezzo unitario, IVA, totale

    for linea in (l.strip() for l in corpo.splitlines()):
        if not linea:
            continue
        if linea.lower() == "(codice fornitore)":
            if descrizione:
                descrizione.pop()       # sopra c'era il codice articolo, non la descrizione
            continue
        if RUMORE.match(linea):
            if INIZIO_ARTICOLO.match(linea):
                descrizione, lotto, scadenza, prezzi = [], None, None, False
            elif (m := re.match(r"Rif\.\s*testo:\s*(.+)", linea, re.I)):
                lotto = m.group(1).strip()
            elif (m := re.match(r"Rif\.\s*data:\s*(\d{2}-\d{2}-\d{4})", linea, re.I)):
                scadenza = dt.datetime.strptime(m.group(1), "%d-%m-%Y").date()
            elif (m := re.match(r"lotto:\s*(.+)", linea, re.I)) and righe:
                righe[-1].lotto = m.group(1).strip()    # questo sta sotto il suo articolo
            continue
        if NUMERO.match(linea):
            if prezzi or not descrizione:
                continue                # prezzi dell'articolo appena chiuso, o riga di servizio
            righe.append(RigaFattura(" ".join(descrizione),
                                     float(linea.replace(".", "").replace(",", ".")),
                                     lotto=lotto, scadenza=scadenza))
            descrizione, lotto, scadenza, prezzi = [], None, None, True
            continue
        if prezzi and UNITA.match(linea):
            righe[-1].unita = linea.upper()
            continue
        prezzi = False
        descrizione.append(linea)

    return Fattura(numero.split("/")[0], data, _fornitore(testo), righe, percorso)


# --------------------------------------------------------------------------- #
# classificazione delle righe
# --------------------------------------------------------------------------- #
@dataclass
class Pesce(acquisti.Pesce):
    """Come quello di FLOTT, ma sulla riga del registro non vanno peso unitario e lotto."""
    fresco: bool = False

    @property
    def riga_cam(self) -> str:
        if self.fresco:
            return f"{self.specie} FRESCHE FAO {self.zona} Tot Kg {self.kg:g}"
        parti = ["FILETTI DI", self.specie, self.conservazione,
                 f"FAO {self.zona}", f"Tot Kg {self.kg:g}"]
        return " ".join(x for x in parti if x)


@dataclass
class Olio(acquisti.Olio):
    @property
    def descrizione(self) -> str:
        return f"{self.nome.upper()} KG {self.kg:g}"


def _misura(alto: str) -> int | None:
    """Il numero che dice il formato, saltando quello del filetto di chiusura.

    "VASO ANFORA GR. 80 TP 45" e' un vaso da 80 con capsula 45; "- Nr. 1 x 6.468" e'
    il conteggio dei colli e non c'entra con la misura.
    """
    ripulito = re.sub(r"-\s*NR\.?\s.*$", "", alto)
    ripulito = re.sub(r"\b(?:TP|TO|DIAM\.?)\s*\.?\s*\d+", " ", ripulito)
    m = re.search(r"\b(\d+)\b", ripulito)
    return int(m.group(1)) if m else None


def _diametro(alto: str, formati: dict) -> int | None:
    """Per le capsule il numero utile e' il primo, che sia dopo TO/DIAM. o no.

    Sui coperchi pero' e' scritta la misura del vaso ("COPERCHIO ... DA 1500"): in quel
    caso la capsula giusta la dice la tabella dei formati.
    """
    ripulito = re.sub(r"-\s*NR\.?\s.*$", "", alto)
    m = re.search(r"\b(\d+)\b", ripulito)
    if not m:
        return None
    valore = int(m.group(1))
    formato = formati.get(str(valore))
    if formato and isinstance(formato.capsula, int):
        return formato.capsula
    return valore


def classifica(riga: RigaFattura, formati: dict, avvisi: list[str]):
    testo = re.sub(r"\s+", " ", riga.descrizione).strip()
    alto = testo.upper()

    if alto.startswith(IGNORA_IN_SILENZIO):
        return None
    if "TONNO" in alto and not alto.startswith(("VASO", "CAPS", "COPERCHIO")):
        avvisi.append(f"prodotto a base di tonno, non registrato: {testo}")
        return None

    if alto.startswith(("VASO", "CAPS", "COPERCHIO")):
        quantita = int(round(riga.quantita))
        if alto.startswith(("CAPS", "COPERCHIO")):
            diametro = _diametro(alto, formati)
            if diametro is None:
                avvisi.append(f"diametro capsula non trovato: {testo}")
                return None
            return Imballaggio("CAPSULA", diametro, quantita)
        misura = _misura(alto)
        if misura is None:
            avvisi.append(f"formato vaso non trovato: {testo}")
            return None
        return Imballaggio("VASO", misura, quantita,
                           ermetico="ERMETIC" in alto,
                           bonta="BONTA" in alto or "BONTÀ" in alto)

    pesce = ("FILETTI" in alto or alto.startswith("FIL ")
             or "ALICI" in alto or "ACCIUGHE" in alto or "SARDE" in alto)
    if pesce:
        specie = "SARDE" if "SARDE" in alto or "SARDINA" in alto else "ACCIUGHE"
        filetti = "FILETTI" in alto or alto.startswith("FIL ")
        conservazione = ("SOTTOVUOTO" if "SOTTOVUOTO" in alto else
                         "OLIO SEMI" if "OLIO SEMI" in alto else None)
        if filetti and conservazione is None:
            avvisi.append(f"conservazione del pesce non riconosciuta: {testo}")
            return None
        z = re.search(r"FAO\s*([\d.]+?)(?=\s|--|$)", alto)
        if z:
            zona = z.group(1).rstrip(".")
        elif filetti:
            avvisi.append(f"zona FAO non trovata: {testo}")
            return None
        else:
            zona = ZONA_PESCATO_FRESCO      # alici fresche: pescato locale
        if riga.unita and riga.unita != "KG":
            avvisi.append(f"pesce fatturato a {riga.unita.lower()} e non a chili: {testo}")
            return None
        if filetti and not riga.lotto:
            avvisi.append(f"lotto non trovato: {testo}")
        if specie == "SARDE" and not filetti:
            specie = "SARDE"
        return Pesce(specie, conservazione, zona, 0, riga.lotto, riga.quantita,
                     fresco=not filetti)

    if alto.startswith("OLIO"):
        if "EVO" in alto or "EXTRAVERGINE" in alto:
            nome = "olio oliva EVO"
        elif "OLIVA" in alto:
            nome = "olio d'oliva"
        elif "SEMI" in alto:
            nome = "olio di semi di girasole"
        else:
            avvisi.append(f"tipo di olio non riconosciuto: {testo}")
            return None
        if not riga.lotto:
            avvisi.append(f"lotto non trovato: {testo}")
        return Olio(nome, riga.lotto, riga.quantita)

    avvisi.append(f"riga non riconosciuta, saltata: {testo}")
    return None


# --------------------------------------------------------------------------- #
# registro giacenze
# --------------------------------------------------------------------------- #
def foglio_giacenze(voce, avvisi: list[str]) -> tuple[str, str | None]:
    """Dove va la riga. Fra gli esterni non c'e' la scheda ALTRO PESCE."""
    if isinstance(voce, acquisti.Pesce):
        if voce.specie == "SARDE":
            return "SARDE FAO 34", None
        return f"PESCE FAO {voce.zona.split('.')[0]}", None
    if isinstance(voce, acquisti.Olio):
        return {"olio di semi di girasole": "OLIO DI SEMI",
                "olio oliva EVO": "OLIO EVO",
                "olio d'oliva": "OLIO D'OLIVA"}[voce.nome], None
    misura = ALIAS_VASI.get(voce.misura, voce.misura)
    if voce.genere == "CAPSULA":
        return f"CAPSULE {misura}", None
    return f"VASI {misura}", "bontà" if voce.bonta else None


# --------------------------------------------------------------------------- #
# scrittura dei registri di accettazione
# --------------------------------------------------------------------------- #
def scrivi_fattura(fattura: Fattura, voci: list, libri: dict, conteggio: dict,
                   sovrascrivi: bool) -> list[str]:
    pesci = [v for v in voci if isinstance(v, acquisti.Pesce)]
    olii = [v for v in voci if isinstance(v, acquisti.Olio)]
    imballaggi = [v for v in voci if isinstance(v, Imballaggio)]
    base = fattura.base
    fornitore = f"Fornitore: {fattura.fornitore}" if fattura.fornitore else None
    fatto: list[str] = []

    def salta(sigla: str, nome: str) -> bool:
        wb, _ = libri[sigla]
        if nome in wb.sheetnames and not sovrascrivi:
            return True
        if nome in wb.sheetnames:
            wb.remove(wb[nome])
        return False

    def testa(ws, sigla: str) -> None:
        ws["A6"] = f"Data ingresso: {fattura.data:%d/%m/%Y}"
        ws["A8"] = f"Data controllo: {fattura.data:%d/%m/%Y}"
        if fornitore:
            ws["H6"] = fornitore

    if pesci and not salta("CAM", base):
        wb, modello = libri["CAM"]
        ws = _nuovo_foglio(wb, modello, base)
        testa(ws, "CAM")
        ws["D6"] = f"Estremi del D.D.T. {fattura.numero}"
        for i, p in enumerate(pesci):
            ws.cell(row=11 + i, column=4).value = p.riga_cam
        conteggio["CAM"] += 1
        fatto.append(f"  + {ETICHETTA['CAM']}: {base}  "
                     f"({plurale(len(pesci), 'lotto', 'lotti')})")

    for sigla, elenco in (("CAO", olii), ("CAI", imballaggi)):
        wb, modello = libri[sigla]
        for voce, suffisso in zip(elenco, _suffissi(len(elenco))):
            nome = base + suffisso
            if salta(sigla, nome):
                continue
            ws = _nuovo_foglio(wb, modello, nome)
            testa(ws, sigla)
            if sigla == "CAO":
                ws["D6"] = sottolinea("Estremi del DDT:", f" {fattura.numero}", suffisso)
                ws["D8"] = sottolinea("Descrizione prodotto: ", voce.descrizione)
                ws["D17"] = int(voce.kg) if float(voce.kg).is_integer() else voce.kg
            else:
                ws["D6"] = f"Estremi del DDT: {fattura.numero}{suffisso}"
                ws["D8"] = f"Descrizione prodotto:  {voce.descrizione}"
            conteggio[sigla] += 1
            fatto.append(f"  + {ETICHETTA[sigla]}: {nome}  {voce.descrizione}")
    return fatto


# --------------------------------------------------------------------------- #
def main() -> int:
    ap = argparse.ArgumentParser(
        description="Compila i registri di accettazione e le giacenze degli esterni.")
    ap.add_argument("pdf", nargs="+", type=Path, help="fatture in PDF o cartelle che le contengono")
    ap.add_argument("-o", "--registri", type=Path, required=True,
                    help="cartella dei tre registri EST-; {anno} e {trimestre} vengono "
                         "sostituiti con quelli della fattura")
    ap.add_argument("-g", "--giacenze", type=Path, required=True,
                    help="EST-REGISTRO GIACENZE da aggiornare; {anno} viene sostituito "
                         "con l'anno della fattura")
    ap.add_argument("-f", "--formati", type=Path, default=FORMATI)
    ap.add_argument("--sovrascrivi", action="store_true")
    ap.add_argument("-l", "--log", type=Path)
    ap.add_argument("-b", "--backup", type=Path,
                    help="cartella dove mettere le copie .bak (predefinito: accanto al file)")
    args = ap.parse_args()

    percorsi = sorted({p for voce in args.pdf
                       for p in (voce.rglob("*.pdf") if voce.is_dir() else [voce])})
    if not percorsi:
        print("Nessun PDF trovato.", file=sys.stderr)
        return 1
    registro = Registro(args.log or risolvi(args.registri, dt.date.today()) / "acquisti.log")
    formati = leggi_formati(args.formati)

    per_cartella: dict[Path, list] = {}
    for percorso in percorsi:
        try:
            fattura = leggi_fattura(percorso)
        except ValueError as errore:
            registro.contesto(percorso.name)
            registro.anomalia(str(errore))
            continue
        avvisi: list[str] = []
        voci = [v for v in (classifica(r, formati, avvisi) for r in fattura.righe)
                if v is not None]
        per_cartella.setdefault(risolvi(args.registri, fattura.data), []).append(
            (percorso, fattura, voci, avvisi))

    gia_visti = registro.percorso.with_name(registro.percorso.stem + "-visti.txt")
    visti = leggi_visti(gia_visti)
    giacenze = GestoreGiacenze(args.giacenze, args.backup, foglio_giacenze)
    conteggio = dict.fromkeys(NOMI_REGISTRI, 0)
    salvati: list[Path] = []

    for cartella in sorted(per_cartella):
        if len(per_cartella) > 1:
            print(f"\n── {Path(*cartella.parts[-2:])} ──")
        prima = {s: conteggio[s] for s in NOMI_REGISTRI}
        libri = {s: _apri(cartella / n, RADICE / f"modello_est_{s}.xlsx", args.backup)
                 for s, n in NOMI_REGISTRI.items()}

        for percorso, fattura, voci, avvisi in per_cartella[cartella]:
            fatto = scrivi_fattura(fattura, voci, libri, conteggio, args.sovrascrivi)
            for posizione, voce in enumerate(voci, 1):
                if giacenze.aggiungi(voce, fattura, posizione, avvisi) == "aggiunta":
                    fatto.append(f"  + giacenze  {foglio_giacenze(voce, avvisi)[0]}")
            da_dire = avvisi if (fattura.base not in visti or fatto) else []
            visti.add(fattura.base)
            if fatto or da_dire:
                print(f"\n  fattura {fattura.numero} del {fattura.data:%d/%m/%Y}"
                      f"  {fattura.fornitore}")
            if fatto:
                print("\n".join(fatto))
            if da_dire:
                registro.contesto(f"fattura {fattura.numero}  ({percorso.name})")
                for a in da_dire:
                    registro.anomalia(a)

        cartella.mkdir(parents=True, exist_ok=True)
        for sigla, nome in NOMI_REGISTRI.items():
            wb, _ = libri[sigla]
            if wb.sheetnames == ["MODELLO"] or conteggio[sigla] == prima[sigla]:
                continue
            if "MODELLO" in wb.sheetnames:
                wb.remove(wb["MODELLO"])
            _ordina(wb)
            salva_libro(wb, cartella / nome)
            salvati.append(cartella / nome)

    salvati += giacenze.salva()
    scrivi_visti(gia_visti, visti)
    riassunto = (f"{plurale(len(percorsi), 'fattura letta', 'fatture lette')}, "
                 f"fogli aggiunti: pesce {conteggio['CAM']}, olio {conteggio['CAO']}, "
                 f"vasi {conteggio['CAI']}; "
                 f"{plurale(registro.anomalie, 'anomalia', 'anomalie')}")
    registro.chiudi(riassunto)
    print(f"\n{riassunto}")
    elenca(salvati)
    registro.avvisa()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
