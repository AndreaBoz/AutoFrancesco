"""Dalle fatture di acquisto FLOTT ai tre registri di accettazione e al registro giacenze.

Ogni riga della fattura viene classificata in pesce, olio o imballaggio e finisce
nel registro che le compete. Vassoi, tonno ed etichette vengono ignorati.
La scheda RIMANENZE del registro giacenze non viene mai toccata.
"""
from __future__ import annotations

import argparse
import collections
import datetime as dt
import re
import shutil
import sys
from copy import copy
from dataclasses import dataclass, field
from pathlib import Path

import openpyxl
import pymupdf

from riepilogo import (elenca, leggi_visti, plurale, salva_copia, salva_libro,
                       scrivi_visti)
from openpyxl.cell.rich_text import CellRichText, TextBlock
from openpyxl.cell.text import InlineFont

RADICE = Path(__file__).resolve().parent

# Materiale di confezionamento che non va nei registri: si scarta in silenzio, senza
# sporcare il log. Si riconosce dalla prima parola, perche' le stesse parole compaiono
# anche dentro descrizioni da registrare, come "VASO VETRO 160 ETICHETTA LINEARE ...".
IGNORA_IN_SILENZIO = ("VASSOIO", "VASSOI", "ETICHETTA", "ETICHETTE",
                      "SIGILLO", "SIGILLI", "SCATOLA", "SCATOLE", "PELLICOLA")

# Nel DR.CAO la parte variabile delle due righe di testa e' sottolineata, mentre
# l'etichetta e l'eventuale suffisso -A/-B restano normali: e' testo formattato a
# tratti dentro la stessa cella.
CARATTERE = dict(rFont="Arial", family=2, sz=10, b=False)
SOTTOLINEATO = InlineFont(u="single", **CARATTERE)
NORMALE = InlineFont(**CARATTERE)


def sottolinea(etichetta: str, valore: str, coda: str = "") -> CellRichText:
    pezzi = [etichetta, TextBlock(SOTTOLINEATO, valore)]
    if coda:
        pezzi.append(TextBlock(NORMALE, coda))
    return CellRichText(*pezzi)


# corrispondenze fra il formato scritto in fattura e la scheda del registro giacenze:
# il vaso di plastica da 40 ML si registra nella scheda VASI 42
ALIAS_VASI = {40: 42}


# --------------------------------------------------------------------------- #
# lettura della fattura
# --------------------------------------------------------------------------- #
# Righe di servizio da scartare: riferimenti al DDT, codici a barre, intestazioni di
# colonna ripetute a ogni pagina e pie' di pagina del browser da cui si stampa il PDF.
RUMORE = re.compile(
    r"^(?:DDT\s+\d+\s+del\s+[\d.-]+"
    r"|Tipo dato:.*|Rif\. (?:testo|numero):.*"
    r"|Cod\. articolo|Descrizione|Quantit\u00e0|Prezzo unitario|UM|Sconto o|magg\.|%IVA"
    r"|Prezzo totale|Firefox|Chrome|Safari|https?://.*|\d+ di \d+"
    r"|\d{2}/\d{2}/\d{2},\s*\d{2}:\d{2})$", re.IGNORECASE)
INIZIO_RIGA = re.compile(r"^DDT\s+\d+\s+del\s+[\d.-]+$", re.IGNORECASE)
@dataclass
class RigaFattura:
    descrizione: str
    quantita: float


@dataclass
class Fattura:
    numero: str            # progressivo depurato, es. "121416"
    data: dt.date
    righe: list[RigaFattura]
    file: Path


def leggi_fattura(percorso: Path) -> Fattura:
    testo = "\n".join(p.get_text() for p in pymupdf.open(percorso))

    # il numero e' del tipo 26121416/V2_26, seguito dalla data del documento:
    # il TD in testa puo' occupare piu' righe, quindi mi aggancio al numero
    m = re.search(r"\n\d{2}(\d+)/V\d+_\d+\s*\n(\d{2}-\d{2}-\d{4})", testo)
    if not m:
        m = re.search(r"TD\d\d[^\n]*\n\d{2}(\d+)\S*\s*\n(\d{2}-\d{2}-\d{4})", testo)
    if not m:
        raise ValueError("non trovo numero e data della fattura")
    numero = m.group(1)
    data = dt.datetime.strptime(m.group(2), "%d-%m-%Y").date()

    corpo = testo.split("Prezzo totale", 1)[-1].split("RIEPILOGHI IVA", 1)[0]
    righe, buffer = [], []
    for linea in (l.strip() for l in corpo.splitlines()):
        if not linea:
            continue
        if RUMORE.match(linea):
            if INIZIO_RIGA.match(linea):
                buffer = []          # comincia un altro articolo
            continue
        if re.fullmatch(r"[\d.]+,\d+", linea):
            # prima cifra dopo la descrizione: e' la quantita', il resto sono i prezzi
            if buffer:
                righe.append(RigaFattura(" ".join(buffer),
                                         float(linea.replace(".", "").replace(",", "."))))
                buffer = []
        else:
            buffer.append(linea)
    return Fattura(numero, data, righe, percorso)


# --------------------------------------------------------------------------- #
# classificazione delle righe
# --------------------------------------------------------------------------- #
@dataclass
class Pesce:
    specie: str            # ACCIUGHE o SARDE
    conservazione: str     # SOTTOVUOTO o OLIO SEMI
    zona: str              # 34, 37.2.2, 87 ...
    peso_unitario: float
    lotto: str
    kg: float

    @property
    def riga_cam(self) -> str:
        peso = f"{self.peso_unitario:g}"
        kg = f"{self.kg:g}"
        return (f"FIL {self.specie} {self.conservazione} FAO {self.zona} "
                f"{peso} KG L {self.lotto} Tot Kg {kg}")


@dataclass
class Olio:
    nome: str              # olio di semi di girasole, olio oliva EVO, olio d'oliva
    lotto: str
    kg: float

    @property
    def descrizione(self) -> str:
        return f"{self.nome} – lotto {self.lotto} kg {self.kg:g}"


@dataclass
class Imballaggio:
    genere: str            # VASO o CAPSULA
    misura: int
    quantita: int
    ermetico: bool = False
    bonta: bool = False

    @property
    def descrizione(self) -> str:
        if self.genere == "CAPSULA":
            return f"CAPS. DIAM. {self.misura} quantità {self.quantita}"
        suffisso = " erm." if self.ermetico else (" bontà" if self.bonta else "")
        return f"VASO {self.misura} G{suffisso} quantità {self.quantita}"


def _lotto(testo: str) -> str | None:
    """L'ultima parentesi dopo LOTTO contiene il lotto buono: IJ188XFOS(D)(26208) -> 26208."""
    m = re.search(r"LOTTO\s+(.*?)(?:\s+SCADENZA\b|\s+Q\.?\s?TA\b|$)", testo)
    if not m:
        return None
    parentesi = re.findall(r"\(([^)]*)\)", m.group(1))
    for valore in reversed(parentesi):
        if valore.isdigit():
            return valore
    grezzo = re.match(r"[\w]+", m.group(1))
    return grezzo.group(0) if grezzo else None


def classifica(riga: RigaFattura, avvisi: list[str]):
    testo = re.sub(r"\s+", " ", riga.descrizione).strip()
    alto = testo.upper()

    if alto.startswith(IGNORA_IN_SILENZIO):
        return None

    # il tonno non si registra, ma va segnalato: i vasi vuoti destinati al tonno invece
    # cominciano con VASO e vengono registrati normalmente
    if alto.startswith("TONNO"):
        avvisi.append(f"prodotto a base di tonno, non registrato: {testo}")
        return None

    if alto.startswith("VASO") or "CAPS" in alto:
        collo = re.search(r"(?:PALLET|CARTONI|CARTONE|SCATOLE|SCATOLA|CONF\.?)\s+DA\s+([\d.]+)\s*PZ", alto)
        if collo:
            moltiplicatore = int(collo.group(1).replace(".", ""))
        elif re.search(r"UNIT\w*\s+SINGOL", alto):
            moltiplicatore = 1          # venduto a pezzo: la quantita' e' gia' quella
        else:
            moltiplicatore = 1
            avvisi.append(f"pezzi per collo non trovati, uso la quantita' cosi' com'e': {testo}")
        quantita = int(round(riga.quantita * moltiplicatore))
        if "CAPS" in alto:
            d = re.search(r"(?:DIAM\.?|T)\s*(\d+)", alto)
            if not d:
                avvisi.append(f"diametro capsula non trovato: {testo}")
                return None
            return Imballaggio("CAPSULA", int(d.group(1)), quantita)
        g = re.search(r"(\d+)\s*G\b", alto) or re.search(r"VASO(?:\s+[A-Z]+)*?\s+(\d+)", alto)
        if not g:
            avvisi.append(f"formato vaso non trovato: {testo}")
            return None
        return Imballaggio("VASO", int(g.group(1)), quantita,
                           ermetico="ERMETIC" in alto,
                           bonta="BONTA" in alto or "BONTÀ" in alto)

    if "FILETTI" in alto or alto.startswith("FIL "):
        if "TONNO" in alto:
            avvisi.append(f"pesce tonno, non registrato: {testo}")
            return None
        specie = "SARDE" if "SARDE" in alto else "ACCIUGHE"
        conservazione = "SOTTOVUOTO" if "SOTTOVUOTO" in alto else (
            "OLIO SEMI" if "OLIO SEMI" in alto else None)
        if conservazione is None:
            avvisi.append(f"conservazione del pesce non riconosciuta: {testo}")
            return None
        z = re.search(r"FAO\s+([\d.]+?)(?=\s|--|$)", alto)
        if not z:
            avvisi.append(f"zona FAO non trovata: {testo}")
            return None
        # il peso unitario e' quello dell'unita' di misura fra i doppi trattini
        u = re.search(r"--\s*(?:CARTONE|BUSTA|SECCHIELLO|SCATOLA|CONF\.?)\s*([\d.,]+)\s*KG", alto)
        if not u:
            avvisi.append(f"peso unitario non trovato: {testo}")
            return None
        peso = float(u.group(1).replace(",", "."))
        lotto = _lotto(alto)
        if lotto is None:
            avvisi.append(f"lotto non trovato: {testo}")
        return Pesce(specie, conservazione, z.group(1).rstrip("."), peso,
                     lotto, round(riga.quantita * peso, 3))

    if alto.startswith("OLIO"):
        if "EVO" in alto:
            nome = "olio oliva EVO"
        elif "OLIVA" in alto:
            nome = "olio d'oliva"
        elif "SEMI" in alto:
            nome = "olio di semi di girasole"
        else:
            avvisi.append(f"tipo di olio non riconosciuto: {testo}")
            return None
        lotto = _lotto(alto)
        if lotto is None:
            avvisi.append(f"lotto non trovato: {testo}")
        return Olio(nome, lotto, riga.quantita)

    avvisi.append(f"riga non riconosciuta, saltata: {testo}")
    return None


# --------------------------------------------------------------------------- #
# scrittura dei registri di accettazione
# --------------------------------------------------------------------------- #
def _apri(percorso: Path, modello: Path, backup: Path | None):
    # rich_text=True e' indispensabile: senza, riaprendo un registro gia' scritto
    # openpyxl appiattirebbe le sottolineature del DR.CAO e andrebbero perse
    if percorso.exists():
        wb = openpyxl.load_workbook(percorso, rich_text=True)
        salva_copia(percorso, backup)
        return wb, wb.worksheets[0]
    wb = openpyxl.load_workbook(modello, rich_text=True)
    return wb, wb["MODELLO"]


def _nuovo_foglio(wb, modello, nome: str):
    """copy_worksheet non porta con se' vista e area di stampa: le ricopio a mano."""
    ws = wb.copy_worksheet(modello)
    ws.title = nome
    ws.sheet_view.showGridLines = modello.sheet_view.showGridLines
    ws.sheet_view.showRowColHeaders = modello.sheet_view.showRowColHeaders
    ws.sheet_view.zoomScale = modello.sheet_view.zoomScale
    ws.sheet_view.zoomScaleNormal = modello.sheet_view.zoomScaleNormal
    if modello.print_area:
        ws.print_area = str(modello.print_area).split("!")[-1]
    return ws


def _ordina(wb) -> None:
    def chiave(ws):
        m = re.match(r"(\d+)_(\d{2})(\d{2})(\d{2})(-([A-Z]))?", ws.title)
        if not m:
            return (99, 99, 99, 10 ** 9, "")
        return (int(m.group(4)), int(m.group(3)), int(m.group(2)),
                int(m.group(1)), m.group(6) or "")
    wb._sheets.sort(key=chiave)


def _suffissi(quanti: int):
    return [""] + [f"-{chr(ord('A') + i)}" for i in range(quanti - 1)]


# --------------------------------------------------------------------------- #
# scrittura del registro giacenze
# --------------------------------------------------------------------------- #
def foglio_giacenze(voce, avvisi: list[str]) -> tuple[str, str | None]:
    """Restituisce (nome del foglio, nota da scrivere accanto alla quantita')."""
    if isinstance(voce, Pesce):
        if voce.specie == "SARDE":
            return "SARDE FAO 34", None
        principale = voce.zona.split(".")[0]
        if principale in ("34", "37", "87"):
            return f"PESCE FAO {principale}", None
        return "ALTRO PESCE", principale
    if isinstance(voce, Olio):
        return {"olio di semi di girasole": "OLIO DI SEMI",
                "olio oliva EVO": "OLIO EVO",
                "olio d'oliva": "OLIO D'OLIVA"}[voce.nome], None
    misura = ALIAS_VASI.get(voce.misura, voce.misura)
    if voce.genere == "CAPSULA":
        return f"CAPSULE {misura}", None
    return f"VASI {misura}", "bontà" if voce.bonta else None


def _colonne(ws) -> tuple[int, int | None, int | None]:
    """Trova (colonna quantita', colonna lotto, colonna nota) leggendo l'intestazione."""
    intestazioni = {str(ws.cell(2, c).value).strip().upper(): c for c in range(1, 7)}
    peso = intestazioni.get("PESO")
    if peso is None:
        peso = 4 if ws.cell(3, 4).value is not None else 3
    return (peso, 3, 5) if peso == 4 else (peso, None, 4)


def _confronta(valore):
    """Normalizza per confrontare righe gia' presenti con quelle da scrivere."""
    if isinstance(valore, dt.datetime):
        return valore.date()
    if isinstance(valore, float) and valore.is_integer():
        return int(valore)
    if isinstance(valore, str):
        return valore.strip().lower()
    return valore


COL_ID = 6          # colonna F: libera in tutte le schede


class Giacenze:
    """Aggiunge righe ai fogli dei materiali, senza mai toccare la scheda RIMANENZE.

    Ogni riga porta in colonna F un identificativo del tipo #121560-2, cioe' documento e
    posizione della voce dentro il documento. Serve a due cose: riconoscere una riga anche
    dopo che l'hai corretta a mano, e permettere due righe identiche quando sono davvero
    due consegne uguali.
    """

    def __init__(self, wb, dove=None):
        self.wb = wb
        self.dove = dove or foglio_giacenze     # gli esterni hanno schede diverse
        self.modificato = False
        self._identificativi: dict[str, dict[str, int]] = {}
        self._contenuti: dict[str, dict[tuple, list[int]]] = {}

    def _indice(self, ws):
        """Legge una volta sola cosa c'e' gia' nel foglio: per identificativo e per contenuto."""
        if ws.title in self._identificativi:
            return self._identificativi[ws.title], self._contenuti[ws.title]
        per_id, per_contenuto = {}, {}
        for r in range(3, ws.max_row + 1):
            riga = tuple(_confronta(ws.cell(r, c).value) for c in range(1, 6))
            if not any(x is not None for x in riga[:4]):
                continue
            marchio = ws.cell(r, COL_ID).value
            if marchio:
                per_id[str(marchio).strip()] = r
            else:
                per_contenuto.setdefault(riga, []).append(r)
        self._identificativi[ws.title] = per_id
        self._contenuti[ws.title] = per_contenuto
        return per_id, per_contenuto

    def _ultima_piena(self, ws) -> int:
        return max((c.row for riga in ws.iter_rows(min_row=3, max_col=4)
                    for c in riga if c.value is not None), default=2)

    def _punto(self, ws, data: dt.date) -> int:
        """Dove va la riga per restare in ordine di data: la prima con data successiva."""
        for r in range(3, ws.max_row + 1):
            valore = ws.cell(r, 1).value
            if isinstance(valore, dt.datetime):
                valore = valore.date()
            if isinstance(valore, dt.date) and valore > data:
                return r
        return self._ultima_piena(ws) + 1

    def _intestazione(self, ws) -> None:
        if ws.cell(2, COL_ID).value is None:
            ws.cell(2, COL_ID).value = "ID"
            ws.cell(2, COL_ID)._style = copy(ws.cell(2, 2)._style)

    def aggiungi(self, voce, fattura: "Fattura", posizione: int, avvisi: list[str]) -> str:
        nome, nota = self.dove(voce, avvisi)
        if nome not in self.wb.sheetnames:
            avvisi.append(f"nel registro giacenze non esiste il foglio {nome!r}: riga non registrata")
            return "saltata"
        ws = self.wb[nome]
        col_peso, col_lotto, col_nota = _colonne(ws)
        marchio = f"#{fattura.numero}-{posizione}"

        valori = {1: fattura.data,
                  2: int(fattura.numero) if fattura.numero.isdigit() else fattura.numero}
        quantita = voce.kg if isinstance(voce, (Pesce, Olio)) else voce.quantita
        valori[col_peso] = int(quantita) if float(quantita).is_integer() else quantita
        if col_lotto and isinstance(voce, (Pesce, Olio)) and voce.lotto:
            valori[col_lotto] = int(voce.lotto) if voce.lotto.isdigit() else voce.lotto
        if nota and col_nota:
            valori[col_nota] = int(nota) if str(nota).isdigit() else nota

        per_id, per_contenuto = self._indice(ws)
        if marchio in per_id:
            return "gia' presente"          # c'e', anche se l'hai corretta a mano

        # riga scritta prima che esistessero gli identificativi: la marchio e la lascio stare
        chiave = tuple(_confronta(valori.get(c)) for c in range(1, 6))
        vecchie = per_contenuto.get(chiave)
        if vecchie:
            r = vecchie.pop(0)
            self._intestazione(ws)
            ws.cell(r, COL_ID).value = marchio
            per_id[marchio] = r
            self.modificato = True      # ho marchiato una riga che non aveva l'id
            return "gia' presente"

        ultima = self._ultima_piena(ws)
        r = self._punto(ws, fattura.data)
        if r <= ultima:
            ws.insert_rows(r)               # in mezzo, per non rompere l'ordine di data
            self._identificativi.pop(ws.title, None)   # gli indici sono scalati
            self._contenuti.pop(ws.title, None)
        modello = r - 1 if r > 3 else (r + 1 if ultima >= 3 else None)
        self._intestazione(ws)
        for c in range(1, COL_ID + 1):
            if modello:
                ws.cell(r, c)._style = copy(ws.cell(modello, c)._style)
            if c in valori:
                ws.cell(r, c).value = valori[c]
        ws.cell(r, COL_ID).value = marchio
        self.modificato = True
        return "aggiunta"


def risolvi(percorso: Path, data: dt.date) -> Path:
    """Sostituisce {anno} e {trimestre} con quelli del documento."""
    return Path(str(percorso)
                .replace("{anno}", str(data.year))
                .replace("{trimestre}", f"{(data.month - 1) // 3 + 1} TRIMESTRE"))


class GestoreGiacenze:
    """Manda ogni riga nel registro giacenze dell'anno del documento.

    Cosi' una fattura di dicembre lavorata a gennaio non finisce nel file dell'anno nuovo.
    """

    def __init__(self, modello_percorso: Path, backup: Path | None = None, dove=None):
        self.modello_percorso = modello_percorso
        self.backup = backup
        self.dove = dove
        self.aperti: dict[int, tuple[Path, object, Giacenze]] = {}
        self.mancanti: set[int] = set()

    def _per(self, anno: int, avvisi: list[str]):
        if anno in self.aperti:
            return self.aperti[anno][2]
        percorso = risolvi(self.modello_percorso, dt.date(anno, 1, 1))
        if not percorso.exists():
            # il registro dell'anno in corso sta nella radice, quelli chiusi nella
            # cartella del loro anno: se qui non c'e', guardo di la'
            archiviato = percorso.parent / str(anno) / percorso.name
            if archiviato.exists():
                percorso = archiviato
        if not percorso.exists():
            if anno not in self.mancanti:
                self.mancanti.add(anno)
                avvisi.append(f"manca il registro giacenze {percorso.name}: "
                              f"le righe del {anno} non vengono registrate")
            return None
        wb = openpyxl.load_workbook(percorso, rich_text=True)
        salva_copia(percorso, self.backup)
        self.aperti[anno] = (percorso, wb, Giacenze(wb, self.dove))
        return self.aperti[anno][2]

    def aggiungi(self, voce, fattura: "Fattura", posizione: int, avvisi: list[str]) -> str:
        libro = self._per(fattura.data.year, avvisi)
        if libro is None:
            return "saltata"
        return libro.aggiungi(voce, fattura, posizione, avvisi)

    def salva(self) -> list[Path]:
        """Salva e restituisce solo i registri giacenze davvero cambiati."""
        scritti = []
        for percorso, wb, libro in self.aperti.values():
            if not libro.modificato:
                continue
            salva_libro(wb, percorso)
            scritti.append(percorso)
        return scritti


# --------------------------------------------------------------------------- #
class Registro:
    def __init__(self, percorso: Path):
        self.percorso, self.righe, self.anomalie = percorso, [], 0

    def anomalia(self, testo: str) -> None:
        self.anomalie += 1
        print(f"      ATTENZIONE {testo}")
        self.righe.append(f"    {testo}")

    def contesto(self, testo: str) -> None:
        self.righe.append(f"  {testo}")

    def avvisa(self) -> None:
        """Rende impossibile non accorgersi che qualcosa e' rimasto fuori."""
        if not self.anomalie:
            return
        quante = "1 cosa inattesa" if self.anomalie == 1 else f"{self.anomalie} cose inattese"
        print("\n" + "!" * 66)
        print(f"!!  ATTENZIONE: {quante}. Qualcosa non e' stato registrato.")
        print(f"!!  Dettagli in: {self.percorso}")
        print("!" * 66)

    def chiudi(self, riassunto: str) -> None:
        self.percorso.parent.mkdir(parents=True, exist_ok=True)
        nuovo = not self.percorso.exists()
        with self.percorso.open("a", encoding="utf-8") as f:
            if nuovo:
                f.write("Registro anomalie - fatture di acquisto\n")
            f.write(f"\n=== {dt.datetime.now():%d/%m/%Y %H:%M}  {riassunto}\n")
            f.write("\n".join(self.righe) + "\n" if self.anomalie else "  nessuna anomalia\n")


# come chiamare i registri a schermo: le sigle dicono poco a colpo d'occhio
ETICHETTA = {"CAM": "pesce", "CAO": "olio", "CAI": "vasi"}

NOMI_REGISTRI = {
    "CAM": "DR.CAM.00- Controllo in accettazione MAT.PRIME e SEMILAVORATI.xlsx",
    "CAO": "DR.CAO.00- Controllo in accettazione OLIO.xlsx",
    "CAI": "DR.CAI.00 - Controllo imballaggi primari.xlsx",
}


def scrivi_fattura(fattura: Fattura, voci: list, libri: dict, conteggio: dict,
                   sovrascrivi: bool) -> list[str]:
    """Scrive i fogli della fattura e restituisce cosa e' stato fatto, riga per riga."""
    pesci = [v for v in voci if isinstance(v, Pesce)]
    olii = [v for v in voci if isinstance(v, Olio)]
    imballaggi = [v for v in voci if isinstance(v, Imballaggio)]
    base = f"{fattura.numero}_{fattura.data:%d%m%y}"
    fatto: list[str] = []

    def salta(sigla: str, nome: str) -> bool:
        wb, _ = libri[sigla]
        if nome in wb.sheetnames and not sovrascrivi:
            return True
        if nome in wb.sheetnames:
            wb.remove(wb[nome])
        return False

    if pesci and not salta("CAM", base):
        wb, modello = libri["CAM"]
        ws = _nuovo_foglio(wb, modello, base)
        ws["A6"] = f"Data ingresso: {fattura.data:%d/%m/%Y}"
        ws["A8"] = f"Data controllo: {fattura.data:%d/%m/%Y}"
        ws["D6"] = f"Estremi del DDT/fattura  {fattura.numero}"
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
            ws["A6"] = f"Data ingresso: {fattura.data:%d/%m/%Y}"
            ws["A8"] = (f"Data ingresso: {fattura.data:%d/%m/%Y}" if sigla == "CAO"
                        else f"Data controllo: {fattura.data:%d/%m/%Y}")
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


def main() -> int:
    ap = argparse.ArgumentParser(description="Compila i registri di accettazione e le giacenze.")
    ap.add_argument("pdf", nargs="+", type=Path, help="fatture in PDF o cartelle che le contengono")
    ap.add_argument("-o", "--registri", type=Path, required=True,
                    help="cartella dei tre registri DR.CAM / DR.CAO / DR.CAI; {anno} e "
                         "{trimestre} vengono sostituiti con quelli della fattura")
    ap.add_argument("-g", "--giacenze", type=Path, required=True,
                    help="file REGISTRO GIACENZE da aggiornare; {anno} viene sostituito "
                         "con l'anno della fattura")
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
    oggi = dt.date.today()
    registro = Registro(args.log or risolvi(args.registri, oggi) / "acquisti.log")

    # prima leggo tutte le fatture, poi le raggruppo per cartella di destinazione
    per_cartella: dict[Path, list[tuple[Path, Fattura, list, list]]] = {}
    for percorso in percorsi:
        try:
            fattura = leggi_fattura(percorso)
        except ValueError as errore:
            registro.contesto(percorso.name)
            registro.anomalia(str(errore))
            continue
        avvisi: list[str] = []
        voci = [v for v in (classifica(r, avvisi) for r in fattura.righe) if v is not None]
        cartella = risolvi(args.registri, fattura.data)
        per_cartella.setdefault(cartella, []).append((percorso, fattura, voci, avvisi))

    gia_visti = registro.percorso.with_name(registro.percorso.stem + "-visti.txt")
    visti = leggi_visti(gia_visti)
    giacenze = GestoreGiacenze(args.giacenze, args.backup)
    conteggio = dict.fromkeys(["CAM", "CAO", "CAI"], 0)
    salvati: list[Path] = []

    for cartella in sorted(per_cartella):
        if len(per_cartella) > 1:
            print(f"\n── {Path(*cartella.parts[-2:])} ──")
        prima = {s: conteggio[s] for s in NOMI_REGISTRI}
        libri = {s: _apri(cartella / n, RADICE / f"modello_{s}.xlsx", args.backup)
                 for s, n in NOMI_REGISTRI.items()}

        for percorso, fattura, voci, avvisi in per_cartella[cartella]:
            fatto = scrivi_fattura(fattura, voci, libri, conteggio, args.sovrascrivi)
            for posizione, voce in enumerate(voci, 1):
                if giacenze.aggiungi(voce, fattura, posizione, avvisi) == "aggiunta":
                    fatto.append(f"  + giacenze  {foglio_giacenze(voce, avvisi)[0]}")
            # le segnalazioni si dicono la prima volta che si legge un documento, o
            # quando quel documento ha comunque prodotto qualcosa di nuovo
            da_dire = avvisi if (fattura.numero not in visti or fatto) else []
            visti.add(fattura.numero)
            if fatto or da_dire:
                print(f"\n  fattura {fattura.numero} del {fattura.data:%d/%m/%Y}")
            if fatto:
                print("\n".join(fatto))
            if da_dire:
                registro.contesto(f"fattura {fattura.numero}  ({percorso.name})")
                for a in da_dire:
                    registro.anomalia(a)

        cartella.mkdir(parents=True, exist_ok=True)
        for sigla, nome in NOMI_REGISTRI.items():
            wb, _ = libri[sigla]
            if wb.sheetnames == ["MODELLO"]:
                continue             # nessun documento di questo tipo: niente file vuoto
            if conteggio[sigla] == prima[sigla]:
                continue             # niente di nuovo: non lo riscrivo nemmeno
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
