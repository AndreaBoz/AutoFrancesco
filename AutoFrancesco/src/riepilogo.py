"""Genera il DR.RPD.00 - Riepilogo produzione a partire dai PDF dei DDT in uscita.

Ogni riga di prodotto del DDT diventa un blocco (5 colonne) su un foglio
chiamato <numero>_<ddmmyy>. I campi ORIGINE e LOTTO DI ENTRATA restano vuoti:
non sono ricavabili dal DDT.
"""
from __future__ import annotations

import argparse
import datetime as dt
import os
import re
import shutil
import sys
import zipfile
from copy import copy
from dataclasses import dataclass, field
from pathlib import Path

import openpyxl
import pymupdf
from openpyxl.cell.cell import MergedCell
from openpyxl.utils import get_column_letter

RADICE = Path(__file__).resolve().parent          # la cartella src
PROGETTO = RADICE.parent
MODELLO = RADICE / "modello_RPD.xlsx"
# La tabella dei formati e' una sola e sta nell'archivio, non nel progetto
FORMATI = Path(os.environ.get("HACCP_FORMATI", Path.home() / "HACCP" / "FORMATI E CAPSULE.xlsx"))

LARGH_BLOCCO = 5          # colonne per prodotto, compresa la colonna di stacco
RIGHE = 28                # ultima riga del modulo
R_PRODOTTO, R_DATA, R_LOTTO, R_SCADENZA = 9, 11, 12, 13
R_PESCE, R_FUSTI, R_OLIO, R_VASI, R_CAPSULE, R_PRODOTTE = 16, 17, 18, 19, 20, 28


# --------------------------------------------------------------------------- #
# tabella formati / capsule / coefficienti
# --------------------------------------------------------------------------- #
@dataclass
class Formato:
    capsula: object          # int (diametro) oppure "ERMETICO"
    coefficiente: float
    varianti: dict = field(default_factory=dict)   # es. {"MARCHIO": 0.245}


def _numero(valore):
    """Legge 0.24 sia da 0.24 sia da '0,24 NORMALE'."""
    if isinstance(valore, (int, float)):
        return float(valore)
    m = re.search(r"\d+[.,]?\d*", str(valore))
    return float(m.group(0).replace(",", ".")) if m else None


def leggi_formati(percorso: Path) -> dict[str, Formato]:
    ws = openpyxl.load_workbook(percorso, data_only=True).worksheets[0]
    tabella = {}
    for riga in ws.iter_rows(min_row=2, max_col=4, values_only=True):
        formato, capsula, coeff = riga[0], riga[1], riga[2]
        if formato is None or coeff is None:
            continue
        varianti = {}
        for extra in riga[3:]:
            if extra is None:
                continue
            etichetta = re.sub(r"[\d.,]+", "", str(extra)).strip()
            if etichetta:
                varianti[etichetta.upper()] = _numero(extra)
        tabella[str(formato).strip().lower()] = Formato(
            capsula=capsula if isinstance(capsula, int) else str(capsula).strip().upper(),
            coefficiente=_numero(coeff),
            varianti=varianti,
        )
    return tabella


# --------------------------------------------------------------------------- #
# lettura del PDF
# --------------------------------------------------------------------------- #
@dataclass
class RigaDDT:
    descrizione: str
    confezioni: int
    unita: str


@dataclass
class DDT:
    numero: int
    data: dt.date
    righe: list[RigaDDT]
    file: Path


def leggi_ddt(percorso: Path) -> DDT:
    testo = "\n".join(p.get_text() for p in pymupdf.open(percorso))

    m = re.search(r"Documento di Trasporto\s+n\.?\s*(\d+)\s+del\s+(\d{2}/\d{2}/\d{4})", testo)
    if not m:
        raise ValueError(f"{percorso.name}: non trovo numero e data del DDT")
    numero, data = int(m.group(1)), dt.datetime.strptime(m.group(2), "%d/%m/%Y").date()

    corpo = testo.split("U.M.", 1)[-1].split("Dati di trasporto", 1)[0]
    righe, buffer = [], []
    for linea in (l.strip() for l in corpo.splitlines()):
        if not linea:
            continue
        if re.fullmatch(r"[\d.]+,\d+", linea):          # es. 228,00 -> quantità
            righe.append([" ".join(buffer), float(linea.replace(".", "").replace(",", ".")), None])
            buffer = []
        elif righe and righe[-1][2] is None and len(linea) <= 4:
            righe[-1][2] = linea                        # unità di misura
        else:
            buffer.append(linea)
    return DDT(numero, data, [RigaDDT(d, int(q), u or "pz") for d, q, u in righe], percorso)


# --------------------------------------------------------------------------- #
# interpretazione della descrizione
# --------------------------------------------------------------------------- #
@dataclass
class Prodotto:
    descrizione_originale: str
    prodotto: str
    formato: str | None
    confezioni: int
    pezzi_per_confezione: int | None
    lotto: int | None
    scadenza: dt.date | None
    data: dt.date | None
    etichetta_pesce: str
    etichetta_fusti: str
    etichetta_olio: str
    etichetta_vasi: str
    etichetta_capsule: str | None
    vasi: int | None
    materia_prima: int | None
    registrabile: bool = True
    avvisi: list[str] = field(default_factory=list)

    @property
    def prodotte(self) -> str | None:
        if self.pezzi_per_confezione is None:
            return None
        return f"QUANTITA' PRODOTTE {self.confezioni} CONF. DA {self.pezzi_per_confezione}"


OLII = [
    (r"\bosg\b|olio di semi", "osg", "Olio di semi di girasole"),
    (r"delicat", "O. delicato", "Olio delicato"),
    (r"oliva", "O.oliva", "Olio d'oliva"),
]


def data_produzione(data_ddt: dt.date, scadenza: dt.date | None,
                    avvisi: list[str]) -> dt.date | None:
    """Il giorno di produzione e' quello della scadenza, riportato sul mese del DDT.

    Se quel giorno cade dopo l'emissione del DDT la produzione e' del mese precedente.
    Il lotto non entra nel calcolo, anche quando coincide.
    """
    if scadenza is None:
        avvisi.append("scadenza mancante: SCADENZA e DATA lasciate vuote")
        return None

    anno, mese = data_ddt.year, data_ddt.month
    if scadenza.day > data_ddt.day:
        mese -= 1
        if mese == 0:
            anno, mese = anno - 1, 12
    try:
        return dt.date(anno, mese, scadenza.day)
    except ValueError:
        avvisi.append(
            f"il {mese:02d}/{anno} non ha il giorno {scadenza.day} della scadenza: "
            "DATA lasciata vuota"
        )
        return None


# Nei DDT diversi da FLOTT ogni riga comincia col codice del cliente: "sg) vasi di ...",
# "et vasi di ..." (la parentesi ogni tanto manca), "pas)Vasi di ..." (lo spazio pure).
CODICE_CLIENTE = re.compile(r"\s*([A-Za-z]{1,4})\)|\s*([A-Za-z]{1,4})\s+(?=[Vv]asi\b)")


def prefisso_cliente(testo: str) -> str | None:
    m = CODICE_CLIENTE.match(testo)
    return (m.group(1) or m.group(2)).lower() if m else None


def interpreta(riga: RigaDDT, formati: dict[str, Formato], data_ddt: dt.date,
               col_prefisso: bool = False) -> Prodotto:
    testo = re.sub(r"\s+", " ", riga.descrizione).strip()
    basso = testo.lower()
    avvisi = []

    if "tonno" in basso:
        avvisi.append(f"prodotto a base di tonno, non registrato: {testo}")
        return Prodotto(testo, testo, None, riga.confezioni, None, None, None, None,
                        "", "", "", "", None, None, None, registrabile=False, avvisi=avvisi)

    sarde = "sarde" in basso
    sigla_olio, etichetta_olio = "", ""
    for schema, sigla, etichetta in OLII:
        if re.search(schema, basso):
            sigla_olio, etichetta_olio = sigla, etichetta
            break
    if not sigla_olio:
        avvisi.append("tipo di olio non riconosciuto")

    # "da gr. 160", "da gr 235", "gr . 600": il "da" non c'e' sempre
    m = re.search(r"\bgr\s*\.?\s*\.?\s*(\d+)", basso)
    grammatura = m.group(1) if m else None
    if grammatura is None:
        # i formati grandi a volte sono scritti in chili: "da kg. 1,5" sono 1500 grammi
        m = re.search(r"\bkg\s*\.?\s*(\d+(?:[.,]\d+)?)", basso)
        if m:
            grammatura = str(round(float(m.group(1).replace(",", ".")) * 1000))
    if grammatura is None:
        avvisi.append("grammatura non trovata: niente quantita' di materiali")
    ermetico = bool(re.search(r"\berm\b|\berm\.", basso))
    bonta = "bontà" in basso or "bonta" in basso

    m = re.search(r"conf\.?\s*da\s*(\d+)", basso)
    pezzi = int(m.group(1)) if m else None
    if pezzi is None:
        avvisi.append("pezzi per confezione non trovati")

    m = re.search(r"\bl\.?\s*(\d{5})\b", basso)
    lotto = int(m.group(1)) if m else None
    if lotto is None:
        avvisi.append("lotto mancante")

    m = re.search(r"\bsc[.:]?\s*(\d{2}/\d{2}/\d{4})", basso)
    scadenza = dt.datetime.strptime(m.group(1), "%d/%m/%Y").date() if m else None

    coda = testo[m.end():].strip() if m else ""
    coda = re.sub(r"^(per|x)\s+", "", coda, flags=re.I).strip()

    data = data_produzione(data_ddt, scadenza, avvisi)

    chiave = None
    if grammatura:
        chiave = f"{grammatura} bontà" if bonta and f"{grammatura} bontà" in formati else grammatura
        if chiave not in formati:
            avvisi.append(f"formato {grammatura} assente in FORMATI E CAPSULE")
            chiave = None

    vasi = riga.confezioni * pezzi if pezzi else None
    materia_prima = capsule = None
    if chiave:
        f = formati[chiave]
        coefficiente = f.coefficiente
        for etichetta, valore in f.varianti.items():
            if etichetta.lower().replace("'", "") in basso.replace("'", ""):
                coefficiente = valore
        if vasi is not None and coefficiente:
            materia_prima = round(vasi * coefficiente)
        capsule = f"Capsule diam. {f.capsula}" if isinstance(f.capsula, int) else None

    suffisso = f"{' erm.' if ermetico else ''}{' bontà' if bonta else ''}"
    pesce = "sarde" if sarde else "acc."
    prefisso = ""
    if col_prefisso:
        codice = prefisso_cliente(testo)
        if codice:
            prefisso = f"{codice})"
        else:
            avvisi.append("codice del cliente non trovato in testa alla riga")
    nome = " ".join(x for x in (prefisso, "Vasi Fil. di", pesce, sigla_olio,
                                f"gr. {grammatura}" if grammatura else "", suffisso.strip(), coda) if x)

    return Prodotto(
        descrizione_originale=testo,
        prodotto=re.sub(r"\s+", " ", nome).strip(),
        formato=chiave,
        confezioni=riga.confezioni,
        pezzi_per_confezione=pezzi,
        lotto=lotto,
        scadenza=scadenza,
        data=data,
        etichetta_pesce="Sarde sottovuoto" if sarde else "Acciughe sottovuoto",
        etichetta_fusti="Fil. di sarde fust. in olio" if sarde else "Fil. di acciughe fust. in olio",
        etichetta_olio=etichetta_olio,
        etichetta_vasi=f"Vasi gr. {grammatura}{suffisso}" if grammatura else "",
        etichetta_capsule=capsule,
        vasi=vasi,
        materia_prima=materia_prima,
        registrabile=True,
        avvisi=avvisi,
    )


# --------------------------------------------------------------------------- #
# scrittura del foglio
# --------------------------------------------------------------------------- #
# celle che cambiano da prodotto a prodotto, in coordinate relative al blocco:
# vanno svuotate, perche' il modello puo' essere un foglio gia' compilato
VARIABILI = ([(R_PRODOTTO, 1), (R_DATA, 1), (R_LOTTO, 1), (R_SCADENZA, 1), (R_PRODOTTE, 0)]
             + [(riga, colonna) for riga in range(16, 27) for colonna in range(0, 4)])


def _replica_blocchi(destinazione, modello, quanti: int) -> None:
    """Ricopia il modulo vuoto (colonne A-E del modello) affiancandolo `quanti` volte.

    Del modello si usa solo il primo blocco: se e' un foglio gia' scritto, le sue celle
    unite arrivano fino all'ultimo blocco e invaderebbero quelli che stiamo creando.
    """
    unioni = [u for u in modello.merged_cells.ranges if u.max_col <= LARGH_BLOCCO]

    for riga in range(1, RIGHE + 1):
        if modello.row_dimensions[riga].height:
            destinazione.row_dimensions[riga].height = modello.row_dimensions[riga].height

    for blocco in range(quanti):
        scarto = blocco * LARGH_BLOCCO
        for colonna in range(1, LARGH_BLOCCO + 1):
            sorgente = modello.column_dimensions[get_column_letter(colonna)]
            arrivo = destinazione.column_dimensions[get_column_letter(colonna + scarto)]
            arrivo.width = sorgente.width
            for riga in range(1, RIGHE + 1):
                da = modello.cell(row=riga, column=colonna)
                a = destinazione.cell(row=riga, column=colonna + scarto)
                a.value = da.value
                a._style = copy(da._style)
        for unione in unioni:
            u = copy(unione)
            u.shift(col_shift=scarto)
            destinazione.merge_cells(str(u))

    # il modello puo' portarsi dietro i dati del prodotto che conteneva
    for blocco in range(quanti):
        base = 2 + blocco * LARGH_BLOCCO
        for riga, colonna in VARIABILI:
            cella = destinazione.cell(row=riga, column=base + colonna)
            if not isinstance(cella, MergedCell):
                cella.value = None

    # colonna di stacco finale, come nei fogli compilati a mano
    coda = get_column_letter(quanti * LARGH_BLOCCO + 1)
    destinazione.column_dimensions[coda].width = modello.column_dimensions["A"].width

    destinazione.sheet_view.showGridLines = modello.sheet_view.showGridLines
    destinazione.sheet_view.zoomScale = modello.sheet_view.zoomScale
    destinazione.page_setup = copy(modello.page_setup)
    destinazione.page_margins = copy(modello.page_margins)
    destinazione.print_options = copy(modello.print_options)
    destinazione.sheet_properties.pageSetUpPr = copy(modello.sheet_properties.pageSetUpPr)


def scrivi_foglio(destinazione, modello, prodotti: list[Prodotto]) -> None:
    _replica_blocchi(destinazione, modello, len(prodotti))
    for blocco, p in enumerate(prodotti):
        b = 2 + blocco * LARGH_BLOCCO          # colonna "B" del blocco
        def scrivi(riga, colonna, valore):
            if valore is not None:
                destinazione.cell(row=riga, column=b + colonna).value = valore

        scrivi(R_PRODOTTO, 1, p.prodotto)
        scrivi(R_DATA, 1, p.data)
        scrivi(R_LOTTO, 1, p.lotto)
        scrivi(R_SCADENZA, 1, p.scadenza)

        scrivi(R_PESCE, 0, p.etichetta_pesce)
        scrivi(R_PESCE, 3, p.materia_prima)
        scrivi(R_FUSTI, 0, p.etichetta_fusti)
        for colonna in (1, 2, 3):
            scrivi(R_FUSTI, colonna, "=")
        scrivi(R_OLIO, 0, p.etichetta_olio or None)
        scrivi(R_OLIO, 3, p.materia_prima)
        scrivi(R_VASI, 0, p.etichetta_vasi or None)
        scrivi(R_VASI, 3, p.vasi)
        if p.etichetta_capsule:
            scrivi(R_CAPSULE, 0, p.etichetta_capsule)
            scrivi(R_CAPSULE, 3, p.vasi)
        scrivi(R_PRODOTTE, 0, p.prodotte)


def _ordina(wb) -> None:
    def chiave(ws):
        m = re.match(r"(\d+)_(\d{6})", ws.title)
        return (int(m.group(2)[4:6]), int(m.group(2)[2:4]), int(m.group(1))) if m else (99, 99, 10**9)
    wb._sheets.sort(key=chiave)


# openpyxl scrive <t> </t> senza dichiarare che gli spazi vanno conservati. Succede solo
# nelle celle con piu' formati dentro (il grassetto di una parola sola, le sottolineature
# del registro olio) ed Excel lo considera contenuto danneggiato: all'apertura chiede di
# riparare il file. Rimedio dopo il salvataggio, riscrivendo i fogli che ne hanno bisogno.
_SPAZI = re.compile(rb'<t( xml:space="preserve")?>([^<]*)</t>')


def _preserva_spazi(xml: bytes) -> bytes:
    def rimedia(m):
        attributi, testo = m.group(1), m.group(2)
        if attributi or testo == testo.strip():
            return m.group(0)
        return b'<t xml:space="preserve">' + testo + b'</t>'
    return _SPAZI.sub(rimedia, xml)


def salva_libro(wb, percorso: Path) -> None:
    """Salva il file e sistema gli spazi del testo formattato."""
    percorso.parent.mkdir(parents=True, exist_ok=True)
    wb.save(percorso)

    with zipfile.ZipFile(percorso) as zip_letto:
        parti = [(i, zip_letto.read(i.filename)) for i in zip_letto.infolist()]
    corrette = [(i, _preserva_spazi(d) if i.filename.startswith(("xl/worksheets/sheet",
                                                                "xl/sharedStrings")) else d)
                for i, d in parti]
    if all(a[1] == b[1] for a, b in zip(parti, corrette)):
        return                       # niente da rimediare, lascio il file com'e'

    temporaneo = percorso.with_suffix(".xlsx.tmp")
    with zipfile.ZipFile(temporaneo, "w", zipfile.ZIP_DEFLATED) as zip_scritto:
        for info, dati in corrette:
            zip_scritto.writestr(info, dati)
    temporaneo.replace(percorso)


def elenca(modificati: list[Path]) -> None:
    """Stampa i file toccati davvero, con le ultime due parti del percorso."""
    if not modificati:
        return
    print("  file modificati:")
    for f in modificati:
        print(f"    {Path(*f.parts[-2:])}")


def leggi_visti(percorso: Path) -> set[str]:
    """I documenti gia' analizzati in passato.

    I documenti che non producono nessun foglio — quelli di solo tonno, le fatture di
    soli imballaggi — non lasciano traccia nei registri: senza questo elenco le loro
    segnalazioni tornerebbero a ogni esecuzione.
    """
    if not percorso.exists():
        return set()
    return {riga.strip() for riga in percorso.read_text(encoding="utf-8").splitlines()
            if riga.strip()}


def scrivi_visti(percorso: Path, visti: set[str]) -> None:
    percorso.parent.mkdir(parents=True, exist_ok=True)
    percorso.write_text("\n".join(sorted(visti)) + "\n", encoding="utf-8")


def plurale(quanti: int, uno: str, molti: str) -> str:
    return f"{quanti} {uno if quanti == 1 else molti}"


def salva_copia(percorso: Path, cartella: Path | None) -> None:
    """Copia di sicurezza del file prima di modificarlo.

    Con `cartella` le copie finiscono tutte li', cosi' le cartelle dei registri
    restano pulite; senza, resta accanto al file come prima.
    """
    if cartella is None:
        shutil.copy2(percorso, percorso.with_suffix(".xlsx.bak"))
        return
    cartella.mkdir(parents=True, exist_ok=True)
    shutil.copy2(percorso, cartella / (percorso.name + ".bak"))


# --------------------------------------------------------------------------- #
# registro delle anomalie
# --------------------------------------------------------------------------- #
class Registro:
    """Scrive a schermo e in coda al file di log le anomalie di ogni esecuzione."""

    def __init__(self, percorso: Path, intestazione: str = "DR.RPD.00 Riepilogo produzione"):
        self.percorso = percorso
        self.intestazione = intestazione
        self.righe: list[str] = []
        self.anomalie = 0

    def anomalia(self, testo: str, dettaglio: str = "") -> None:
        self.anomalie += 1
        print(f"      ATTENZIONE {testo}")
        if dettaglio:
            print(f"        {dettaglio}")
        self.righe.append(f"    {testo}")
        if dettaglio:
            self.righe.append(f"        {dettaglio}")

    def contesto(self, testo: str) -> None:
        self.righe.append(f"  {testo}")

    def avvisa(self) -> None:
        """Rende impossibile non accorgersi che qualcosa e' rimasto fuori."""
        if not self.anomalie:
            return
        quante = "1 cosa inattesa" if self.anomalie == 1 else f"{self.anomalie} cose inattese"
        print("\n" + "!" * 66)
        print(f"!!  ATTENZIONE: {quante}. Controlla i fogli interessati.")
        print(f"!!  Dettagli in: {self.percorso}")
        print("!" * 66)

    def chiudi(self, riassunto: str) -> None:
        self.percorso.parent.mkdir(parents=True, exist_ok=True)
        nuovo = not self.percorso.exists()
        with self.percorso.open("a", encoding="utf-8") as f:
            if nuovo:
                f.write(f"Registro anomalie - {self.intestazione}\n")
            f.write(f"\n=== {dt.datetime.now():%d/%m/%Y %H:%M}  {riassunto}\n")
            if self.anomalie:
                f.write("\n".join(self.righe) + "\n")
            else:
                f.write("  nessuna anomalia\n")


# --------------------------------------------------------------------------- #
def risolvi(percorso: Path, data: dt.date) -> Path:
    """Sostituisce {anno} e {trimestre} con quelli del documento."""
    return Path(str(percorso)
                .replace("{anno}", str(data.year))
                .replace("{trimestre}", f"{(data.month - 1) // 3 + 1} TRIMESTRE"))


def main() -> int:
    ap = argparse.ArgumentParser(description="Compila il Riepilogo produzione dai DDT in PDF.")
    ap.add_argument("pdf", nargs="+", type=Path, help="file PDF o cartelle che li contengono")
    ap.add_argument("-o", "--output", type=Path, required=True,
                    help="xlsx da creare o aggiornare; {anno} e {trimestre} vengono "
                         "sostituiti con quelli del DDT, cosi' un documento di dicembre "
                         "lavorato a gennaio finisce nel file giusto")
    ap.add_argument("-f", "--formati", type=Path, default=FORMATI)
    ap.add_argument("--esterni", action="store_true",
                    help="DDT di clienti diversi da FLOTT: al nome del prodotto va davanti "
                         "il codice del cliente che sta in testa alla riga (sg), et), pas))")
    ap.add_argument("--sovrascrivi", action="store_true",
                    help="rifai i fogli dei DDT gia' presenti invece di saltarli")
    ap.add_argument("-l", "--log", type=Path,
                    help="file delle anomalie (predefinito: accanto all'xlsx, con estensione .log)")
    ap.add_argument("-b", "--backup", type=Path,
                    help="cartella dove mettere le copie .bak (predefinito: accanto al file)")
    args = ap.parse_args()

    registro = Registro(args.log or risolvi(args.output, dt.date.today()).with_suffix(".log"))

    percorsi = sorted({p for voce in args.pdf
                       for p in (voce.rglob("*.pdf") if voce.is_dir() else [voce])})
    if not percorsi:
        print("Nessun PDF trovato.", file=sys.stderr)
        return 1

    formati = leggi_formati(args.formati)

    # prima leggo tutti i DDT, poi li raggruppo per file di destinazione
    per_file: dict[Path, list[tuple[Path, DDT]]] = {}
    for percorso in percorsi:
        try:
            ddt = leggi_ddt(percorso)
        except ValueError as errore:
            print(f"  ! {errore}", file=sys.stderr)
            registro.contesto(percorso.name)
            registro.anomalia(str(errore).split(": ", 1)[-1])
            continue
        per_file.setdefault(risolvi(args.output, ddt.data), []).append((percorso, ddt))

    gia_visti = registro.percorso.with_name(registro.percorso.stem + "-visti.txt")
    visti = leggi_visti(gia_visti)
    aggiunti = 0
    prodotti_file = []
    for uscita in sorted(per_file):
        if len(per_file) > 1:
            print(f"\n── {Path(*uscita.parent.parts[-2:])} ──")
        prima = aggiunti
        if uscita.exists():
            wb = openpyxl.load_workbook(uscita)
            modello = wb.worksheets[0]
            salva_copia(uscita, args.backup)
        else:
            wb = openpyxl.load_workbook(MODELLO)
            modello = wb["MODELLO"]

        for percorso, ddt in per_file[uscita]:
            nome = f"{ddt.numero}_{ddt.data:%d%m%y}"
            if nome in wb.sheetnames:
                if not args.sovrascrivi:
                    continue
                wb.remove(wb[nome])

            letti = [interpreta(r, formati, ddt.data, args.esterni) for r in ddt.righe]
            prodotti = [p for p in letti if p.registrabile]
            # un DDT che non produce nessun foglio non lascia traccia: senza la memoria
            # dei documenti gia' letti le sue segnalazioni tornerebbero a ogni esecuzione
            da_dire = any(p.avvisi for p in letti) and (prodotti or nome not in visti)
            visti.add(nome)
            if da_dire:
                registro.contesto(f"foglio {nome}  ({percorso.name})"
                                  if prodotti else f"DDT {nome}  ({percorso.name})")
                for p in letti:
                    for avviso in p.avvisi:
                        registro.anomalia(avviso, f"riga: {p.descrizione_originale}")
            if not prodotti:
                continue                # niente da registrare: nessun foglio da creare
            scrivi_foglio(wb.create_sheet(nome), modello, prodotti)
            aggiunti += 1
            print(f"  + {nome}  ({plurale(len(prodotti), 'prodotto', 'prodotti')})")

        if aggiunti == prima:
            continue                    # niente di nuovo: non lo riscrivo nemmeno
        if "MODELLO" in wb.sheetnames and len(wb.sheetnames) > 1:
            wb.remove(wb["MODELLO"])
        _ordina(wb)
        salva_libro(wb, uscita)
        prodotti_file.append(uscita)

    scrivi_visti(gia_visti, visti)
    riassunto = (f"{plurale(len(percorsi), 'PDF letto', 'PDF letti')}, "
                 f"{plurale(aggiunti, 'foglio aggiunto', 'fogli aggiunti')}, "
                 f"{plurale(registro.anomalie, 'anomalia', 'anomalie')}")
    registro.chiudi(riassunto)
    print(f"\n{riassunto}")
    elenca(prodotti_file)
    registro.avvisa()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
