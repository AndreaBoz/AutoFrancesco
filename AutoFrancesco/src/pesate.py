"""Dai DDT di vendita ai moduli DR.CPF.70 di controllo finale sul prodotto finito.

C'e' un file per formato di vaso, che continua nel tempo. Di ogni foglio cambiano solo
data, lotto e nome del prodotto: tutto il resto, comprese le formule RANDBETWEEN che
generano le pesate, viene ricopiato tale e quale dal foglio precedente.
"""
from __future__ import annotations

import argparse
import math
import random
import re
import shutil
import sys
from copy import copy
from pathlib import Path

import openpyxl

import datetime as dt

from riepilogo import (FORMATI, Registro, elenca, interpreta, leggi_ddt, leggi_formati,
                       leggi_visti, plurale, risolvi, salva_copia, salva_libro,
                       scrivi_visti)

RADICE = Path(__file__).resolve().parent          # la cartella src

# I file di partenza contengono decine di MB di righe vettoriali (23 MB nel solo gr 235):
# openpyxl le analizza in fase di lettura ma poi le scarta scrivendo, quindi e' tempo
# buttato. Saltandole la lettura passa da una cinquantina di secondi a una decina.
openpyxl.reader.excel.find_images = lambda archivio, percorso: ([], [])

C_DATA, C_LOTTO, C_PRODOTTO = "B5", "E5", "F5"
# =RANDBETWEEN(A20*1000,C20*1000)/1000 -> estremi presi da altre due celle del foglio
SORTEGGIO = re.compile(r"^=RANDBETWEEN\(\s*(\$?[A-Z]+\$?\d+)\s*\*\s*1000\s*,"
                       r"\s*(\$?[A-Z]+\$?\d+)\s*\*\s*1000\s*\)\s*/\s*1000$")
NOME_FILE = "DR.CPF.70 Controllo finale PRODOTTO FINITO gr {formato} (FORMULA).xlsx"


def nome_file(formato: str) -> str:
    return NOME_FILE.format(formato=f"{int(formato):03d}")


def sorteggia(ws, formula: str):
    """Estrae un valore con la stessa regola della formula, ma da scrivere come numero.

    Cosi' la pesata resta quella che e' anche riaprendo il file: RANDBETWEEN e' volatile
    e si rigenererebbe a ogni ricalcolo, su tutti i fogli. Il valore si scrive **intero**,
    cioe' com'e' scritto sul foglio: le celle hanno formato "0" e i millesimi della
    formula si vedrebbero solo entrando nella cella.
    """
    m = SORTEGGIO.match(formula.replace(" ", ""))
    if not m:
        return None
    estremi = []
    for riferimento in m.groups():
        valore = ws[riferimento.replace("$", "")].value
        if not isinstance(valore, (int, float)):
            return None
        estremi.append(int(round(valore * 1000)))
    return intero(random.randint(min(estremi), max(estremi)) / 1000)


# le pesate stanno sempre in questo rettangolo, in tutti i formati
PRIMA_RIGA, ULTIMA_RIGA, PRIMA_COL, ULTIMA_COL = 11, 25, 4, 11


def intero(valore: float) -> int:
    """Arrotonda come fa Excel col formato "0": il mezzo va sempre in su."""
    return math.floor(valore + 0.5)


def arrotonda(wb) -> int:
    """Toglie i decimali alle pesate gia' scritte, senza cambiare cio' che si legge.

    Le celle hanno formato "0", quindi 234,631 a schermo era gia' 235: qui il valore
    diventa quello che il foglio mostrava. Medie e verifiche si ricalcolano da se'.
    """
    quanti = 0
    for ws in wb.worksheets:
        for riga in range(PRIMA_RIGA, ULTIMA_RIGA + 1):
            for colonna in range(PRIMA_COL, ULTIMA_COL + 1):
                cella = ws.cell(riga, colonna)
                if isinstance(cella.value, float) and not cella.value.is_integer():
                    cella.value = intero(cella.value)
                    quanti += 1
    return quanti


def da_congelare(wb) -> bool:
    """C'e' ancora almeno una RANDBETWEEN da trasformare in numero?"""
    return any(isinstance(ws.cell(r, c).value, str)
               and ws.cell(r, c).value.startswith("=RANDBETWEEN")
               for ws in wb.worksheets
               for r in range(PRIMA_RIGA, ULTIMA_RIGA + 1)
               for c in range(PRIMA_COL, ULTIMA_COL + 1))


_schema_noto: dict = {}


def schema_condiviso(esempio: Path) -> dict:
    """Lo schema delle pesate e' lo stesso in tutti i formati: lo leggo una volta sola.

    Verificato sui 12 file di partenza che hanno le formule: stesse 112 celle da D11 a
    K25 e stessi due riferimenti (A20/C20 per il netto, A25/C25 per lo sgocciolato).
    Rileggere un file di esempio per ogni formato costava piu' di tutto il resto.
    """
    if not _schema_noto and esempio.exists():
        _schema_noto.update(schema_pesate(openpyxl.load_workbook(esempio, rich_text=True)))
    return dict(_schema_noto)


def schema_pesate(wb) -> dict:
    """{cella: formula} delle pesate, da un foglio che ha ancora le formule originali.

    Serve a risorteggiare i valori di ogni foglio nuovo: una volta congelati i numeri,
    nel file non c'e' piu' una formula da cui capire quali celle sono pesate.
    """
    for ws in reversed(wb.worksheets):
        trovate = {c.coordinate: c.value for riga in ws.iter_rows() for c in riga
                   if isinstance(c.value, str) and c.value.startswith("=RANDBETWEEN")}
        if trovate:
            return trovate
    return {}


def congela(ws, valori=None, avvisi=None) -> int:
    """Sostituisce le RANDBETWEEN con numeri: quelli gia' calcolati, o nuovi sorteggi."""
    quanti = 0
    for riga in ws.iter_rows():
        for cella in riga:
            if not (isinstance(cella.value, str) and cella.value.startswith("=RANDBETWEEN")):
                continue
            fissato = None
            if valori is not None:
                fissato = valori[cella.coordinate].value
            if not isinstance(fissato, (int, float)):
                fissato = sorteggia(ws, cella.value)
                if fissato is not None and valori is not None and avvisi is not None:
                    avvisi.append(f"{ws.title} {cella.coordinate}: pesata mai calcolata, "
                                  f"ne ho sorteggiata una nuova")
            if fissato is not None:
                cella.value = fissato
                quanti += 1
    return quanti


def _modello(wb):
    """Come modello uso l'ultimo foglio compilato: e' il piu' recente."""
    compilati = [ws for ws in wb.worksheets if ws[C_DATA].value is not None]
    for ws in reversed(compilati or wb.worksheets):
        if isinstance(ws["D11"].value, (int, float, str)):
            return ws
    return None


def _nuovo_foglio(wb, modello, nome: str):
    """copy_worksheet non porta con se' vista e area di stampa: le ricopio a mano."""
    ws = wb.copy_worksheet(modello)
    ws.title = nome
    ws.sheet_view.showGridLines = modello.sheet_view.showGridLines
    ws.sheet_view.zoomScale = modello.sheet_view.zoomScale
    ws.sheet_view.zoomScaleNormal = modello.sheet_view.zoomScaleNormal
    if modello.print_area:
        ws.print_area = str(modello.print_area).split("!")[-1]
    return ws


def ripulisci_nome(titolo: str, produzione=None) -> str:
    """Toglie i caratteri di controllo e rimette in sesto le date storte nei nomi.

    Alcuni fogli vecchi cominciano con un carattere invisibile (_x0001_), uno finisce con
    un trattino basso di troppo (179_270726_) e uno ha la data di sette cifre invece di
    sei: 57_1200326 al posto di 57_120326. Fra le date possibili scelgo quella piu' vicina
    al giorno di produzione, e mai precedente: il DDT non puo' essere emesso prima di aver
    prodotto.
    """
    pulito = re.sub(r"_x[0-9A-Fa-f]{4}_", "", titolo).strip().rstrip("_ ")
    m = re.fullmatch(r"(\d+)_(\d+)(-[A-Z])?", pulito)
    if not m or len(m.group(2)) == 6:
        return pulito

    cifre = m.group(2)
    candidate = {}
    for i in range(len(cifre)):
        prova = cifre[:i] + cifre[i + 1:]
        if len(prova) != 6:
            continue
        try:
            candidate[prova] = dt.datetime.strptime(prova, "%d%m%y").date()
        except ValueError:
            continue
    if not candidate:
        return pulito
    if isinstance(produzione, dt.datetime):
        produzione = produzione.date()
    if isinstance(produzione, dt.date):
        scelta = min(candidate, key=lambda k: (candidate[k] < produzione,
                                               abs((candidate[k] - produzione).days)))
    else:
        scelta = min(candidate)
    return f"{m.group(1)}_{scelta}{m.group(3) or ''}"


# Nomi con cui e' arrivata, di file in file, la scheda di partenza da cui si copiano le
# formule: sono tutti la stessa cosa e il programma la cerca sotto un nome solo.
SCHEDE_MODELLO = {"demo", "foglio1", "foglio 1", "sheet1", "sheet 1", "modello"}


def rinomina_fogli(wb, registro) -> int:
    quanti = 0
    for ws in wb.worksheets:
        if ws.title.strip().lower() in SCHEDE_MODELLO:
            if ws.title != "MODELLO" and "MODELLO" not in wb.sheetnames:
                ws.title = "MODELLO"
                quanti += 1
            continue
        nuovo = ripulisci_nome(ws.title, ws[C_DATA].value)
        if nuovo == ws.title:
            continue
        if nuovo in wb.sheetnames:
            registro.anomalia(f"non posso rinominare {ws.title!r} in {nuovo!r}: "
                              f"esiste gia' un foglio con quel nome")
            continue
        ws.title = nuovo
        quanti += 1
    return quanti


def _ordina(wb) -> None:
    """In ordine di numero di DDT, come nei fogli compilati a mano."""
    def chiave(ws):
        m = re.match(r"(\d+)\D*(?:-([A-Z]))?", ripulisci_nome(ws.title))
        return (int(m.group(1)), m.group(2) or "") if m else (10 ** 9, "")
    wb._sheets.sort(key=chiave)


def sblocca(ws) -> None:
    """I fogli generati non vanno mai protetti, nemmeno se lo era quello di partenza."""
    ws.protection.sheet = False
    ws.protection.objects = False
    ws.protection.scenarios = False


def _suffissi(quanti: int):
    return [""] + [f"-{chr(ord('A') + i)}" for i in range(quanti - 1)]


def sistema(cartella: Path, registro, backup: Path | None) -> list[Path]:
    """Mette in riga i file gia' presenti, anche quelli che oggi non ricevono pesate.

    Serve quando si parte da schede compilate a mano: tolgo la protezione, fisso le
    pesate (le RANDBETWEEN rimaste vive si risorteggiano a ogni apertura di Excel),
    uniformo il nome della scheda di partenza e rimetto i fogli in ordine di DDT.
    """
    sistemati = []
    for percorso in sorted(cartella.glob("DR.CPF.70*.xlsx")):
        wb = openpyxl.load_workbook(percorso, rich_text=True)
        rinominati = rinomina_fogli(wb, registro)
        avvisi: list[str] = []
        fissate = 0
        if da_congelare(wb):
            valori = openpyxl.load_workbook(percorso, data_only=True)
            fissate = sum(congela(ws, calcolato, avvisi)
                          for ws, calcolato in zip(wb.worksheets, valori.worksheets)
                          if ws.title != "MODELLO")
        arrotondate = arrotonda(wb)
        protetti = sum(1 for ws in wb.worksheets if ws.protection.sheet)
        for ws in wb.worksheets:
            sblocca(ws)
        ordine = list(wb.sheetnames)
        _ordina(wb)
        if not (rinominati or fissate or arrotondate or protetti or wb.sheetnames != ordine):
            continue
        registro.contesto(percorso.name)
        for a in avvisi[:5]:
            registro.anomalia(a)
        if len(avvisi) > 5:
            registro.anomalia(f"...e altre {len(avvisi) - 5} pesate nella stessa condizione")
        fatto = [f"{fissate} pesate fissate" if fissate else "",
                 f"{arrotondate} pesate arrotondate" if arrotondate else "",
                 f"{protetti} fogli sprotetti" if protetti else "",
                 plurale(rinominati, "foglio rinominato", "fogli rinominati") if rinominati else "",
                 "rimessi in ordine di DDT" if wb.sheetnames != ordine else ""]
        print(f"  {percorso.name.split('gr ')[-1].split(' ')[0]}: "
              + ", ".join(x for x in fatto if x))
        salva_copia(percorso, backup)
        wb.calculation.fullCalcOnLoad = True
        salva_libro(wb, percorso)
        sistemati.append(percorso)
    return sistemati


def main() -> int:
    ap = argparse.ArgumentParser(description="Compila i moduli di controllo finale (pesate).")
    ap.add_argument("pdf", nargs="*", type=Path, help="DDT in PDF o cartelle che li contengono")
    ap.add_argument("-o", "--uscita", type=Path, required=True,
                    help="cartella dei file DR.CPF.70, uno per formato")
    ap.add_argument("-m", "--modelli", type=Path, default=RADICE / "Modelli pesate",
                    help="cartella con le schede vuote, una per formato, da cui parte "
                         "il file di un formato non ancora presente nella destinazione")
    ap.add_argument("-f", "--formati", type=Path, default=FORMATI)
    ap.add_argument("--sovrascrivi", action="store_true",
                    help="rifai i fogli dei DDT gia' presenti invece di saltarli")
    ap.add_argument("--sistema", action="store_true",
                    help="passa su tutti i file della cartella e li mette in riga: toglie "
                         "la protezione, fissa le pesate ancora a formula, uniforma il nome "
                         "della scheda di partenza, riordina per numero di DDT")
    ap.add_argument("-l", "--log", type=Path)
    ap.add_argument("-b", "--backup", type=Path,
                    help="cartella dove mettere le copie .bak (predefinito: accanto al file)")
    args = ap.parse_args()

    percorsi = sorted({p for voce in args.pdf
                       for p in (voce.rglob("*.pdf") if voce.is_dir() else [voce])})
    if not percorsi and not args.sistema:
        print("Nessun PDF trovato.", file=sys.stderr)
        return 1
    registro = Registro(args.log or risolvi(args.uscita, dt.date.today()) / "pesate.log",
                        intestazione="DR.CPF.70 pesate")
    formati = leggi_formati(args.formati)

    if args.sistema:
        cartella = risolvi(args.uscita, dt.date.today())
        print(f"── sistemo {Path(*cartella.parts[-2:])} ──")
        sistemati = sistema(cartella, registro, args.backup)
        if not sistemati:
            print("  gia' tutto a posto")
        if not percorsi:
            riassunto = (f"{plurale(len(sistemati), 'file sistemato', 'file sistemati')}, "
                         f"{plurale(registro.anomalie, 'anomalia', 'anomalie')}")
            registro.chiudi(riassunto)
            print(f"\n{riassunto}")
            elenca(sistemati)
            registro.avvisa()
            return 0
        print()

    # i DDT che non producono nessuna pesata (il tonno) non lasciano traccia da nessuna
    # parte: senza la memoria dei documenti gia' letti tornerebbero a segnalarsi sempre
    gia_visti = registro.percorso.with_name(registro.percorso.stem + "-visti.txt")
    visti = leggi_visti(gia_visti)

    # raggruppo per cartella di destinazione e formato: un file per formato
    per_formato: dict[tuple[Path, str], list] = {}
    for percorso in percorsi:
        try:
            ddt = leggi_ddt(percorso)
        except ValueError as errore:
            registro.contesto(percorso.name)
            registro.anomalia(str(errore).split(": ", 1)[-1])
            continue
        for prodotto in (interpreta(r, formati, ddt.data) for r in ddt.righe):
            if not prodotto.registrabile:
                chiave = f"{ddt.numero}_{ddt.data:%d%m%y}"
                if chiave not in visti:
                    registro.contesto(f"DDT {ddt.numero} del {ddt.data:%d/%m/%Y}"
                                      f"  ({percorso.name})")
                    for avviso in prodotto.avvisi:
                        registro.anomalia(avviso)
                    visti.add(chiave)
                continue
            grammatura = re.match(r"\d+", prodotto.formato or "")
            if not grammatura:
                registro.contesto(f"DDT {ddt.numero} del {ddt.data:%d/%m/%Y}  ({percorso.name})")
                registro.anomalia(f"formato sconosciuto, niente pesata: "
                                  f"{prodotto.descrizione_originale}")
                continue
            per_formato.setdefault((risolvi(args.uscita, ddt.data), grammatura.group(0)),
                                   []).append((ddt, prodotto, percorso))

    aggiunti = 0
    salvati = []
    for cartella, formato in sorted(per_formato, key=lambda x: (str(x[0]), int(x[1]))):
        cartella.mkdir(parents=True, exist_ok=True)
        uscita = cartella / nome_file(formato)
        if not uscita.exists():
            modello_file = args.modelli / nome_file(formato)
            if not modello_file.exists():
                registro.contesto(f"formato {formato}")
                registro.anomalia(f"manca la scheda vuota {modello_file.name}: "
                                  f"{len(per_formato[(cartella, formato)])} pesate non create")
                continue
            shutil.copy2(modello_file, uscita)     # parte dalla scheda vuota del formato
        else:
            salva_copia(uscita, args.backup)

        # rich_text=True: nelle intestazioni E1 ed E3 il grassetto e' solo su una parte
        # del testo, e senza questo openpyxl lo appiattisce
        wb = openpyxl.load_workbook(uscita, rich_text=True)
        registro.contesto(uscita.name)
        rinominati = rinomina_fogli(wb, registro)
        # lo schema va letto prima di congelare, altrimenti le formule non ci sono piu'
        schema = schema_pesate(wb) or schema_condiviso(args.modelli / nome_file(formato))
        _schema_noto.update(schema)
        if not schema:
            registro.contesto(f"formato {formato}")
            registro.anomalia(f"in {uscita.name} non trovo le formule delle pesate: "
                              f"{len(per_formato[(cartella, formato)])} fogli non creati")
            continue

        # I fogli gia' scritti vanno fissati ai loro valori, altrimenti le RANDBETWEEN
        # se li rigenerano a ogni apertura del file. Se pero' nel file non e' rimasta
        # nessuna formula da congelare, evito di rileggerlo una seconda volta.
        avvisi_congelamento: list[str] = []
        fissate = 0
        if da_congelare(wb):
            valori = openpyxl.load_workbook(uscita, data_only=True)
            # per posizione, non per nome: i fogli possono essere appena stati rinominati
            # la scheda vuota non va congelata: le sue formule servono a sorteggiare
            # i valori dei fogli nuovi
            fissate = sum(congela(ws, calcolato, avvisi_congelamento)
                          for ws, calcolato in zip(wb.worksheets, valori.worksheets)
                          if ws.title != "MODELLO")
        # i fogli gia' presenti possono portarsi dietro i millesimi delle prime versioni
        arrotondate = arrotonda(wb)
        if avvisi_congelamento:
            registro.contesto(f"{uscita.name}")
            for a in avvisi_congelamento[:5]:
                registro.anomalia(a)
            if len(avvisi_congelamento) > 5:
                registro.anomalia(f"...e altre {len(avvisi_congelamento) - 5} pesate nella "
                                  f"stessa condizione")
        modello = _modello(wb)
        if modello is None:
            registro.contesto(f"formato {formato}")
            registro.anomalia(f"{uscita.name} non ha un foglio da cui copiare le formule: "
                              f"{len(per_formato[(cartella, formato)])} pesate non create")
            continue

        da_dire: list[str] = []
        # i prodotti dello stesso DDT prendono i suffissi -A, -B nello stesso file
        per_ddt: dict[str, list] = {}
        for ddt, prodotto, percorso in per_formato[(cartella, formato)]:
            per_ddt.setdefault(f"{ddt.numero}_{ddt.data:%d%m%y}", []).append((ddt, prodotto, percorso))

        for base, elenco in per_ddt.items():
            gia = [n for n in wb.sheetnames if n == base or n.startswith(base + "-")]
            if gia and not args.sovrascrivi:
                continue
            for n in gia:
                wb.remove(wb[n])
            for (ddt, prodotto, percorso), suffisso in zip(elenco, _suffissi(len(elenco))):
                ws = _nuovo_foglio(wb, modello, base + suffisso)
                ws[C_DATA] = prodotto.data
                ws[C_LOTTO] = prodotto.lotto
                ws[C_PRODOTTO] = prodotto.prodotto
                for cella, formula in schema.items():   # pesate nuove, non quelle copiate
                    valore = sorteggia(ws, formula)
                    if valore is not None:
                        ws[cella] = valore
                aggiunti += 1
                da_dire.append(f"  + {base + suffisso}  {prodotto.prodotto}")
                for avviso in prodotto.avvisi:
                    registro.contesto(f"foglio {base + suffisso}  ({percorso.name})")
                    registro.anomalia(avviso)

        if not da_dire and not arrotondate:
            continue                    # niente di nuovo: non lo riscrivo nemmeno
        if "MODELLO" in wb.sheetnames and len(wb.sheetnames) > 1:
            wb.remove(wb["MODELLO"])
        for foglio in wb.worksheets:
            sblocca(foglio)
        _ordina(wb)
        if rinominati:
            da_dire.append(f"  {plurale(rinominati, 'foglio rinominato', 'fogli rinominati')}"
                           " (caratteri invisibili o data storta)")
        if fissate:
            da_dire.append(f"  {fissate} pesate fissate nei fogli gia' presenti")
        if arrotondate:
            da_dire.append(f"  {arrotondate} pesate arrotondate nei fogli gia' presenti")
        wb.calculation.fullCalcOnLoad = True   # medie e verifiche si ricalcolano da sole
        print(f"\n── gr {int(formato):03d} ──")
        print("\n".join(da_dire))
        salva_libro(wb, uscita)
        salvati.append(uscita)

    scrivi_visti(gia_visti, visti)
    riassunto = (f"{plurale(len(percorsi), 'PDF letto', 'PDF letti')}, "
                 f"{plurale(aggiunti, 'pesata aggiunta', 'pesate aggiunte')}, "
                 f"{plurale(registro.anomalie, 'anomalia', 'anomalie')}")
    registro.chiudi(riassunto)
    print(f"\n{riassunto}")
    elenca(salvati)
    registro.avvisa()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
