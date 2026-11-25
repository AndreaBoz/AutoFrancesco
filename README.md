# AutoFrancesco — HACCP Register Automation

Python automation suite that reads sales delivery notes (DDT) and purchase invoices in PDF format and populates local Excel HACCP registers (`~/HACCP`).

This repository brings together both operational modules sharing the same processing engine:

| Folder | Scope | Input Documents | Generated / Updated Registers |
|---|---|---|---|
| [`AutoFrancesco/`](./AutoFrancesco/README.md) | **FLOTT** workflow | Sales DDTs and FLOTT invoices (`PDF Vendite`, `PDF Acquisti`) | `DR.RPD.00`, `DR.CPF.70` (weights), `DR.CAM`, `DR.CAO`, `DR.CAI`, `REGISTRO GIACENZE` |
| [`AutoFrancesco Esterni/`](./AutoFrancesco%20Esterni/README.md) | **External** clients & suppliers | External client DDTs and multi-supplier invoices (`EST PDF Vendite`, `EST PDF Acquisti`) | `EST-Riepilogo produzione`, `DR.CPF.70` (weights), `EST-Controllo` (MAT.PRIME / OLIO / IMBALLAGGI), `EST-REGISTRO GIACENZE` |

---

## Project Structure

```text
.
├── README.md
├── requirements.txt
├── .gitignore
├── AutoFrancesco/
│   ├── README.md                       Operational guide (FLOTT, Italian)
│   ├── genera.command                  Double-click launcher (FLOTT HACCP archive)
│   ├── PDF Vendite/                    Input folder for sales DDT PDFs
│   ├── PDF Acquisti/                   Input folder for purchase invoice PDFs
│   ├── Prova/
│   │   └── prova a vuoto.command       Dry-run execution into Prova/Generati/
│   └── src/
│       ├── riepilogo.py                DDT -> DR.RPD.00 / EST-Riepilogo produzione
│       ├── pesate.py                   DDT -> DR.CPF.70 finished-product weight checks
│       ├── acquisti.py                 FLOTT invoices -> DR.CAM, DR.CAO, DR.CAI + inventory
│       ├── acquisti_esterni.py         External supplier invoices -> EST registers + inventory
│       ├── modello_*.xlsx              Blank Excel templates (FLOTT & External)
│       ├── Modelli pesate/             Blank DR.CPF.70 templates per jar size (FLOTT)
│       └── Modelli pesate esterni/     Blank DR.CPF.70 templates per jar size (External)
└── AutoFrancesco Esterni/
    ├── README.md                       Operational guide (External, Italian)
    ├── genera esterni.command          Double-click launcher (HACCP/ESTERNI archive)
    ├── EST PDF Vendite/                Input folder for external sales DDT PDFs
    ├── EST PDF Acquisti/               Input folder for external purchase invoice PDFs
    └── Prova/
        └── prova a vuoto.command       Dry-run execution into Prova/Generati/
```

---

## Python Modules (`AutoFrancesco/src/`)

- **`riepilogo.py`**: Parses sales DDT PDFs and creates one worksheet per document (`number_ddmmyy`) in the annual production summary workbook, computing jar counts, caps, and fish/oil quantities from `FORMATI E CAPSULE.xlsx`. Supports `--esterni` (prepends the client code to the product name) and `--sovrascrivi` (rebuilds existing sheets).
- **`pesate.py`**: Populates `DR.CPF.70` finished-product weight workbooks per jar size, updating date, batch/lot, and product name. Replaces volatile `RANDBETWEEN` formulas with frozen integer weights sampled using the same bounds, removes sheet protection, and keeps sheets sorted by DDT number. Includes `--sistema` to normalize pre-existing weight workbooks.
- **`acquisti.py`**: Parses FLOTT purchase invoices, classifies each line item into fish (`DR.CAM`), oil (`DR.CAO`), or primary packaging (`DR.CAI`), and appends entries chronologically into `REGISTRO GIACENZE` tagged with unique `#document-position` IDs (leaving the `RIMANENZE` sheet untouched).
- **`acquisti_esterni.py`**: Adapts invoice parsing and classification to heterogeneous external supplier PDF layouts, populating `EST-*` acceptance registers and `EST-REGISTRO GIACENZE`.

All pipelines are **incremental and idempotent**: already-recorded documents are skipped, a `.bak` backup is created in `Backup/` before modifying any workbook, and warnings/anomalies (e.g., tuna items or missing jar templates) are logged once in `Log/`.

---

## Setup & Installation

Requirements: **macOS** (for the `.command` launchers) and **Python 3** (built on Python 3.13).

Create the virtual environment inside `AutoFrancesco/.venv` (shared by `AutoFrancesco Esterni` as well):

```bash
cd AutoFrancesco
python3 -m venv .venv
.venv/bin/pip install -r ../requirements.txt
```

Dependencies (`requirements.txt`):
- `openpyxl==3.1.5` — Excel workbook reading/writing (with rich-text support and XML space preservation)
- `pymupdf==1.28.2` — PDF text extraction for DDTs and invoices

### Local Archive Configuration

The `.command` launchers and `riepilogo.py` expect the local Excel archive at `~/HACCP` by default (not tracked in git). You can override the archive location via environment variables:
- `HACCP_DIR`: root directory of the local HACCP archive (defaults to `$HOME/HACCP`; external registers live under `$HACCP_DIR/ESTERNI`)
- `HACCP_FORMATI`: path to `FORMATI E CAPSULE.xlsx` (defaults to `~/HACCP/FORMATI E CAPSULE.xlsx`)

---

## Usage

1. Place new PDF files into the corresponding input folders (`PDF Vendite` / `PDF Acquisti` for FLOTT, or `EST PDF Vendite` / `EST PDF Acquisti` for external partners).
2. To preview changes without touching the live archive, run **`Prova/prova a vuoto.command`**.
3. To update the live HACCP archive, double-click:
   - **`AutoFrancesco/genera.command`** (FLOTT)
   - **`AutoFrancesco Esterni/genera esterni.command`** (External)

For detailed parsing rules, quarter/year rollover instructions, and CLI flags, see the operational docs in [`AutoFrancesco/README.md`](./AutoFrancesco/README.md) and [`AutoFrancesco Esterni/README.md`](./AutoFrancesco%20Esterni/README.md).

---

## Excluded Data

`.gitignore` excludes business PDFs (DDTs and invoices), log files, `*-visti.txt` tracking files, `.bak` backups, dry-run outputs in `Prova/Generati/`, and `.venv/`.
