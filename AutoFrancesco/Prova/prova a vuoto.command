#!/bin/bash
# Prova a vuoto: compila i registri dentro Generati, senza toccare l'archivio HACCP.
# Riproduce la stessa struttura dell'archivio, cosi' si vede il risultato prima di
# riversarlo davvero con genera.command.
cd "$(dirname "$0")" || exit 1      # la cartella Prova
PROGETTO="$(cd .. && pwd)"
CODICE="$PROGETTO/src"
LOG="$CODICE/Log"
BACKUP="$CODICE/Backup/Prova"       # separata da quella dell'archivio vero
PY="$PROGETTO/.venv/bin/python"
HACCP="${HACCP_DIR:-$HOME/HACCP}"
mkdir -p Generati "$LOG" "$BACKUP"

# Al primo avvio prende una copia del registro giacenze dall'archivio: e' solo una
# copia, l'originale non viene mai toccato.
if ! ls Generati/REGISTRO\ GIACENZE\ *.xlsx >/dev/null 2>&1; then
    echo "Primo avvio: copio il registro giacenze dall'archivio."
    cp "$HACCP/"REGISTRO\ GIACENZE\ *.xlsx Generati/ || exit 1
fi

echo "── Vendite ──────────────────────────────"
"$PY" "$CODICE/riepilogo.py" "$PROGETTO/PDF Vendite" \
    -o "Generati/{anno}/DR.RPD.00 - Riepilogo produzione {anno}.xlsx" \
    -l "$LOG/prova-vendite.log" -b "$BACKUP"

echo
echo "── Pesate ───────────────────────────────"
"$PY" "$CODICE/pesate.py" "$PROGETTO/PDF Vendite" \
    -o "Generati/{anno}/PESATE" \
    -l "$LOG/prova-pesate.log" -b "$BACKUP"

echo
echo "── Acquisti ─────────────────────────────"
"$PY" "$CODICE/acquisti.py" "$PROGETTO/PDF Acquisti" \
    -o "Generati/{anno}" \
    -g "Generati/REGISTRO GIACENZE {anno}.xlsx" \
    -l "$LOG/prova-acquisti.log" -b "$BACKUP"

echo
read -r -p "Premi Invio per chiudere."
