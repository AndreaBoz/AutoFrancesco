#!/bin/bash
# Prova a vuoto: compila i registri degli esterni dentro Generati, senza toccare
# l'archivio HACCP. Serve a vedere il risultato prima di riversarlo davvero.
cd "$(dirname "$0")" || exit 1      # la cartella Prova
PROGETTO="$(cd .. && pwd)"
CODICE="$PROGETTO/../AutoFrancesco/src"
PY="$PROGETTO/../AutoFrancesco/.venv/bin/python"
LOG="$PROGETTO/Log"
BACKUP="$PROGETTO/Backup/Prova"     # separata da quella dell'archivio vero
HACCP="${HACCP_DIR:-$HOME/HACCP}/ESTERNI"
mkdir -p Generati "$LOG" "$BACKUP"

# Al primo avvio prende una copia del registro giacenze e delle pesate dall'archivio:
# sono copie, gli originali non vengono mai toccati.
if ! ls Generati/EST-REGISTRO\ GIACENZE\ *.xlsx >/dev/null 2>&1; then
    echo "Primo avvio: copio giacenze e pesate dall'archivio."
    cp "$HACCP/"EST-REGISTRO\ GIACENZE\ *.xlsx Generati/ || exit 1
    mkdir -p "Generati/$(date +%Y)"
    cp -R "$HACCP/$(date +%Y)/PESATE" "Generati/$(date +%Y)/" 2>/dev/null
fi

echo "── Vendite ──────────────────────────────"
"$PY" "$CODICE/riepilogo.py" "$PROGETTO/EST PDF Vendite" --esterni \
    -o "Generati/{anno}/EST-Riepilogo produzione {anno}.xlsx" \
    -l "$LOG/prova-vendite.log" -b "$BACKUP"

echo
echo "── Pesate ───────────────────────────────"
"$PY" "$CODICE/pesate.py" "$PROGETTO/EST PDF Vendite" \
    -o "Generati/{anno}/PESATE" -m "$CODICE/Modelli pesate esterni" \
    -l "$LOG/prova-pesate.log" -b "$BACKUP"

echo
echo "── Acquisti ─────────────────────────────"
"$PY" "$CODICE/acquisti_esterni.py" "$PROGETTO/EST PDF Acquisti" \
    -o "Generati/{anno}" \
    -g "Generati/EST-REGISTRO GIACENZE {anno}.xlsx" \
    -l "$LOG/prova-acquisti.log" -b "$BACKUP"

echo
read -r -p "Premi Invio per chiudere."
