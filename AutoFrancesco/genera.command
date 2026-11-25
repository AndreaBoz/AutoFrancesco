#!/bin/bash
# Doppio clic per compilare i registri HACCP dai PDF di questa cartella.
#   PDF Vendite  -> DR.RPD.00 Riepilogo produzione + DR.CPF.70 pesate
#   PDF Acquisti -> DR.CAM / DR.CAO / DR.CAI + REGISTRO GIACENZE
#
# {anno} viene sostituito con l'anno del documento, non con la data di oggi: un DDT del
# 29 dicembre lavorato a gennaio finisce comunque nella cartella del 2026, e la cartella
# dell'anno nuovo viene creata da sola quando serve.
HACCP="${HACCP_DIR:-$HOME/HACCP}"

cd "$(dirname "$0")" || exit 1
CODICE="src"
LOG="$CODICE/Log"
BACKUP="$CODICE/Backup"
PY=".venv/bin/python"
mkdir -p "$LOG" "$BACKUP"

if [ ! -x "$PY" ]; then
    echo "Non trovo l'ambiente Python in .venv"
    read -r -p "Premi Invio per chiudere."; exit 1
fi
if [ ! -d "$HACCP" ]; then
    echo "Non trovo l'archivio in $HACCP"
    read -r -p "Premi Invio per chiudere."; exit 1
fi

# Il registro giacenze dell'anno in corso sta nella radice di HACCP; quando lo archivi
# nella cartella dell'anno, lo script lo cerca li'.
if ls "$HACCP"/REGISTRO\ GIACENZE\ *.xlsx >/dev/null 2>&1; then
    GIACENZE="$HACCP/REGISTRO GIACENZE {anno}.xlsx"
else
    GIACENZE="$HACCP/{anno}/REGISTRO GIACENZE {anno}.xlsx"
fi

echo "── Vendite ──────────────────────────────"
"$PY" "$CODICE/riepilogo.py" "PDF Vendite" \
    -o "$HACCP/{anno}/DR.RPD.00 - Riepilogo produzione {anno}.xlsx" \
    -l "$LOG/vendite.log" -b "$BACKUP"

echo
echo "── Pesate ───────────────────────────────"
"$PY" "$CODICE/pesate.py" "PDF Vendite" \
    -o "$HACCP/{anno}/PESATE" -l "$LOG/pesate.log" -b "$BACKUP"

echo
echo "── Acquisti ─────────────────────────────"
"$PY" "$CODICE/acquisti.py" "PDF Acquisti" \
    -o "$HACCP/{anno}" -g "$GIACENZE" -l "$LOG/acquisti.log" -b "$BACKUP"

echo
read -r -p "Premi Invio per chiudere."
