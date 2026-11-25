#!/bin/bash
# Doppio clic per compilare i registri HACCP degli esterni dai PDF di questa cartella.
#   EST PDF Vendite  -> EST-Riepilogo produzione + DR.CPF.70 pesate
#   EST PDF Acquisti -> EST-Controllo MAT.PRIME / OLIO / IMBALLAGGI + EST-REGISTRO GIACENZE
#
# I programmi sono quelli di AutoFrancesco: qui dentro ci sono solo i PDF, i log e le
# copie di sicurezza. {anno} viene sostituito con l'anno del documento, non con la data
# di oggi, cosi' un DDT di fine dicembre lavorato a gennaio finisce nella cartella giusta.
HACCP="${HACCP_DIR:-$HOME/HACCP}/ESTERNI"

cd "$(dirname "$0")" || exit 1
PROGETTO="$(pwd)"
CODICE="$PROGETTO/../AutoFrancesco/src"
PY="$PROGETTO/../AutoFrancesco/.venv/bin/python"
LOG="$PROGETTO/Log"
BACKUP="$PROGETTO/Backup"
mkdir -p "$LOG" "$BACKUP"

if [ ! -x "$PY" ]; then
    echo "Non trovo l'ambiente Python: la cartella AutoFrancesco deve stare qui accanto."
    read -r -p "Premi Invio per chiudere."; exit 1
fi
if [ ! -d "$HACCP" ]; then
    echo "Non trovo l'archivio in $HACCP"
    read -r -p "Premi Invio per chiudere."; exit 1
fi

# Il registro giacenze dell'anno in corso sta nella radice di ESTERNI; quando lo archivi
# nella cartella dell'anno, lo script lo cerca li'.
if ls "$HACCP"/EST-REGISTRO\ GIACENZE\ *.xlsx >/dev/null 2>&1; then
    GIACENZE="$HACCP/EST-REGISTRO GIACENZE {anno}.xlsx"
else
    GIACENZE="$HACCP/{anno}/EST-REGISTRO GIACENZE {anno}.xlsx"
fi

echo "── Vendite ──────────────────────────────"
"$PY" "$CODICE/riepilogo.py" "EST PDF Vendite" --esterni \
    -o "$HACCP/{anno}/EST-Riepilogo produzione {anno}.xlsx" \
    -l "$LOG/vendite.log" -b "$BACKUP"

echo
echo "── Pesate ───────────────────────────────"
"$PY" "$CODICE/pesate.py" "EST PDF Vendite" \
    -o "$HACCP/{anno}/PESATE" -m "$CODICE/Modelli pesate esterni" \
    -l "$LOG/pesate.log" -b "$BACKUP"

echo
echo "── Acquisti ─────────────────────────────"
"$PY" "$CODICE/acquisti_esterni.py" "EST PDF Acquisti" \
    -o "$HACCP/{anno}" -g "$GIACENZE" \
    -l "$LOG/acquisti.log" -b "$BACKUP"

echo
read -r -p "Premi Invio per chiudere."
