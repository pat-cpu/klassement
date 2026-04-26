#!/bin/bash

# Zorg dat het script altijd werkt vanuit zijn eigen map
cd "$(dirname "$0")"

PYTHON_EXEC=python3

# === INVOER ===
read -p "Jaar (bv. 2025): " YEAR
read -p "PDF's genereren? (Ja/Nee): " PDF

YEAR=${YEAR:-2025}
PDF=${PDF:-Ja}

echo "== Gebruik Python: $PYTHON_EXEC =="
echo "Jaar = $YEAR"
echo "PDF  = $PDF"
echo

# === CONTROLE BESTANDEN ===
if [ ! -f scores_ophalen.py ]; then
    echo "[FOUT] scores_ophalen.py ontbreekt!"
    exit 1
fi

if [ ! -f main.py ]; then
    echo "[FOUT] main.py ontbreekt!"
    exit 1
fi

# === CSV's ophalen ===
$PYTHON_EXEC scores_ophalen.py --jaar "$YEAR"
if [ $? -ne 0 ]; then
    echo "[FOUT] scores_ophalen.py faalde."
    exit 1
fi

echo
echo "---- CSV's in data/$YEAR ----"
ls "data/$YEAR"/*.csv 2>/dev/null || echo "Geen CSV's gevonden."

# === LOGBESTAND voorbereiden ===
TIMESTAMP=$(date +%Y%m%d_%H%M)
LOGFILE="output/$YEAR/run_${TIMESTAMP}.log"

echo
echo "== Uitvoer van main.py wordt gelogd naar: $LOGFILE =="
mkdir -p "output/$YEAR"  # zorg dat map bestaat

# === MAIN.PY uitvoeren met logging ===
printf "%s\n%s\n" "$YEAR" "$PDF" | "$PYTHON_EXEC" main.py | tee "$LOGFILE"
if [ ${PIPESTATUS[0]} -ne 0 ]; then
    echo "[FOUT] main.py faalde. Zie $LOGFILE"
    exit 1
fi

echo
echo "✅ Klaar. HTML staat in: output/$YEAR/html/"
[[ "$PDF" =~ ^[Jj]a$ ]] && echo "📄 PDF's staan in: output/$YEAR/pdf/"
