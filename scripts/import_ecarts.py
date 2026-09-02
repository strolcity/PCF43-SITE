#!/usr/bin/env python3
"""
Import du fichier immigration_ecarts.xlsx vers un JSON exploitable
par Chart.js (écarts de niveau de vie 2019, avant/après transferts).

Usage : python3 scripts/import_ecarts.py data/immigration/immigration_ecarts.xlsx > data/immigration/ecarts.json
"""
import sys
import json
import openpyxl

def main(path):
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb["Ecarts_niveau_vie"]
    labels, avant, apres = [], [], []
    for row in ws.iter_rows(min_row=2, values_only=True):
        menage = row[0]
        avant_val, apres_val = row[1], row[2]
        if menage is None or not isinstance(avant_val, (int, float)) or not isinstance(apres_val, (int, float)):
            continue
        labels.append(menage)
        avant.append(row[1])
        apres.append(row[2])

    data = {
        "source": "INSEE, Enquête Revenus fiscaux et sociaux 2019",
        "annee": 2019,
        "labels": labels,
        "series": [
            {"label": "Écart avant transferts (%)", "data": avant, "backgroundColor": "rgba(255, 85, 85, 0.7)", "borderColor": "#ff5555"},
            {"label": "Écart après transferts (%)", "data": apres, "backgroundColor": "rgba(85, 255, 85, 0.7)", "borderColor": "#55ff55"},
        ],
    }
    print(json.dumps(data, ensure_ascii=False, indent=2))

if __name__ == "__main__":
    if len(sys.argv) != 2:
        print("Usage: import_ecarts.py <chemin_vers_xlsx>", file=sys.stderr)
        sys.exit(1)
    main(sys.argv[1])
