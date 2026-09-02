#!/usr/bin/env python3
"""
Import du fichier immigration_population.xlsx vers un JSON exploitable
par Chart.js (graphique en courbes 1921-2025).

Usage : python3 scripts/import_population.py data/immigration/immigration_population.xlsx > data/immigration/population.json
"""
import sys
import json
import openpyxl

def main(path):
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb["Population"]
    annees, etrangers, fr_acquisition, fr_naissance = [], [], [], []
    for row in ws.iter_rows(min_row=2, values_only=True):
        annee = row[0]
        if annee is None or not str(annee).strip():
            continue
        # stop at the notes block (annee cell contains a sentence, not a year)
        if not (str(annee)[:4].isdigit()):
            continue
        annees.append(annee)
        etrangers.append(row[1])
        fr_acquisition.append(row[2])
        fr_naissance.append(row[3])

    data = {
        "source": "INSEE, https://www.insee.fr/fr/statistiques/3633212",
        "unite": "milliers",
        "annees": annees,
        "series": [
            {"label": "Étrangers", "data": etrangers, "color": "#ff5555"},
            {"label": "Français par acquisition", "data": fr_acquisition, "color": "#ffaa00"},
            {"label": "Français de naissance", "data": fr_naissance, "color": "#55ff55"},
        ],
    }
    print(json.dumps(data, ensure_ascii=False, indent=2))

if __name__ == "__main__":
    if len(sys.argv) != 2:
        print("Usage: import_population.py <chemin_vers_xlsx>", file=sys.stderr)
        sys.exit(1)
    main(sys.argv[1])
