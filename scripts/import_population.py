#!/usr/bin/env python3
"""
Import du fichier immigration_population.xlsx vers un JSON exploitable
par Chart.js (graphique en courbes 1921-2025).

Usage :
  python scripts/import_population.py data/immigration/immigration_population.xlsx data/immigration/population.json
  (le 2e argument, le chemin de sortie, est optionnel : sans lui, le JSON
  part sur la sortie standard - mais en PowerShell, rediriger avec '>'
  encode en UTF-16 et casse le JSON. Toujours donner le 2e argument.)
"""
import sys
import json
import openpyxl


def write_output(data, out_path=None):
    """Ecrit le JSON en UTF-8 sans BOM, sur disque si un chemin est fourni,
    sinon sur stdout. Evite le piege PowerShell qui encode '>' en UTF-16."""
    text = json.dumps(data, ensure_ascii=False, indent=2)
    if out_path:
        with open(out_path, "w", encoding="utf-8", newline="\n") as f:
            f.write(text)
    else:
        sys.stdout.buffer.write(text.encode("utf-8"))
        sys.stdout.write("\n")


def main(path, out_path=None):
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb["Population"]
    annees, etrangers, fr_acquisition, fr_naissance = [], [], [], []
    for row in ws.iter_rows(min_row=2, values_only=True):
        annee = row[0]
        if annee is None or not str(annee).strip():
            continue
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
    write_output(data, out_path)


if __name__ == "__main__":
    if len(sys.argv) not in (2, 3):
        print("Usage: import_population.py <chemin_vers_xlsx> [chemin_sortie.json]", file=sys.stderr)
        sys.exit(1)
    main(sys.argv[1], sys.argv[2] if len(sys.argv) > 2 else None)
