#!/usr/bin/env python3
"""
Import du fichier immigration_ecarts.xlsx vers un JSON exploitable
par Chart.js (écarts de niveau de vie 2019, avant/après transferts).

Usage :
  python scripts/import_ecarts.py data/immigration/immigration_ecarts.xlsx data/immigration/ecarts.json
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
    ws = wb["Ecarts_niveau_vie"]
    labels, avant, apres = [], [], []
    for row in ws.iter_rows(min_row=2, values_only=True):
        menage = row[0]
        avant_val, apres_val = row[1], row[2]
        if menage is None or not isinstance(avant_val, (int, float)) or not isinstance(apres_val, (int, float)):
            continue
        labels.append(menage)
        avant.append(avant_val)
        apres.append(apres_val)

    data = {
        "source": "INSEE, Enquête Revenus fiscaux et sociaux 2019",
        "annee": 2019,
        "labels": labels,
        "series": [
            {"label": "Écart avant transferts (%)", "data": avant, "backgroundColor": "rgba(255, 85, 85, 0.7)", "borderColor": "#ff5555"},
            {"label": "Écart après transferts (%)", "data": apres, "backgroundColor": "rgba(85, 255, 85, 0.7)", "borderColor": "#55ff55"},
        ],
    }
    write_output(data, out_path)


if __name__ == "__main__":
    if len(sys.argv) not in (2, 3):
        print("Usage: import_ecarts.py <chemin_vers_xlsx> [chemin_sortie.json]", file=sys.stderr)
        sys.exit(1)
    main(sys.argv[1], sys.argv[2] if len(sys.argv) > 2 else None)
