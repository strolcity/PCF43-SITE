#!/usr/bin/env python3
"""
Import du fichier immigration_origines.xlsx vers un JSON exploitable
par Chart.js (camembert 4 continents + détail pays pour le pop-up).

Usage :
  python scripts/import_origines.py data/immigration/immigration_origines.xlsx data/immigration/origines.json
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


def read_continents(ws):
    rows = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        continent, effectifs, part = row[0], row[1], row[2]
        if continent is None or continent == "Ensemble":
            continue
        rows.append({"continent": continent, "effectifs": effectifs, "part_pct": part})
    return rows


def read_detail(ws):
    rows = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        continent, pays, effectifs, part = row[0], row[1], row[2], row[3]
        if continent is None:
            continue
        rows.append({"continent": continent, "pays": pays, "effectifs": effectifs, "part_pct": part})
    return rows


def main(path, out_path=None):
    wb = openpyxl.load_workbook(path, data_only=True)
    data = {
        "source": "INSEE, https://www.insee.fr/fr/statistiques/2861345#tableau-figure1_radio1",
        "annee": 2024,
        "champ": "France, flux d'entrées 2024",
        "immigres": {
            "continents": read_continents(wb["Continents_Immigres"]),
            "detail_pays": read_detail(wb["Detail_pays_Immigres"]),
        },
        "etrangers": {
            "continents": read_continents(wb["Continents_Etrangers"]),
            "detail_pays": read_detail(wb["Detail_pays_Etrangers"]),
        },
    }
    write_output(data, out_path)


if __name__ == "__main__":
    if len(sys.argv) not in (2, 3):
        print("Usage: import_origines.py <chemin_vers_xlsx> [chemin_sortie.json]", file=sys.stderr)
        sys.exit(1)
    main(sys.argv[1], sys.argv[2] if len(sys.argv) > 2 else None)
