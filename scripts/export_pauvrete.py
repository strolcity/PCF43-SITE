"""
Export de pauvrete_niveau_vie.xlsx vers un JSON exploitable par le site.

Ce fichier vient d'une source differente de delinquance.xlsx (INSEE Filosofi,
une seule annee 2021 disponible), on le garde donc separe, dans data/economie/,
pour pouvoir le reutiliser sur d'autres pages (emploi, economie...).
"""

import json
import pandas as pd

FICHIER_EXCEL = "../data/economie/pauvrete_niveau_vie.xlsx"
FEUILLE = "pauvrete_niveau_vie_2021"
FICHIER_JSON = "../data/economie/pauvrete_niveau_vie.json"

df = pd.read_excel(FICHIER_EXCEL, sheet_name=FEUILLE)
lignes = df.to_dict(orient="records")

with open(FICHIER_JSON, "w", encoding="utf-8") as f:
    json.dump(lignes, f, ensure_ascii=False)

print(f"{len(lignes)} departements ecrits dans {FICHIER_JSON}")
