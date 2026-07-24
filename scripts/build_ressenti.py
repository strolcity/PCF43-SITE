"""
Import et export du ressenti d'insecurite / de la perception de la delinquance
comme probleme prioritaire, pour alimenter le futur compteur/jauge du site.

Contrairement a delinquance.xlsx (mis a jour automatiquement depuis un flux
ouvert du SSMSI), ressenti.xlsx est lui-meme la source de reference : il n'y
a pas de fichier source unique et stable a re-telecharger chaque annee, donc
on ajoute une ligne a la main dans ressenti.xlsx quand un nouveau bulletin
SSMSI/INSEE sort, puis on relance ce script.

Ce script :
- lit ressenti.xlsx (onglet "ressenti_national")
- l'importe dans une table "ressenti" de securite.db
- exporte un fichier JSON pret pour le site (data/securite/ressenti.json)
"""

import sqlite3
import json
import pandas as pd

FICHIER_EXCEL = "../data/securite/ressenti.xlsx"
FEUILLE_EXCEL = "ressenti_national"
FICHIER_BASE = "../data/securite/securite.db"
FICHIER_JSON = "../data/securite/ressenti.json"

print(f"Lecture de {FICHIER_EXCEL}...")
df = pd.read_excel(FICHIER_EXCEL, sheet_name=FEUILLE_EXCEL)
print(f"{len(df)} lignes lues, indicateurs : {sorted(df['indicateur'].unique())}")

conn = sqlite3.connect(FICHIER_BASE)
df.to_sql("ressenti", conn, if_exists="replace", index=False)
conn.commit()

lignes = df.to_dict(orient="records")
with open(FICHIER_JSON, "w", encoding="utf-8") as f:
    json.dump(lignes, f, ensure_ascii=False)

nb = conn.execute("SELECT COUNT(*) FROM ressenti").fetchone()[0]
print(f"Table 'ressenti' creee dans {FICHIER_BASE} : {nb} lignes.")
print(f"{len(lignes)} lignes ecrites dans {FICHIER_JSON}")

conn.close()
print("Termine.")
