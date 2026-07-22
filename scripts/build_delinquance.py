"""
Etape 1 : import du fichier Excel "base_donnees_propre.xlsx" dans une base SQLite.

Ce script ne fait qu'une seule chose : lire le fichier Excel et recreer
la meme table dans un fichier site.db, pour qu'on puisse ensuite faire
des requetes SQL dessus (total France, comparaisons entre departements, etc.)

Rien n'est transforme, rien n'est filtre : c'est une copie fidele.
"""

import sqlite3
import pandas as pd

# --- Reglages ---
# Le script est dans scripts/, le fichier Excel est dans data/securite/
# donc on remonte d'un cran (..) avant de redescendre dans data/securite/
FICHIER_EXCEL = "../data/securite/delinquance.xlsx"
FEUILLE_EXCEL = "donnee-dep-data.gouv BRUT"  # verifie que l'onglet s'appelle bien ainsi dans ton fichier renomme
FICHIER_BASE = "../data/securite/site.db"

# --- 1. Lecture du fichier Excel ---
print(f"Lecture de {FICHIER_EXCEL}...")
df = pd.read_excel(FICHIER_EXCEL, sheet_name=FEUILLE_EXCEL)

# On uniformise les noms de colonnes en minuscules, plus simples a utiliser en SQL
df = df.rename(columns={
    "Code_departement": "code_dep",
    "annee": "annee",
    "indicateur": "indicateur",
    "nombre": "nombre",
    "insee_pop": "insee_pop",
})

print(f"{len(df)} lignes lues, {df['indicateur'].nunique()} indicateurs, "
      f"{df['code_dep'].nunique()} departements, "
      f"annees {df['annee'].min()} a {df['annee'].max()}")

# --- 2. Ecriture dans la base SQLite ---
conn = sqlite3.connect(FICHIER_BASE)
df.to_sql("delinquance", conn, if_exists="replace", index=False)

# Un index sur (annee, indicateur) pour que les requetes du site soient rapides
conn.execute("CREATE INDEX IF NOT EXISTS idx_delinquance_annee_ind ON delinquance(annee, indicateur);")
conn.execute("CREATE INDEX IF NOT EXISTS idx_delinquance_dep ON delinquance(code_dep);")
conn.commit()

# --- 3. Petite verification ---
nb_lignes = conn.execute("SELECT COUNT(*) FROM delinquance").fetchone()[0]
print(f"Table 'delinquance' creee dans {FICHIER_BASE} : {nb_lignes} lignes.")

conn.close()
print("Termine.")