"""
Etape 1 et 2 : import du fichier Excel "delinquance.xlsx" dans une base SQLite,
plus creation des vues pour le total France et le total delinquance.

- Table "delinquance"      : copie fidele des donnees brutes (departement x annee x indicateur)
- Table "indicateurs_ref"  : la liste des indicateurs, avec Oui/Non pour dire s'ils comptent dans le total
                             (pour eviter de compter en double Usage stup AFD + hors AFD)
- Vue "total_france"       : pour chaque annee+indicateur, la somme sur tous les departements
- Vue "total_delinquance"  : pour chaque departement+annee, la somme des indicateurs "Oui" uniquement
- Vue "total_delinquance_france" : la meme chose mais au niveau France entiere
"""

import sqlite3
import pandas as pd

# --- Reglages ---
# Le script est dans scripts/, le fichier Excel est dans data/securite/
FICHIER_EXCEL = "../data/securite/delinquance.xlsx"
FEUILLE_DONNEES = "donnee-dep-data.gouv BRUT"
FEUILLE_REF = "Indicateurs_ref"
FICHIER_BASE = "../data/securite/securite.db"

# --- 1. Lecture des donnees brutes ---
print(f"Lecture de {FICHIER_EXCEL} (onglet {FEUILLE_DONNEES})...")
df = pd.read_excel(FICHIER_EXCEL, sheet_name=FEUILLE_DONNEES)
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

# --- 2. Lecture de la table de reference des indicateurs ---
print(f"Lecture de l'onglet {FEUILLE_REF}...")
df_ref = pd.read_excel(FICHIER_EXCEL, sheet_name=FEUILLE_REF)
df_ref = df_ref.rename(columns={"indicateur": "indicateur", "compte_dans_total": "compte_dans_total"})
nb_oui = (df_ref["compte_dans_total"] == "Oui").sum()
nb_non = (df_ref["compte_dans_total"] == "Non").sum()
print(f"{len(df_ref)} indicateurs references : {nb_oui} comptes dans le total, {nb_non} exclus.")

# --- 3. Ecriture dans la base SQLite ---
conn = sqlite3.connect(FICHIER_BASE)
df.to_sql("delinquance", conn, if_exists="replace", index=False)
df_ref.to_sql("indicateurs_ref", conn, if_exists="replace", index=False)

conn.execute("CREATE INDEX IF NOT EXISTS idx_delinquance_annee_ind ON delinquance(annee, indicateur);")
conn.execute("CREATE INDEX IF NOT EXISTS idx_delinquance_dep ON delinquance(code_dep);")

# --- 4. Creation des vues ---
conn.execute("DROP VIEW IF EXISTS total_france;")
conn.execute("""
    CREATE VIEW total_france AS
    SELECT annee, indicateur, SUM(nombre) AS nombre, SUM(insee_pop) AS insee_pop
    FROM delinquance
    WHERE code_dep != 'FR'
    GROUP BY annee, indicateur
    UNION ALL
    SELECT annee, indicateur, nombre, insee_pop
    FROM delinquance
    WHERE code_dep = 'FR';
""")

conn.execute("DROP VIEW IF EXISTS total_delinquance;")
conn.execute("""
    CREATE VIEW total_delinquance AS
    SELECT d.code_dep, d.annee, SUM(d.nombre) AS nombre, MAX(d.insee_pop) AS insee_pop
    FROM delinquance d
    JOIN indicateurs_ref r ON d.indicateur = r.indicateur
    WHERE r.compte_dans_total = 'Oui' AND d.code_dep != 'FR'
    GROUP BY d.code_dep, d.annee;
""")

conn.execute("DROP VIEW IF EXISTS total_delinquance_france;")
conn.execute("""
    CREATE VIEW total_delinquance_france AS
    SELECT annee, SUM(nombre) AS nombre, SUM(insee_pop) AS insee_pop
    FROM total_delinquance
    GROUP BY annee;
""")

conn.execute("DROP VIEW IF EXISTS total_categorie;")
conn.execute("""
    CREATE VIEW total_categorie AS
    SELECT d.code_dep, d.annee, r.categorie, SUM(d.nombre) AS nombre, MAX(d.insee_pop) AS insee_pop
    FROM delinquance d
    JOIN indicateurs_ref r ON d.indicateur = r.indicateur
    WHERE r.compte_dans_total = 'Oui' AND d.code_dep != 'FR'
    GROUP BY d.code_dep, d.annee, r.categorie;
""")

conn.execute("DROP VIEW IF EXISTS homicides_historique_france;")
conn.execute("""
    CREATE VIEW homicides_historique_france AS
    SELECT annee, nombre, insee_pop, (nombre * 100000.0 / insee_pop) AS taux_100k
    FROM delinquance
    WHERE code_dep = 'FR' AND indicateur = 'Homicides'
    ORDER BY annee;
""")

conn.commit()

# --- 5. Petites verifications ---
nb_lignes = conn.execute("SELECT COUNT(*) FROM delinquance").fetchone()[0]
print(f"Table 'delinquance' creee dans {FICHIER_BASE} : {nb_lignes} lignes.")

test_43 = conn.execute(
    "SELECT annee, nombre FROM total_delinquance WHERE code_dep='43' ORDER BY annee"
).fetchall()
print("Total delinquance Haute-Loire (verif) :", test_43)

test_fr = conn.execute(
    "SELECT annee, nombre FROM total_delinquance_france ORDER BY annee"
).fetchall()
print("Total delinquance France (verif) :", test_fr)

conn.close()
print("Termine.")