"""
Etape 3 : export de la base SQLite vers un fichier JSON pour securite.html

GitHub Pages ne sait pas executer Python ni interroger SQLite directement :
un site statique a besoin d'un fichier JSON tout pret que le JavaScript peut
charger avec fetch(). Ce script cree ce fichier a partir de securite.db.

Ce qu'il fait :
- reprend toutes les lignes par departement (table delinquance)
- ajoute des lignes "FR" (France entiere) calculees a partir de la vue total_france
- harmonise les codes departement sur 2 caracteres (1 -> 01) pour que les filtres
  du site (qui utilisent "01", "13", etc.) matchent correctement
- ecrit tout ca dans data/securite/delinquance.json
"""

import sqlite3
import json

FICHIER_BASE = "../data/securite/securite.db"
FICHIER_JSON = "../data/securite/delinquance.json"


def formater_code_dep(code):
    """Harmonise le format des codes departement : '1' -> '01', '2A'/'2B' inchanges, '971' inchange."""
    code = str(code)
    if code.isdigit() and len(code) == 1:
        return "0" + code
    return code


conn = sqlite3.connect(FICHIER_BASE)
conn.row_factory = sqlite3.Row

lignes = []

# --- 1. Les donnees par departement (le pseudo-departement FR est gere a part, voir plus bas) ---
for row in conn.execute("SELECT code_dep, annee, indicateur, nombre, insee_pop FROM delinquance WHERE code_dep != 'FR'"):
    lignes.append({
        "Code_departement": formater_code_dep(row["code_dep"]),
        "annee": row["annee"],
        "indicateur": row["indicateur"],
        "nombre": row["nombre"],
        "insee_pop": row["insee_pop"],
    })

# --- 2. Les totaux France (calcules dans la vue total_france) ---
for row in conn.execute("SELECT annee, indicateur, nombre, insee_pop FROM total_france"):
    lignes.append({
        "Code_departement": "FR",
        "annee": row["annee"],
        "indicateur": row["indicateur"],
        "nombre": row["nombre"],
        "insee_pop": row["insee_pop"],
    })

conn.close()

with open(FICHIER_JSON, "w", encoding="utf-8") as f:
    json.dump(lignes, f, ensure_ascii=False)

print(f"{len(lignes)} lignes ecrites dans {FICHIER_JSON}")
print("Dont lignes France :", sum(1 for l in lignes if l["Code_departement"] == "FR"))
print("Codes departement 01 a 09 presents :",
      sorted(set(l["Code_departement"] for l in lignes if l["Code_departement"] in
                  ["01","02","03","04","05","06","07","08","09"])))

# --- 3. Les sous-totaux par categorie (Violences / Vols / Autres) ---
# Un second fichier JSON, plus leger, pour l'affichage par defaut du site
# (3 courbes/barres au lieu de 18)
conn = sqlite3.connect(FICHIER_BASE)
conn.row_factory = sqlite3.Row
lignes_categorie = []

for row in conn.execute("SELECT code_dep, annee, categorie, nombre, insee_pop FROM total_categorie"):
    lignes_categorie.append({
        "Code_departement": formater_code_dep(row["code_dep"]),
        "annee": row["annee"],
        "categorie": row["categorie"],
        "nombre": row["nombre"],
        "insee_pop": row["insee_pop"],
    })

# France par categorie : on recalcule a partir de la vue total_categorie
# (somme sur tous les departements, population recalculee aussi)
france_cat = {}
for row in conn.execute("SELECT annee, categorie, nombre, insee_pop FROM total_categorie"):
    cle = (row["annee"], row["categorie"])
    if cle not in france_cat:
        france_cat[cle] = {"nombre": 0, "insee_pop": 0}
    france_cat[cle]["nombre"] += row["nombre"]

# la population France par annee (une seule fois, pas par categorie)
pop_france = {}
for row in conn.execute("SELECT annee, SUM(insee_pop) AS pop FROM total_delinquance GROUP BY annee"):
    pop_france[row["annee"]] = row["pop"]

for (annee, categorie), valeurs in france_cat.items():
    lignes_categorie.append({
        "Code_departement": "FR",
        "annee": annee,
        "categorie": categorie,
        "nombre": valeurs["nombre"],
        "insee_pop": pop_france[annee],
    })

conn.close()

FICHIER_JSON_CATEGORIE = "../data/securite/delinquance_categories.json"

# On ajoute l'historique France 2008-2015 (estimation reconstruite, voir Methodologie_historique),
# calculee directement depuis les lignes brutes FR (jamais mélangée aux vues officielles 2016-2025)
conn = sqlite3.connect(FICHIER_BASE)
conn.row_factory = sqlite3.Row
historique_categorie = {}
for row in conn.execute("""
    SELECT d.annee, r.categorie, d.nombre, d.insee_pop
    FROM delinquance d
    JOIN indicateurs_ref r ON d.indicateur = r.indicateur
    WHERE d.code_dep = 'FR' AND d.annee < 2016 AND r.compte_dans_total = 'Oui'
"""):
    cle = (row["annee"], row["categorie"])
    if cle not in historique_categorie:
        historique_categorie[cle] = {"nombre": 0, "insee_pop": row["insee_pop"]}
    historique_categorie[cle]["nombre"] += row["nombre"]
conn.close()

for (annee, categorie), valeurs in historique_categorie.items():
    lignes_categorie.append({
        "Code_departement": "FR",
        "annee": annee,
        "categorie": categorie,
        "nombre": valeurs["nombre"],
        "insee_pop": valeurs["insee_pop"],
        "estimation": True
    })

with open(FICHIER_JSON_CATEGORIE, "w", encoding="utf-8") as f:
    json.dump(lignes_categorie, f, ensure_ascii=False)

print(f"{len(lignes_categorie)} lignes ecrites dans {FICHIER_JSON_CATEGORIE}")
print(f"Dont {len(historique_categorie)} lignes d'estimation historique (2008-2015)")

# --- 4. La liste des indicateurs avec leur categorie (pour le detail deplie du site) ---
conn = sqlite3.connect(FICHIER_BASE)
conn.row_factory = sqlite3.Row
ref = []
for row in conn.execute(
    "SELECT indicateur, categorie FROM indicateurs_ref WHERE compte_dans_total = 'Oui' ORDER BY categorie, indicateur"
):
    ref.append({"indicateur": row["indicateur"], "categorie": row["categorie"]})
conn.close()

FICHIER_JSON_REF = "../data/securite/indicateurs_ref.json"
with open(FICHIER_JSON_REF, "w", encoding="utf-8") as f:
    json.dump(ref, f, ensure_ascii=False)

print(f"{len(ref)} indicateurs ecrits dans {FICHIER_JSON_REF}")

# --- 5. Historique Homicides France (seule serie qui remonte avant 2016) ---
conn = sqlite3.connect(FICHIER_BASE)
conn.row_factory = sqlite3.Row
historique = [dict(row) for row in conn.execute(
    "SELECT annee, nombre, insee_pop, taux_100k FROM homicides_historique_france ORDER BY annee"
)]
conn.close()

FICHIER_JSON_HOMICIDES = "../data/securite/homicides_historique.json"
with open(FICHIER_JSON_HOMICIDES, "w", encoding="utf-8") as f:
    json.dump(historique, f, ensure_ascii=False)

print(f"{len(historique)} annees ecrites dans {FICHIER_JSON_HOMICIDES}")