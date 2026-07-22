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

# --- 1. Les donnees par departement ---
for row in conn.execute("SELECT code_dep, annee, indicateur, nombre, insee_pop FROM delinquance"):
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
with open(FICHIER_JSON_CATEGORIE, "w", encoding="utf-8") as f:
    json.dump(lignes_categorie, f, ensure_ascii=False)

print(f"{len(lignes_categorie)} lignes ecrites dans {FICHIER_JSON_CATEGORIE}")