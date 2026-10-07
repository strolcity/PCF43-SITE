# -*- coding: utf-8 -*-
"""
build_emploi.py — pipeline de la page Emploi (PCF 43)
=====================================================

Principe identique aux autres pages du site :
    data/emploi/emploi.xlsx  -->  scripts/build_emploi.py  -->  data/emploi/emploi_data.json

Le classeur emploi.xlsx contient une feuille par bloc de données.
Chaque feuille a UNE ligne d'en-tête ; le script lit tout ce qui suit.

MODES D'UTILISATION
------------------
1) PREMIÈRE FOIS (génère le xlsx initial à partir du JSON déjà collecté) :
       python scripts/build_emploi.py init
   -> crée data/emploi/emploi.xlsx à partir de data/emploi/emploi_data.json

2) MISE À JOUR ANNUELLE (votre routine habituelle) :
   - ouvrir data/emploi/emploi.xlsx,
   - ajouter les nouvelles lignes (nouveau trimestre, nouveau mois, etc.),
   - puis :
       python scripts/build_emploi.py
   -> régénère data/emploi/emploi_data.json

N'hésitez pas à ajouter des lignes de COMMENTAIRES dans les feuilles
(colonne A commençant par "#") : elles sont ignorées par le script et
peuvent servir de notes de sourcing (source, date de récupération...).

Dépendance : openpyxl  (pip install openpyxl)
"""

import json
import os
import sys

try:
    from openpyxl import Workbook, load_workbook
    from openpyxl.utils import get_column_letter
except ImportError:
    sys.exit("Il faut openpyxl :  pip install openpyxl")

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))  # racine du site
XLSX = os.path.join(BASE, "data", "emploi", "emploi.xlsx")
JSON_OUT = os.path.join(BASE, "data", "emploi", "emploi_data.json")

# Liste ordonnée des feuilles du classeur : (nom_feuille, colonnes attendues)
SHEETS = {
    # Taux de chômage INSEE (BIT / localisés) — une ligne par trimestre
    "taux_chomage":   ["periode", "france", "hauteLoire"],
    # Inscrits France Travail (DARES, stocks fin de mois, bruts) — une ligne par mois
    "defm_france":    ["date", "A", "B", "C", "D", "E", "F", "G", "ABC", "ABCDE"],
    "defm_43":        ["date", "A", "B", "C", "D", "E", "F", "G", "ABC", "ABCDE"],
    # Flux annuels France, par motif (DARES) — une ligne par année
    "flux_france":    ["annee", "entrees", "sorties", "radiations",
                       "reprisesEmploi", "cessationsDefautActualisation"],
    # Flux mensuels Haute-Loire, catégories ABC — une ligne par mois
    "flux_43":        ["date", "entrees", "sorties"],
    # Profil des inscrits 43 (ventilation libre) — une ligne par case
    "profil_43":      ["date", "categorie", "sexe", "tranche", "nombre"],
    # CAF — RSA par ancienneté d'inscription (une colonne geo = "43" ou "France")
    "caf_rsa":        ["geo", "anciennete", "foyers"],
    # CAF — RSA 43 par type et tranche de montant
    "caf_rsa_montants_43": ["type", "montant", "foyers"],
    # CAF — Prime d'activité par type et situation familiale
    "caf_ppa":        ["geo", "type", "situation", "foyers", "personnes", "montantTotal"],
    # CAF — Prime d'activité 43 par tranche d'âge
    "caf_ppa_age_43": ["age", "foyers", "personnes"],
    # CAF — AAH (geo = "43" avec sexe+age, ou "France" avec sexe et age="Total")
    "caf_aah":        ["geo", "sexe", "age", "nombre"],
    # DARES — temps partiel, série annuelle France
    "temps_partiel":  ["annee", "nombreMilliers", "partPourcent"],
    # DARES — raisons du temps partiel (une ligne par raison, mettre l'année en colonne "annee")
    "temps_partiel_raisons": ["annee", "raison", "part"],
    # DARES — emplois vacants France (champ "Ensemble")
    "emplois_vacants": ["trimestre", "emploisVacants", "taux"],
}

META = {
    "description": ("Données officielles pour la page Emploi — chômage, RSA, "
                    "prime d'activité, AAH, temps partiel, emplois vacants. "
                    "France et Haute-Loire."),
    "sources": {
        "tauxChomage": "INSEE - séries 001515908 (Haute-Loire) et 001688527 (France), insee.fr",
        "defm": ("DARES - Inscrits à France Travail, "
                 "data.dares.travail-emploi.gouv.fr (stocks et flux, données brutes)"),
        "caf": ("CAF open data data.caf.fr — RSA, prime d'activité, AAH "
                "(régime général, hors MSA)"),
        "tempsPartiel": "DARES - temps partiel (série et raisons)",
        "emploisVacants": "DARES - emplois vacants",
    },
    "notes": ("Champ CAF = régime général (hors MSA). DEFM = stocks bruts fin de mois. "
              "La hausse 2025 des catégories F/G est en partie mécanique "
              "(inscription automatique RSA / jeunes, loi plein emploi)."),
}


def read_sheet(wb, name):
    """Lit une feuille en liste de dicts. Ignore lignes vides et commentaires (#)."""
    if name not in wb.sheetnames:
        return []
    ws = wb[name]
    rows = list(ws.iter_rows(values_only=True))
    if not rows:
        return []
    header = [str(c).strip() if c is not None else "" for c in rows[0]]
    out = []
    for r in rows[1:]:
        if r is None or r[0] is None:
            continue
        first = str(r[0]).strip()
        if first == "" or first.startswith("#"):
            continue
        d = {}
        for i, key in enumerate(header):
            if key == "":
                continue
            v = r[i] if i < len(r) else None
            d[key] = v
        out.append(d)
    return out


def num(v):
    """Convertit en nombre propre (int si possible), null si vide."""
    if v is None or (isinstance(v, str) and v.strip() == ""):
        return None
    if isinstance(v, str):
        v = v.replace(",", ".").replace(" ", "").replace("\u00a0", "")
    try:
        f = float(v)
        return int(f) if f == int(f) else f
    except (ValueError, TypeError):
        return None


def txt(v):
    return "" if v is None else str(v).strip()


# ----------------------------------------------------------------------
# MODE XLSX -> JSON  (mise à jour annuelle)
# ----------------------------------------------------------------------
def xlsx_to_json():
    wb = load_workbook(XLSX, data_only=True)

    tx = [{"periode": txt(r["periode"]), "france": num(r.get("france")),
           "hauteLoire": num(r.get("hauteLoire"))} for r in read_sheet(wb, "taux_chomage")]

    cats = ["A", "B", "C", "D", "E", "F", "G", "ABC", "ABCDE"]
    defm_fr = [{"date": txt(r["date"]), **{c: num(r.get(c)) for c in cats}}
               for r in read_sheet(wb, "defm_france")]
    defm_43 = [{"date": txt(r["date"]), **{c: num(r.get(c)) for c in cats}}
               for r in read_sheet(wb, "defm_43")]

    flux_fr = [{"annee": txt(r["annee"]), "entrees": num(r.get("entrees")),
                "sorties": num(r.get("sorties")), "radiations": num(r.get("radiations")),
                "reprisesEmploi": num(r.get("reprisesEmploi")),
                "cessationsDefautActualisation": num(r.get("cessationsDefautActualisation"))}
               for r in read_sheet(wb, "flux_france")]
    flux_43 = [{"date": txt(r["date"]), "entrees": num(r.get("entrees")),
                "sorties": num(r.get("sorties"))}
               for r in read_sheet(wb, "flux_43")]

    prof = read_sheet(wb, "profil_43")
    par_age = [{"categorie": txt(r["categorie"]), "tranche": txt(r["tranche"]),
                "nombre": num(r.get("nombre"))}
               for r in prof if txt(r.get("sexe")) == "Total" and txt(r.get("tranche")) != "Total"]
    par_sexe = [{"categorie": txt(r["categorie"]), "sexe": txt(r["sexe"]),
                 "nombre": num(r.get("nombre"))}
                 for r in prof if txt(r.get("tranche")) == "Total" and txt(r.get("sexe")) != "Total"]

    rsa_rows = read_sheet(wb, "caf_rsa")
    rsa_43 = {txt(r["anciennete"]): num(r.get("foyers")) for r in rsa_rows if txt(r.get("geo")) == "43"}
    rsa_fr = {txt(r["anciennete"]): num(r.get("foyers")) for r in rsa_rows if txt(r.get("geo")) == "France"}
    rsa_mtt = {f"{txt(r['type'])} ({txt(r['montant'])})": num(r.get("foyers"))
               for r in read_sheet(wb, "caf_rsa_montants_43")}

    ppa_rows = read_sheet(wb, "caf_ppa")
    ppa_43 = {f"{txt(r['type'])} - {txt(r['situation'])}": num(r.get("foyers"))
              for r in ppa_rows if txt(r.get("geo")) == "43"}
    ppa_fr = {f"{txt(r['type'])} - {txt(r['situation'])}":
              {"foyers": num(r.get("foyers")), "personnes": num(r.get("personnes")),
               "montantTotal": num(r.get("montantTotal"))}
              for r in ppa_rows if txt(r.get("geo")) == "France"}
    ppa_age = {txt(r["age"]): num(r.get("foyers")) for r in read_sheet(wb, "caf_ppa_age_43")}

    aah_rows = read_sheet(wb, "caf_aah")
    aah_43 = {f"{txt(r['sexe'])} - {txt(r['age'])}": num(r.get("nombre"))
              for r in aah_rows if txt(r.get("geo")) == "43"}
    fr_h = sum(r.get("nombre") or 0 for r in aah_rows
               if txt(r.get("geo")) == "France" and txt(r.get("sexe")) == "Homme")
    fr_f = sum(r.get("nombre") or 0 for r in aah_rows
               if txt(r.get("geo")) == "France" and txt(r.get("sexe")) == "Femme")
    aah_fr = {"Hommes": fr_h, "Femmes": fr_f, "Total": fr_h + fr_f}

    tps = [{"annee": txt(r["annee"]), "nombreMilliers": num(r.get("nombreMilliers")),
            "partPourcent": num(r.get("partPourcent"))}
           for r in read_sheet(wb, "temps_partiel")]
    raisons = {txt(r["raison"]): num(r.get("part"))
               for r in read_sheet(wb, "temps_partiel_raisons")
               if txt(r.get("raison")) and txt(r.get("raison")) != "Total"}
    vac = [{"trimestre": txt(r["trimestre"]), "emploisVacants": num(r.get("emploisVacants")),
            "taux": num(r.get("taux"))}
           for r in read_sheet(wb, "emplois_vacants")]

    data = {
        "_meta": META,
        "tauxChomage": tx,
        "defmCategories": {"france": defm_fr, "hauteLoire": defm_43},
        "defmFluxFranceAnnuel": flux_fr,
        "defmFluxHauteLoire": flux_43,
        "defmProfilHauteLoire": {"parAge": par_age, "parSexe": par_sexe},
        "caf": {
            "rsa": {"hauteLoire": rsa_43, "france": rsa_fr},
            "rsaMontantsHauteLoire": rsa_mtt,
            "primeActivite": {"hauteLoire": ppa_43, "france": ppa_fr,
                              "parAgeHauteLoire": ppa_age},
            "aah": {"hauteLoire": aah_43, "franceParSexe": aah_fr},
        },
        "tempsPartiel": {"serieFrance": tps, "raisons": raisons},
        "emploisVacantsFrance": vac,
    }

    with open(JSON_OUT, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    print(f"OK -> {JSON_OUT}")
    print(f"  taux: {len(tx)} lignes | defm: {len(defm_fr)} FR / {len(defm_43)} HL")
    print(f"  flux: {len(flux_fr)} années FR / {len(flux_43)} mois HL | vacants: {len(vac)}")


# ----------------------------------------------------------------------
# MODE INIT : JSON -> XLSX  (une seule fois, pour créer le classeur)
# ----------------------------------------------------------------------
def json_to_xlsx():
    with open(JSON_OUT, encoding="utf-8") as f:
        j = json.load(f)

    wb = Workbook()
    wb.remove(wb.active)

    def new_sheet(name, header, rows):
        ws = wb.create_sheet(name)
        ws.append(header)
        for r in rows:
            ws.append(r)
        ws.column_dimensions["A"].width = 14
        for i in range(2, len(header) + 1):
            ws.column_dimensions[get_column_letter(i)].width = 16
        # note de sourcing en fin de feuille
        ws.append([])
        ws.append(["# Feuille : " + name])

    # taux_chomage
    new_sheet("taux_chomage", SHEETS["taux_chomage"],
              [[t["periode"], t["france"], t["hauteLoire"]] for t in j["tauxChomage"]])

    cats = ["A", "B", "C", "D", "E", "F", "G", "ABC", "ABCDE"]
    new_sheet("defm_france", SHEETS["defm_france"],
              [[d["date"]] + [d.get(c) for c in cats]
               for d in j["defmCategories"]["france"]])
    new_sheet("defm_43", SHEETS["defm_43"],
              [[d["date"]] + [d.get(c) for c in cats]
               for d in j["defmCategories"]["hauteLoire"]])

    new_sheet("flux_france", SHEETS["flux_france"],
              [[f["annee"], f["entrees"], f["sorties"], f["radiations"],
                f["reprisesEmploi"], f["cessationsDefautActualisation"]]
               for f in j["defmFluxFranceAnnuel"]])
    new_sheet("flux_43", SHEETS["flux_43"],
              [[f["date"], f["entrees"], f["sorties"]] for f in j["defmFluxHauteLoire"]])

    prof_rows = []
    # reconstruit une table plate date|categorie|sexe|tranche|nombre impossible depuis le JSON
    # agrégé : on met juste une feuille d'accueil expliquant la mise à jour.
    new_sheet("profil_43", SHEETS["profil_43"],
              [["2026-01", p["categorie"], "Total", p["tranche"], p["nombre"]]
               for p in j["defmProfilHauteLoire"]["parAge"]] +
              [["2026-01", p["categorie"], p["sexe"], "Total", p["nombre"]]
               for p in j["defmProfilHauteLoire"]["parSexe"]])

    rsa_rows = [["43", k, v] for k, v in j["caf"]["rsa"]["hauteLoire"].items()] + \
               [["France", k, v] for k, v in j["caf"]["rsa"]["france"].items()]
    new_sheet("caf_rsa", SHEETS["caf_rsa"], rsa_rows)
    new_sheet("caf_rsa_montants_43", SHEETS["caf_rsa_montants_43"],
              [[k.rsplit(" (", 1)[0], k.rsplit(" (", 1)[1].rstrip(")"), v]
               for k, v in j["caf"]["rsaMontantsHauteLoire"].items()])

    ppa_rows = []
    for k, v in j["caf"]["primeActivite"]["hauteLoire"].items():
        typ, sit = k.split(" - ", 1)
        ppa_rows.append(["43", typ, sit, v, None, None])
    for k, v in j["caf"]["primeActivite"]["france"].items():
        typ, sit = k.split(" - ", 1)
        ppa_rows.append(["France", typ, sit, v.get("foyers"), v.get("personnes"),
                         v.get("montantTotal")])
    new_sheet("caf_ppa", SHEETS["caf_ppa"], ppa_rows)
    new_sheet("caf_ppa_age_43", SHEETS["caf_ppa_age_43"],
              [[k, v, None] for k, v in j["caf"]["primeActivite"]["parAgeHauteLoire"].items()])

    aah_rows = []
    for k, v in j["caf"]["aah"]["hauteLoire"].items():
        sexe, age = k.split(" - ", 1)
        aah_rows.append(["43", sexe, age, v])
    aah_fr = j["caf"]["aah"]["franceParSexe"]
    if aah_fr.get("Hommes") is not None:
        aah_rows.append(["France", "Homme", "Total", aah_fr.get("Hommes")])
        aah_rows.append(["France", "Femme", "Total", aah_fr.get("Femmes")])
    new_sheet("caf_aah", SHEETS["caf_aah"], aah_rows)

    new_sheet("temps_partiel", SHEETS["temps_partiel"],
              [[t["annee"], t["nombreMilliers"], t["partPourcent"]]
               for t in j["tempsPartiel"]["serieFrance"]])
    # raisons : reprend la dernière année disponible et la met en "annee"
    raisons = j["tempsPartiel"].get("raisons") or j["tempsPartiel"].get("raisons2025")
    annee_raisons = "2025"
    new_sheet("temps_partiel_raisons", SHEETS["temps_partiel_raisons"],
              [[annee_raisons, k, v] for k, v in raisons.items()])
    new_sheet("emplois_vacants", SHEETS["emplois_vacants"],
              [[v["trimestre"], v["emploisVacants"], v["taux"]]
               for v in j["emploisVacantsFrance"]])

    os.makedirs(os.path.dirname(XLSX), exist_ok=True)
    wb.save(XLSX)
    print(f"OK -> {XLSX} ({len(wb.sheetnames)} feuilles)")
    for s in wb.sheetnames:
        print("   -", s)


if __name__ == "__main__":
    mode = sys.argv[1] if len(sys.argv) > 1 else "build"
    if mode == "init":
        if not os.path.exists(JSON_OUT):
            sys.exit(f"Introuvable : {JSON_OUT}\nMettez d'abord le JSON collecté à cet endroit.")
        json_to_xlsx()
    else:
        if not os.path.exists(XLSX):
            sys.exit(f"Introuvable : {XLSX}\nLancez d'abord :  python {sys.argv[0]} init")
        xlsx_to_json()