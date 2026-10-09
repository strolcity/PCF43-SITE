# -*- coding: utf-8 -*-
"""
build_emploi.py v2 — pipeline de la page Emploi (PCF 43)
========================================================

Principe identique aux autres pages du site :
    data/emploi/emploi.xlsx  -->  scripts/build_emploi.py  -->  data/emploi/emploi_data.json

NOUVEAUTÉS v2 :
  - feuille defm_departements (tous les départements, janvier de chaque année)
    -> permet le menu déroulant « votre département » sur la page
  - feuille departements_noms (liste des départements)
  - feuille defm_france_janvier (France entière, janvier de chaque année)
  - feuille caf_ppa désormais en 4 groupes (Seul avec enfants, etc.) avec
    personnes et montantTotal ; le montant estimé pour la Haute-Loire est
    calculé automatiquement (foyers 43 x montant moyen France du groupe).
  - feuille temps_partiel_quotite (répartition par quotité travaillée)
  - feuille constantes (ARE, RSA...) pour les calculettes de la page.

MODES D'UTILISATION
-------------------
1) PREMIÈRE FOIS (génère le xlsx à partir du JSON déjà collecté) :
       python scripts\build_emploi.py init
   -> crée data\emploi\emploi.xlsx (18 feuilles) à partir de data\emploi\emploi_data.json

2) MISE À JOUR ANNUELLE (votre routine habituelle) :
   - ouvrir data\emploi\emploi.xlsx,
   - ajouter les nouvelles lignes (nouveau trimestre, nouveau mois, etc.),
   - puis :
       python scripts\build_emploi.py
   -> régénère data\emploi\emploi_data.json

Pour defm_departements : ajouter une ligne par département et par année
(colonne annee = 2027, 2028...) à partir des fichiers DARES sur
data.dares.travail-emploi.gouv.fr (inscrits à France Travail par département).

N'hésitez pas à ajouter des lignes de COMMENTAIRES dans les feuilles
(colonne A commençant par "#") : elles sont ignorées par le script.

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

# Liste ordonnée des feuilles du classeur : nom_feuille -> colonnes attendues
SHEETS = {
    "taux_chomage":   ["periode", "france", "hauteLoire"],
    "defm_france":    ["date", "A", "B", "C", "D", "E", "F", "G", "ABC", "ABCDE"],
    "defm_43":        ["date", "A", "B", "C", "D", "E", "F", "G", "ABC", "ABCDE"],
    # Tous les départements, janvier de chaque année (menu déroulant de la page)
    "defm_departements": ["annee", "dep", "A", "B", "C", "D", "E", "ABC"],
    "departements_noms": ["dep", "nom"],
    "defm_france_janvier": ["annee", "A", "B", "C", "D", "E", "ABC"],
    "flux_france":    ["annee", "entrees", "sorties", "radiations",
                       "reprisesEmploi", "cessationsDefautActualisation"],
    "flux_43":        ["date", "entrees", "sorties"],
    "profil_43":      ["date", "categorie", "sexe", "tranche", "nombre"],
    "caf_rsa":        ["geo", "anciennete", "foyers"],
    "caf_rsa_montants_43": ["type", "montant", "foyers"],
    # PPA en 4 groupes : geo = "43" ou "France"
    "caf_ppa":        ["geo", "groupe", "foyers", "personnes", "montantTotal"],
    "caf_ppa_age_43": ["age", "foyers", "personnes"],
    "caf_aah":        ["geo", "sexe", "age", "nombre"],
    "temps_partiel":  ["annee", "nombreMilliers", "partPourcent"],
    "temps_partiel_raisons": ["annee", "raison", "part"],
    "temps_partiel_quotite": ["annee", "quotite", "part"],
    "emplois_vacants": ["trimestre", "emploisVacants", "taux"],
    # Cr\u00e9ations d'entreprises (INSEE) : bloc/cle/valeur
    "creations_entreprises": ["bloc", "cle", "valeur"],
    # Constantes des calculettes (ARE, RSA...)
    "constantes":     ["cle", "valeur", "unite", "description"],
}

GROUPES_PPA = ["Seul avec enfants", "Couple avec enfants",
               "Seul sans enfant", "Couple sans enfant"]

META = {
    "description": ("Données officielles pour la page Emploi — chômage, RSA, "
                    "prime d'activité, AAH, temps partiel, emplois vacants. "
                    "France et Haute-Loire."),
    "version": 2,
    "sources": {
        "tauxChomage": "INSEE - séries 001515908 (Haute-Loire) et 001688527 (France), insee.fr",
        "defm": ("DARES - Inscrits à France Travail, "
                 "data.dares.travail-emploi.gouv.fr (stocks et flux, données brutes)"),
        "caf": ("CAF open data data.caf.fr — RSA, prime d'activité, AAH "
                "(régime général, hors MSA)"),
        "tempsPartiel": "DARES - temps partiel (série, raisons, quotité)",
        "emploisVacants": "DARES - emplois vacants",
    },
    "notes": ("Champ CAF = régime général (hors MSA). DEFM = stocks bruts fin de mois ; "
              "comparaison départementale = mois de janvier de chaque année. "
              "Montants PPA Haute-Loire : estimation = foyers 43 x montant moyen "
              "France du groupe (la CAF ne publie pas ce montant au niveau "
              "départemental par situation familiale). "
              "La hausse 2025 des catégories F/G est en partie mécanique "
              "(inscription automatique RSA / jeunes, loi plein emploi)."),
}

CONSTANTES_INIT = [
    # Assurance chômage (convention d'assurance chômage / règlement UNÉDIC)
    ["are_coef_partie_variable", 0.404, "taux", "ARE : part du SJR (40,4 %)"],
    ["are_partie_fixe", 13.19, "euros/jour", "ARE : partie fixe par jour (2026)"],
    ["are_coef_min", 0.57, "taux", "ARE : 57 % du SJR (minimum garanti)"],
    ["are_plancher_jour", 31.97, "euros/jour", "ARE : allocation journalière minimale (2026)"],
    ["are_plafond_coef", 0.75, "taux", "ARE : allocation <= 75 % du salaire journalier de référence"],
    ["are_jours_min", 130, "jours", "ARE : jours travaillés exigés sur les 24 derniers mois"],
    ["cotis_chomage_salarie_pct", 0.0, "%", "Cotisation chômage des salariés : 0 % depuis 2018"],
    ["cotis_chomage_employeur_pct", 4.05, "%", "Cotisation chômage payée uniquement par l'employeur"],
    ["rsa_socle", 651.69, "euros/mois", "RSA socle pour une personne seule (2026)"],
    ["smic_horaire_brut", 11.88, "euros/h", "SMIC horaire brut (2026) \u2014 \u00e0 ajuster chaque ann\u00e9e"],
    ["ppa_exonerations_patronales_md", 80, "Md€/an", "Exonérations de cotisations patronales (voir page économie)"],
]


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

    # --- Départements (menu déroulant de la page) ---
    dep_data = {}
    for r in read_sheet(wb, "defm_departements"):
        a = txt(r["annee"]); dp = txt(r["dep"])
        if not a or not dp:
            continue
        dep_data.setdefault(dp, {})[a] = {c: num(r.get(c)) for c in ["A", "B", "C", "D", "E", "ABC"]}
    noms = {txt(r["dep"]): txt(r["nom"])
            for r in read_sheet(wb, "departements_noms") if txt(r.get("dep"))}
    fr_janv = {txt(r["annee"]): {c: num(r.get(c)) for c in ["A", "B", "C", "D", "E", "ABC"]}
               for r in read_sheet(wb, "defm_france_janvier") if txt(r.get("annee"))}

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

    # --- PPA en 4 groupes, avec montants ---
    ppa_rows = read_sheet(wb, "caf_ppa")
    ppa_fr = {}
    ppa_43 = {}
    for r in ppa_rows:
        g = txt(r.get("groupe"))
        if g not in GROUPES_PPA:
            continue
        if txt(r.get("geo")) == "France":
            ppa_fr[g] = {"foyers": num(r.get("foyers")), "personnes": num(r.get("personnes")),
                         "montantTotal": num(r.get("montantTotal"))}
        else:
            ppa_43[g] = {"foyers": num(r.get("foyers"))}
    # montant moyen France par groupe + estimation pour le 43
    for g in GROUPES_PPA:
        if g in ppa_fr and ppa_fr[g].get("foyers") and ppa_fr[g].get("montantTotal"):
            moy = int(ppa_fr[g]["montantTotal"] / ppa_fr[g]["foyers"])
            ppa_fr[g]["montantMoyenFoyer"] = moy
            if g in ppa_43 and ppa_43[g].get("foyers"):
                ppa_43[g]["montantMoyenFoyerFrance"] = moy
                ppa_43[g]["montantEstimeMensuel"] = ppa_43[g]["foyers"] * moy
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
    # raisons : on garde la dernière année présente dans la feuille
    rais_rows = read_sheet(wb, "temps_partiel_raisons")
    if rais_rows:
        derniere = max(txt(r["annee"]) for r in rais_rows)
        raisons = {txt(r["raison"]): num(r.get("part"))
                   for r in rais_rows if txt(r.get("annee")) == derniere
                   and txt(r.get("raison")) and txt(r.get("raison")) != "Total"}
        raisons_annee = derniere
    else:
        raisons, raisons_annee = {}, ""
    quot_rows = read_sheet(wb, "temps_partiel_quotite")
    if quot_rows:
        derniere_q = max(txt(r["annee"]) for r in quot_rows)
        quotite = {txt(r["quotite"]): num(r.get("part"))
                   for r in quot_rows if txt(r.get("annee")) == derniere_q
                   and txt(r.get("quotite")) and txt(r.get("quotite")) != "Total"}
        quotite_annee = derniere_q
    else:
        quotite, quotite_annee = {}, ""

    vac = [{"trimestre": txt(r["trimestre"]), "emploisVacants": num(r.get("emploisVacants")),
            "taux": num(r.get("taux"))}
           for r in read_sheet(wb, "emplois_vacants")]

    # --- Cr\u00e9ations d'entreprises (INSEE) : bloc / cle / valeur ---
    cre = {}
    for r in read_sheet(wb, "creations_entreprises"):
        b, c, v = txt(r.get("bloc")), txt(r.get("cle")), num(r.get("valeur"))
        if b and c:
            cre.setdefault(b, {})[c] = v

    const = {}
    for r in read_sheet(wb, "constantes"):
        k = txt(r.get("cle"))
        if k:
            const[k] = {"valeur": num(r.get("valeur")), "unite": txt(r.get("unite")),
                        "description": txt(r.get("description"))}

    data = {
        "_meta": META,
        "constantes": const,
        "tauxChomage": tx,
        "defmCategories": {"france": defm_fr, "hauteLoire": defm_43},
        "defmNomsDepartements": noms,
        "defmDepartements": dep_data,
        "defmFranceJanvier": fr_janv,
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
        "tempsPartiel": {"serieFrance": tps, "raisons": raisons,
                         "raisonsAnnee": raisons_annee, "quotite": quotite,
                         "quotiteAnnee": quotite_annee},
        "emploisVacantsFrance": vac,
        "creationsEntreprises": cre,
    }

    with open(JSON_OUT, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    print(f"OK -> {JSON_OUT}")
    print(f"  taux: {len(tx)} lignes | defm: {len(defm_fr)} FR / {len(defm_43)} HL")
    print(f"  departements: {len(dep_data)} deps | france janvier: {len(fr_janv)} annees")
    print(f"  flux: {len(flux_fr)} annees FR / {len(flux_43)} mois HL | vacants: {len(vac)}")


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
            ws.column_dimensions[get_column_letter(i)].width = 18
        ws.append([])
        ws.append(["# Feuille : " + name])

    new_sheet("taux_chomage", SHEETS["taux_chomage"],
              [[t["periode"], t["france"], t["hauteLoire"]] for t in j["tauxChomage"]])

    cats = ["A", "B", "C", "D", "E", "F", "G", "ABC", "ABCDE"]
    new_sheet("defm_france", SHEETS["defm_france"],
              [[d["date"]] + [d.get(c) for c in cats]
               for d in j["defmCategories"]["france"]])
    new_sheet("defm_43", SHEETS["defm_43"],
              [[d["date"]] + [d.get(c) for c in cats]
               for d in j["defmCategories"]["hauteLoire"]])

    dep_rows = []
    dep_data = j.get("defmDepartements", {})
    for dp, annees in dep_data.items():
        for a in sorted(annees.keys()):
            v = annees[a]
            dep_rows.append([a, dp, v.get("A"), v.get("B"), v.get("C"),
                             v.get("D"), v.get("E"), v.get("ABC")])
    dep_rows.sort(key=lambda r: (str(r[1]).zfill(3), str(r[0])))
    new_sheet("defm_departements", SHEETS["defm_departements"], dep_rows)
    new_sheet("departements_noms", SHEETS["departements_noms"],
              [[k, v] for k, v in sorted(j.get("defmNomsDepartements", {}).items(),
                                          key=lambda kv: kv[0].zfill(3))])
    new_sheet("defm_france_janvier", SHEETS["defm_france_janvier"],
              [[a] + [v.get(c) for c in ["A", "B", "C", "D", "E", "ABC"]]
               for a, v in sorted(j.get("defmFranceJanvier", {}).items())])

    new_sheet("flux_france", SHEETS["flux_france"],
              [[f["annee"], f["entrees"], f["sorties"], f["radiations"],
                f["reprisesEmploi"], f["cessationsDefautActualisation"]]
               for f in j["defmFluxFranceAnnuel"]])
    new_sheet("flux_43", SHEETS["flux_43"],
              [[f["date"], f["entrees"], f["sorties"]] for f in j["defmFluxHauteLoire"]])

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

    pa = j["caf"]["primeActivite"]
    ppa_rows = []
    for g in GROUPES_PPA:
        v = pa["france"].get(g, {})
        ppa_rows.append(["France", g, v.get("foyers"), v.get("personnes"),
                         v.get("montantTotal")])
    for g in GROUPES_PPA:
        v = pa["hauteLoire"].get(g, {})
        ppa_rows.append(["43", g, v.get("foyers"), None, None])
    new_sheet("caf_ppa", SHEETS["caf_ppa"], ppa_rows)
    # la clé parAgeHauteLoire est optionnelle (donn\u00e9es par \u00e2ge PPA 43)
    par_age_ppa = pa.get("parAgeHauteLoire") or {}
    new_sheet("caf_ppa_age_43", SHEETS["caf_ppa_age_43"],
              [[k, v, None] for k, v in par_age_ppa.items()])

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
    tp = j["tempsPartiel"]
    raisons = tp.get("raisons") or tp.get("raisons2025") or {}
    annee_r = tp.get("raisonsAnnee", "") or "2025"
    new_sheet("temps_partiel_raisons", SHEETS["temps_partiel_raisons"],
              [[annee_r, k, v] for k, v in raisons.items()])
    quotite = tp.get("quotite") or tp.get("quotite2025") or {}
    annee_q = tp.get("quotiteAnnee", "") or "2025"
    new_sheet("temps_partiel_quotite", SHEETS["temps_partiel_quotite"],
              [[annee_q, k, v] for k, v in quotite.items()])
    new_sheet("emplois_vacants", SHEETS["emplois_vacants"],
              [[v["trimestre"], v["emploisVacants"], v["taux"]]
               for v in j["emploisVacantsFrance"]])

    ce_rows = []
    ce = j.get("creationsEntreprises") or {}
    for b, paires in ce.items():
        for c, v in paires.items():
            ce_rows.append([b, c, v])
    new_sheet("creations_entreprises", SHEETS["creations_entreprises"], ce_rows)

    new_sheet("constantes", SHEETS["constantes"], CONSTANTES_INIT)

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