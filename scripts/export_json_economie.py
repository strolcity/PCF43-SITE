# -*- coding: utf-8 -*-
"""
export_json_economie.py
------------------------
Lit economie.db et écrit economie_data.json (data/economie/), le fichier que
economie.html va charger avec fetch(). Un seul fichier JSON pour l'instant :
le volume de données reste raisonnable (quelques centaines de points).

À lancer APRÈS import_economie.py et import_dette.py.
Rejouable autant de fois que nécessaire.
"""

import sqlite3
import json
from pathlib import Path

RACINE = Path(__file__).resolve().parent.parent
DOSSIER_ECONOMIE = RACINE / "data" / "economie"
CHEMIN_DB = DOSSIER_ECONOMIE / "economie.db"
CHEMIN_JSON = DOSSIER_ECONOMIE / "economie_data.json"

PAYS_COLONNES = ["France", "Allemagne", "Etats-Unis", "Japon", "Royaume-Uni", "Zone_euro", "OCDE_moyenne"]


def main():
    conn = sqlite3.connect(CHEMIN_DB)
    cur = conn.cursor()
    data = {}

    # ---- PIB annuel (somme des 4 trimestres) ----
    cur.execute("""
        SELECT annee, SUM(pib_valeur_mdeur), SUM(pib_volume_mdeur), COUNT(*)
        FROM pib_niveau GROUP BY annee HAVING COUNT(*) = 4 ORDER BY annee
    """)
    data["pib_annuel"] = [
        {"annee": a, "valeur": round(v, 1), "volume": round(vol, 1)}
        for a, v, vol, n in cur.fetchall()
    ]

    # ---- Comparaison internationale (depuis 1990) ----
    cur.execute("""
        SELECT annee, pays, croissance_pct, type FROM pib_international
        WHERE annee >= 1990 ORDER BY annee
    """)
    intl = {}
    for annee, pays, val, typ in cur.fetchall():
        intl.setdefault(pays, []).append({"annee": annee, "v": round(val * 100, 2), "type": typ})
    data["international"] = intl

    # ---- Financement de la Sécu ----
    cur.execute("""
        SELECT annee, total_prestations_pct_pib, retraites_pct_pib, total_prestations_meuros, vieillesse_survie_meuros
        FROM vue_protection_sociale_pct_pib ORDER BY annee
    """)
    data["secu_pct_pib"] = [
        {"annee": a, "total": round(t, 2), "retraites": round(r, 2),
         "total_mdeur": round(tm/1000, 1), "retraites_mdeur": round(rm/1000, 1)}
        for a, t, r, tm, rm in cur.fetchall()
    ]

    cur.execute("""
        SELECT annee, total_meuros, reduction_generale_bas_salaires_meuros, total_exonerations_pct_pib
        FROM vue_exonerations_pct_pib ORDER BY annee
    """)
    data["exonerations"] = [
        {"annee": a, "total_meuros": round(t, 0), "reduction_fillon_meuros": round(r, 0), "pct_pib": round(p, 2)}
        for a, t, r, p in cur.fetchall()
    ]

    # ---- Pauvreté / IFI ----
    cur.execute("""
        SELECT zone, indicateur, valeur FROM pauvrete_filosofi
        WHERE indicateur IN ('Niveau de vie médian (€)', 'Taux de pauvreté (%)')
    """)
    pauvrete = {}
    for zone, ind, val in cur.fetchall():
        pauvrete.setdefault(zone, {})[ind] = val
    data["pauvrete"] = pauvrete

    cur.execute("SELECT annee, nombre_redevables FROM patrimoine_ifi ORDER BY annee")
    data["ifi"] = [{"annee": a, "redevables": r} for a, r in cur.fetchall()]

    # ---- Compteurs dette (section "extra") ----
    extra = {}

    cur.execute("SELECT trimestre, dette_mdeur, dette_pct_pib FROM dette_publique_pib ORDER BY trimestre")
    extra["dette_pib_trim"] = [{"t": t, "md": md, "pct": pct} for t, md, pct in cur.fetchall()]

    cur.execute("SELECT annee, charge_mdeur, type FROM charge_dette_etat ORDER BY annee")
    extra["charge_dette"] = [
        {"annee": a, "v": v, "prevision": ("prévision" in (typ or "").lower() or "projection" in (typ or "").lower())}
        for a, v, typ in cur.fetchall()
    ]

    cur.execute("""
        SELECT annee, deficit_mdeur, statut FROM deficit_secu
        WHERE statut NOT LIKE 'Révisé%' ORDER BY annee
    """)
    extra["deficit_secu"] = [
        {"annee": a, "v": v, "prevision": ("prévision" in (statut or "").lower())}
        for a, v, statut in cur.fetchall()
    ]

    cur.execute("SELECT annee, endettement_mdeur, mesure FROM dette_unedic ORDER BY annee")
    extra["unedic_dette"] = [
        {"annee": a, "v": v, "prevision": ("prévision" in (m or "").lower())} for a, v, m in cur.fetchall()
    ]

    cur.execute("SELECT annee, total_prestations_pct_pib FROM vue_protection_sociale_pct_pib ORDER BY annee")
    extra["depenses_sociales_pct_pib"] = [{"annee": a, "v": round(v, 1)} for a, v in cur.fetchall()]

    cur.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='smic_boeuf'")
    if cur.fetchone():
        cur.execute("""
            SELECT annee, smic_horaire_brut_eur, prix_boeuf_filet_eur, heures_smic_necessaires
            FROM smic_boeuf ORDER BY annee
        """)
        extra["smic_boeuf"] = [
            {"annee": a, "smic": s, "boeuf": b, "heures": round(h, 2)} for a, s, b, h in cur.fetchall()
        ]

    cur.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='salaires_prix_indices'")
    if cur.fetchone():
        cur.execute("""
            SELECT annee, smic_indice, prix_indice, salaire_median_indice, smic_eur_heure, salaire_median_eur_an
            FROM salaires_prix_indices ORDER BY annee
        """)
        extra["salaires_prix"] = [
            {"annee": a, "smic": s, "prix": p, "salaire_median": sm, "smic_eur": se, "salaire_median_eur": sme}
            for a, s, p, sm, se, sme in cur.fetchall()
        ]

    cur.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='partage_capital_travail'")
    if cur.fetchone():
        cur.execute("SELECT annee, taux_marge_pct, part_salariale_pct FROM partage_capital_travail ORDER BY annee")
        extra["partage_capital_travail"] = [
            {"annee": a, "taux_marge": tm, "part_salariale": ps} for a, tm, ps in cur.fetchall()
        ]

    cur.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='dividendes_investissement'")
    if cur.fetchone():
        cur.execute("""
            SELECT annee, dividendes_mdeur, fbcf_mdeur, dividendes_indice, fbcf_indice
            FROM dividendes_investissement ORDER BY annee
        """)
        extra["dividendes_investissement"] = [
            {"annee": a, "dividendes": d, "fbcf": f, "dividendes_indice": di, "fbcf_indice": fi}
            for a, d, f, di, fi in cur.fetchall()
        ]

    cur.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='redistribution_2024'")
    if cur.fetchone():
        cur.execute("""
            SELECT tranche, ordre, niveau_vie_avant_eur, prelevements_eur, prestations_eur, niveau_vie_apres_eur, taux_redistribution_pct
            FROM redistribution_2024 ORDER BY ordre
        """)
        extra["redistribution_2024"] = [
            {"tranche": t, "avant": a, "prelevements": p, "prestations": pr, "apres": ap, "taux": tx}
            for t, o, a, p, pr, ap, tx in cur.fetchall()
        ]

    # ---- 4 branches de la protection sociale, en % du PIB (calculé, pas stocké en dur) ----
    cur.execute("""
        SELECT p.annee, p.sante_meuros, p.famille_meuros, p.emploi_meuros, pa.pib_valeur_mdeur_annuel
        FROM protection_sociale p JOIN vue_pib_annuel pa ON pa.annee = p.annee
        ORDER BY p.annee
    """)
    extra["secu_branches_pct_pib"] = [
        {
            "annee": a,
            "sante": round(100 * (sante / 1000) / pib, 2),
            "famille": round(100 * (famille / 1000) / pib, 2),
            "chomage": round(100 * (emploi / 1000) / pib, 2),
            "sante_mdeur": round(sante/1000, 1),
            "famille_mdeur": round(famille/1000, 1),
            "chomage_mdeur": round(emploi/1000, 1),
        }
        for a, sante, famille, emploi, pib in cur.fetchall()
    ]

    # ---- Pauvreté vs IFI : personnes pauvres (France) vs foyers IFI (France) ----
    extra["pauvrete_vs_ifi"] = {
        "personnes_pauvres": {"annee": 2023, "valeur": 9800000, "unite": "personnes"},
        "foyers_ifi": {"annee": 2025, "valeur": 193500, "unite": "foyers fiscaux"},
    }

    data["extra"] = extra

    conn.close()

    with open(CHEMIN_JSON, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False)

    print(f"Terminé. Fichier écrit : {CHEMIN_JSON}")


if __name__ == "__main__":
    main()