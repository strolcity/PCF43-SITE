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
        SELECT annee, total_prestations_pct_pib, retraites_pct_pib
        FROM vue_protection_sociale_pct_pib ORDER BY annee
    """)
    data["secu_pct_pib"] = [
        {"annee": a, "total": round(t, 2), "retraites": round(r, 2)} for a, t, r in cur.fetchall()
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

    data["extra"] = extra

    conn.close()

    with open(CHEMIN_JSON, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False)

    print(f"Terminé. Fichier écrit : {CHEMIN_JSON}")


if __name__ == "__main__":
    main()
