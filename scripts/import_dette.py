# -*- coding: utf-8 -*-
"""
import_dette.py
----------------
Lit Data_Dette_Compteurs.xlsx (data/economie/) et charge les 4 feuilles de
données dans economie.db (la même base que import_economie.py, tables en plus).

À lancer APRÈS import_economie.py (qui crée la base si elle n'existe pas encore).
Rejouable autant de fois que nécessaire.
"""

import sqlite3
import openpyxl
from pathlib import Path

RACINE = Path(__file__).resolve().parent.parent
DOSSIER_ECONOMIE = RACINE / "data" / "economie"
CHEMIN_DB = DOSSIER_ECONOMIE / "economie.db"
FICHIER_DETTE = DOSSIER_ECONOMIE / "Data_Dette_Compteurs.xlsx"


def lire_lignes(fichier, feuille, ligne_debut=2):
    wb = openpyxl.load_workbook(fichier, data_only=True, read_only=True)
    ws = wb[feuille]
    for row in ws.iter_rows(min_row=ligne_debut, values_only=True):
        if row[0] is None:
            continue
        yield row


def import_dette(cur):
    cur.execute("DROP TABLE IF EXISTS dette_publique_pib")
    cur.execute("""
        CREATE TABLE dette_publique_pib (
            trimestre TEXT PRIMARY KEY,
            dette_mdeur REAL,
            dette_pct_pib REAL
        )
    """)
    for trimestre, md, pct in lire_lignes(FICHIER_DETTE, "Dette_publique_PIB"):
        cur.execute("INSERT INTO dette_publique_pib VALUES (?,?,?)", (trimestre, md, pct))

    cur.execute("DROP TABLE IF EXISTS charge_dette_etat")
    cur.execute("""
        CREATE TABLE charge_dette_etat (
            annee INTEGER PRIMARY KEY,
            charge_mdeur REAL,
            type TEXT
        )
    """)
    for annee, val, typ in lire_lignes(FICHIER_DETTE, "Charge_dette_Etat"):
        if not isinstance(annee, (int, float)):
            continue  # ignore la ligne de note en bas de la feuille
        cur.execute("INSERT INTO charge_dette_etat VALUES (?,?,?)", (int(annee), val, typ))

    cur.execute("DROP TABLE IF EXISTS deficit_secu")
    cur.execute("""
        CREATE TABLE deficit_secu (
            annee INTEGER,
            deficit_mdeur REAL,
            statut TEXT,
            PRIMARY KEY (annee, statut)
        )
    """)
    for annee, val, statut in lire_lignes(FICHIER_DETTE, "Deficit_secu"):
        cur.execute("INSERT INTO deficit_secu VALUES (?,?,?)", (int(annee), val, statut))

    cur.execute("DROP TABLE IF EXISTS dette_unedic")
    cur.execute("""
        CREATE TABLE dette_unedic (
            annee INTEGER PRIMARY KEY,
            endettement_mdeur REAL,
            mesure TEXT
        )
    """)
    for annee, val, mesure in lire_lignes(FICHIER_DETTE, "Dette_Unedic"):
        cur.execute("INSERT INTO dette_unedic VALUES (?,?,?)", (int(annee), val, mesure))


# ----------------------------------------------------------------------
# Ajout : Data_PouvoirAchat.xlsx (Smic / prix du bœuf)
# Regroupé dans ce même script pour éviter un fichier de plus à retenir.
# ----------------------------------------------------------------------
FICHIER_POUVOIR_ACHAT = DOSSIER_ECONOMIE / "Data_PouvoirAchat.xlsx"


def import_pouvoir_achat(cur):
    if not FICHIER_POUVOIR_ACHAT.exists():
        print(f"  (ignoré : {FICHIER_POUVOIR_ACHAT.name} absent)")
        return
    cur.execute("DROP TABLE IF EXISTS smic_boeuf")
    cur.execute("""
        CREATE TABLE smic_boeuf (
            annee INTEGER PRIMARY KEY,
            smic_horaire_brut_eur REAL,
            prix_boeuf_filet_eur REAL,
            heures_smic_necessaires REAL
        )
    """)
    for annee, smic, boeuf, heures in lire_lignes(FICHIER_POUVOIR_ACHAT, "Smic_Boeuf"):
        cur.execute("INSERT INTO smic_boeuf VALUES (?,?,?,?)", (int(annee), smic, boeuf, heures))


# ----------------------------------------------------------------------
# Ajout : Data_Salaires_Prix.xlsx (Smic / salaire médian / indice des prix)
# ----------------------------------------------------------------------
FICHIER_SALAIRES_PRIX = DOSSIER_ECONOMIE / "Data_Salaires_Prix.xlsx"


def import_salaires_prix(cur):
    if not FICHIER_SALAIRES_PRIX.exists():
        print(f"  (ignoré : {FICHIER_SALAIRES_PRIX.name} absent)")
        return
    cur.execute("DROP TABLE IF EXISTS salaires_prix_indices")
    cur.execute("""
        CREATE TABLE salaires_prix_indices (
            annee INTEGER PRIMARY KEY,
            smic_indice REAL,
            prix_indice REAL,
            salaire_median_indice REAL,
            smic_eur_heure REAL,
            salaire_median_eur_an REAL
        )
    """)
    for annee, smic, prix, sal, smic_eur, sal_eur in lire_lignes(FICHIER_SALAIRES_PRIX, "Salaires_Prix_Indices"):
        if not isinstance(annee, (int, float)):
            continue
        cur.execute("INSERT INTO salaires_prix_indices VALUES (?,?,?,?,?,?)", (int(annee), smic, prix, sal, smic_eur, sal_eur))


# ----------------------------------------------------------------------
# Ajout : Data_Capital_Travail.xlsx (taux de marge / dividendes vs FBCF)
# ----------------------------------------------------------------------
FICHIER_CAPITAL_TRAVAIL = DOSSIER_ECONOMIE / "Data_Capital_Travail.xlsx"


def import_capital_travail(cur):
    if not FICHIER_CAPITAL_TRAVAIL.exists():
        print(f"  (ignoré : {FICHIER_CAPITAL_TRAVAIL.name} absent)")
        return
    cur.execute("DROP TABLE IF EXISTS partage_capital_travail")
    cur.execute("""
        CREATE TABLE partage_capital_travail (
            annee INTEGER PRIMARY KEY,
            taux_marge_pct REAL,
            part_salariale_pct REAL
        )
    """)
    for annee, tm, ps in lire_lignes(FICHIER_CAPITAL_TRAVAIL, "Partage_Capital_Travail"):
        cur.execute("INSERT INTO partage_capital_travail VALUES (?,?,?)", (int(annee), tm, ps))

    cur.execute("DROP TABLE IF EXISTS dividendes_investissement")
    cur.execute("""
        CREATE TABLE dividendes_investissement (
            annee INTEGER PRIMARY KEY,
            dividendes_mdeur REAL,
            fbcf_mdeur REAL,
            dividendes_indice REAL,
            fbcf_indice REAL
        )
    """)
    for annee, div, fbcf, di, fi in lire_lignes(FICHIER_CAPITAL_TRAVAIL, "Dividendes_Investissement"):
        cur.execute("INSERT INTO dividendes_investissement VALUES (?,?,?,?,?)", (int(annee), div, fbcf, di, fi))


# ----------------------------------------------------------------------
# Ajout : Data_Redistribution.xlsx (niveau de vie avant/après, par tranche)
# ----------------------------------------------------------------------
FICHIER_REDISTRIBUTION = DOSSIER_ECONOMIE / "Data_Redistribution.xlsx"


def import_redistribution(cur):
    if not FICHIER_REDISTRIBUTION.exists():
        print(f"  (ignoré : {FICHIER_REDISTRIBUTION.name} absent)")
        return
    cur.execute("DROP TABLE IF EXISTS redistribution_2024")
    cur.execute("""
        CREATE TABLE redistribution_2024 (
            tranche TEXT,
            ordre INTEGER PRIMARY KEY,
            niveau_vie_avant_eur REAL,
            prelevements_eur REAL,
            prestations_eur REAL,
            niveau_vie_apres_eur REAL,
            taux_redistribution_pct REAL
        )
    """)
    for tranche, ordre, avant, prel, prest, apres, taux in lire_lignes(FICHIER_REDISTRIBUTION, "Redistribution_2024"):
        if not isinstance(ordre, (int, float)):
            continue  # ignore la ligne de note en bas de la feuille
        cur.execute("INSERT INTO redistribution_2024 VALUES (?,?,?,?,?,?,?)",
                     (tranche, int(ordre), avant, prel, prest, apres, taux))


def main():
    if not CHEMIN_DB.exists():
        raise SystemExit(
            f"{CHEMIN_DB} n'existe pas. Lancez d'abord import_economie.py."
        )
    conn = sqlite3.connect(CHEMIN_DB)
    cur = conn.cursor()
    print("Import compteurs dette...")
    import_dette(cur)
    print("Import pouvoir d'achat (Smic/bœuf)...")
    import_pouvoir_achat(cur)
    print("Import salaires/prix (indices)...")
    import_salaires_prix(cur)
    print("Import capital/travail...")
    import_capital_travail(cur)
    print("Import redistribution 2024...")
    import_redistribution(cur)
    conn.commit()
    conn.close()
    print(f"Terminé. Tables ajoutées à : {CHEMIN_DB}")


if __name__ == "__main__":
    main()