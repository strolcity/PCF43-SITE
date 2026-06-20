import os
import csv
from collections import defaultdict

# Dossier contenant les CSV
output_dir = r"C:\Users\PC\Desktop\SITE_PCF_43\outputs"

# Dictionnaires pour stocker les nuances
nuances_t1 = defaultdict(list)  # {nuance: [fichiers où elle apparaît en T1]}
nuances_t2 = defaultdict(list)  # {nuance: [fichiers où elle apparaît en T2]}
doublons = defaultdict(list)    # {nuance: [fichiers où elle est en doublon]}

# Parcourir tous les fichiers CSV
for filename in os.listdir(output_dir):
    if filename.endswith(".csv"):
        filepath = os.path.join(output_dir, filename)
        with open(filepath, "r", encoding="utf-8") as f:
            reader = csv.reader(f, delimiter=";")
            for row_idx, row in enumerate(reader):
                # Vérifier T1 (colonne A = index 0)
                if row_idx >= 8 and len(row) > 0 and row[0].strip():  # On saute les en-têtes
                    nuance_t1 = row[0].strip()
                    if nuance_t1:  # Si la cellule n'est pas vide
                        nuances_t1[nuance_t1].append(filename)
                # Vérifier T2 (colonne H = index 7)
                if row_idx >= 8 and len(row) > 7 and row[7].strip():
                    nuance_t2 = row[7].strip()
                    if nuance_t2:
                        nuances_t2[nuance_t2].append(filename)

# Trouver les doublons (nuances présentes dans plusieurs fichiers)
for nuance in nuances_t1:
    if len(nuances_t1[nuance]) > 1:
        doublons[nuance] = nuances_t1[nuance]
for nuance in nuances_t2:
    if len(nuances_t2[nuance]) > 1:
        doublons[nuance].extend(nuances_t2[nuance])

# Générer le rapport
with open(os.path.join(output_dir, "rapport_nuances.txt"), "w", encoding="utf-8") as rapport:
    rapport.write("=== RAPPORT DES NUANCES ===\n\n")

    # Toutes les nuances uniques
    rapport.write("1. LISTE DE TOUTES LES NUANCES UNIQUES (T1 et T2) :\n")
    toutes_nuances = set(nuances_t1.keys()).union(set(nuances_t2.keys()))
    for nuance in sorted(toutes_nuances):
        rapport.write(f"- {nuance}\n")

    # Doublons
    rapport.write("\n2. DOUBLONS (nuances présentes dans plusieurs fichiers) :\n")
    if doublons:
        for nuance, fichiers in doublons.items():
            rapport.write(f"- {nuance} : {fichiers}\n")
    else:
        rapport.write("Aucun doublon détecté.\n")

    # Nuances par fichier (T1)
    rapport.write("\n3. NUANCES PAR FICHIER (T1) :\n")
    for filename in sorted(os.listdir(output_dir)):
        if filename.endswith(".csv"):
            rapport.write(f"\n{filename} :\n")
            filepath = os.path.join(output_dir, filename)
            with open(filepath, "r", encoding="utf-8") as f:
                reader = csv.reader(f, delimiter=";")
                for row_idx, row in enumerate(reader):
                    if row_idx >= 7 and len(row) > 0 and row[0].strip():
                        rapport.write(f"  - {row[0].strip()}\n")

    # Nuances par fichier (T2)
    rapport.write("\n4. NUANCES PAR FICHIER (T2) :\n")
    for filename in sorted(os.listdir(output_dir)):
        if filename.endswith(".csv"):
            rapport.write(f"\n{filename} :\n")
            filepath = os.path.join(output_dir, filename)
            with open(filepath, "r", encoding="utf-8") as f:
                reader = csv.reader(f, delimiter=";")
                for row_idx, row in enumerate(reader):
                    if row_idx >= 7 and len(row) > 7 and row[7].strip():
                        rapport.write(f"  - {row[7].strip()}\n")

print("✅ Rapport généré : outputs/rapport_nuances.txt")