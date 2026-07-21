import pandas as pd
import json

# Chemin vers ton fichier XLSX
input_file = "../data/securite/delinquance.xlsx"
# Chemin vers le fichier JSON de sortie
output_file = "../data/securite/delinquance.json"

# Lire le XLSX
df = pd.read_excel(input_file)

# Convertir en JSON (format "records" = liste de dictionnaires)
data = df.to_dict(orient="records")

# Sauvegarder en JSON
with open(output_file, "w", encoding="utf-8") as f:
    json.dump(data, f, ensure_ascii=False, indent=4)

print(f"Fichier JSON généré : {output_file}")