import pandas as pd
import json

input_file = "../data/securite/delinquance.xlsx"
output_file = "../data/securite/delinquance.json"

df = pd.read_excel(input_file)
data = df.to_dict(orient="records")

with open(output_file, "w", encoding="utf-8") as f:
    json.dump(data, f, ensure_ascii=False, indent=4)

print(f"✅ Fichier JSON généré : {output_file}")