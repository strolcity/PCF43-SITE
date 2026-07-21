import pandas as pd
import sqlite3

XLSX = r"C:\Users\PC\Documents\GitHub\SITE_PCF_43\data\elections\Data_Election.xlsx"
DB   = r"C:\Users\PC\Documents\GitHub\SITE_PCF_43\data\elections\elections.db"

conn = sqlite3.connect(DB)

# Import feuille Data
df_data = pd.read_excel(XLSX, sheet_name="Data")
df_data["Zone"] = df_data["Zone"].astype(str)
df_data.to_sql("Data", conn, if_exists="replace", index=False)
print(f"✓ Data : {len(df_data)} lignes importées")

# Import feuille Nuances
df_map = pd.read_excel(XLSX, sheet_name="Nuances")
df_map = df_map[["Année", "Nuance", "Brique", "Nb Elu", "Poids", "Bloc electoral"]]
df_map = df_map[df_map["Nuance"].notna()]
df_map["Année"] = df_map["Année"].fillna(0).astype(int)
df_map.to_sql("Nuances", conn, if_exists="replace", index=False)
print(f"✓ Nuances : {len(df_map)} lignes importées")

conn.close()
print("✓ Base elections.db créée !")