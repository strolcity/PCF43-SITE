import pandas as pd
import sqlite3

XLSX = r"C:\Users\PC\Documents\GitHub\SITE_PCF_43\Data_Election.xlsx"
DB   = r"C:\Users\PC\Documents\GitHub\SITE_PCF_43\elections.db"

conn = sqlite3.connect(DB)

# Import feuille Data
df_data = pd.read_excel(XLSX, sheet_name="Data")
df_data["Zone"] = df_data["Zone"].astype(str)
df_data.to_sql("Data", conn, if_exists="replace", index=False)
print(f"✓ Data : {len(df_data)} lignes importées")

# Import feuille Nuances
df_map = pd.read_excel(XLSX, sheet_name="Nuances")

# Garde uniquement les colonnes utiles (exclut la table de référence H/I)
df_map = df_map[["Année", "Nuance", "Brique", "Nb Elu",
                  "Poid Brique dans la Nuance", "Bloc electoral"]]

# Garde uniquement les lignes avec une nuance valide
df_map = df_map[df_map["Nuance"].notna()]

df_map.to_sql("Nuances", conn, if_exists="replace", index=False)
print(f"✓ Nuances : {len(df_map)} lignes importées")

conn.close()
print("✓ Base elections.db créée !")