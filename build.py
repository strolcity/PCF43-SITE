import sqlite3
import json
import os

DB   = r"C:\Users\PC\Documents\GitHub\SITE_PCF_43\elections.db"
OUT  = r"C:\Users\PC\Documents\GitHub\SITE_PCF_43\data\ventillees.json"

# Crée le dossier data/ si nécessaire
os.makedirs(os.path.dirname(OUT), exist_ok=True)

conn = sqlite3.connect(DB)
conn.row_factory = sqlite3.Row
cur = conn.cursor()

cur.execute("SELECT * FROM Data_Ventillees ORDER BY Année, TypesScrutin, Tour, Zone, Bloc")
rows = [dict(r) for r in cur.fetchall()]

conn.close()

with open(OUT, "w", encoding="utf-8") as f:
    json.dump(rows, f, ensure_ascii=False, indent=2)

print(f"✓ {len(rows)} lignes exportées dans data/ventillees.json")