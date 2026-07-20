import sqlite3
import json
import os

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DB   = os.path.join(BASE, "data", "elections", "elections.db")
OUT  = os.path.join(BASE, "data", "elections", "ventillees.json")

conn = sqlite3.connect(DB)
conn.row_factory = sqlite3.Row
cur = conn.cursor()

cur.execute("SELECT * FROM Data_Ventillees ORDER BY Année, TypesScrutin, Tour, Zone, Bloc")
rows = [dict(r) for r in cur.fetchall()]
conn.close()

with open(OUT, "w", encoding="utf-8") as f:
    json.dump(rows, f, ensure_ascii=False, indent=2)

print(f"✓ {len(rows)} lignes exportées dans data/elections/ventillees.json")