import os, csv, re
from bs4 import BeautifulSoup
import html

def clean(s):
    # Nettoie le texte : entités HTML, espaces, etc.
    s = html.unescape(s)  # Convertit &#169; en ©, &nbsp; en espace, etc.
    s = re.sub(r'\s+', ' ', s).strip()  # Remplace les espaces multiples par un seul
    return s

def extract_nuance_from_name(name):
    # Extrait la nuance depuis un nom comme "M. Guy VOCANSON (SOC)"
    match = re.search(r'\(([A-Z]{2,4})\)', name)  # Cherche (SOC), (UMP), etc.
    if match:
        return match.group(1)  # Retourne "SOC", "UMP", etc.
    return "INCONNUE"  # Si pas de nuance trouvée

def parse_tables(html):
    soup = BeautifulSoup(html, "html.parser")
    tables = soup.select("table")
    outputs = []
    for table in tables:
        rows = []
        for tr in table.select("tr"):
            cells = [clean(cell.get_text()) for cell in tr.select("th, td")]
            # Si c'est une ligne de candidat (ex: "M. Guy VOCANSON (SOC)")
            if cells and len(cells) >= 1 and "(" in cells[0] and ")" in cells[0]:
                # Extrait la nuance et le nom
                nuance = extract_nuance_from_name(cells[0])
                name = cells[0].replace(f"({nuance})", "").strip()
                # Reconstruit la ligne avec la nuance en première colonne
                cells = [nuance, name] + cells[1:]
            rows.append(cells)
        if rows:
            outputs.append(rows)
    return outputs

def write_csv(tables, fname):
    with open(fname, "w", newline="", encoding="utf-8") as f:
        w = csv.writer(f, delimiter=';')
        for table in tables:
            for row in table:
                w.writerow(row)

def main():
    html_dir = r"C:\Users\PC\Desktop\HTML_BRUTS"
    output_dir = r"C:\Users\PC\Desktop\SITE_PCF_43\outputs"
    os.makedirs(output_dir, exist_ok=True)

    html_files = [
        "LEG_12_4301.html",
        "LEG_12_4302.html",
        "LEG_12_FR.html",
        "LEG_17_4301.html",
        "LEG_17_4302.html",
        "LEG_17_FR.html",
        "LEG_22_4301.html",
        "LEG_22_4302.html",
        "LEG_22_FR.html",
        "LEG_24_4301.html",
        "LEG_24_4302.html",
        "LEG_24_FR.html",
        "PRES_12_4300.html",
        "PRES_12_FR.html",
        "PRES_17_4300.html",
        "PRES_17_FR.html",
        "PRES_22_4300.html",
        "PRES_22_FR.html"
    ]

    for file in html_files:
        file_path = os.path.join(html_dir, file)
        if os.path.exists(file_path):
            with open(file_path, "r", encoding="utf-8") as f:
                html_content = f.read()
            tables = parse_tables(html_content)
            if tables:
                fname = os.path.join(output_dir, file.replace(".html", ".csv"))
                write_csv(tables, fname)
                print(f"✅ Généré: {fname}")
            else:
                print(f"⚠️ Aucune table trouvée dans {file}")
        else:
            print(f"❌ Fichier {file} introuvable dans {html_dir}")

if __name__ == "__main__":
    main()