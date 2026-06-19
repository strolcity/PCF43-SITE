import os, csv
from bs4 import BeautifulSoup

def clean(s):
    return ' '.join(s.replace('\xa0', ' ').strip().split())

def parse_tables(html):
    soup = BeautifulSoup(html, "html.parser")
    tables = soup.select("table")
    outputs = []
    for table in tables:
        rows = []
        for tr in table.select("tr"):
            cells = [clean(cell.get_text()) for cell in tr.select("th, td")]
            if cells:
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
    html_dir = "pages"
    output_dir = "outputs"
    os.makedirs(output_dir, exist_ok=True)

    html_files = [
        "LEG_24_4301.html",
        "LEG_24_4302.html",
        "PRES_22_4300.html"
    ]

    for file in html_files:
        file_path = os.path.join(html_dir, file)
        if os.path.exists(file_path):
            with open(file_path, "r", encoding="utf-8") as f:
                html = f.read()
            tables = parse_tables(html)
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