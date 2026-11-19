"""
Lit le fichier des statistiques ICC Longueuil (secret SOURCE_URL_STATS_LGL)
et écrit seulement sa STRUCTURE dans stats/structure.txt :
noms des onglets, taille, et textes des premières lignes (les en-têtes).
Les nombres sont remplacés par « # » : aucune donnée n'est publiée.
"""
import io
import os
import sys
from pathlib import Path

import openpyxl

sys.path.insert(0, str(Path(__file__).resolve().parent))
from build_ics import download  # noqa: E402

ROOT = Path(__file__).resolve().parent.parent


def short(v):
    if v is None:
        return ""
    if isinstance(v, (int, float)):
        return "#"
    if hasattr(v, "isoformat"):
        return "<date>"
    return str(v).strip().replace("\n", " ")[:40]


def main():
    url = os.environ.get("SOURCE_URL_STATS_LGL", "").strip()
    if not url:
        sys.exit("Secret SOURCE_URL_STATS_LGL absent.")
    wb = openpyxl.load_workbook(io.BytesIO(download(url)), data_only=True)
    out = []
    for ws in wb.worksheets:
        out.append(f"== Onglet « {ws.title} » : {ws.max_row} lignes x {ws.max_column} colonnes")
        for i, row in enumerate(ws.iter_rows(max_row=6, values_only=True), 1):
            cells = [short(v) for v in row]
            while cells and not cells[-1]:
                cells.pop()
            if cells:
                out.append(f"  L{i}: " + " | ".join(cells))
    text = "\n".join(out) + "\n"
    (ROOT / "stats").mkdir(exist_ok=True)
    (ROOT / "stats" / "structure.txt").write_text(text, encoding="utf-8")
    print(text)


if __name__ == "__main__":
    main()
