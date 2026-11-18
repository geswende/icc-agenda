"""Assemble docs/icc-*.json en un seul fichier pour le tableau de bord Claude.
Usage : python3 scripts/dashboard_json.py SORTIE.json"""
import json, sys
from pathlib import Path
docs = Path(__file__).resolve().parent.parent / "docs"
rows = []
for f, eg, niv in [("icc-mtl", "Montréal", "grand"), ("icc-lgl", "Longueuil", "grand"),
                   ("icc-mtl-detail", "Montréal", "detail"), ("icc-lgl-detail", "Longueuil", "detail")]:
    p = docs / f"{f}.json"
    if not p.exists():
        continue
    for e in json.loads(p.read_text(encoding="utf-8"))["events"]:
        rows.append({"id": e["uid"].replace("@icc-agenda", ""), "eglise": eg, "niveau": niv,
                     "titre": e["title"], "jour": e["date"], "fin_jour": e["end_date"],
                     "debut": e["start"] or "", "fin": e["end"] or "", "lieu": e.get("location", "")})
out = {"rows": rows, "meta": [{"label": "Source", "value": "https://geswende.github.io/icc-agenda/"}]}
Path(sys.argv[1]).write_text(json.dumps(out, ensure_ascii=False), encoding="utf-8")
print(len(rows), "activités")
