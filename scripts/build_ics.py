"""
Télécharge le fichier Excel des événements ICC (lien SharePoint public)
et génère :
  docs/icc.ics   -> flux d'agenda (abonnement Google / Outlook / iPhone)
  docs/icc.json  -> mêmes événements en JSON (pour un tableau de bord)

Variables d'environnement :
  SOURCE_URL  lien public SharePoint du fichier Excel (obligatoire)
  LOCAL_FILE  (optionnel) chemin d'un .xlsx local, pour tester sans réseau

Le script devine les colonnes (Date, Événement, Heure, Lieu...).
Si ça ne marche pas, remplir scripts/config.json (voir README).
"""
import datetime as dt
import hashlib
import io
import json
import os
import re
import sys
import unicodedata
from collections import Counter
import urllib.request
from pathlib import Path

import openpyxl

sys.path.insert(0, str(Path(__file__).resolve().parent))
import parsers  # noqa: E402

ROOT = Path(__file__).resolve().parent.parent
OUT = ROOT / "docs"
CONFIG = Path(__file__).resolve().parent / "config.json"
TZ = "America/Toronto"
CAL_NAME = "Événements ICC"

# Mots-clés pour reconnaître les colonnes (sans accents, en minuscules)
KEYS = {
    "date": ["date", "jour", "day"],
    "end_date": ["date de fin", "date fin"],
    "title": ["evenement", "event", "activite", "titre", "nom", "programme", "intitule"],
    "start": ["heure de debut", "heure debut", "debut", "heure", "start", "horaire"],
    "end": ["heure de fin", "heure fin", "fin", "end"],
    "location": ["lieu", "endroit", "adresse", "salle", "location"],
    "description": ["description", "details", "detail", "note", "commentaire", "responsable", "theme"],
}

MOIS = {
    "janvier": 1, "janv": 1, "fevrier": 2, "fevr": 2, "fev": 2, "mars": 3,
    "avril": 4, "avr": 4, "mai": 5, "juin": 6, "juillet": 7, "juil": 7,
    "aout": 8, "septembre": 9, "sept": 9, "octobre": 10, "oct": 10,
    "novembre": 11, "nov": 11, "decembre": 12, "dec": 12,
}


def norm(s):
    s = unicodedata.normalize("NFD", str(s or "")).encode("ascii", "ignore").decode()
    return re.sub(r"\s+", " ", s).strip().lower()


# ---------- Téléchargement ----------
def candidate_urls(url):
    """Plusieurs façons connues de télécharger un lien de partage SharePoint."""
    base = url.split("?")[0]
    yield f"{base}?download=1"
    # https://hote/:x:/g/personal/USER/TOKEN -> download.aspx?share=TOKEN
    m = re.match(r"(https://[^/]+)/:\w:/\w/(personal/[^/]+|sites/[^/]+)/([^/?]+)", base)
    if m:
        yield f"{m[1]}/{m[2]}/_layouts/15/download.aspx?share={m[3]}"
    yield url


def download(url):
    # Pas d'URL dans les messages : les journaux du dépôt public sont visibles.
    import http.cookiejar
    notes = []
    for i, u in enumerate(candidate_urls(url), 1):
        jar = http.cookiejar.CookieJar()  # SharePoint pose un cookie avant le fichier
        opener = urllib.request.build_opener(urllib.request.HTTPCookieProcessor(jar))
        opener.addheaders = [("User-Agent", "Mozilla/5.0"), ("Accept", "*/*")]
        try:
            with opener.open(u, timeout=60) as r:
                data, ctype = r.read(), r.headers.get("Content-Type", "?")
        except Exception as exc:
            notes.append(f"méthode {i}: {type(exc).__name__} {getattr(exc, 'code', '')}")
            continue
        if data.startswith(b"PK"):
            print(f"  téléchargé (méthode {i}, {len(data)} octets)")
            return data
        title = re.search(rb"<title>(.*?)</title>", data, re.S | re.I)
        notes.append(f"méthode {i}: {ctype}, {len(data)} octets, page "
                     f"'{title[1].decode(errors='ignore').strip()[:60] if title else '?'}'")
    raise RuntimeError("le lien n'a pas renvoyé de fichier Excel ("
                       + " | ".join(notes) + "). Vérifiez que le partage est "
                       "'Toute personne disposant du lien'.")


# ---------- Lecture des valeurs ----------
def parse_date(v):
    if isinstance(v, dt.datetime):
        return v.date()
    if isinstance(v, dt.date):
        return v
    s = norm(v)
    if not s:
        return None
    m = re.search(r"(\d{4})-(\d{1,2})-(\d{1,2})", s)
    if m:
        return dt.date(int(m[1]), int(m[2]), int(m[3]))
    m = re.search(r"(\d{1,2})[/.-](\d{1,2})[/.-](\d{2,4})", s)
    if m:
        y = int(m[3]) + (2000 if len(m[3]) == 2 else 0)
        return dt.date(y, int(m[2]), int(m[1]))  # format québécois jj/mm/aaaa
    m = re.search(r"(\d{1,2})(?:er)?\s+([a-z]+)\.?\s*(\d{4})?", s)
    if m and m[2] in MOIS:
        y = int(m[3]) if m[3] else dt.date.today().year
        return dt.date(y, MOIS[m[2]], int(m[1]))
    return None


def parse_time(v):
    if isinstance(v, dt.datetime):
        return v.time() if (v.hour or v.minute) else None
    if isinstance(v, dt.time):
        return v
    s = norm(v)
    m = re.search(r"(\d{1,2})\s*(?:h|:)\s*(\d{2})?", s)
    if m and int(m[1]) < 24:
        return dt.time(int(m[1]), int(m[2] or 0))
    return None


def parse_time_range(v):
    """'19h-21h' ou '9:00 à 12:00' -> (début, fin)"""
    if not isinstance(v, str):
        return parse_time(v), None
    parts = re.split(r"\s*(?:-|–|à|a|to)\s*", norm(v))
    times = [t for t in (parse_time(p) for p in parts) if t]
    return (times[0] if times else None), (times[1] if len(times) > 1 else None)


# ---------- Détection des colonnes ----------
def find_columns(rows):
    cfg = json.loads(CONFIG.read_text(encoding="utf-8")) if CONFIG.exists() else {}
    for i, row in enumerate(rows[:15]):
        headers = [norm(c) for c in row]
        cols = {}
        for field, words in KEYS.items():
            if field in cfg:
                wanted = norm(cfg[field])
                idx = [j for j, h in enumerate(headers) if h == wanted]
            else:
                idx = [j for j, h in enumerate(headers)
                       if h and j not in cols.values()
                       and any(h == w or h.startswith(w) for w in words)]
            if idx:
                cols[field] = idx[0]
        if "date" in cols and "title" in cols:
            return i, cols
    return None, None


# ---------- Construction des événements ----------
def cell(row, cols, key):
    j = cols.get(key)
    return row[j] if j is not None and j < len(row) else None


def extract_events(wb):
    events = []
    for ws in wb.worksheets:
        rows = [list(r) for r in ws.iter_rows(values_only=True)]
        h, cols = find_columns(rows)
        if cols is None:
            print(f"  feuille '{ws.title}' ignorée (colonnes Date/Événement introuvables)")
            continue
        print(f"  feuille '{ws.title}': colonnes {cols}")
        for row in rows[h + 1:]:
            day = parse_date(cell(row, cols, "date"))
            title = str(cell(row, cols, "title") or "").strip()
            if not day or not title:
                continue
            start, end = parse_time_range(cell(row, cols, "start"))
            if cell(row, cols, "end") is not None:
                end = parse_time(cell(row, cols, "end")) or end
            if start is None:  # heure parfois écrite dans la cellule date
                start = parse_time(cell(row, cols, "date")) if isinstance(
                    cell(row, cols, "date"), dt.datetime) else None
            end_day = parse_date(cell(row, cols, "end_date")) or day
            events.append({
                "title": title,
                "date": day.isoformat(),
                "end_date": end_day.isoformat(),
                "start": start.strftime("%H:%M") if start else None,
                "end": end.strftime("%H:%M") if end else None,
                "location": str(cell(row, cols, "location") or "").strip(),
                "description": str(cell(row, cols, "description") or "").strip(),
                "sheet": ws.title,
            })
    events.sort(key=lambda e: (e["date"], e["start"] or ""))
    for e in events:  # identifiant stable -> pas de doublons dans l'agenda
        key = f"{e['date']}|{norm(e['title'])}"
        e["uid"] = hashlib.sha1(key.encode()).hexdigest()[:16] + "@icc-agenda"
    return events


# ---------- Génération ICS ----------
def esc(s):
    return (s.replace("\\", "\\\\").replace(";", "\\;")
             .replace(",", "\\,").replace("\n", "\\n"))


def fold(line):
    b = line.encode("utf-8")
    out = []
    while len(b) > 75:
        cut = 75 if not out else 74
        while (b[cut] & 0xC0) == 0x80:  # ne pas couper un caractère UTF-8
            cut -= 1
        out.append(b[:cut].decode())
        b = b[cut:]
    out.append(b.decode())
    return "\r\n ".join(out)


VTIMEZONE = """BEGIN:VTIMEZONE
TZID:America/Toronto
BEGIN:DAYLIGHT
TZOFFSETFROM:-0500
TZOFFSETTO:-0400
TZNAME:EDT
DTSTART:19700308T020000
RRULE:FREQ=YEARLY;BYMONTH=3;BYDAY=2SU
END:DAYLIGHT
BEGIN:STANDARD
TZOFFSETFROM:-0400
TZOFFSETTO:-0500
TZNAME:EST
DTSTART:19701101T020000
RRULE:FREQ=YEARLY;BYMONTH=11;BYDAY=1SU
END:STANDARD
END:VTIMEZONE"""


def to_ics(events):
    now = dt.datetime.now(dt.timezone.utc).strftime("%Y%m%dT%H%M%SZ")
    L = ["BEGIN:VCALENDAR", "VERSION:2.0", "PRODID:-//ICC//Agenda//FR",
         "CALSCALE:GREGORIAN", "METHOD:PUBLISH", f"X-WR-CALNAME:{CAL_NAME}",
         f"X-WR-TIMEZONE:{TZ}", "REFRESH-INTERVAL;VALUE=DURATION:PT6H",
         "X-PUBLISHED-TTL:PT6H"]
    L += VTIMEZONE.split("\n")
    for e in events:
        d = dt.date.fromisoformat(e["date"])
        d2 = dt.date.fromisoformat(e["end_date"])
        L += ["BEGIN:VEVENT", f"UID:{e['uid']}", f"DTSTAMP:{now}",
              f"SUMMARY:{esc(e['title'])}"]
        if e["start"]:
            s = dt.datetime.combine(d, dt.time.fromisoformat(e["start"]))
            if e["end"]:
                f = dt.datetime.combine(d2, dt.time.fromisoformat(e["end"]))
                if f <= s:
                    f = s + dt.timedelta(hours=2)
            else:
                f = s + dt.timedelta(hours=2)  # durée par défaut
            L += [f"DTSTART;TZID={TZ}:{s:%Y%m%dT%H%M%S}",
                  f"DTEND;TZID={TZ}:{f:%Y%m%dT%H%M%S}"]
        else:  # journée entière
            L += [f"DTSTART;VALUE=DATE:{d:%Y%m%d}",
                  f"DTEND;VALUE=DATE:{d2 + dt.timedelta(days=1):%Y%m%d}"]
        if e["location"]:
            L.append(f"LOCATION:{esc(e['location'])}")
        if e["description"]:
            L.append(f"DESCRIPTION:{esc(e['description'])}")
        L.append("END:VEVENT")
    L.append("END:VCALENDAR")
    return "\r\n".join(fold(x) for x in L) + "\r\n"


NAMES = {"MTL": "Montréal", "LGL": "Longueuil"}
# Quels agendas produire pour chaque église :
#   "liste"  = onglet « Grands évènements »
#   "grille" = onglets mensuels (toutes les activités)
MODES = {
    "MTL": {"": "liste", "-detail": "grille"},
    "LGL": {"": "liste", "-detail": "grille"},
}
SUFFIX_LABEL = {"": "", "-detail": " (détaillé)"}


def build(wb, mode):
    if mode == "liste":
        year = parsers.workbook_year(wb)
        sheets = [ws for ws in wb.worksheets if parsers.is_list(ws)]
        events = [e for ws in sheets for e in parsers.parse_list(ws, year)]
    else:
        sheets = [ws for ws in wb.worksheets if parsers.is_grid(ws)]
        events = [e for ws in sheets for e in parsers.parse_grid(ws)]
    print(f"  mode {mode}: {len(sheets)} onglet(s) -> {len(events)} événements")
    if not events:  # dernier recours : tableau classique Date / Événement
        events = extract_events(wb)
    events.sort(key=lambda e: (e["date"], e["start"] or ""))
    seen = Counter()
    for e in events:  # identifiant stable -> pas de doublons dans l'agenda
        key = f"{e['date']}|{e['start']}|{norm(e['title'])}"
        seen[key] += 1
        e["uid"] = (hashlib.sha1(f"{key}|{seen[key]}".encode()).hexdigest()[:16]
                    + "@icc-agenda")
    return events


def sources():
    """Chaque variable SOURCE_URL_XXX (ou LOCAL_FILE_XXX pour tester)
    devient une source nommée XXX."""
    found = {}
    for k, v in sorted(os.environ.items()):
        for prefix in ("SOURCE_URL_", "LOCAL_FILE_"):
            if k.startswith(prefix) and v.strip():
                found[k[len(prefix):].upper()] = (prefix, v.strip())
    return found


def write(name, label, events):
    global CAL_NAME
    CAL_NAME = label
    (OUT / f"{name}.ics").write_text(to_ics(events), encoding="utf-8", newline="")
    (OUT / f"{name}.json").write_text(json.dumps(
        {"name": label, "updated": dt.datetime.now(dt.timezone.utc).isoformat(),
         "events": events}, ensure_ascii=False, indent=1), encoding="utf-8")
    print(f"  -> docs/{name}.ics ({len(events)} événements)")


def main():
    srcs = sources()
    if not srcs:
        sys.exit("Aucun secret SOURCE_URL_XXX trouvé.")
    OUT.mkdir(exist_ok=True)
    combined, errors = [], []
    for code, (prefix, value) in srcs.items():
        base = f"ICC {NAMES.get(code, code)}"
        print(f"== {base}")
        try:
            data = (Path(value).read_bytes() if prefix == "LOCAL_FILE_"
                    else download(value))
            if os.environ.get("RAW_DIR"):  # copie brute pour diagnostic
                Path(os.environ["RAW_DIR"]).mkdir(exist_ok=True)
                (Path(os.environ["RAW_DIR"]) / f"{code}.xlsx").write_bytes(data)
            wb = openpyxl.load_workbook(io.BytesIO(data), data_only=True)
        except Exception as exc:  # une source en panne n'arrête pas l'autre
            errors.append(f"{base}: {exc}")
            print(f"  ERREUR: {exc}")
            continue
        for suffix, mode in MODES.get(code, {"": "grille"}).items():
            label = base + SUFFIX_LABEL.get(suffix, suffix)
            events = build(wb, mode)
            if not events:
                errors.append(f"{label}: aucun événement")
                continue
            for e in events:
                e["source"] = code
                e["uid"] = f"{code.lower()}{suffix}-{e['uid']}"
            write(f"icc-{code.lower()}{suffix}", label, events)
            if suffix == "":
                combined += [{**e, "title": f"[{code}] {e['title']}"} for e in events]
    if combined:
        combined.sort(key=lambda e: (e["date"], e["start"] or ""))
        write("icc", "ICC (grands événements)", combined)
    if errors:
        sys.exit("Problèmes :\n" + "\n".join(errors))


if __name__ == "__main__":
    main()
