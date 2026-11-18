"""
Lecteurs adaptés aux calendriers ICC :
  - grille : un onglet par mois, des cellules-dates, les activités en dessous
  - liste  : onglet « Grands évènements » (mois en colonne A, texte en B/C)
"""
import datetime as dt
import re
import unicodedata
from collections import Counter

MOIS = {
    "janvier": 1, "janv": 1, "jan": 1, "january": 1,
    "fevrier": 2, "fevr": 2, "fev": 2, "february": 2,
    "mars": 3, "march": 3, "avril": 4, "avr": 4, "april": 4,
    "mai": 5, "may": 5, "juin": 6, "june": 6, "juillet": 7, "juil": 7,
    "july": 7, "aout": 8, "august": 8, "septembre": 9, "sept": 9,
    "september": 9, "octobre": 10, "oct": 10, "october": 10,
    "novembre": 11, "nov": 11, "november": 11, "decembre": 12, "dec": 12,
    "december": 12,
}
STOP = ("notes",)


def norm(s):
    s = unicodedata.normalize("NFD", str(s or "")).encode("ascii", "ignore").decode()
    return re.sub(r"\s+", " ", s).strip().lower()


def clean(s):
    s = str(s).replace("**", "").replace("\n", " ")
    return re.sub(r"\s+", " ", s).strip(" -–")


def as_date(v):
    if isinstance(v, dt.datetime):
        return v.date()
    if isinstance(v, dt.date):
        return v
    return None


# ---------- heures écrites dans le texte : (21h à 3h), (14h-17h), 20h 30 ----------
H = r"(\d{1,2})\s*h\s*(\d{2})?"
RANGE = re.compile(H + r"\s*(?:à|a|-|–|au)\s*" + H, re.I)
SINGLE = re.compile(r"(?<!\d)" + H + r"(?!\s*\d)", re.I)


def times_in(text):
    m = RANGE.search(text)
    if m and int(m[1]) < 24 and int(m[3]) < 25:
        return (dt.time(int(m[1]), int(m[2] or 0)),
                dt.time(int(m[3]) % 24, int(m[4] or 0)))
    m = SINGLE.search(text)
    if m and int(m[1]) < 24:
        return dt.time(int(m[1]), int(m[2] or 0)), None
    return None, None


def location_in(text):
    m = re.search(r"(Salles?\s+[A-ZÀ-Ýa-zà-ÿ0-9]+(?:\s+et\s+[A-ZÀ-Ý][a-zà-ÿ]+)?|Zoom|Hors CDV)",
                  text, re.I)
    return m[1] if m else ""


def make_event(title, day, end_day=None, sheet=""):
    title = clean(title)
    start, end = times_in(title)
    if start:  # l'heure va dans l'agenda, pas besoin de la garder dans le titre
        title = clean(re.sub(r"\(\s*\)", "", RANGE.sub("", title)))
    end_day = end_day or day
    if start and end and end <= start:
        end_day = day + dt.timedelta(days=1)  # finit après minuit
    return {
        "title": title, "date": day.isoformat(), "end_date": end_day.isoformat(),
        "start": start.strftime("%H:%M") if start else None,
        "end": end.strftime("%H:%M") if end else None,
        "location": location_in(title), "description": "", "sheet": sheet,
    }


# ---------- GRILLE ----------
def is_grid(ws):
    n = 0
    for row in ws.iter_rows(max_row=60, values_only=True):
        n += sum(1 for v in row if as_date(v))
        if n >= 20:
            return True
    return False


def grid_month(ws):
    """Mois de l'onglet : d'après son nom (Janvier, January...), sinon le mois
    le plus fréquent avant la zone « Notes » (les mini-calendriers trompent)."""
    c = Counter()
    for row in ws.iter_rows(max_row=80, values_only=True):
        if any(isinstance(v, str) and norm(v).startswith(STOP) for v in row):
            break
        for v in row:
            d = as_date(v)
            if d:
                c[(d.year, d.month)] += 1
    if not c:
        return None
    by_title = MOIS.get(norm(ws.title).split(" ")[0])
    if by_title:
        years = Counter({y: n for (y, m), n in c.items() if m == by_title})
        if years:
            return years.most_common(1)[0][0], by_title
    return c.most_common(1)[0][0]


def parse_grid(ws):
    ym = grid_month(ws)
    if not ym:
        return []
    current = {}          # colonne -> date en cours
    per_day = {}          # (date, colonne) -> [textes]
    for row in ws.iter_rows():
        cells = [c for c in row if c.value not in (None, "")]
        if any(norm(c.value).startswith(STOP) for c in cells
               if isinstance(c.value, str)):
            break
        dates = {c.column: as_date(c.value) for c in cells if as_date(c.value)}
        if dates:
            low = min(dates.values())
            for col in list(current):
                if col not in dates and current[col] < low - dt.timedelta(days=1):
                    del current[col]  # nouvelle semaine : on oublie l'ancienne date
            current.update(dates)
            continue
        for c in cells:
            if not isinstance(c.value, str) or c.column not in current:
                continue
            text = clean(c.value)
            if not text or text.startswith("..") or len(text) < 3:
                continue
            day = current[c.column]
            if (day.year, day.month) != ym:
                continue  # jours du mois voisin : ils sont dans l'autre onglet
            per_day.setdefault((day, c.column), []).append(text)
    events = []
    for (day, _), texts in sorted(per_day.items()):
        for t in texts:
            events.append(make_event(t, day, sheet=ws.title))
    return merge_consecutive(events)


def merge_consecutive(events):
    """'3 Jours de J&P' les 14, 15 et 16 -> un seul événement du 14 au 16."""
    out, last = [], {}
    for e in sorted(events, key=lambda e: (norm(e["title"]), e["date"])):
        k = norm(e["title"])
        prev = last.get(k)
        if (prev and not e["start"] and not prev["start"] and not e["location"]
                and dt.date.fromisoformat(e["date"])
                == dt.date.fromisoformat(prev["end_date"]) + dt.timedelta(days=1)):
            prev["end_date"] = e["date"]
            continue
        out.append(e)
        last[k] = e
    return sorted(out, key=lambda e: (e["date"], e["start"] or ""))


# ---------- LISTE « Grands évènements » ----------
def month_label(v):
    w = norm(v).split(" ")[0] if v else ""
    return MOIS.get(w)


def parse_spec(spec, month, year):
    """'16' | '14 au 16' | '27 au 1er mars' | 'du 29 janv. au 7 fév.' | '18-21'"""
    toks = re.findall(r"(\d{1,2})(?:er)?\s*([a-zA-Zéèûôà\.]+)?", spec)
    parts = []
    for d, m in toks:
        mm = MOIS.get(norm(m).strip(".")) if m else None
        parts.append((int(d), mm))
    if not parts:
        return None, None
    d1, m1 = parts[0]
    m1 = m1 or month
    d2, m2 = parts[-1] if len(parts) > 1 else (d1, m1)
    m2 = m2 or m1
    if (m2, d2) < (m1, d1):
        m2 = m1 % 12 + 1
    try:
        start = dt.date(year, m1, d1)
        end = dt.date(year + (1 if m2 < m1 else 0), m2, d2)
    except ValueError:
        return None, None
    return start, end


def is_list(ws):
    return "grand" in norm(ws.title)


def parse_list(ws, year):
    events, month, prev_blank, prev_day = [], None, False, 0
    for row in ws.iter_rows(values_only=True):
        vals = list(row) + [None] * 6
        a, b, c = vals[0], vals[1], vals[2]
        lab = month_label(a)
        texts = [v for v in (b, c) if isinstance(v, (str, int, float)) and str(v).strip()]
        if lab:
            month, prev_day = lab, 0
        elif texts and prev_blank and month:
            month = month % 12 + 1  # bloc sans nom de mois après une ligne vide
        prev_blank = not texts and not lab
        if not month or not texts:
            continue
        if c not in (None, "") and b not in (None, ""):
            # format Longueuil : B = date, C = événement
            s, e = parse_spec(str(b), month, year)
            if s and s.day <= prev_day - 15:
                # ligne placée avant l'étiquette du mois suivant (ex. 26 avril puis 3)
                month = month % 12 + 1
                s, e = parse_spec(str(b), month, year)
            if s:
                prev_day = s.day
                events.append(make_event(str(c), s, e, ws.title))
            continue
        text = str(b if b not in (None, "") else c)
        # format Montréal : 'Titre (16)', plusieurs par cellule possible
        found = list(re.finditer(r"([^()]+?)\s*\(([^)]*\d[^)]*)\)", text))
        for m in found:
            title = m[1].strip(" ,;")
            s, e = parse_spec(m[2], month, year)
            if s and title:
                ev = make_event(title, s, e, ws.title)
                ev["start"] = ev["end"] = None  # date seule dans la liste
                events.append(ev)
        if found:
            rest = clean(text[found[-1].end():])
            if rest and events:
                events[-1]["title"] += f" ({rest})"
    return sorted(events, key=lambda e: e["date"])


def workbook_year(wb):
    c = Counter()
    for ws in wb.worksheets:
        if is_grid(ws):
            ym = grid_month(ws)
            if ym:
                c[ym[0]] += 1
    return c.most_common(1)[0][0] if c else dt.date.today().year
