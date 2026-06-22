import re
import unicodedata
from config import (
    GREEK_TO_LATIN, LOCATION_GR, LATIN_TO_GREEK_DIGRAPHS,
    LATIN_TO_GREEK_SINGLE, GREEK_STRONG_DIGRAPHS, PLACEHOLDER_CLUBS,
    PERFORMANCE_RE, TEXT_FIELDS,
)


def perf_float(p):
    if not p:
        return 999.0
    clean = re.sub(r'[^\d.]', '', p.replace("(", "").split()[0])
    try:
        return float(clean) if clean else 999.0
    except ValueError:
        return 999.0


def normalize_name(name):
    return name.translate(GREEK_TO_LATIN).upper().strip()


def nk_latin(name):
    n = normalize_name(name)
    n = n.replace("OY", "OU")
    parts = n.split()
    return parts[-1] if parts else ""


def norm_full(name):
    n = normalize_name(name)
    n = n.replace("OY", "OU").replace("GG", "NG")
    return n


def latin_to_greek(name):
    result = []
    i = 0
    upper = name.upper()
    while i < len(upper):
        matched = False
        for digraph, greek in LATIN_TO_GREEK_DIGRAPHS:
            if upper[i:i+len(digraph)] == digraph:
                result.append(greek)
                i += len(digraph)
                matched = True
                break
        if matched:
            continue
        ch = upper[i]
        result.append(LATIN_TO_GREEK_SINGLE.get(ch, ch))
        i += 1
    out = "".join(result)
    parts = out.split()
    if parts and len(parts[-1]) > 2 and parts[-1].endswith("ΚΙ"):
        parts[-1] = parts[-1][:-1] + "Η"
        out = " ".join(parts)
    return out


def has_greek_chars(name):
    return any('\u0370' <= c <= '\u03FF' for c in name)


def _no_tonos(s):
    return ''.join(c for c in unicodedata.normalize('NFD', s) if unicodedata.category(c) != 'Mn').upper()


def _club_key(n):
    k = norm_full(n)
    return k, k.replace("V", "Y").replace("B", "Y")


def is_wind_legal(r):
    if 'w' in r['performance']:
        return False
    w = r['wind'].strip().lstrip('+')
    if w.upper() == 'NWI':
        return True
    try:
        return float(w) <= 2.0
    except ValueError:
        return True


def fmt_wind(w):
    w = w.strip()
    if not w or w.upper() == "NWI":
        return "NWI"
    return w.lstrip("+")


def fmt_comp(r):
    return r["competition"]


def fmt_loc(r):
    loc = r["location"]
    if loc.endswith(", GREECE"):
        loc = loc[:-8]
    return loc


def clean_competition_location(r):
    r["competition"] = re.sub(r'^Roster Athletics\s*·\s*', '', r["competition"]).strip()
    r["location"] = re.sub(r'^\d+\s*', '', r["location"]).strip()
    r["location"] = re.sub(r', ([^,]+), \1$', r', \1', r["location"])
    m = re.search(r'& (.+?) \(.*?(?:ΔΡ[ΟΌ]ΜΟΙ|[ΑΆ]ΛΜΑΤΑ).*?\), (.+)', r["location"], re.IGNORECASE)
    if m:
        r["location"] = m.group(1).strip() + ", " + m.group(2).strip()
    loc_city = r["location"].split(",")[0].strip()
    if loc_city:
        comp = r["competition"]
        comp_flat = _no_tonos(comp)
        city_flat = _no_tonos(loc_city)
        idx = comp.find(",")
        if idx >= 0 and city_flat in comp_flat[idx:idx+50]:
            r["competition"] = comp[:idx].strip()


def translate_location(r):
    loc_up = r["location"].upper()
    if loc_up in LOCATION_GR:
        r["location"] = LOCATION_GR[loc_up]


def clean_performance(p):
    if "(" in p or ")" in p:
        m = re.search(PERFORMANCE_RE, p)
        if m:
            clean = m.group(1)
            if "w" in p:
                clean += "w"
            if "h" in p:
                clean += "h"
            return clean
    return p


def uppercase_text_fields(r):
    for k in TEXT_FIELDS:
        v = r.get(k)
        if v and isinstance(v, str):
            r[k] = v.upper()


def uppercase_all(results):
    for r in results:
        uppercase_text_fields(r)


def clear_placeholder_clubs(results):
    for r in results:
        if r.get("club", "").strip() in PLACEHOLDER_CLUBS:
            r["club"] = ""


def clear_lane_placeholders(results):
    for r in results:
        lane = r.get("lane", "").strip()
        if lane in ("", "-", "--"):
            r["lane"] = ""
