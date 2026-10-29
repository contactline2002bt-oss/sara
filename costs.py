"""Költségszámítás a usage_log tábla soraiból (Streamlittől független, ezért tesztelhető).

Az árak dollárban, millió tokenenként (platform.claude.com, 2026. október). Ha az Anthropic árat módosít,
csak a PRICES táblát kell átírni.
"""
from datetime import datetime, timedelta, timezone

# modell-előtag: (bemenet, kimenet, gyorsítótár-olvasás szorzója)
PRICES = {
    "claude-fable-5-1": (10.0, 50.0, 0.025),
    "claude-mythos-5-1": (10.0, 50.0, 0.025),
    "claude-opus-5-5": (4.0, 20.0, 0.05),
    "claude-sonnet-5-5": (2.0, 10.0, 0.05),
    "claude-haiku-5-5": (0.10, 0.50, 0.10),
}
HAIKU_LONG = (0.50, 2.50, 0.10)      # Haiku 5.5, ha a kérés 100 ezer tokennél hosszabb
DEFAULT_PRICE = (2.0, 10.0, 0.10)    # ismeretlen modellnél a Sonnet árával számolunk
CACHE_WRITE = 2.0                    # egy órás tárolás: az alap bemeneti ár kétszerese
WEB_SEARCH = 10.0 / 1000             # dollár / keresés

KIND_NAMES = {"chat": "Beszélgetés", "kivonat": "Dokumentumkivonat", "osszefoglalo": "Összefoglaló",
              "kereso": "Kereső kolléga", "kereso_web": "Kereső kolléga (web)", "ocr": "Szövegfelismerés",
              "ellenorzo": "Ellenőrző kolléga",
              "redflag": "Red Flag", "diff": "Verzió-összehasonlítás"}


def price_for(model, prompt_tokens):
    for prefix, price in PRICES.items():
        if str(model or "").startswith(prefix):
            if prefix == "claude-haiku-5-5" and prompt_tokens > 100_000:
                return HAIKU_LONG
            return price
    return DEFAULT_PRICE


def row_cost(row):
    """Egy hívás: (tényleges költség, költség gyorsítótár nélkül) dollárban."""
    inp = int(row.get("input_tokens") or 0)
    write = int(row.get("cache_write_tokens") or 0)
    read = int(row.get("cache_read_tokens") or 0)
    out = int(row.get("output_tokens") or 0)
    web = int(row.get("web_searches") or 0) * WEB_SEARCH
    p_in, p_out, read_mult = price_for(row.get("model"), inp + write + read)
    actual = (inp * p_in + write * p_in * CACHE_WRITE + read * p_in * read_mult + out * p_out) / 1e6 + web
    without = ((inp + write + read) * p_in + out * p_out) / 1e6 + web
    return actual, without


def period_start(key, now=None):
    now = now or datetime.now(timezone.utc)
    if key == "today":
        local = now.astimezone()
        return local.replace(hour=0, minute=0, second=0, microsecond=0).astimezone(timezone.utc)
    if key == "7d":
        return now - timedelta(days=7)
    if key == "month":
        local = now.astimezone()
        return local.replace(day=1, hour=0, minute=0, second=0, microsecond=0).astimezone(timezone.utc)
    return None


def summarize(rows, project_names):
    """Összesítés a nézethez. Visszaad: dict összesítőkkel és táblázatsorokkal."""
    total = {"actual": 0.0, "without": 0.0, "calls": 0, "tokens_in": 0, "cache_read": 0, "cache_write": 0,
             "tokens_out": 0, "web": 0}
    by_project, by_kind, by_day, calls = {}, {}, {}, []
    for r in rows:
        actual, without = row_cost(r)
        tin = int(r.get("input_tokens") or 0) + int(r.get("cache_write_tokens") or 0) + int(r.get("cache_read_tokens") or 0)
        total["actual"] += actual
        total["without"] += without
        total["calls"] += 1
        total["tokens_in"] += tin
        total["cache_read"] += int(r.get("cache_read_tokens") or 0)
        total["cache_write"] += int(r.get("cache_write_tokens") or 0)
        total["tokens_out"] += int(r.get("output_tokens") or 0)
        total["web"] += int(r.get("web_searches") or 0)
        pname = project_names.get(r.get("project_id"), "(projekt nélkül)" if not r.get("project_id") else "(törölt projekt)")
        kind = KIND_NAMES.get(r.get("kind"), r.get("kind") or "egyéb")
        day = str(r.get("created_at") or "")[:10]
        for bucket, key in ((by_project, pname), (by_kind, kind)):
            b = bucket.setdefault(key, {"actual": 0.0, "without": 0.0, "calls": 0, "tokens": 0})
            b["actual"] += actual
            b["without"] += without
            b["calls"] += 1
            b["tokens"] += tin + int(r.get("output_tokens") or 0)
        by_day[day] = by_day.get(day, 0.0) + actual
        calls.append({"time": r.get("created_at"), "project": pname, "kind": kind, "model": r.get("model"),
                      "tokens_in": tin, "tokens_out": int(r.get("output_tokens") or 0), "cost": actual})
    total["saved"] = total["without"] - total["actual"]
    total["hit_rate"] = total["cache_read"] / total["tokens_in"] if total["tokens_in"] else 0.0
    return {"total": total, "by_project": by_project, "by_kind": by_kind, "by_day": by_day,
            "top_calls": sorted(calls, key=lambda c: c["cost"], reverse=True)[:10]}
