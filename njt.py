"""Hatályos jogszabályszöveg eseti lekérése a Nemzeti Jogszabálytárból (njt.hu) – úgy, ahogy egy felhasználó a
böngészőben megnyit egy jogszabályt: egy oldal, és ha a keresett szakasz a később betöltődő részben van, annak
blokkjai. Nem tölt le és nem tárol adatbázist; a lekért oldalt egy napig gyorsítótárban tartja, hogy ugyanazt ne
kérje le újra. Csak a Wikiből használjuk, ügyadat nem kerül bele.

A szerkezet felismerése a nyílt forrású magyar-jogszabaly MCP-szerver (Apache 2.0) njt.hu-feldolgozóját követi:
- szakaszjel: <span class="szakasz-jel">6:272. §</span>
- blokkhatár: <span class="jhId" id="…"></span>
- később betöltődő blokkok: class="pH borderStart"data-show-order="N" → POST /ajax/njtGetBlock.json
"""
import html as htmllib
import json
import re
import time
import urllib.request
from datetime import date, datetime


def _today():
    """A mai nap magyar idő szerint (a szerver UTC-ben fut)."""
    try:
        from zoneinfo import ZoneInfo
        return datetime.now(ZoneInfo("Europe/Budapest")).date()
    except Exception:
        return date.today()

HOSTS = ["https://njt.jog.gov.hu", "https://njt.hu"]   # az új hivatalos cím és a régi
BASE = HOSTS[-1] + "/jogszabaly/"   # a linkekben
USER_AGENT = ("Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) "
              "Chrome/124.0 Safari/537.36")
MIN_DELAY = 1.2
MAX_SECTION_CHARS = 8000

TOOL = {
    "name": "njt_jogszabaly",
    "description": (
        "Hatályos magyar jogszabályszöveg lekérése a Nemzeti Jogszabálytárból (njt.hu). Megadható a jogszabály "
        "rövid neve (Ptk., Ctv., Pp., Itv., Inytv., Üttv., Vht., Mt., Art., Áfa tv., Ákr., Szjt., Infotv., Pmt., "
        "Btk., Csődtv., Fgytv.) vagy a teljes megjelölése (pl. „2013. évi V. törvény”, „338/2011. (XII. 29.) Korm. "
        "rendelet”). Paragrafusszámmal (pl. „6:272”, „50”, „12/A”) a szakasz teljes szövegét adja; kulcsszóval "
        "megkeresi, mely szakaszokban szerepel. A kérésbe csak jogszabályt, paragrafust és jogi fogalmat írj, "
        "soha nevet vagy ügyadatot."),
    "input_schema": {"type": "object", "properties": {
        "jogszabaly": {"type": "string"},
        "paragrafus": {"type": "string", "description": "pl. 6:272, 50, 12/A (elhagyható, ha kulcsszóval keresel)"},
        "kulcsszo": {"type": "string", "description": "jogi fogalom a jogszabályon belüli kereséshez"}},
        "required": ["jogszabaly"]},
}

# Tartalék: ha a mi szerverünkről az njt nem érhető el, az Anthropic szerveroldali letöltője nyitja meg az oldalt
# (csak az njt címein; a címet az njt_jogszabaly hibaüzenete adja meg, így a letöltő elfogadja)
WEB_FETCH = {"type": "web_fetch_20250910", "name": "web_fetch", "max_uses": 3,
             "allowed_domains": ["njt.jog.gov.hu", "njt.hu"], "max_content_tokens": 60000}

SHORT = {   # rövid név → njt azonosító
    "ptk": "2013-5-00-00", "polgári törvénykönyv": "2013-5-00-00",
    "ctv": "2006-5-00-00", "cégtörvény": "2006-5-00-00",
    "gt": "2006-4-00-00",
    "pp": "2016-130-00-00", "polgári perrendtartás": "2016-130-00-00",
    "itv": "1990-93-00-00", "illetéktörvény": "1990-93-00-00",
    "inytv": "2021-100-00-00", "ingatlan-nyilvántartási törvény": "2021-100-00-00",
    "üttv": "2017-78-00-00", "ügyvédi tevékenységről szóló törvény": "2017-78-00-00",
    "vht": "1994-53-00-00", "mt": "2012-1-00-00", "munka törvénykönyve": "2012-1-00-00",
    "art": "2017-150-00-00", "áfa tv": "2007-127-00-00", "áfa": "2007-127-00-00",
    "ákr": "2016-150-00-00", "szjt": "1999-76-00-00", "infotv": "2011-112-00-00",
    "pmt": "2017-53-00-00", "btk": "2012-100-00-00", "csődtv": "1991-49-00-00",
    "fgytv": "1997-155-00-00", "tftv": "2013-122-00-00", "földforgalmi törvény": "2013-122-00-00",
    "étv": "1997-78-00-00", "társasházi törvény": "2003-133-00-00", "tht": "2003-133-00-00",
}

ROMAN = {"I": 1, "V": 5, "X": 10, "L": 50, "C": 100, "D": 500, "M": 1000}


class NjtError(Exception):
    pass


def roman_to_int(s):
    s = s.upper()
    total = 0
    for i, ch in enumerate(s):
        v = ROMAN[ch]
        total += -v if i + 1 < len(s) and ROMAN[s[i + 1]] > v else v
    return total


def document_id(citation):
    """A jogszabály njt-azonosítója a rövid névből vagy a teljes megjelölésből."""
    c = " ".join(str(citation or "").split())
    key = re.sub(r"^(a|az)\s+", "", c.lower()).rstrip(".").strip()
    if key in SHORT:
        return SHORT[key]
    for name, ident in SHORT.items():   # „a Ptk.-ban”, „Ptk. 6:272. §” stb.
        if re.match(rf"^{re.escape(name)}\b", key):
            return ident
    m = re.search(r"(\d{4})\.\s*évi\s+([IVXLCDM]+)\.?\s*törvény", c, re.I)
    if m:
        return f"{m.group(1)}-{roman_to_int(m.group(2))}-00-00"
    m = re.search(r"(\d+)/(\d{4})\.\s*\([^)]*\)\s*Korm\.\s*rendelet", c, re.I)
    if m:
        return f"{m.group(2)}-{m.group(1)}-20-22"
    m = re.fullmatch(r"(\d{4}-\d+-\d{2}-\d{2})", c)
    if m:
        return m.group(1)
    raise NjtError(f"Nem ismerem fel a jogszabályt („{citation}”). Add meg a teljes megjelölését, pl. "
                   "„2013. évi V. törvény”.")


_last = [0.0]
_cache = {}


def _wait():
    delta = time.time() - _last[0]
    if delta < MIN_DELAY:
        time.sleep(MIN_DELAY - delta)
    _last[0] = time.time()


def _request(url, payload=None, opener=None):
    _wait()
    data = json.dumps(payload).encode("utf-8") if payload is not None else None
    req = urllib.request.Request(url, data=data, headers={
        "User-Agent": USER_AGENT, "Accept": "text/html,application/json,*/*", "Accept-Language": "hu-HU,hu;q=0.9",
        **({"Content-Type": "application/json; charset=utf-8"} if data else {})})
    with (opener or urllib.request.urlopen)(req, timeout=30) as resp:
        return resp.read().decode("utf-8", errors="replace")


def html_to_text(fragment):
    t = re.sub(r'<sup[^>]*class="fnSup"[^>]*>[\s\S]*?</sup>', "", fragment, flags=re.I)
    t = re.sub(r"<br\s*/?>", "\n", t, flags=re.I)
    t = re.sub(r"</?(?:p|div|li|ul|ol|tr|td|th|table|tbody|h[1-6])[^>]*>", "\n", t, flags=re.I)
    t = htmllib.unescape(re.sub(r"<[^>]+>", "", t))
    t = re.sub(r"[ \t\xa0]+", " ", t)
    return re.sub(r"\n\s*\n+", "\n", t).strip()


def _norm_section(s):
    return re.sub(r"[\s.§]", "", str(s or "")).upper()


def _sections(page):
    """[(szakaszjel normalizálva, kezdőpozíció)] a szakaszjelek sorrendjében."""
    return [(_norm_section(html_to_text(m.group(1))), m.start())
            for m in re.finditer(r'<span class="szakasz-jel">([\s\S]*?)</span>', page, re.I)]


def _block_body(raw):
    """A blokk-végpont válasza HTML vagy JSON-ba csomagolt HTML lehet."""
    try:
        data = json.loads(raw)
    except ValueError:
        return raw
    parts = []
    def walk(x):
        if isinstance(x, str):
            parts.append(x)
        elif isinstance(x, dict):
            for v in x.values():
                walk(v)
        elif isinstance(x, list):
            for v in x:
                walk(v)
    walk(data)
    return "\n".join(parts)


def load(doc_id, want=None, opener=None):
    """A jogszabály oldala; ha a keresett szakasz (want) a később betöltődő részben van, a blokkokat is
    betölti, de csak addig, amíg megtalálja. Egy napig gyorsítótárban tartja."""
    key = (doc_id, _today())
    page = _cache.get(key)
    if page is None:
        errors = []
        for host in HOSTS:
            try:
                page = _request(f"{host}/jogszabaly/{doc_id}", opener=opener)
            except Exception as e:   # kapcsolódási hiba: a másik címmel próbáljuk
                errors.append(f"{host}: {e}")
                continue
            if "jogszabalyMainTitle" in page or "szakasz-jel" in page:
                _cache[(doc_id, "host")] = host
                break
            errors.append(f"{host}: nem jogszabályszöveg jött vissza")
            page = None
        if page is None:
            links = " vagy ".join(f"{h}/jogszabaly/{doc_id}" for h in HOSTS)
            raise NjtError("A Nemzeti Jogszabálytár most nem érhető el a szerverünkről (" + "; ".join(errors)
                           + f"). Próbáld a web_fetch eszközzel ezt a címet: {links}")
        _cache[key] = page
    starts = sorted(int(n) for n in re.findall(r'class="pH borderStart"\s*data-show-order="(\d+)"', page))
    done = _cache.setdefault((doc_id, _today(), "blocks"), set())
    pending = [s for s in starts if s not in done]
    if pending and (want is None or want not in {s for s, _ in _sections(page)}):
        ranges = [{"start": s, **({"last": starts[i + 1]} if i + 1 < len(starts) else {})}
                  for i, s in enumerate(starts) if s in pending]
        for i in range(0, len(ranges), 20):
            chunk = ranges[i:i + 20]
            host = _cache.get((doc_id, "host"), HOSTS[0])
            page += "\n" + _block_body(_request(f"{host}/ajax/njtGetBlock.json",
                                                 {"documentId": doc_id, "data": chunk}, opener))
            done.update(r["start"] for r in chunk)
            _cache[key] = page
            if want and want in {s for s, _ in _sections(page)}:
                break
    return page


def title_of(page):
    m = re.search(r'<h1[^>]*class="[^"]*jogszabalyMainTitle[^"]*"[^>]*>([\s\S]*?)</h1>', page, re.I)
    sub = re.search(r'<h2[^>]*class="[^"]*jogszabalySubtitle[^"]*"[^>]*>([\s\S]*?)</h2>', page, re.I)
    return " ".join(html_to_text(x.group(1)) for x in (m, sub) if x).strip()


def section_text(page, section):
    marks = _sections(page)
    want = _norm_section(section)
    for i, (s, pos) in enumerate(marks):
        if s == want:
            end = marks[i + 1][1] if i + 1 < len(marks) else len(page)
            # a szakaszjel előtti blokkhatártól, hogy a jel is benne legyen
            start = page.rfind('<span class="jhId"', 0, pos)
            return html_to_text(page[start if start >= 0 else pos:end])
    return None


def lookup(citation, paragraph=None, keyword=None, opener=None):
    doc_id = document_id(citation)
    want = _norm_section(paragraph) if paragraph else None
    page = load(doc_id, want, opener)
    url = f"{_cache.get((doc_id, 'host'), HOSTS[-1])}/jogszabaly/{doc_id}"
    title = title_of(page) or str(citation)
    head = f"{title}\nForrás: {url} (Nemzeti Jogszabálytár, hatályos állapot, lekérve: {_today():%Y.%m.%d.})"
    if paragraph:
        text = section_text(page, paragraph)
        if not text:
            return (f"{head}\n\nA(z) {paragraph}. § nem található ebben a jogszabályban (lehet, hogy hatályon kívül "
                    "van, vagy más a számozása). Kulcsszavas kereséssel megtalálható a helyes szakasz.")
        if len(text) > MAX_SECTION_CHARS:
            text = text[:MAX_SECTION_CHARS] + " […]"
        return f"{head}\n\n{text}"
    if keyword:
        page = load(doc_id, None, opener)   # a kereséshez a teljes szöveg kell
        kw = str(keyword).lower()
        marks = _sections(page)
        hits = []
        for i, (s, pos) in enumerate(marks):
            end = marks[i + 1][1] if i + 1 < len(marks) else len(page)
            body = html_to_text(page[pos:end])
            k = body.lower().find(kw)
            if k >= 0:
                hits.append(f"- {s}. §: …{body[max(0, k - 120):k + 200].strip()}…")
            if len(hits) >= 12:
                break
        return f"{head}\n\nA(z) „{keyword}” kifejezés előfordulásai:\n" + ("\n".join(hits) if hits else "nincs találat")
    return f"{head}\n\nAdj meg paragrafust vagy kulcsszót."
