"""FEOR-08 (Foglalkozások Egységes Osztályozási Rendszere, KSH) a munkaszerződés-mintákhoz.

A 485 foglalkozás kódja és megnevezése a feor08.tsv fájlban van (a KSH nyilvános FEOR-08 listája,
https://www.ksh.hu/docs/szolgaltatasok/hun/feor08/feorlista.html). Egy foglalkozás KSH-leírását (feladatok,
ide tartozó munkakörök) a program igény szerint a KSH oldaláról olvassa be; ha az nem érhető el, a javaslat
csak a megnevezésből készül. A FEOR-kód a biztosítotti bejelentéshez (NAV, 08E / T1041) kötelező adat.
"""
import difflib
import os
import re
import unicodedata
import urllib.request

KSH_URL = "https://www.ksh.hu/docs/szolgaltatasok/hun/feor08/{group}/{code}.html"
UA = ("Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) "
      "Chrome/124.0 Safari/537.36")


def _load():
    path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "feor08.tsv")
    out = {}
    with open(path, encoding="utf-8") as fh:
        for line in fh:
            code, _, name = line.rstrip("\n").partition("\t")
            if re.fullmatch(r"\d{4}", code):
                out[code] = name.strip()
    return out


CODES = _load()


def label(code):
    return f"{code} – {CODES[code]}" if code in CODES else str(code or "")


def code_of(value):
    """A beírt vagy kiválasztott értékből a négyjegyű kód (pl. „4111 – Titkár(nő)” → 4111)."""
    m = re.search(r"\b(\d{4})\b", str(value or ""))
    return m.group(1) if m and m.group(1) in CODES else None


def _fold(text):
    t = unicodedata.normalize("NFKD", str(text or "").lower())
    return "".join(c for c in t if not unicodedata.combining(c))


def suggest(text, n=4):
    """Munkakör szövegéhez a legvalószínűbb FEOR-kódok (AI nélkül, szóegyezés és hasonlóság alapján)."""
    q = _fold(text)
    words = [w for w in re.findall(r"\w+", q) if len(w) >= 3]
    if not words:
        return []
    scored = []
    for code, name in CODES.items():
        nf = _fold(name)
        nwords = re.findall(r"\w+", nf)
        hit = 0.0
        for w in words:
            if w in nwords:
                hit += 1
            elif any(len(x) >= 5 and (x.startswith(w) or w.startswith(x)) for x in nwords):
                hit += 0.8
            elif any(len(x) >= 5 and x[:max(5, len(x) - 3)] == w[:max(5, len(x) - 3)] for x in nwords):
                hit += 0.5
        score = hit / len(words) + 0.5 * difflib.SequenceMatcher(None, q, nf).ratio()
        if nf.startswith("egyeb"):
            score -= 0.05   # a gyűjtőkategória csak akkor nyerjen, ha nincs jobb
        if hit:
            scored.append((score, code))
    scored.sort(reverse=True)
    return [c for _, c in scored[:n]]


def _html_text(data):
    try:
        from lxml import html as lhtml
        doc = lhtml.fromstring(data)
        for bad in doc.xpath("//script|//style|//nav|//header|//footer"):
            bad.getparent().remove(bad)
        text = doc.text_content()
    except Exception:
        text = re.sub(r"<[^>]+>", " ", data.decode("utf-8", "replace"))
    return "\n".join(" ".join(line.split()) for line in text.splitlines() if line.strip())


def describe(code, opener=None, timeout=15):
    """A foglalkozás KSH-leírása (szöveg) vagy None, ha az oldal nem érhető el."""
    code = code_of(code)
    if not code:
        return None
    req = urllib.request.Request(KSH_URL.format(group=code[0], code=code), headers={"User-Agent": UA})
    try:
        with (opener or urllib.request.urlopen)(req, timeout=timeout) as resp:
            data = resp.read()
    except Exception:
        return None
    text = _html_text(data)
    i = text.find(code)   # a fejléc és a menü utáni rész
    if i > 0:
        text = text[i:]
    text = re.split(r"\bVissza\b", text)[0]
    return text[:4000].strip() or None


TASKS_SYSTEM = (
    "Magyar munkaszerződéshez a munkakör feladatainak (munkaköri leírás) javaslatát írod. A KSH FEOR-08 leírásából "
    "és a megadott munkakörből dolgozz; a feladatok legyenek konkrétak, egyenként egy sorban, gondolatjellel "
    "kezdve, 6–12 pont. A végére írj egy záró pontot: „– a munkáltató által a munkakörébe utalt egyéb, "
    "képzettségének megfelelő feladatok ellátása.” Csak a listát add vissza, bevezető és magyarázat nélkül.")


def suggest_tasks(client, model, code, job_title="", ksh_text=None, log=None):
    code = code_of(code)
    if not code:
        raise ValueError("Előbb válassz FEOR-kódot.")
    user = (f"FEOR-08: {label(code)}\nMunkakör a szerződésben: {job_title or CODES[code]}\n"
            + (f"A KSH leírása:\n{ksh_text}\n" if ksh_text else
               "(A KSH leírása most nem érhető el; a megnevezésből dolgozz, általános, óvatos feladatokkal.)\n"))
    kwargs = dict(model=model, max_tokens=1500, system=TASKS_SYSTEM, messages=[{"role": "user", "content": user}])
    try:
        response = client.messages.create(thinking={"type": "disabled"}, **kwargs)
    except Exception as e:
        if getattr(e, "status_code", None) != 400 or "thinking" not in str(e).lower():
            raise
        response = client.messages.create(**kwargs)
    if log:
        log(model, getattr(response, "usage", None))
    return "".join(getattr(b, "text", "") for b in getattr(response, "content", [])
                   if getattr(b, "type", "") == "text").strip()
