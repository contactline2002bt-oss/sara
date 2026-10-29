"""Irat készítése az Iratminta Tár Word-mintájából űrlappal.

A mintában a kitöltendő helyek jelölése:
  [Munkavállaló neve]                  adatmező (szöveg)
  [Kezdés napja|dátum]                 dátum (2026-10-17 → 2026. október 17.)
  [Munkabér|összeg]                    összeg (350000 → 350.000)
  [Munkabér|betűvel]                   az azonos nevű összeg betűvel, magától (háromszázötvenezer)
  [Próbaidő: 3 hónap / 6 hónap / nincs] választás a felsoroltakból
  [Feladatkör|hosszú]                  hosszabb szöveg (több sor)
  [[BLOKK: Felelősség]]                szabad blokk: külön bekezdésben; a felhasználó írja, vagy Sara javasol

A mezőket a program olvassa ki (AI nélkül), Sara csak előtölti őket a projekt és a törzsakták adataiból, és a szabad blokkra ad
javaslatot. A Word-fájlt a program tölti ki (docx_edit), így a minta fix szövege és formázása betűre megmarad.
"""
import io
import json
import re
import zipfile
from datetime import date

from lxml import etree

import docx_edit

BLOCK = re.compile(r"\[\[\s*blokk\s*:\s*([^\[\]]{1,80}?)\s*\]\]", re.I)
FIELD = re.compile(r"(?<!\[)\[([^\[\]]{1,90})\](?!\])")
TYPES = {"dátum": "date", "datum": "date", "összeg": "amount", "osszeg": "amount", "szám": "amount",
         "betűvel": "words", "betuvel": "words", "hosszú": "long", "hosszu": "long", "szöveg": "text",
         "feor": "feor"}
DOTS = "……………………"
# pénznemek: a mintában az összeg után álló pénznem (Ft, forint, EUR…) az űrlapon választhatóvá válik
CURRENCIES = {"HUF": ("Ft", "forint"), "EUR": ("EUR", "euró"), "USD": ("USD", "amerikai dollár"),
              "CHF": ("CHF", "svájci frank"), "GBP": ("GBP", "angol font")}
_CUR_WORDS = {"ft": "HUF", "huf": "HUF", "forint": "HUF", "eur": "EUR", "euró": "EUR", "euro": "EUR",
              "€": "EUR", "usd": "USD", "dollár": "USD", "chf": "CHF", "gbp": "GBP"}
CUR_RE = re.compile(r"(\s*)(Ft|HUF|forint|EUR|euró|euro|€|USD|dollár|CHF|GBP)(?![\wáéíóöőúüű])", re.I)


def currency_code(token):
    return _CUR_WORDS.get(str(token or "").strip().lower())


def currency_token(code, like):
    """Az új pénznem a minta írásmódja szerint: rövid (Ft → EUR) vagy kiírt (forint → euró)."""
    short, long_ = CURRENCIES.get(code, (code, code))
    return long_ if str(like).strip().lower() in ("forint", "euró", "euro", "dollár") else short
MONTHS = ["január", "február", "március", "április", "május", "június", "július", "augusztus", "szeptember",
          "október", "november", "december"]


class TemplateError(Exception):
    pass


# ---------------------------------------------------------------------------
# Számok betűvel (a magyar helyesírás szerint: 2000 fölött a hármas csoportok közt kötőjel)
# ---------------------------------------------------------------------------
_ONES = ["", "egy", "kettő", "három", "négy", "öt", "hat", "hét", "nyolc", "kilenc"]
_TENS = ["", "tíz", "húsz", "harminc", "negyven", "ötven", "hatvan", "hetven", "nyolcvan", "kilencven"]
_TENS_C = ["", "tizen", "huszon", "harminc", "negyven", "ötven", "hatvan", "hetven", "nyolcvan", "kilencven"]
_SCALES = ["", "ezer", "millió", "milliárd", "billió"]


def _below_1000(n, final):
    h, rest = divmod(n, 100)
    t, o = divmod(rest, 10)
    out = ""
    if h:
        out += ("" if h == 1 else ("két" if h == 2 else _ONES[h])) + "száz"
    if t:
        out += _TENS_C[t] if o else _TENS[t]
    if o:
        out += "kettő" if (o == 2 and final) else ("két" if o == 2 else _ONES[o])
    return out


def number_words(n):
    n = int(n)
    if n == 0:
        return "nulla"
    if n < 0:
        return "mínusz " + number_words(-n)
    groups = []
    k = 0
    while n_rest := (n // 1000 ** k):
        groups.append(n_rest % 1000)
        k += 1
    parts = []
    for i in range(len(groups) - 1, -1, -1):
        g = groups[i]
        if not g:
            continue
        if i == 1 and g == 1 and n < 2000:
            parts.append("ezer")
        else:
            parts.append(_below_1000(g, final=(i == 0)) + _SCALES[i])
    return ("-" if n > 2000 else "").join(parts)


def amount_digits(value):
    digits = re.sub(r"[\s.\u00a0\u202f]", "", str(value or "")).replace(",-", "")
    return digits if re.fullmatch(r"\d{1,15}", digits) else None


def fmt_amount(value):
    d = amount_digits(value)
    return f"{int(d):,}".replace(",", ".") if d else str(value or "").strip()


_MON = {m[:3]: i for i, m in enumerate(MONTHS, 1)}
_MON.update({"szep": 9, "már": 3, "mar": 3, "ápr": 4, "apr": 4, "máj": 5, "maj": 5, "jún": 6, "jun": 6,
             "júl": 7, "jul": 7, "okt": 10, "febr": 2})


def parse_date(value, today=None):
    """Rugalmas dátum: 2026-11-01, 2026.11.01., 2026. november 1., nov 1., 11.01., ma, holnap, tegnap.
    Év nélkül az idei év, ha az több mint fél éve elmúlt volna, akkor a következő. Visszaad: date vagy None."""
    from datetime import timedelta
    today = today or date.today()
    s = " ".join(str(value or "").strip().lower().split())
    if not s:
        return None
    rel = {"ma": 0, "mai nap": 0, "holnap": 1, "holnapután": 2, "tegnap": -1}
    if s.rstrip(".") in rel:
        return today + timedelta(days=rel[s.rstrip(".")])
    y = mo = d = None
    m = re.fullmatch(r"(\d{4})\s*[.\-/]\s*(\d{1,2})\s*[.\-/]\s*(\d{1,2})\.?", s)
    if m:
        y, mo, d = map(int, m.groups())
    else:
        m = re.fullmatch(r"(?:(\d{4})\.?\s*)?([a-záéíóöőúüű]{3,})\.?\s*(\d{1,2})\.?(?:\s*\(?\w*\)?)?", s)
        if m:
            word = m.group(2)
            mo = next((v for k, v in sorted(_MON.items(), key=lambda kv: -len(kv[0])) if word.startswith(k)), None)
            y, d = (int(m.group(1)) if m.group(1) else None), int(m.group(3))
        else:
            m = re.fullmatch(r"(\d{1,2})\s*[.\-/]\s*(\d{1,2})\.?", s)
            if m:
                mo, d = map(int, m.groups())
    if not (mo and d):
        return None
    try:
        if y:
            return date(y, mo, d)
        out = date(today.year, mo, d)
        return date(today.year + 1, mo, d) if (today - out).days > 183 else out
    except ValueError:
        return None


def fmt_date(value, today=None):
    dt = parse_date(value, today)
    return f"{dt.year}. {MONTHS[dt.month - 1]} {dt.day}." if dt else str(value or "").strip()


# ---------------------------------------------------------------------------
# A minta mezőinek kiolvasása
# ---------------------------------------------------------------------------
def _body(data):
    try:
        zin = zipfile.ZipFile(io.BytesIO(data))
        root = etree.fromstring(zin.read("word/document.xml"))
    except Exception:
        raise TemplateError("A minta nem olvasható Word-dokumentumként (.docx).")
    return root.find(docx_edit.W + "body")


def _parse_field(inner):
    inner = " ".join(inner.split())
    if re.fullmatch(r"M\d+|\d+|[a-zA-Z]|…*|\.*", inner):   # megjegyzésjel, hivatkozás, pontozás: nem mező
        return None
    if "|" in inner:
        parts = [x.strip() for x in inner.split("|")]
        name, typ = parts[0], (parts[1] if len(parts) > 1 else "")
        kind = TYPES.get(typ.lower(), "text")
        hint = " | ".join(parts[2:]) if len(parts) > 2 else ("" if typ.lower() in TYPES else typ)
        return {"name": name, "kind": kind, "options": [], "hint": hint}
    m = re.fullmatch(r"([^:]{1,60}):\s*(.+/.+)", inner)
    if m:
        opts = [o.strip() for o in m.group(2).split("/") if o.strip()]
        return {"name": m.group(1).strip(), "kind": "choice", "options": opts}
    low = inner.lower()
    if low.endswith(" betűvel") or low.endswith(" betuvel"):
        return {"name": inner.rsplit(" ", 1)[0].strip(), "kind": "words", "options": []}
    if re.search(r"\bfeor", low):
        return {"name": inner, "kind": "feor", "options": []}
    kind = "date" if (re.search(r"dátum|\bkelt|napja\b|születési idő", low) and "hely" not in low) else "text"
    return {"name": inner, "kind": kind, "options": []}


def scan(data):
    """Visszaad: {"fields": [{name, kind, options, derived}], "blocks": [{name, raw, before, after}],
    "placeholders": [(raw, name, kind)]}. Az azonos nevű mező egyszer szerepel az űrlapon."""
    body = _body(data)
    paras = [p for p in docx_edit._paragraphs(body)]
    texts = [docx_edit._para_text(p) for p in paras]
    fields, blocks, placeholders, dotted, cur_after, cur_by_name = {}, [], [], {}, {}, {}
    for i, text in enumerate(texts):
        for m in BLOCK.finditer(text):
            before = next((t for t in reversed(texts[:i]) if t.strip()), "")
            after = next((t for t in texts[i + 1:] if t.strip()), "")
            blocks.append({"name": m.group(1).strip(), "raw": m.group(0), "before": before[:400],
                           "after": after[:400]})
        for m in FIELD.finditer(BLOCK.sub("", text)):
            f = _parse_field(m.group(1))
            if not f:
                continue
            plain = BLOCK.sub("", text)
            ctx = plain[max(0, m.start() - 60):m.start()] + "▢" + plain[m.end():m.end() + 60]
            f["context"] = ("…" if m.start() > 60 else "") + " ".join(ctx.split()) + (
                "…" if m.end() + 60 < len(plain) else "")
            placeholders.append((m.group(0), f["name"], f["kind"]))
            cm = CUR_RE.match(plain[m.end():])
            if cm and f["kind"] in ("amount", "words", "text"):
                tail = cm.group(1) + cm.group(2)
                cur_after.setdefault(m.group(0), {}).setdefault(tail, 0)
                cur_after[m.group(0)][tail] += 1
                cur_by_name.setdefault(f["name"].lower(), currency_code(cm.group(2)))
            stat = dotted.setdefault(m.group(0), [0, 0])   # [pont követi, nem követi] – a dátum végi ponthoz
            stat[0 if text[m.end():m.end() + 1] == "." else 1] += 1
            key = (f["name"].lower(), f["kind"] == "words")
            if key not in fields:
                fields[key] = f
            elif f["kind"] == "choice" and not fields[key]["options"]:
                fields[key] = f
    base = {name for (name, words) in fields if not words}
    out = []
    for (name, words), f in fields.items():
        f = dict(f)
        f["derived"] = bool(words and name in base)   # a betűvel írt összeg a számmal megadottból jön
        if f["kind"] == "amount" or (f["kind"] == "words" and not f["derived"]):
            f["currency"] = cur_by_name.get(name)
        out.append(f)
    return {"fields": out, "blocks": blocks, "placeholders": list(dict.fromkeys(placeholders)),
            "dotted": dotted, "cur_after": cur_after}


def form_fields(info):
    return [f for f in info["fields"] if not f["derived"]]


def render_value(name, kind, values):
    """A mező végleges szövege az űrlap értékeiből (a betűvel írt összeg a számmal megadottból)."""
    lower = {k.lower(): v for k, v in values.items()}
    if kind == "words":
        src = lower.get(name.lower() + "|words") or lower.get(name.lower())
        d = amount_digits(src)
        return number_words(d) if d else str(src or "").strip()
    v = str(lower.get(name.lower()) or "").strip()
    if kind == "amount":
        return fmt_amount(v)
    if kind == "date":
        return fmt_date(v)
    if kind == "feor":
        import feor
        return feor.code_of(v) or v
    return v


def value_key(f):
    return f["name"] + ("|words" if f["kind"] == "words" else "")


def missing(info, values, blocks_text):
    out = [f["name"] for f in form_fields(info) if not str(values.get(value_key(f)) or "").strip()]
    out += [f"{b['name']} (blokk)" for b in info["blocks"] if not str(blocks_text.get(b["name"]) or "").strip()]
    return out


def accept_revisions(data):
    """A mintában maradt követett változások elfogadása (a törölt szöveg kimarad, a beszúrt marad), hogy a kész
    iratban ne legyen áthúzás és változásjelölés."""
    W = docx_edit.W
    try:
        zin = zipfile.ZipFile(io.BytesIO(data))
        root = etree.fromstring(zin.read("word/document.xml"))
    except Exception:
        return data
    changed = False
    for tag in ("del", "moveFrom", "rPrChange", "pPrChange", "sectPrChange", "tblPrChange", "trPrChange",
                "tcPrChange", "numberingChange"):
        for el in list(root.iter(W + tag)):
            if tag in ("del", "moveFrom") and el.getparent() is not None and el.getparent().tag == W + "rPr":
                el.getparent().remove(el)   # törölt bekezdésjel jelölése: csak a jelölést vesszük le
            else:
                el.getparent().remove(el)
            changed = True
    for tag in ("ins", "moveTo"):
        for el in list(root.iter(W + tag)):
            parent = el.getparent()
            if parent is None:
                continue
            if parent.tag == W + "rPr":
                parent.remove(el)
            else:
                idx = parent.index(el)
                for child in list(el):
                    parent.insert(idx, child)
                    idx += 1
                parent.remove(el)
            changed = True
    for tag in ("moveFromRangeStart", "moveFromRangeEnd", "moveToRangeStart", "moveToRangeEnd"):
        for el in list(root.iter(W + tag)):
            el.getparent().remove(el)
    if not changed:
        return data
    out = io.BytesIO()
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            payload = (etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
                       if item.filename == "word/document.xml" else zin.read(item.filename))
            zout.writestr(item, payload)
    return out.getvalue()


def fill(data, info, values, blocks_text, dots_for_missing=True):
    """A kitöltött Word-fájl bájtjai. values: {mezőkulcs: érték}; blocks_text: {blokknév: szöveg vagy ''}."""
    data = accept_revisions(data)
    edits = []
    lower = {k.lower(): v for k, v in values.items()}
    for raw, name, kind in info["placeholders"]:
        text = render_value(name, kind, values)
        if not text:
            if not dots_for_missing:
                continue
            text = DOTS
        dot, plain = (info.get("dotted") or {}).get(raw, [0, 1])
        tails = (info.get("cur_after") or {}).get(raw) or {}
        chosen = str(lower.get(name.lower() + "|currency") or "").upper()
        if tails and chosen and text != DOTS:   # „[Vételár|összeg] Ft” → „120.000 EUR”, ha EUR-t választott
            changed = 0
            for tail, count in tails.items():
                tok = tail.strip()
                if currency_code(tok) != chosen:
                    sep = tail[:len(tail) - len(tail.lstrip())]
                    edits.append({"action": "replace", "find": raw + tail,
                                  "new_text": text + sep + currency_token(chosen, tok), "all": True})
                    changed += count
            if changed >= dot + plain:
                continue
        if text.endswith(".") and dot:   # „[Kezdés napja].” → ne legyen dupla pont
            edits.append({"action": "replace", "find": raw + ".", "new_text": text, "all": True})
            if not plain:
                continue
        edits.append({"action": "replace", "find": raw, "new_text": text, "all": True})
    for b in info["blocks"]:
        text = str(blocks_text.get(b["name"]) or "").strip()
        if text:
            edits.append({"action": "insert_after", "find": b["raw"], "new_text": text.replace("\r", ""),
                          "all": True})
            edits.append({"action": "delete", "find": b["raw"], "all": True})
        else:
            edits.append({"action": "replace", "find": b["raw"], "new_text": DOTS, "all": True})
    if not edits:
        return data
    try:
        out, _ = docx_edit.apply_edits(data, edits, tracked=False, author="Sara")
    except docx_edit.EditError as e:
        raise TemplateError(str(e))
    return out


# ---------------------------------------------------------------------------
# Sara: előtöltés és blokkjavaslat
# ---------------------------------------------------------------------------
PREFILL_SYSTEM = (
    "Egy cégvezető iratmintájának mezőit töltöd elő a projekt és a kapcsolt törzsakták (cég, ingatlan, személy, jármű adatlapja) és a felhasználó profilja alapján. Csak olyan adatot írj be, "
    "amely az iratok SZÖVEGÉBEN szó szerint vagy egyértelműen szerepel, és add meg, melyik iratból vetted. A projekt "
    "nevéből, a fájlnevekből és az „=== … ===” fejlécekből SOHA ne vegyél adatot (a felhasználó saját elnevezései, "
    "nem okmányadatok). Amit nem találsz, azt hagyd ki (ne írj üres értéket). SOHA ne találj ki adatot, ne becsülj, ne számolj ki dátumot. "
    "Választós mezőnél csak a felsorolt lehetőségek egyikét add meg, pontosan úgy írva. Dátum alakja: "
    "ÉÉÉÉ-HH-NN; összegé csak számjegy. Neveket, címeket az okmány szerinti pontos írásmóddal (ékezetekkel). "
    "Ha egy mezőhöz több személy vagy adat is szóba jön (pl. több ügyfél-adatlap, több képviselő), és a "
    "kijelölt iratokból vagy az utasításból nem dönthető el, melyik kell, NE válassz: az ertek maradjon üres, és a "
    "jeloltek listában add meg a lehetőségeket (legfeljebb 4). "
    "Válasz csak JSON: {\"mezok\": {\"<mező neve pontosan>\": {\"ertek\": \"…\", \"forras\": \"…\", "
    "\"jeloltek\": [\"…\"]}}}")

BLOCK_SYSTEM = (
    "Egy cégvezető iratmintájának egy szabad blokkját fogalmazod meg a felhasználó utasítása szerint, "
    "a hatályos magyar jog és a szokásos szerződési nyelvezet szerint, a minta stílusában. Csak a blokk szövegét "
    "add vissza (cím, magyarázat, Markdown-jelölés nélkül); a bekezdéseket üres sor válassza el. Amit nem "
    "tudsz (adat, összeg), azt [kitöltendő: …] jelöléssel hagyd nyitva; ne találj ki adatot.")


def _call(client, model, system, user, log=None, max_tokens=3000):
    kwargs = dict(model=model, max_tokens=max_tokens, system=system, messages=[{"role": "user", "content": user}])
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


def prefill(client, model, template_name, info, context, instructions="", log=None):
    """Visszaad: {mezőkulcs: (érték, forrás)} – csak a talált adatok."""
    fields = form_fields(info)
    if not fields:
        return {}
    lines = []
    for f in fields:
        desc = {"feor": "FEOR-08 kód, 4 számjegy", "date": "dátum", "amount": "összeg", "words": "összeg (számmal)", "long": "hosszabb szöveg",
                "choice": "választás: " + " / ".join(f["options"])}.get(f["kind"], "szöveg")
        lines.append(f"- {f['name']} ({desc})")
    user = (f"Minta: {template_name}\n"
            + (f"A felhasználó utasításai ehhez a mintához:\n{instructions}\n" if instructions else "")
            + "Kitöltendő mezők:\n" + "\n".join(lines) + "\n\nAz anyagok:\n" + context)
    raw = _call(client, model, PREFILL_SYSTEM, user, log, max_tokens=4000)
    data = _loose_json(raw)
    if isinstance(data, dict) and isinstance(data.get("mezok"), dict):
        data = data["mezok"]
    if not isinstance(data, dict):
        return {}
    by_name = {_key(f["name"]): f for f in fields}
    out, rejected = {}, []
    LAST_CANDIDATES.clear()
    for name, item in data.items():
        f = by_name.get(_key(name))
        if not f:
            continue
        if isinstance(item, dict):
            value, source = item.get("ertek") or item.get("érték") or item.get("value"), item.get("forras") or \
                item.get("forrás") or item.get("source")
            cands = [str(c).strip() for c in (item.get("jeloltek") or []) if str(c).strip()]
            ctx_key = _key(context)
            cands = [c for c in cands if grounded(c, f, context) and in_body(c, f, context)
                     and _key(c) in ctx_key][:4]   # csak a forrásból
            if len(cands) > 1 and not str(value or "").strip():
                LAST_CANDIDATES[value_key(f)] = cands
        else:
            value, source = item, ""
        value = str(value or "").strip()
        if not value or value.lower() in ("null", "none", "-", "–", "nincs adat", "ismeretlen"):
            continue
        if f["kind"] == "choice":
            value = next((o for o in f["options"] if o.lower() == value.lower()), "")
            if not value:
                continue
        if not grounded(value, f, context) or not in_body(value, f, context):   # csak ami a szövegben szerepel
            rejected.append(f["name"])
            continue
        out[value_key(f)] = (value, str(source or "").strip())
    LAST_REJECTED[:] = rejected
    return out


def _text_body(context):
    """A kontextus az „=== fájlnév ===” fejlécek nélkül: ami csak fájlnévben/projektnévben áll, az nem forrás."""
    return "\n".join(line for line in str(context or "").splitlines()
                     if not (line.startswith("===") and line.rstrip().endswith("===")))


def in_body(value, field, context):
    """Szöveges mező (név, cím, cégnév) minden legalább háromjegyű szava szerepeljen az iratok szövegében (nem
    elég, ha csak a fájlnévben vagy a projekt nevében van). Választós, dátum, összeg és hosszú szöveg kivétel."""
    if field.get("kind") not in ("text", None, ""):
        return True
    words = [w for w in _key(value).split() if len(w) >= 3 and not w.isdigit()]
    if not words:
        return True
    body = " " + _key(_text_body(context)) + " "
    return all(w in body for w in words)


LAST_REJECTED = []   # az utolsó előtöltésből kiszűrt (forrásban nem talált) mezők neve
LAST_CANDIDATES = {}   # mezőkulcs → több lehetséges érték, ha az aktából nem dönthető el (pl. két ügyfél)


def grounded(value, field, context):
    """Számjegyes adat (adóazonosító, adószám, cégjegyzékszám, irányítószám, összeg, dátum) csak akkor fogadható
    el, ha a számjegyei – elválasztóktól függetlenül – a forrásban is megvannak. Ez szűri ki a kitalált számot."""
    digits = re.sub(r"\D", "", str(value or ""))
    if len(digits) < 4:
        return True
    ctx = re.sub(r"[\s.\-/,]", "", str(context or ""))
    if digits in ctx:
        return True
    # cím és hasonló vegyes adat (2131 Göd, Petőfi Sándor utca 37/B): a számcsoportok külön-külön legyenek meg
    groups = re.findall(r"\d+", str(value or ""))
    if len(groups) > 1 and re.search(r"[A-Za-zÁÉÍÓÖŐÚÜŰáéíóöőúüű]{3,}", str(value or "")):
        ctx_groups = set(re.findall(r"\d+", str(context or "")))
        if all(g in ctx_groups for g in groups):
            return True
    if field.get("kind") == "date":   # dátum más alakban: 1997.10.26 ↔ 1997. október 26.
        dt = parse_date(value)
        if dt:
            low = ctx.lower()
            for form in (f"{dt.year}{dt.month:02d}{dt.day:02d}", f"{dt.year}{MONTHS[dt.month - 1]}{dt.day}",
                         f"{dt.year}{dt.month}{dt.day}"):
                if form in low:
                    return True
    return False


def _key(name):
    """Mezőnév összevetéshez: a „|típus” és a zárójeles megjegyzés nélkül, ékezet és kis-nagybetű nélkül."""
    import unicodedata
    n = str(name or "").split("|")[0]
    n = re.sub(r"\(.*?\)", " ", n)
    n = unicodedata.normalize("NFKD", n.lower())
    n = "".join(c for c in n if not unicodedata.combining(c))
    return " ".join(re.sub(r"[^\w ]+", " ", n).split())


def _loose_json(raw):
    """JSON a modell válaszából akkor is, ha kódblokkban vagy szöveg között jön."""
    raw = re.sub(r"```(?:json)?", "", str(raw or ""))
    start = raw.find("{")
    if start < 0:
        return None
    depth, in_str, esc = 0, False, False
    for i, ch in enumerate(raw[start:], start):
        if in_str:
            esc = (ch == "\\") and not esc
            if ch == '"' and not esc:
                in_str = False
            continue
        if ch == '"':
            in_str = True
        elif ch == "{":
            depth += 1
        elif ch == "}":
            depth -= 1
            if depth == 0:
                try:
                    return json.loads(raw[start:i + 1])
                except ValueError:
                    return None
    return None


def suggest_block(client, model, template_name, block, values, instruction, context="", log=None):
    filled = "\n".join(f"- {k.split('|')[0]}: {v}" for k, v in values.items() if str(v or "").strip())
    user = (f"Minta: {template_name}\nA blokk: {block['name']}\n"
            f"A blokk előtti bekezdés: {block.get('before') or '–'}\nA blokk utáni bekezdés: "
            f"{block.get('after') or '–'}\n"
            + (f"Az irat kitöltött adatai:\n{filled}\n" if filled else "")
            + (f"Az ügy háttere:\n{context[:6000]}\n" if context else "")
            + f"\nA felhasználó utasítása erre a blokkra: {instruction or 'a szokásos, kiegyensúlyozott szöveg'}\n"
            "Csak a blokk saját szövegét írd: az előtte álló címet és az utána következő bekezdést NE ismételd meg.")
    return strip_neighbours(_call(client, model, BLOCK_SYSTEM, user, log, max_tokens=2500), block)


def _norm(t):
    return re.sub(r"[\W_]+", " ", str(t or "")).strip().lower()


def strip_neighbours(text, block):
    """A modell néha a blokk előtti címet és az utána álló bekezdést is visszaírja – ezek ne duplázódjanak."""
    paras = [p for p in re.split(r"\n\s*\n|\n", str(text or "").strip()) if p.strip()]
    paras = [p.strip().strip("*#").strip() for p in paras]
    before, after = _norm(block.get("before")), _norm(block.get("after"))
    while paras and before and _norm(paras[0]) == before:
        paras.pop(0)
    while paras and after and _norm(paras[-1]) == after:
        paras.pop()
    return "\n\n".join(paras)


def validate(name, kind, value):
    """Gyors, AI nélküli ellenőrzés a mező neve alapján. Visszaad: figyelmeztetés vagy None."""
    v = str(value or "").strip()
    if not v:
        return None
    low = name.lower()
    digits = re.sub(r"\D", "", v)
    if kind == "feor":
        import feor
        return None if feor.code_of(v) else "nem FEOR-08 kód (4 számjegy a KSH listájából)"
    if "próbaidő" in low:
        m = re.search(r"(\d+)\s*hónap", v)
        if m and int(m.group(1)) > 3:
            return "az Mt. 45. § (5) szerint legfeljebb 3 hónap (kollektív szerződés legfeljebb 6 hónapot engedhet)"
    if kind == "date":
        dt = parse_date(v)
        if not dt:
            return "nem értem a dátumot (pl. 2026-11-01, nov 1., ma)"
        if not 1900 <= dt.year <= 2100:
            return f"furcsa évszám: {dt.year}"
        if "születési" in low and dt >= date.today():
            return "a születési idő nem lehet mai vagy jövőbeli"
    if "adóazonosító" in low:
        if len(digits) != 10 or digits[0] != "8" or sum(int(d) * i for i, d in enumerate(digits[:9], 1)) % 11 % 10 != int(digits[9]) \
                or sum(int(d) * i for i, d in enumerate(digits[:9], 1)) % 11 == 10:
            return "hibás adóazonosító jel (10 jegy, 8-cal kezdődik, ellenőrző számjegy)"
    elif "adószám" in low:
        if len(digits) not in (8, 11):
            return "az adószám 8 vagy 11 számjegy"
        total = sum(int(d) * w for d, w in zip(digits[:7], (9, 7, 3, 1, 9, 7, 3)))
        if (10 - total % 10) % 10 != int(digits[7]):
            return "hibás adószám (az ellenőrző számjegy nem stimmel)"
    elif "cégjegyzékszám" in low and not re.fullmatch(r"\d{2}-\d{2}-\d{6}", v.replace(" ", "")):
        return "a cégjegyzékszám alakja: 01-09-123456"
    elif kind in ("amount", "words") and not amount_digits(v):
        return "nem szám (pl. 350000)"
    return None


def company_name_for(field_name, values):
    """Az adószám-mezőhöz tartozó cégnév az űrlapon: ugyanazzal a szóval kezdődő „neve” mező
    (pl. Munkáltató adószáma → Munkáltató neve)."""
    head = field_name.split()[0].lower() if field_name.split() else ""
    for key, value in values.items():
        k = key.lower()
        if ("neve" in k or "név" in k or "cégnév" in k) and (k.startswith(head) or head in ("adószám", "adószáma")):
            if str(value or "").strip():
                return str(value).strip()
    return ""


def names_match(a, b):
    """Cégnév-egyezés laza összevetéssel (rövid és teljes név, cégforma nélkül, ékezet és kis-nagybetű mindegy)."""
    import difflib
    import unicodedata
    forms = r"\b(kft|bt|zrt|nyrt|kkt|rt|korlátolt felelősségű társaság|betéti társaság|zártkörűen működő " \
            r"részvénytársaság|részvénytársaság|közkereseti társaság)\b\.?"

    def norm(x):
        x = unicodedata.normalize("NFKD", str(x or "").lower())
        x = "".join(c for c in x if not unicodedata.combining(c))
        x = re.sub(forms.encode().decode(), " ", unicodedata.normalize("NFKD", x))
        return " ".join(re.sub(r"[^\w ]+", " ", x).split())
    na, nb = norm(a), norm(b)
    if not na or not nb:
        return False
    return na == nb or na in nb or nb in na or difflib.SequenceMatcher(None, na, nb).ratio() > 0.8


# ---------------------------------------------------------------------------
# Mintakészítő: általános / kétértelmű mezőnevek rendbetétele
# ---------------------------------------------------------------------------
GENERIC = {"adat", "adatok", "név", "nev", "szöveg", "szoveg", "dátum", "datum", "x", "xx", "érték", "ertek", "szám",
           "szam", "cím", "cim", "hely", "összeg", "osszeg", "mező", "mezo", "kitöltendő", "kitoltendo", "…", "..."}


def problems(info):
    """A minta gyanús mezőnevei: általános (pl. [adat], [név]) vagy furcsa (zárójeles formátumleírás)."""
    out = []
    for raw, name, kind in info.get("placeholders") or []:
        low = name.strip().lower()
        if low in GENERIC or "(" in name or re.fullmatch(r"[\W\d_]*", name):
            out.append(raw)
    return list(dict.fromkeys(out))


CLEANUP_SYSTEM = (
    "Egy ügyvédi Word-iratminta kitöltendő helyeit teszed egyértelművé. A mintában a kitöltendő hely jelölése "
    "[Mező neve]; típus is megadható: [Név|dátum], [Név|összeg], [Név|betűvel], [Név|hosszú], választás: "
    "[Név: A / B]. A gond: egyes mezők általános nevet kaptak (pl. [adat], [név], [dátum(…)]), így az űrlap "
    "egy mezőnek látja őket, pedig mást jelentenek. Feladat: minden ilyen előfordulásra adj cserét.\n"
    "Szabályok: 1) A find a minta szövegéből SZÓ SZERINT másolt részlet legyen, benne a jelöléssel, és annyi "
    "környezettel, hogy a mintában pontosan EGYSZER forduljon elő (pl. „anyja neve: [adat]”); ha két bekezdés szövege teljesen azonos (pl. két „Budapest, [dátum] napján” sor), add meg a bekezdés sorszámát is (bekezdes), és a környező bekezdésekből döntsd el, melyik mi (pl. megbízás dátuma / elfogadás dátuma). 2) A replace "
    "ugyanez a részlet, csak a szögletes zárójelen belül más: a zárójelen kívüli szöveget betűre hagyd meg. "
    "3) A mezőnév rövid, beszédes, magyar (pl. [Meghatalmazó anyja neve], [Cég székhelye], [Kelt|dátum]); "
    "ugyanaz a jelentés mindenhol ugyanazt a nevet kapja (pl. a meghatalmazó neve mindenütt [Meghatalmazó neve]). "
    "4) Dátum: |dátum, pénzösszeg: |összeg. 5) Csak a felsorolt gyanús jelöléseket cseréld. "
    "Válasz csak JSON: {\"csere\": [{\"find\": \"…\", \"replace\": \"…\", \"bekezdes\": 12, \"miert\": \"…\"}]} "
    "(a bekezdes csak azonos szövegű helyeknél kell).")


def _outside(text):
    return re.sub(r"\[[^\[\]]*\]", "[]", str(text or ""))


def suggest_cleanup(client, model, data, info, log=None):
    """Javasolt cserék [{find, replace, miert}] – csak az ellenőrzött (egyedi, a zárójelen kívül változatlan)."""
    bad = problems(info)
    if not bad:
        return []
    body = _body(data)
    texts = [docx_edit._para_text(p) for p in docx_edit._paragraphs(body)]
    numbered = "\n".join(f"{i + 1}. {t}" for i, t in enumerate(texts) if t.strip())
    user = (f"Gyanús jelölések: {', '.join(bad)}\n\nA minta bekezdései:\n{numbered[:30000]}")
    raw = _call(client, model, CLEANUP_SYSTEM, user, log, max_tokens=4000)
    data_json = _loose_json(raw) or {}
    items = data_json.get("csere") if isinstance(data_json, dict) else None
    out, seen = [], set()
    full = "\n".join(texts)
    for it in items or []:
        if not isinstance(it, dict):
            continue
        find, repl = str(it.get("find") or ""), str(it.get("replace") or "")
        if not find or not repl or find == repl or "\n" in find:
            continue
        if _outside(find) != _outside(repl):
            continue   # a zárójelen kívüli szöveghez nem nyúlhat
        para = None
        if full.count(find) != 1:   # azonos szövegű helyek: csak a megadott bekezdésben, ott egyszer
            try:
                para = int(it.get("bekezdes"))
            except (TypeError, ValueError):
                continue
            if not (1 <= para <= len(texts)) or texts[para - 1].count(find) != 1:
                continue
        key = (find, para)
        if key in seen:
            continue
        seen.add(key)
        where = texts[para - 1] if para else ""
        out.append({"find": find, "replace": repl, "paragraph": para, "why": str(it.get("miert") or "").strip(),
                    "where": where})
    return out


def apply_cleanup(data, edits):
    data = accept_revisions(data)
    if not edits:
        return data
    try:
        out, _ = docx_edit.apply_edits(data, [dict({"action": "replace", "find": e["find"], "new_text": e["replace"]},
                                                   **({"paragraph": e["paragraph"]} if e.get("paragraph") else {}))
                                              for e in edits], tracked=False, author="Sara")
    except docx_edit.EditError as e:
        raise TemplateError(str(e))
    return out
