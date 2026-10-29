"""Törzsakták: cég, ingatlan, személy, jármű – tartós nyilvántartás, mindegyik élén egy adatlappal.

A cég- és az ingatlanadatlap a registry modulból jön (cégkivonat, tulajdoni lap). Itt a személy (személyi
igazolvány, lakcímkártya, útlevél, vezetői engedély) és a jármű (forgalmi engedély, törzskönyv) adatlapja készül:
a modell kiolvas, a program ellenőrzi a formátumokat (dátum, alvázszám, rendszám), és ami bizonytalan, azt jelzi.
A személyes adatokból csak annyit gyűjtünk, amennyi iratokhoz (meghatalmazás, nyilatkozat) kell; a személyi
azonosítót (személyi számot) szándékosan nem.
Streamlittől független, ezért tesztelhető.
"""
import base64
import io
import json
import re

RECORD_TYPES = {   # típus: (megnevezés, ikon, többes szám)
    "ceg": ("Cég", ":material/apartment:", "Cégek"),
    "ingatlan": ("Ingatlan", ":material/home:", "Ingatlanok"),
    "szemely": ("Személy", ":material/person:", "Személyek"),
    "jarmu": ("Jármű", ":material/directions_car:", "Járművek"),
}
SHEET_PREFIX = "Adatlap – "
MAX_SIDE = 1568

PERSON_SYSTEM = (
    "Személyazonosító okmányokból (személyi igazolvány, lakcímkártya, útlevél, vezetői engedély; elő- és hátoldal) "
    "karakterpontosan kiolvasod egy személy adatait iratok (meghatalmazás, nyilatkozat) elkészítéséhez.\n"
    "Szabályok: SOHA ne találj ki adatot; ami nem látszik tisztán, az legyen null. A nevet és a többi adatot a "
    "nyomtatott (ékezetes) mezőkből vedd, ne az alsó gépi sorból (MRZ). Dátum: ÉÉÉÉ.HH.NN. A lakcímkártya "
    "hátoldalán lévő személyi azonosítót (személyi számot) NE add meg. Csak egy JSON-objektumot adj vissza:\n"
    '{"nev": "", "szuletesi_nev": "", "anyja_neve": "", "szuletesi_hely": "", "szuletesi_ido": "", '
    '"allampolgarsag": "", "lakcim": "", "tartozkodasi_hely": null, '
    '"okmanyok": [{"tipus": "Személyazonosító igazolvány|Útlevél|Vezetői engedély|Lakcímkártya", "szam": "", '
    '"ervenyes": "ÉÉÉÉ.HH.NN. vagy null"}], "megjegyzes": "csak valódi bizonytalanság esetén, egy mondatban"}')

VEHICLE_SYSTEM = (
    "Magyar forgalmi engedélyből (és ha van, törzskönyvből) karakterpontosan kiolvasod egy jármű adatait. A "
    "forgalmi engedély mezőkódjai: A rendszám, B első nyilvántartásba vétel, C.1 üzembentartó, C.4 tulajdonos "
    "(ha jelölve), D.1 gyártmány, D.2 típus, D.3 kereskedelmi név, E alvázszám (azonosító szám), P.1 "
    "hengerűrtartalom, P.2 teljesítmény, P.3 hajtóanyag, R szín, S.1 ülések száma, F.1 megengedett "
    "össztömeg, és a műszaki érvényesség. SOHA ne találj ki adatot; ami nem látszik, az legyen null. Dátum: "
    "ÉÉÉÉ.HH.NN. Csak egy JSON-objektumot adj vissza:\n"
    '{"rendszam": "", "alvazszam": "", "gyartmany": "", "tipus": "", "kereskedelmi_nev": "", '
    '"elso_nyilvantartasba_vetel": "", "gyartasi_ev": "", "uzembentarto": "", "tulajdonos": "", '
    '"szin": "", "hengerurtartalom": "", "teljesitmeny": "", "hajtoanyag": "", "ulesek": "", '
    '"ossztomeg": "", "muszaki_ervenyes": "", "forgalmi_szama": "", "torzskonyv_szama": "", '
    '"megjegyzes": "csak valódi bizonytalanság esetén"}')

VIN_RE = re.compile(r"^[A-HJ-NPR-Z0-9]{17}$")
PLATE_RE = re.compile(r"^([A-Z]{3}-\d{3}|[A-Z]{2} ?[A-Z]{2}-\d{3}|[A-Z]{1,4}-?\d{1,4})$")
DATE_RE = re.compile(r"^(19|20)\d{2}\.\d{2}\.\d{2}\.?$")


# --- képek előkészítése (az okmány kivágva, nagy felbontásban) --------------------------------------------
def _encode(img):
    from PIL import Image
    img = img.convert("RGB")
    if max(img.size) > MAX_SIDE:
        r = MAX_SIDE / max(img.size)
        img = img.resize((int(img.width * r), int(img.height * r)), Image.LANCZOS)
    buf = io.BytesIO()
    img.save(buf, "JPEG", quality=92)
    return buf.getvalue()


def _images(name, mime, data):
    """A PDF-be ágyazott szkennelt képek vagy a feltöltött fotó JPEG-ként; üres lista, ha nem sikerül."""
    try:
        from PIL import Image
        images = []
        if mime == "application/pdf":
            from pypdf import PdfReader
            for page in PdfReader(io.BytesIO(data)).pages[:4]:
                for im in page.images:
                    pil = Image.open(io.BytesIO(im.data))
                    if min(pil.size) >= 300:
                        images.append(pil)
        elif mime.startswith("image/"):
            images.append(Image.open(io.BytesIO(data)))
        return [_encode(i) for i in images[:6]]
    except Exception:
        return []


def _content(files):
    content = []
    for name, mime, data in files:
        jpgs = _images(name, mime, data)
        if jpgs:
            for j in jpgs:
                content.append({"type": "image", "source": {"type": "base64", "media_type": "image/jpeg",
                                                             "data": base64.b64encode(j).decode("ascii")}})
        elif mime == "application/pdf":
            content.append({"type": "document", "title": name, "source": {
                "type": "base64", "media_type": "application/pdf", "data": base64.b64encode(data).decode("ascii")}})
        else:
            continue
        content.append({"type": "text", "text": f"(Az előző a(z) {name} fájl.)"})
    return content


def _ask(client, model, system, content, log=None):
    kwargs = dict(model=model, max_tokens=3000, system=system,
                  messages=[{"role": "user", "content": content + [{"type": "text", "text": "Olvasd ki az adatokat."}]}])
    try:
        response = client.messages.create(thinking={"type": "disabled"}, **kwargs)
    except Exception as e:
        if getattr(e, "status_code", None) != 400 or "thinking" not in str(e).lower():
            raise
        response = client.messages.create(**kwargs)
    if log:
        log(model, getattr(response, "usage", None))
    text = "".join(getattr(b, "text", "") for b in response.content if getattr(b, "type", "") == "text")
    m = re.search(r"\{.*\}", text, re.S)
    if not m:
        raise ValueError("a modell nem adott értelmezhető adatokat; jobb minőségű kép kellene")
    return json.loads(m.group(0))


def _v(x):
    return str(x).strip() if x not in (None, "", "null") else "–"


# --- személy ---------------------------------------------------------------------------------------------
def person_checks(p):
    warns = []
    for key, label in (("szuletesi_ido", "születési idő"),):
        v = str(p.get(key) or "").strip()
        if v and not DATE_RE.match(v):
            warns.append(f"a(z) {label} („{v}”) szokatlan alakú")
    for d in p.get("okmanyok") or []:
        v = str(d.get("ervenyes") or "").strip()
        if v and DATE_RE.match(v):
            from datetime import date
            y, mo, da = (int(x) for x in v.rstrip(".").split("."))
            try:
                if date(y, mo, da) < date.today():
                    warns.append(f"a(z) {d.get('tipus') or 'okmány'} érvényessége lejárt ({v})")
            except ValueError:
                pass
    if not str(p.get("nev") or "").strip():
        warns.append("a név nem olvasható")
    return warns


def person_markdown(p, source):
    warns = person_checks(p)
    out = [f"# {SHEET_PREFIX}{_v(p.get('nev'))}", f"Forrás: {source}", ""]
    if warns:
        out += ["## ⚠️ Ellenőrizendő"] + [f"- {w}" for w in warns] + [""]
    out += ["## Személyes adatok", "| Adat | Érték |", "|---|---|",
            f"| Név | {_v(p.get('nev'))} |", f"| Születési név | {_v(p.get('szuletesi_nev'))} |",
            f"| Anyja neve | {_v(p.get('anyja_neve'))} |",
            f"| Születési hely, idő | {_v(p.get('szuletesi_hely'))}, {_v(p.get('szuletesi_ido'))} |",
            f"| Állampolgárság | {_v(p.get('allampolgarsag'))} |", f"| Lakcím | {_v(p.get('lakcim'))} |"]
    if p.get("tartozkodasi_hely"):
        out.append(f"| Tartózkodási hely | {_v(p.get('tartozkodasi_hely'))} |")
    docs = p.get("okmanyok") or []
    if docs:
        out += ["", "## Okmányok", "| Okmány | Szám | Érvényes |", "|---|---|---|"] + [
            f"| {_v(d.get('tipus'))} | {_v(d.get('szam'))} | {_v(d.get('ervenyes'))} |" for d in docs]
    if p.get("megjegyzes"):
        out += ["", "## Megjegyzés", str(p["megjegyzes"])]
    out += ["", "_Az okmányból Sara olvasta ki; iratba vétel előtt az ⚠️ pontokat az eredetin ellenőrizd. A személyi "
                "azonosító szándékosan nem szerepel._"]
    return "\n".join(out)


def person_sheet(client, model, files, source, log=None):
    content = _content(files)
    if not content:
        raise ValueError("nincs feldolgozható okmánykép (képként vagy PDF-ként töltsd fel)")
    data = _ask(client, model, PERSON_SYSTEM, content, log)
    return data, person_markdown(data, source)


# --- jármű -----------------------------------------------------------------------------------------------
def vehicle_checks(v):
    warns = []
    vin = re.sub(r"\s", "", str(v.get("alvazszam") or "")).upper()
    if not vin:
        warns.append("az alvázszám nem olvasható")
    elif not VIN_RE.match(vin):
        warns.append(f"az alvázszám („{vin}”) nem szabályos: 17 karakter kell, I, O és Q betű nélkül")
    plate = str(v.get("rendszam") or "").strip().upper()
    if plate and not PLATE_RE.match(plate):
        warns.append(f"a rendszám („{plate}”) szokatlan alakú")
    for key, label in (("elso_nyilvantartasba_vetel", "első nyilvántartásba vétel"),
                       ("muszaki_ervenyes", "műszaki érvényesség")):
        s = str(v.get(key) or "").strip()
        if s and not DATE_RE.match(s):
            warns.append(f"a(z) {label} („{s}”) szokatlan alakú")
    s = str(v.get("muszaki_ervenyes") or "").strip()
    if s and DATE_RE.match(s):
        from datetime import date
        y, mo, da = (int(x) for x in s.rstrip(".").split("."))
        try:
            if date(y, mo, da) < date.today():
                warns.append(f"a műszaki érvényesség lejárt ({s})")
        except ValueError:
            pass
    return warns


def vehicle_markdown(v, source):
    warns = vehicle_checks(v)
    title = " ".join(x for x in (_v(v.get("rendszam")), _v(v.get("gyartmany")), _v(v.get("tipus"))) if x != "–")
    out = [f"# {SHEET_PREFIX}{title or 'jármű'}", f"Forrás: {source}", ""]
    if warns:
        out += ["## ⚠️ Ellenőrizendő"] + [f"- {w}" for w in warns] + [""]
    rows = [("Rendszám", "rendszam"), ("Alvázszám", "alvazszam"), ("Gyártmány", "gyartmany"), ("Típus", "tipus"),
            ("Kereskedelmi név", "kereskedelmi_nev"), ("Első nyilvántartásba vétel", "elso_nyilvantartasba_vetel"),
            ("Gyártási év", "gyartasi_ev"), ("Üzembentartó", "uzembentarto"), ("Tulajdonos", "tulajdonos"),
            ("Szín", "szin"), ("Hengerűrtartalom", "hengerurtartalom"), ("Teljesítmény", "teljesitmeny"),
            ("Hajtóanyag", "hajtoanyag"), ("Ülések", "ulesek"), ("Megengedett össztömeg", "ossztomeg"),
            ("Műszaki érvényesség", "muszaki_ervenyes"), ("Forgalmi engedély száma", "forgalmi_szama"),
            ("Törzskönyv száma", "torzskonyv_szama")]
    out += ["## Járműadatok", "| Adat | Érték |", "|---|---|"] + [f"| {label} | {_v(v.get(key))} |" for label, key in rows]
    if v.get("megjegyzes"):
        out += ["", "## Megjegyzés", str(v["megjegyzes"])]
    out += ["", "_A forgalmi engedélyből Sara olvasta ki; iratba vétel előtt az ⚠️ pontokat az eredetin ellenőrizd._"]
    return "\n".join(out)


def vehicle_sheet(client, model, files, source, log=None):
    content = _content(files)
    if not content:
        raise ValueError("nincs feldolgozható kép vagy PDF (a forgalmi engedélyt képként vagy PDF-ként töltsd fel)")
    data = _ask(client, model, VEHICLE_SYSTEM, content, log)
    return data, vehicle_markdown(data, source)
