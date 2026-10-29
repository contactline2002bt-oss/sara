"""Tulajdoni lap, cégkivonat és civil szervezeti (alapítványi) kivonat elemzése: a TakarNetről, az e-cégjegyzékből vagy a cégbíróságtól letöltött irat
(PDF vagy .es3-ból kibontott PDF) felismerése, a lényeges adatok strukturált adatlapba gyűjtése, és a kockázatok
(terhek, széljegyek, végrehajtás, felszámolás) kiemelése.

A modell (Haiku) gyűjti ki az adatokat; a program ellenőrzi őket a szövegben (helyrajzi szám, cégjegyzékszám,
adószám betű szerint), összeadja a tulajdoni hányadokat, és kulcsszavas kereséssel is keresi a kockázati
bejegyzéseket, hogy a modell tévedése esetén se maradjon ki egy jelzálog vagy egy felszámolás.
Streamlittől független, ezért tesztelhető.
"""
import json
import re
from datetime import date, datetime
from fractions import Fraction

MODEL = "claude-haiku-5-5"
MAX_CHARS = 120_000
FRESH_DAYS = 30   # ennél régebbi lekérésnél figyelmeztet, hogy szerződéshez érdemes frisset kérni

TOOL = {
    "name": "registry_sheet",
    "description": (
        "Tulajdoni lap vagy cégkivonat (cégmásolat) adatlapja: a dokumentumtárban lévő irat lényeges adatai táblázatban "
        "(ingatlannál település, helyrajzi szám, jelleg, terület, tulajdonosok és hányadaik, terhek, széljegyek; "
        "cégnél cégnév, cégjegyzékszám, adószám, székhely, képviselők, tagok, folyamatban lévő eljárások), a "
        "kockázatok kiemelésével és a lekérés frissességével. Feltöltéskor magától elkészül; ezzel a meglévő "
        "iratokra is lefuttatható."),
    "input_schema": {"type": "object", "properties": {
        "document": {"type": "string", "description": "A tulajdoni lap vagy cégkivonat neve a projektben vagy törzsaktában."}},
        "required": ["document"]},
}

LAND_SYSTEM = (
    "Magyar tulajdoni lapot (TakarNet / E-ING, hiteles vagy nem hiteles, teljes vagy szemle másolat) dolgozol fel. "
    "Gyűjtsd ki az adatokat pontosan úgy, ahogy a lapon állnak (neveket, számokat betű szerint). Ne találj ki "
    "semmit; ami nincs a lapon, az legyen null. A III. rész HATÁLYOS bejegyzéseit sorold fel a terhek közé "
    "(jelzálogjog, végrehajtási jog, elidegenítési és terhelési tilalom, haszonélvezet, szolgalom, vezeték- és "
    "használati jog, elővásárlási jog stb.), a jogosulttal és az összeggel együtt. A már TÖRÖLT bejegyzéseket "
    "(„törölve”, „törlés”, áthúzott vagy törlő határozattal megszüntetett) ne a terhek közé, hanem a torolt_terhek "
    "közé tedd. A II. részben is csak a jelenlegi tulajdonosokat add meg. A széljegyeket (folyamatban lévő, még el nem "
    "intézett beadványok) külön sorold fel. Csak egy JSON-objektumot adj vissza:\n"
    '{"lekeres_ideje": "ÉÉÉÉ.HH.NN. vagy null", "hiteles": true/false/null, "telepules": "", "helyrajzi_szam": "", '
    '"cim": "", "jelleg": "pl. lakóház, udvar / társasházi lakás", "terulet": "m2 vagy ha szövege", '
    '"tulajdonosok": [{"nev": "", "hanyad": "pl. 1/2", "jogcim": "", "bejegyzo_hatarozat": ""}], '
    '"terhek": [{"tipus": "", "jogosult": "", "osszeg": "", "bejegyzes": "sorszám, határozat"}], '
    '"torolt_terhek": [{"tipus": "", "torles": "a törlő határozat"}], '
    '"szeljegyek": [{"leiras": "", "erkezes": ""}], "megjegyzes": "a II. és III. rész egyéb fontos tényei"}')

COMPANY_SYSTEM = (
    "Magyar cégkivonatot vagy cégmásolatot (e-cégjegyzék / Céginformációs Szolgálat) dolgozol fel. Gyűjtsd ki "
    "az adatokat pontosan úgy, ahogy az iratban állnak (neveket, számokat betű szerint), mindig a HATÁLYOS "
    "(nem törölt) bejegyzést véve; a már törölt vagy lezárult eljárást ne tedd az eljárások közé. A "
    "lekeres_ideje az a nap, amelyre az irat az adatokat igazolja (pl. „…napján hatályos adatok”, a lekérés vagy "
    "kiállítás dátuma); az utolso_valtozas a legutóbbi cégjegyzék-bejegyzés (változásbejegyzés) dátuma. Ne találj ki semmit; ami nincs benne, az legyen null. Csak egy JSON-objektumot "
    "adj vissza:\n"
    '{"lekeres_ideje": "ÉÉÉÉ.HH.NN. vagy null", "utolso_valtozas": "ÉÉÉÉ.HH.NN. vagy null", "cegnev": "", "rovid_nev": "", "cegjegyzekszam": "", '
    '"adoszam": "", "szekhely": "", "cegforma": "", "jegyzett_toke": "", '
    '"kepviselok": [{"nev": "", "tisztseg": "pl. ügyvezető", "kepviselet_modja": "önálló / együttes", '
    '"lakcim": "", "jogviszony_kezdete": ""}], "tagok": [{"nev": "", "torzsbetet_vagy_reszesedes": ""}], '
    '"eljarasok": ["folyamatban lévő felszámolás, végelszámolás, csődeljárás, kényszertörlés, végrehajtás, '
    'törvényességi felügyeleti eljárás – a bejegyzés szövegével"], "megjegyzes": ""}')

CIVIL_SYSTEM = (
    "Magyar civil szervezet (alapítvány, egyesület) bírósági nyilvántartási kivonatát (civilsznev.birosag.hu) "
    "dolgozod fel. Gyűjtsd ki az adatokat pontosan úgy, ahogy az iratban állnak (neveket, számokat betű szerint); "
    "csak a hatályos (meg nem szűnt) képviselőket és tagokat vedd fel. A lekeres_ideje a lekérdezés napja (az oldal "
    "tetején álló dátum); az utolso_valtozas a legutóbbi bejegyzés dátuma. A tagok a kezelő szerv (pl. kuratórium) "
    "tagjai, a tisztségükkel (pl. elnök, ha az irat így jelöli). Ne találj ki semmit; ami nincs benne, az legyen null. "
    "Csak egy JSON-objektumot adj vissza:\n"
    '{"lekeres_ideje": "ÉÉÉÉ.HH.NN. vagy null", "utolso_valtozas": "ÉÉÉÉ.HH.NN. vagy null", "cegnev": "a megnevezés", '
    '"rovid_nev": "", "cegjegyzekszam": "a nyilvántartási szám", "adoszam": "", "szekhely": "", '
    '"cegforma": "pl. vagyonkezelő alapítvány (zárt)", "jegyzett_toke": "az alapítványi vagyonra vonatkozó szabály, '
    'pl. a vagyon nem csökkenhet X Ft alá", "allapot": "pl. Nyilvántartásba vett", "ugyszam": "", '
    '"kepviselok": [{"nev": "", "tisztseg": "pl. képviselő", "kepviselet_modja": "önálló / együttes", '
    '"lakcim": "", "jogviszony_kezdete": ""}], "tagok": [{"nev": "", "torzsbetet_vagy_reszesedes": '
    '"a kezelő szerv és a tisztség, pl. kuratórium – elnök"}], "eljarasok": ["folyamatban lévő eljárás, '
    'biztosítási vagy végrehajtási intézkedés – csak ha VAN"], "megjegyzes": "a cél rövid összefoglalása és más '
    'fontos tény"}')

LAND_RISK = [   # (minta, megnevezés) – a szövegben kulcsszóval is keressük
    (r"jelzálogjog", "jelzálogjog"), (r"végrehajtási jog", "végrehajtási jog"),
    (r"elidegenítési (?:és terhelési )?tilalom", "elidegenítési/terhelési tilalom"),
    (r"terhelési tilalom", "terhelési tilalom"), (r"haszonélvezeti jog", "haszonélvezeti jog"),
    (r"zár(?:lat)?\b|zárlat", "zárlat"), (r"perfeljegyzés|per feljegyzés", "per feljegyzése"),
    (r"széljegy", "széljegy (folyamatban lévő beadvány)"), (r"vételi jog", "vételi jog"),
    (r"elővásárlási jog", "elővásárlási jog"), (r"bírósági végrehajtás", "bírósági végrehajtás"),
]
CIVIL_RISK = [
    (r"felszámolás", "felszámolási eljárás"), (r"végelszámolás", "végelszámolás"),
    (r"kényszertörlés", "kényszertörlési eljárás"), (r"törvényességi felügyeleti", "törvényességi felügyeleti eljárás"),
]
COMPANY_RISK = [
    (r"felszámolás", "felszámolási eljárás"), (r"végelszámolás", "végelszámolás"),
    (r"csődeljárás", "csődeljárás"), (r"kényszertörlés", "kényszertörlési eljárás"),
    (r"végrehajtás", "végrehajtás"), (r"törvényességi felügyeleti", "törvényességi felügyeleti eljárás"),
    (r"adószám(?:ot)? felfüggeszt|adószám törl", "adószám felfüggesztése/törlése"),
    (r"megszűn", "megszűnés"),
]


def detect(text):
    """„land” (tulajdoni lap), „company” (cégkivonat/cégmásolat) vagy None."""
    t = str(text or "")[:20000].lower()
    land = sum(k in t for k in ("tulajdoni lap", "helyrajzi szám", "i. rész", "ii. rész", "iii. rész",
                                "földhivatal", "e-ing", "takarnet", "művelési ág"))
    company = sum(k in t for k in ("cégjegyzékszám", "cégkivonat", "cégmásolat", "a cég elnevezése",
                                   "a cég székhelye", "cégbíróság", "képviseletre jogosult", "jegyzett tőke"))
    civil = sum(k in t for k in ("civilsznev", "civil szervezet", "szervezet nyilvántartási száma",
                                 "a szervezet képviselői", "szervezet típusa", "ügyvezető szerv", "alapítvány"))
    if land >= 3 and land >= company:
        return "land"
    if civil >= 3 and civil >= company:
        return "civil"
    if company >= 3:
        return "company"
    return None


def _text_of(response):
    return "".join(getattr(b, "text", "") for b in getattr(response, "content", [])
                   if getattr(b, "type", "") == "text").strip()


def extract(client, kind, text, log=None, model=MODEL, pdf=None):
    system = {"land": LAND_SYSTEM, "civil": CIVIL_SYSTEM}.get(kind, COMPANY_SYSTEM)
    content = str(text)[:MAX_CHARS]
    if pdf and "\ufffd" in content and len(pdf) < 20 * 1024 * 1024:
        # a PDF egyes betűi (ő, ű) kódolás nélküliek: a modell a PDF képéről olvassa a helyes betűt
        import base64
        content = [{"type": "document", "source": {"type": "base64", "media_type": "application/pdf",
                                                    "data": base64.b64encode(pdf).decode("ascii")}},
                   {"type": "text", "text": "A fenti PDF kinyert szövege következik; ahol „\ufffd” jel áll benne, "
                                            "a helyes betűt (pl. ő, ű) a PDF képéről olvasd:\n\n" + content}]
    kwargs = dict(model=model, max_tokens=4000, system=system,
                  messages=[{"role": "user", "content": content}])
    try:
        response = client.messages.create(thinking={"type": "disabled"}, **kwargs)
    except Exception as e:
        if getattr(e, "status_code", None) != 400 or "thinking" not in str(e).lower():
            raise
        response = client.messages.create(**kwargs)
    if log:
        log(model, getattr(response, "usage", None))
    m = re.search(r"\{.*\}", _text_of(response), re.S)
    if not m:
        raise ValueError("a modell nem adott értelmezhető adatlapot")
    return json.loads(m.group(0))


def _norm(s):
    return re.sub(r"\s+", "", str(s or "")).lower()


def _found(value, text):
    return bool(value) and _norm(value) in _norm(text)


def _age_days(when):
    m = re.search(r"(\d{4})\.\s?(\d{1,2})\.\s?(\d{1,2})", str(when or ""))
    if not m:
        return None
    try:
        return (date.today() - date(int(m.group(1)), int(m.group(2)), int(m.group(3)))).days
    except ValueError:
        return None


def _share(text):
    m = re.search(r"(\d+)\s*/\s*(\d+)", str(text or ""))
    if not m or int(m.group(2)) == 0:
        return None
    return Fraction(int(m.group(1)), int(m.group(2)))


def checks(kind, data, text):
    """A program saját ellenőrzései: (kockázatok, figyelmeztetések)."""
    risks, warns = [], []
    low = str(text or "").lower()
    # a „nincs” értékű rovatok ne adjanak hamis riasztást (pl. „Széljegy: nincs”, „III. rész: –”)
    low = re.sub(r"(széljegy\w*|teher|terhek|végrehajtás\w*|felszámolás\w*|eljárás\w*)\s*:?\s*(nincs|-+|–)", " ", low)
    low = re.sub(r"(nincs|nem áll)\s+(folyamatban\s+)?(széljegy\w*|felszámolás\w*|végelszámolás\w*|"
                 r"csődeljárás\w*|végrehajtás\w*|kényszertörlés\w*)", " ", low)
    deleted = re.compile(r"törl|töröl|megszűnt|megszüntet|befejez|lezár", re.I)
    model_has = " ".join(json.dumps(data.get(k) or [], ensure_ascii=False).lower()
                         for k in (("terhek", "szeljegyek") if kind == "land" else ("eljarasok",)))
    for pattern, label in {"land": LAND_RISK, "civil": CIVIL_RISK}.get(kind, COMPANY_RISK):
        hits = [m for m in re.finditer(pattern, low)]
        if not hits:
            continue
        # a kulcsszó csak akkor kockázat, ha van olyan előfordulása, amely körül nincs törlésre utaló szó
        live = [m for m in hits if not deleted.search(low[max(0, m.start() - 250):m.end() + 250])]
        if re.search(pattern, model_has):
            risks.append(label)                 # a modell is hatályosnak látja
        elif live:
            warns.append(f"a(z) „{label}” kifejezés szerepel az iratban, de nem találtam hatályos bejegyzést – "
                         "nézd meg, nem törölt-e")
    if kind == "land":
        age = _age_days(data.get("lekeres_ideje"))
        if age is None:
            warns.append("a lekérés dátuma nem állapítható meg – nézd meg, mennyire friss a tulajdoni lap")
        elif age > FRESH_DAYS:
            warns.append(f"a tulajdoni lap {age} napos; szerződéshez érdemes friss példányt lekérni")
    elif data.get("utolso_valtozas") and data.get("lekeres_ideje"):
        ch, req = _age_days(data["utolso_valtozas"]), _age_days(data["lekeres_ideje"])
        if ch is not None and req is not None and ch < req:
            warns.append("a lekérés óta változásbejegyzés történt – friss cégkivonat kell")
    if kind == "land":
        hrsz = data.get("helyrajzi_szam")
        if hrsz and not _found(hrsz, text):
            warns.append(f"a helyrajzi szám („{hrsz}”) betű szerint nem található a szövegben – ellenőrizd")
        shares = [_share(o.get("hanyad")) for o in data.get("tulajdonosok") or []]
        if shares and all(s is not None for s in shares):
            total = sum(shares)
            if total != 1:
                warns.append(f"a tulajdoni hányadok összege {total}, nem 1/1 – lehet, hogy egy tulajdonos "
                             "kimaradt, vagy közös tulajdon van más bontásban")
        if data.get("hiteles") is False:
            warns.append("nem hiteles tulajdoni lap: tájékoztatásra jó, földhivatali eljáráshoz hiteles kell")
    else:
        if kind == "civil" and data.get("allapot") and "nyilvántartásba vett" not in str(data["allapot"]).lower():
            risks.append(f"a szervezet állapota: {data['allapot']}")
        for key, label, pattern in (("cegjegyzekszam", "nyilvántartási szám" if kind == "civil" else "cégjegyzékszám",
                                     r"^\d{2}-\d{2}-\d{6,7}$"),
                                    ("adoszam", "adószám", r"^\d{8}-\d-\d{2}$")):
            v = str(data.get(key) or "").strip()
            if v and not _found(v, text):
                warns.append(f"a(z) {label} („{v}”) betű szerint nem található a szövegben – ellenőrizd")
            elif v and not re.match(pattern, v):
                warns.append(f"a(z) {label} („{v}”) szokatlan alakú")
    return list(dict.fromkeys(risks)), warns


def _cell(v):
    return str(v or "–").replace("|", "/").replace("\n", " ")


def to_markdown(kind, data, risks, warns, source):
    out = []
    if kind == "land":
        out += [f"# Tulajdoni lap adatlap – {_cell(data.get('telepules'))} {_cell(data.get('helyrajzi_szam'))} hrsz.",
                f"Forrás: {source} · lekérés: {_cell(data.get('lekeres_ideje'))} · "
                + ("hiteles" if data.get("hiteles") else "nem hiteles" if data.get("hiteles") is False else "hitelesség: ?"),
                ""]
        if risks or warns:
            out += ["## ⚠️ Figyelem"] + [f"- **{r}** szerepel a lapon" for r in risks] + [f"- {w}" for w in warns] + [""]
        out += ["## Ingatlan", "| Adat | Érték |", "|---|---|",
                f"| Település | {_cell(data.get('telepules'))} |", f"| Helyrajzi szám | {_cell(data.get('helyrajzi_szam'))} |",
                f"| Cím | {_cell(data.get('cim'))} |", f"| Jelleg | {_cell(data.get('jelleg'))} |",
                f"| Terület | {_cell(data.get('terulet'))} |", "", "## Tulajdonosok (II. rész)"]
        owners = data.get("tulajdonosok") or []
        out += (["| Név | Hányad | Jogcím | Bejegyző határozat |", "|---|---|---|---|"]
                + [f"| {_cell(o.get('nev'))} | {_cell(o.get('hanyad'))} | {_cell(o.get('jogcim'))} | "
                   f"{_cell(o.get('bejegyzo_hatarozat'))} |" for o in owners]) if owners else ["Nincs adat."]
        out += ["", "## Terhek (III. rész)"]
        burdens = data.get("terhek") or []
        out += (["| Teher | Jogosult | Összeg | Bejegyzés |", "|---|---|---|---|"]
                + [f"| {_cell(b.get('tipus'))} | {_cell(b.get('jogosult'))} | {_cell(b.get('osszeg'))} | "
                   f"{_cell(b.get('bejegyzes'))} |" for b in burdens]) if burdens else ["A III. részben nincs bejegyzés (tehermentes)."]
        gone = data.get("torolt_terhek") or []
        if gone:
            out += ["", f"_Törölt (már nem hatályos) bejegyzések: {len(gone)} – "
                        + "; ".join(_cell(g.get("tipus")) for g in gone) + "_"]
        out += ["", "## Széljegyek"]
        marg = data.get("szeljegyek") or []
        out += [f"- {_cell(s.get('leiras'))} ({_cell(s.get('erkezes'))})" for s in marg] if marg else ["Nincs széljegy."]
    else:
        civil = kind == "civil"
        out += [f"# {'Szervezeti adatlap' if civil else 'Cégadatlap'} – {_cell(data.get('cegnev'))}",
                f"Forrás: {source} · az adatok napja: {_cell(data.get('lekeres_ideje'))}"
                + (f" · utolsó változásbejegyzés: {_cell(data.get('utolso_valtozas'))}" if data.get("utolso_valtozas") else ""),
                ""]
        if risks or warns:
            out += ["## ⚠️ Figyelem"] + [f"- **{r}** szerepel az iratban – nézd meg, hatályos-e" for r in risks] \
                   + [f"- {w}" for w in warns] + [""]
        out += ["## Alapadatok", "| Adat | Érték |", "|---|---|",
                f"| {'Megnevezés' if civil else 'Cégnév'} | {_cell(data.get('cegnev'))} |",
                f"| Rövid név | {_cell(data.get('rovid_nev'))} |",
                f"| {'Nyilvántartási szám' if civil else 'Cégjegyzékszám'} | {_cell(data.get('cegjegyzekszam'))} |",
                f"| Adószám | {_cell(data.get('adoszam'))} |",
                f"| Székhely | {_cell(data.get('szekhely'))} |",
                f"| {'Típus' if civil else 'Cégforma'} | {_cell(data.get('cegforma'))} |",
                f"| {'Vagyon' if civil else 'Jegyzett tőke'} | {_cell(data.get('jegyzett_toke'))} |"]
        if civil:
            out += [f"| Állapot | {_cell(data.get('allapot'))} |", f"| Ügyszám | {_cell(data.get('ugyszam'))} |"]
        out += ["", "## Képviseletre jogosultak"]
        reps = data.get("kepviselok") or []
        out += (["| Név | Tisztség | Képviselet módja | Kezdete |", "|---|---|---|---|"]
                + [f"| {_cell(r.get('nev'))} | {_cell(r.get('tisztseg'))} | {_cell(r.get('kepviselet_modja'))} | "
                   f"{_cell(r.get('jogviszony_kezdete'))} |" for r in reps]) if reps else ["Nincs adat."]
        members = data.get("tagok") or []
        if members:
            out += ["", "## Kezelő szerv (kuratórium) tagjai" if civil else "## Tagok",
                    "| Név | Tisztség |" if civil else "| Név | Részesedés |", "|---|---|"] + [
                f"| {_cell(m.get('nev'))} | {_cell(m.get('torzsbetet_vagy_reszesedes'))} |" for m in members]
        procs = data.get("eljarasok") or []
        out += ["", "## Folyamatban lévő eljárások"] + ([f"- {_cell(p)}" for p in procs] if procs else ["Nincs."])
    if data.get("megjegyzes"):
        out += ["", "## Megjegyzés", str(data["megjegyzes"])]
    out += ["", "_Az adatlapot Sara készítette az iratból; fontos iratba vétel előtt az eredetivel "
                "egyeztesd. A kulcsszavas kockázatjelzés óvatos: a törölt bejegyzéseket is jelezheti._"]
    return "\n".join(out)


def sheet(client, text, source, log=None, model=MODEL, pdf=None):
    """Visszaad: (fajta, markdown, kockázatok, figyelmeztetések) vagy None, ha nem tulajdoni lap / cégkivonat."""
    kind = detect(text)
    if not kind:
        return None
    data = extract(client, kind, text, log, model, pdf)
    if "\ufffd" in json.dumps(data, ensure_ascii=False):   # ha mégis maradt, legalább jelezzük
        data["megjegyzes"] = ((data.get("megjegyzes") or "") + " Egyes ékezetes betűk (ő, ű) a PDF-ben "
                              "kódolás nélküliek, ezért „�” jelként maradtak; a neveket ellenőrizd.").strip()
    risks, warns = checks(kind, data, text)
    return kind, to_markdown(kind, data, risks, warns, source), risks, warns
