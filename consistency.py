"""Sara ellenőrző kollégája: több dokumentum egymásnak ellentmondó adatainak felderítése (szinkronizálás).

Tipikus eset: egy projekt dokumentációja (rendszerterv, biztonsági terv, szerződés, ajánlat, jegyzőkönyv) több
helyen ugyanazt az adatot tartalmazza – verziószám, dátum, határidő, összeg, felelős, cím, IP-cím, darabszám –,
és ezek idővel elcsúsznak egymástól.

Olcsó modell (Haiku), két lépésben:
1. kiolvasás, dokumentumonként párhuzamosan: a dokumentum adatpontjai szó szerinti idézettel;
   a program ellenőrzi, hogy az idézet tényleg benne van-e a dokumentumban (a kitalált adat kiesik);
2. összevetés: az azonos jelentésű adatpontok csoportosítása, és csak az eltérések visszaadása, javaslattal
   (ha van mértékadó dokumentum, az nyer; különben a felhasználó dönt).
A javítás nem itt történik: Sara a jóváhagyás után az edit_word_document eszközzel, követett változással végzi.
Streamlittől független, ezért tesztelhető.
"""
import json
import re
from concurrent.futures import ThreadPoolExecutor

CHECK_MODEL = "claude-haiku-5-5"
MAX_DOCS = 15            # egy ellenőrzésben legfeljebb ennyi dokumentum
MAX_DOC_CHARS = 120_000  # egy dokumentumból ennyi szöveget kap a kiolvasó
MAX_PARALLEL = 5         # egyszerre ennyi kiolvasás fut
MAX_FACTS = 120          # dokumentumonként legfeljebb ennyi adatpont

TOOL = {
    "name": "consistency_check",
    "description": (
        "Az ellenőrző kolléga: több dokumentum (pl. egy projekt dokumentációja) ugyanarra vonatkozó adatait veti "
        "össze, és kigyűjti az eltéréseket: verziószám, dátum, határidő, összeg, darabszám, felelős, név, cím, "
        "azonosító, műszaki paraméter. Mindegyiknél megmutatja, melyik dokumentumban mi áll (szó szerint), és "
        "javaslatot tesz az egységes értékre. Akkor használd, ha a felhasználó dokumentumok szinkronba hozását, "
        "egyeztetését, ellentmondások keresését kéri. A javítást csak jóváhagyás után végezd."),
    "input_schema": {"type": "object", "properties": {
        "documents": {"type": "array", "items": {"type": "string"}, "description": (
            "Az összevetendő dokumentumok nevei (a kijelölt munkaanyagok vagy a projekt dokumentumai közül). Ha "
            "elhagyod, a projekt összes dokumentuma (legfeljebb 15).")},
        "focus": {"type": "string", "description": (
            "Mire figyeljen elsősorban, pl. „verziószámok, határidők és felelősök” vagy „IP-címek és "
            "szervernevek”. Elhagyható: ilyenkor minden konkrét adatot összevet.")},
        "reference": {"type": "string", "description": (
            "A mértékadó dokumentum neve, ha van (pl. a legfrissebb, jóváhagyott terv): eltérésnél ennek az "
            "értéke a javaslat.")}}},
}

EXTRACT_SYSTEM = (
    "Egy dokumentum konkrét, ellenőrizhető adatait gyűjtöd ki, hogy később más dokumentumokkal össze lehessen "
    "vetni őket: verziószám, dátum, határidő, időtartam, összeg, díj, darabszám, mennyiség, százalék, személy- és "
    "cégnév, szerepkör és felelős, cím, azonosító (szerződésszám, cégjegyzékszám, adószám), elérhetőség, műszaki "
    "paraméter (szervernév, IP-cím, kapacitás, verzió, port), helyszín. Általános szöveget, véleményt ne gyűjts. "
    "Minden adatponthoz: „adat” = rövid, általános megnevezés, amely más dokumentumban is ugyanígy hívható (pl. "
    "„projekt befejezési határideje”, „adatbázis-szerver IP-címe”, „szerződés teljes díja”, „projektvezető”), "
    "„ertek” = az érték pontosan, „idezet” = a dokumentumból SZÓ SZERINT kimásolt rövid részlet (legfeljebb 200 "
    "karakter), amely az értéket tartalmazza, „hely” = fejezet/pont/táblázat, ha kiderül. "
    "Csak JSON: {\"adatok\": [{\"adat\": \"…\", \"ertek\": \"…\", \"idezet\": \"…\", \"hely\": \"…\"}]}")

COMPARE_SYSTEM = (
    "Több dokumentum kigyűjtött adatpontjait kapod meg. Csoportosítsd az azonos jelentésű adatpontokat (a "
    "megnevezésük eltérhet: „befejezési határidő” = „projekt vége” = „átadás napja”, ha ugyanarra vonatkozik), "
    "és keresd meg, ahol a dokumentumok ELLENTMONDANAK egymásnak. Ne jelezd eltérésnek, ami csak írásmód "
    "(2026.10.01. = 2026. október 1.; 1 000 000 Ft = 1.000.000 Ft), és azt sem, ami szándékosan más (pl. két "
    "különböző szerver IP-címe, egy terv két változatának jelzett különbsége). Csak a valódi eltéréseket add "
    "vissza, súlyosság szerint (KRITIKUS: összeg, határidő, felelős, azonosító, műszaki paraméter; ELLENŐRIZENDŐ: "
    "minden más). Mindegyikhez javasolj egységes értéket: ha van mértékadó dokumentum, annak értékét; különben, "
    "ha egyértelmű (pl. a többség és a legfrissebb dokumentum egyezik), azt, indoklással; ha nem dönthető el, "
    "„javaslat” maradjon üres, és az indoklásban írd meg, mit kell eldönteni. Ha egy adat csak egy "
    "dokumentumban szerepel, az nem eltérés. "
    "Csak JSON: {\"elteresek\": [{\"adat\": \"…\", \"sulyossag\": \"KRITIKUS|ELLENŐRIZENDŐ\", \"ertekek\": "
    "[{\"dokumentum\": \"…\", \"ertek\": \"…\", \"idezet\": \"…\", \"hely\": \"…\"}], \"javaslat\": \"…\", "
    "\"indoklas\": \"…\"}]}")


def _text_of(response):
    return "".join(getattr(b, "text", "") for b in getattr(response, "content", [])
                   if getattr(b, "type", "") == "text").strip()


def _ask(client, system, user, max_tokens, log):
    kwargs = dict(model=CHECK_MODEL, max_tokens=max_tokens, system=system,
                  messages=[{"role": "user", "content": user}])
    try:
        response = client.messages.create(thinking={"type": "disabled"}, **kwargs)
    except Exception as e:
        if getattr(e, "status_code", None) != 400 or "thinking" not in str(e).lower():
            raise
        response = client.messages.create(**kwargs)
    if log:
        log(CHECK_MODEL, getattr(response, "usage", None))
    return _text_of(response)


def _json(text):
    text = re.sub(r"^```(?:json)?|```$", "", str(text or "").strip(), flags=re.MULTILINE).strip()
    try:
        return json.loads(text)
    except ValueError:
        m = re.search(r"\{.*\}", text, flags=re.DOTALL)
        if m:
            try:
                return json.loads(m.group(0))
            except ValueError:
                return None
    return None


def _squash(text):
    """Összevetéshez: szóközök, kis-nagybetű és tipográfiai jelek nélkül."""
    t = str(text or "").lower().replace(" ", " ")
    t = t.replace("„", '"').replace("”", '"').replace("’", "'").replace("–", "-").replace("—", "-")
    return re.sub(r"\s+", "", t)


def grounded(fact, doc_text_squashed):
    """Az adatpont akkor marad, ha az idézete (és benne az érték) tényleg szerepel a dokumentumban."""
    quote = _squash(fact.get("idezet"))
    value = _squash(fact.get("ertek"))
    if not value:
        return False
    if quote and len(quote) >= 4 and quote in doc_text_squashed:
        return True
    return len(value) >= 3 and value in doc_text_squashed


def extract(client, name, text, focus="", log=None):
    """Egy dokumentum adatpontjai (csak a dokumentumban ténylegesen szereplők). Visszaad: (adatpontok, kiszűrt db)."""
    text = str(text or "")
    user = ((f"Elsősorban ezekre figyelj: {focus}\n\n" if focus else "")
            + f"A DOKUMENTUM ({name}):\n{text[:MAX_DOC_CHARS]}")
    data = _json(_ask(client, EXTRACT_SYSTEM, user, 8000, log)) or {}
    facts = data.get("adatok") if isinstance(data, dict) else data
    facts = [f for f in (facts or []) if isinstance(f, dict)][:MAX_FACTS]
    squashed = _squash(text)
    kept = [f for f in facts if grounded(f, squashed)]
    return kept, len(facts) - len(kept)


def check(client, docs, focus="", reference="", log=None):
    """docs: [(név, szöveg)]. Visszaad: Sarának szóló jelentés (szöveg)."""
    docs = [(n, t) for n, t in docs if str(t or "").strip()][:MAX_DOCS]
    if len(docs) < 2:
        return ("Legalább két olvasható szövegű dokumentum kell az összevetéshez. Jelöld ki őket, vagy add meg a "
                "nevüket.")
    with ThreadPoolExecutor(max_workers=min(MAX_PARALLEL, len(docs))) as pool:
        results = list(pool.map(lambda d: extract(client, d[0], d[1], focus, log), docs))
    lines, dropped = [], 0
    for (name, _), (facts, lost) in zip(docs, results):
        dropped += lost
        if not facts:
            continue
        lines.append(f"### {name}" + (" [MÉRTÉKADÓ]" if reference and name == reference else ""))
        for f in facts:
            lines.append(json.dumps({k: f.get(k, "") for k in ("adat", "ertek", "idezet", "hely")},
                                    ensure_ascii=False))
    if not lines:
        return "Az ellenőrző kolléga egyik dokumentumban sem talált összevethető konkrét adatot."
    user = ((f"Mértékadó dokumentum: {reference}\n" if reference else "Nincs megjelölt mértékadó dokumentum.\n")
            + (f"A felhasználót elsősorban ez érdekli: {focus}\n" if focus else "")
            + "\nAZ ADATPONTOK DOKUMENTUMONKÉNT:\n" + "\n".join(lines))
    data = _json(_ask(client, COMPARE_SYSTEM, user, 8000, log)) or {}
    diffs = data.get("elteresek") if isinstance(data, dict) else data
    diffs = [d for d in (diffs or []) if isinstance(d, dict) and len(d.get("ertekek") or []) >= 2]
    diffs.sort(key=lambda d: str(d.get("sulyossag", "")).upper() != "KRITIKUS")
    return report(docs, diffs, dropped, reference)


def report(docs, diffs, dropped=0, reference=""):
    out = ["ELLENŐRZŐ JELENTÉS – dokumentumok összhangja",
           "Összevetett dokumentumok: " + ", ".join(n for n, _ in docs)
           + (f" (mértékadó: {reference})" if reference else ""), ""]
    if not diffs:
        out.append("Nincs eltérés: a közös adatok minden dokumentumban egyeznek.")
    for i, d in enumerate(diffs, 1):
        out.append(f"{i}. {d.get('adat', '?')} – {str(d.get('sulyossag') or 'ELLENŐRIZENDŐ').upper()}")
        for v in d.get("ertekek") or []:
            hely = f", {v['hely']}" if v.get("hely") else ""
            out.append(f"   - {v.get('dokumentum', '?')}{hely}: „{v.get('ertek', '')}”"
                       + (f" – „{v['idezet']}”" if v.get("idezet") and v.get("idezet") != v.get("ertek") else ""))
        if d.get("javaslat"):
            out.append(f"   → Javaslat: „{d['javaslat']}”" + (f" ({d['indoklas']})" if d.get("indoklas") else ""))
        else:
            out.append("   → Döntés kell" + (f": {d['indoklas']}" if d.get("indoklas") else ""))
    if dropped:
        out += ["", f"({dropped} adatpontot a kolléga nem tudott szó szerint visszaigazolni a dokumentumokból; "
                    "ezeket kihagytam.)"]
    out += ["", "A jelentést mutasd meg a felhasználónak táblázatosan vagy pontokba szedve. Javítani csak az ő "
                "jóváhagyása után javíts: .docx-nél dokumentumonként az edit_word_document eszközzel, követett "
                "változással (minden eltérő értéket a jóváhagyott értékre); más formátumnál új változatot készíts. "
                "Ha a felhasználó kéri, az eltéréslistát Excel-táblázatba is elkészítheted."]
    return "\n".join(out)
