"""Red Flag jelentés: egy szerződés- vagy okirat-tervezet átvizsgálása a felhasználó (vagy akit képvisel) szemszögéből.

Kigyűjti a megbízóra nézve előnytelen, egyoldalú vagy kockázatos kikötéseket (kötbér, elállás, felmondás,
felelősségkorlátozás, biztosítékok, fizetési feltételek, határidők, joglemondás, illetékesség, titoktartás,
versenytilalom stb.), és a hiányzó védelmeket. Minden megállapításhoz szó szerinti idézet tartozik, amelyet a
program ellenőriz: ha az idézet nem található meg betű szerint a tervezetben, a jelentés jelzi.
Streamlittől független, ezért tesztelhető.
"""
import json
import re

MAX_CHARS = 150_000

TOOL = {
    "name": "red_flag_review",
    "description": (
        "Red Flag jelentés (kockázatelemzés): egy szerződés- vagy okirat-tervezet (pl. az ellenérdekű fél által "
        "küldött tervezet) átvizsgálása a megbízó szemszögéből. Kigyűjti az előnytelen vagy egyoldalú "
        "kikötéseket (kötbér, elállási és felmondási jog, felelősségkorlátozás, biztosítékok, fizetési "
        "feltételek, határidők, joglemondás, illetékesség), és a hiányzó védelmeket; mindegyikhez szó szerinti "
        "idézetet, súlyosságot, indoklást és javaslatot ad. Akkor használd, ha a felhasználó egy tervezet átnézését, "
        "kockázatait, „red flag”-jeit kéri. Ha nem tudod, kinek az oldalán áll a felhasználó (melyik cége, milyen szerepben), előbb kérdezd meg."),
    "input_schema": {"type": "object", "properties": {
        "document": {"type": "string", "description": (
            "A tervezet neve a projekt dokumentumai vagy a beszélgetésbe csatolt fájlok közül.")},
        "client_role": {"type": "string", "description": (
            "Kinek az oldalán áll a felhasználó, kinek a szemszögéből kell vizsgálni (pl. „vevő”, „bérlő”, „alapítók”, "
            "„a Béta Kft. mint megrendelő”).")},
        "focus": {"type": "string", "description": "Nem kötelező: amire külön figyelni kell (pl. fizetési biztosíték)."},
        "context": {"type": "string", "description": (
            "Nem kötelező: az ügy háttere a beszélgetésből (ügylet tárgya, vételár, a felhasználó kérései).")}},
        "required": ["document", "client_role"]},
}

SYSTEM = (
    "Tapasztalt magyar ügyvéd vagy (társasági és ingatlanjog), aki egy szerződés- vagy okirat-tervezetet néz át "
    "a megbízója szemszögéből, a magyar jog (különösen a Ptk.) alapján. Feladatod a Red Flag jelentés: a "
    "megbízóra nézve előnytelen, egyoldalú, kiegyensúlyozatlan, homályos vagy jogszabályba ütköző kikötések, és "
    "a hiányzó, de szokásos védelmek összegyűjtése.\n"
    "Nézd végig legalább ezeket: fizetési feltételek és biztosítékok; kötbér, kártérítés, felelősségkorlátozás; "
    "elállás, felmondás, megszűnés és a vagyon sorsa megszűnéskor; határidők és késedelem; képviselet, "
    "aláírási és bankszámla-rendelkezési jog; testületi döntéshozatal (határozatképesség, szavazategyenlőség, "
    "döntő szavazat, minősített többség, vétójog); tagok, tisztségviselők kijelölése, visszahívása, "
    "összeférhetetlenség; ellenőrzési és tájékoztatási jogok; joglemondás, illetékesség, jogvita; "
    "titoktartás, versenytilalom; kitöltetlen vagy ellentmondó adatok (nevek, számozás).\n"
    "Szabályok:\n"
    "- Az idézet (quote) a tervezet szövegéből betű szerint másolt, összefüggő részlet legyen (legfeljebb 2–3 "
    "mondat). Ne foglald össze, ne javítsd ki, ne rövidítsd „…”-tal. Hiányzó kikötésnél az idézet üres.\n"
    "- A hivatkozás (ref) a pont vagy szakasz száma, ahogy a tervezetben áll (pl. „9.3.”).\n"
    "- Súlyosság: MAGAS (jelentős anyagi vagy jogvesztési kockázat, jogszabályba ütközés, érvénytelenség), "
    "KÖZEPES (kiegyensúlyozatlan, tárgyalni kell), ALACSONY (pontosítás, szövegezés).\n"
    "- Az indoklás konkrét legyen: miért hátrányos a megbízónak, milyen helyzetben üt vissza.\n"
    "- A javaslat egy mondatban a szokásos, kiegyensúlyozott megoldás vagy a javasolt szövegezés iránya.\n"
    "- Jogszabályhelyre csak akkor hivatkozz, ha biztos vagy benne; ha nem, írd, hogy ellenőrizendő.\n"
    "- Ne találj ki tényt, ami nincs a tervezetben. A kitöltetlen helyeket (… vagy [ ]) egy tételben jelezd.\n"
    "- Súlyosság szerint rendezve add meg a tételeket, a legsúlyosabbal kezdve.\n"
    "Csak egy JSON-objektumot adj vissza, más szöveget nem, ebben a szerkezetben:\n"
    '{"summary": "2–4 mondatos összkép", "findings": [{"ref": "", "title": "rövid cím", "quote": "", '
    '"severity": "MAGAS|KÖZEPES|ALACSONY", "issue": "miért kockázatos", "suggestion": "javaslat"}], '
    '"missing": [{"title": "hiányzó védelem", "issue": "miért kellene", "suggestion": "javaslat"}]}')

SEVERITY_ORDER = {"MAGAS": 0, "KÖZEPES": 1, "ALACSONY": 2}


def _norm(text):
    text = str(text or "").replace(" ", " ").replace("„", '"').replace("”", '"').replace("“", '"')
    text = text.replace("’", "'").replace("–", "-").replace("—", "-")
    return re.sub(r"\s+", " ", text).strip().lower()


def quote_found(quote, text):
    """Szó szerint szerepel-e az idézet (szóközöktől, idézőjel- és kötőjelfajtától függetlenül)."""
    q = _norm(quote).strip(' ."')
    return bool(q) and q in _norm(text)


def _parse_json(raw):
    raw = str(raw or "").strip()
    m = re.search(r"\{.*\}", raw, re.S)
    if not m:
        raise ValueError("a modell nem adott értelmezhető jelentést")
    return json.loads(m.group(0))


def _text_of(response):
    return "".join(getattr(b, "text", "") for b in getattr(response, "content", [])
                   if getattr(b, "type", "") == "text").strip()


def analyze(client, model, doc_name, text, client_role, focus="", context="", log=None):
    """A modell elemzése szerkezetben: {"summary", "findings", "missing"}; az idézetek ellenőrizve."""
    text = str(text or "")
    if not text.strip():
        raise ValueError(f"A(z) {doc_name} szövege nem olvasható ki, ezért nem vizsgálható.")
    user = (f"A megbízó (akinek az oldalán állunk): {client_role}\n"
            + (f"Külön figyelj erre: {focus}\n" if focus else "")
            + (f"Az ügy háttere: {context}\n" if context else "")
            + f"\nA TERVEZET ({doc_name}):\n{text[:MAX_CHARS]}")
    kwargs = dict(model=model, max_tokens=8000, system=SYSTEM, messages=[{"role": "user", "content": user}])
    try:
        response = client.messages.create(thinking={"type": "disabled"}, **kwargs)
    except Exception as e:
        if getattr(e, "status_code", None) != 400 or "thinking" not in str(e).lower():
            raise
        response = client.messages.create(**kwargs)
    if log:
        log(model, getattr(response, "usage", None))
    data = _parse_json(_text_of(response))
    findings = []
    for f in data.get("findings") or []:
        if not isinstance(f, dict):
            continue
        sev = str(f.get("severity", "")).upper().replace("KOZEPES", "KÖZEPES")
        f["severity"] = sev if sev in SEVERITY_ORDER else "KÖZEPES"
        f["verified"] = quote_found(f.get("quote"), text) if f.get("quote") else None
        findings.append(f)
    findings.sort(key=lambda f: SEVERITY_ORDER[f["severity"]])
    return {"summary": str(data.get("summary") or "").strip(), "findings": findings,
            "missing": [m for m in (data.get("missing") or []) if isinstance(m, dict)],
            "truncated": len(text) > MAX_CHARS}


def report_text(result, doc_name, client_role):
    """A jelentés szövegként (Larának és a chatbe)."""
    f = result["findings"]
    counts = {s: sum(1 for x in f if x["severity"] == s) for s in SEVERITY_ORDER}
    lines = [f"RED FLAG JELENTÉS – {doc_name}", f"Szemszög: {client_role}",
             f"Kockázatok: {counts['MAGAS']} magas, {counts['KÖZEPES']} közepes, {counts['ALACSONY']} alacsony; "
             f"hiányzó védelem: {len(result['missing'])}", ""]
    if result["summary"]:
        lines += ["Összkép: " + result["summary"], ""]
    for n, x in enumerate(f, 1):
        lines.append(f"{n}. [{x['severity']}] {x.get('title') or ''}" + (f" ({x['ref']})" if x.get("ref") else ""))
        if x.get("quote"):
            mark = "" if x["verified"] else "  (FIGYELEM: ez az idézet szó szerint nem található a tervezetben – ellenőrizd)"
            lines.append(f"   Idézet: „{x['quote']}”{mark}")
        if x.get("issue"):
            lines.append(f"   Kockázat: {x['issue']}")
        if x.get("suggestion"):
            lines.append(f"   Javaslat: {x['suggestion']}")
    if result["missing"]:
        lines += ["", "HIÁNYZÓ VÉDELMEK:"]
        for x in result["missing"]:
            lines.append(f"- {x.get('title') or ''}: {x.get('issue') or ''}"
                         + (f" Javaslat: {x['suggestion']}" if x.get("suggestion") else ""))
    if result["truncated"]:
        lines += ["", f"Megjegyzés: a tervezet hosszú, csak az első {MAX_CHARS // 1000} ezer karaktert vizsgáltam."]
    lines += ["", "A jelentést mutasd meg a felhasználónak (a legsúlyosabb tételekkel kezdve). Word-jelentést csak kérésre "
                  "készíts, a create_document eszközzel.",
              "A jelentés nem jogi tanácsadás: fontos ügyben, aláírás előtt javasold az ügyvédi egyeztetést."]
    return "\n".join(lines)


def review(client, model, doc_name, text, client_role, focus="", context="", log=None):
    result = analyze(client, model, doc_name, text, client_role, focus, context, log)
    return report_text(result, doc_name, client_role)
