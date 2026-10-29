"""Sara kereső kollégája: adatot keres a dokumentumtárban, előbb az aktuális projektben, aztán a többiben.

Olcsó modell (Haiku) végzi, két lépésben:
1. kiválasztás: a dokumentumok kivonatai alapján megmondja, melyikben lehet a keresett adat, és kulcsszavakat ad
   a szöveges kereséshez (ha az adat a kivonatban nem látszik, pl. egy cégjegyzékszám egy szerződés közepén);
2. olvasás: a kiválasztott dokumentumokból szó szerint kiidézi a releváns részeket, a helyükkel együtt.
Sara ezt a néhány idézetet kapja vissza a teljes dokumentumok helyett.

A bizalmas projektek (projects.confidential) dokumentumai más projektből nem kereshetők.
Streamlittől független, ezért tesztelhető.
"""
import json
import re
from concurrent.futures import ThreadPoolExecutor

import search_tools as stx

FINDER_MODEL = "claude-haiku-5-5"
MAX_CATALOG = 400        # ennyi dokumentum kivonatát látja a kiválasztó
MAX_READ = 4             # legfeljebb ennyi dokumentumot olvas el egy keresés
MAX_DOC_CHARS = 150_000  # egy dokumentumból ennyi szöveget kap az olvasó (kb. 40-50 ezer token)

TOOL = {
    "name": "find_information",
    "description": (
        "Adat keresése a dokumentumtárban a kereső kollégával: előbb az aktuális projekt dokumentumaiban, aztán a "
        "többi projektében (pl. cégadatok egy cégkivonatból, egy szerződés feltétele, egy összeg, határidő, "
        "elérhetőség). Szó szerinti idézeteket ad vissza a dokumentum nevével és a projekttel. Olcsóbb és gyorsabb, "
        "mint egy teljes dokumentum beolvasása, ezért adatkereséshez ezt használd; a read_project_document-et "
        "akkor, ha a teljes szöveg kell (szerkesztés, teljes elemzés, összevetés)."),
    "input_schema": {"type": "object", "properties": {
        "question": {"type": "string", "description": (
            "Mit keressen, pontosan és önállóan megfogalmazva, pl. „A Realion2001 Zrt. cégjegyzékszáma, székhelye "
            "és adószáma”.")}},
        "required": ["question"]},
}

SELECT_SYSTEM = (
    "Egy cég dokumentumtárában segítesz adatot keresni. Megkapod a kérdést és a dokumentumok listáját "
    "(azonosító, projekt, név, rövid kivonat). Válaszd ki azokat a dokumentumokat, amelyekben a kérdésre a válasz a "
    "legvalószínűbben benne van (legfeljebb 3, a legígéretesebb elöl; ha egyformán ígéretesek, az aktuális projektét "
    "részesítsd előnyben). Az [ELAVULT] jelölésű dokumentum régi változat: csak akkor válaszd, ha a kérdés "
    "kifejezetten régi változatra vonatkozik. Adj meg 1–3 kulcsszót is, amelyekre a dokumentumok teljes szövegében érdemes rákeresni "
    "(jellegzetes szavak vagy nevek, nem általános szavak). Csak JSON-t írj, más szöveget ne: "
    '{"ids": [12, 7], "keywords": ["cégjegyzékszám", "Realion"]}')

READ_SYSTEM = (
    "Egy dokumentumból keresel ki adatot. Idézd szó szerint azokat a részeket, amelyek a kérdésre választ adnak "
    "(legfeljebb 5 idézet, egyenként legfeljebb kb. 600 karakter), mindegyik elé írd a helyét, ha kiderül "
    "(pont- vagy fejezetszám, táblázat sora, oldal). Ne foglalj össze, ne értelmezz, ne egészíts ki semmit. Ha a "
    "dokumentumban nincs a kérdéshez tartozó adat, csak ennyit írj: NINCS")


def _text_of(response):
    return "".join(getattr(b, "text", "") for b in getattr(response, "content", [])
                   if getattr(b, "type", "") == "text").strip()


def _ask(client, system, user, max_tokens, log):
    kwargs = dict(model=FINDER_MODEL, max_tokens=max_tokens, system=system,
                  messages=[{"role": "user", "content": user}])
    try:
        response = client.messages.create(thinking={"type": "disabled"}, **kwargs)
    except Exception as e:
        if getattr(e, "status_code", None) != 400 or "thinking" not in str(e).lower():
            raise
        response = client.messages.create(**kwargs)
    if log:
        log(FINDER_MODEL, getattr(response, "usage", None))
    return _text_of(response)


def _parse_selection(text):
    """A kiválasztó válaszából: (azonosítók, kulcsszavak). Hibás JSON esetén a számokat szedi ki."""
    try:
        start, end = text.index("{"), text.rindex("}") + 1
        data = json.loads(text[start:end])
        ids = [int(i) for i in data.get("ids", []) if str(i).strip().lstrip("-").isdigit()]
        keywords = [str(k).strip() for k in data.get("keywords", []) if str(k).strip()]
        return ids, keywords
    except (ValueError, TypeError, AttributeError):
        return [int(n) for n in re.findall(r"\b\d+\b", text)][:3], []


def searchable_projects(projects, current_pid):
    """Az aktuális projekt és a nem bizalmas projektek; az aktuális elöl."""
    current = [p for p in projects if p["id"] == current_pid]
    others = [p for p in projects if p["id"] != current_pid and not p.get("confidential")]
    return current + others


def find(client, question, current_pid, projects, get_docs_meta, get_text, keyword_search=None, log=None):
    """Visszaad: Sarának szóló szöveg az idézetekkel, vagy hogy nem található."""
    question = str(question or "").strip()
    if not question:
        raise ValueError("Add meg, mit keressen (question).")
    scope = searchable_projects(projects, current_pid)
    names = {p["id"]: p["name"] for p in scope}
    docs = get_docs_meta([p["id"] for p in scope])
    docs.sort(key=lambda d: str(d.get("created_at") or ""), reverse=True)   # a legújabb elöl…
    docs.sort(key=lambda d: d["project_id"] != current_pid)                  # …az aktuális projekté a lista elején
    if not docs:
        return "A dokumentumtárban (az aktuális és a kereshető projektekben) nincs egyetlen dokumentum sem."
    by_id = {d["id"]: d for d in docs}

    catalog = []
    for d in docs[:MAX_CATALOG]:
        mark = (" [AKTUÁLIS PROJEKT]" if d["project_id"] == current_pid else "") + (" [ELAVULT]" if d.get("obsolete") else "")
        summary = " ".join(str(d.get("summary") or "nincs kivonat").split())[:400]
        catalog.append(f"[{d['id']}] Projekt: {names.get(d['project_id'], '?')}{mark} | {d['name']} – {summary}")
    ids, keywords = _parse_selection(_ask(
        client, SELECT_SYSTEM, f"Kérdés: {question}\n\nDokumentumok:\n" + "\n".join(catalog), 400, log))
    chosen = [i for i in dict.fromkeys(ids) if i in by_id][:3]

    # Szöveges keresés a kulcsszavakra: az is előkerül, ami a kivonatból nem látszik
    if keyword_search and keywords:
        hits = {}
        for kw in keywords[:3]:
            terms = stx.query_terms(kw)
            if not terms:
                continue
            for d in keyword_search([p["id"] for p in scope], stx.filter_stems(terms)):
                if d["id"] in by_id and stx.matches(f"{d['name']}\n{d.get('content_text') or ''}", terms):
                    hits[d["id"]] = hits.get(d["id"], 0) + stx.count_hits(d.get("content_text") or "", terms)
        ranked = sorted(hits, key=lambda i: (by_id[i]["project_id"] != current_pid, -hits[i]))
        for i in ranked:
            if len(chosen) >= MAX_READ:
                break
            if i not in chosen:
                chosen.append(i)

    if not chosen:
        return ("A kereső nem talált olyan dokumentumot, amelyben ez az adat szerepelhet "
                f"({len(docs)} dokumentumot nézett át {len(scope)} projektben). Kérdezd meg a felhasználót.")

    def read(doc_id):
        d = by_id[doc_id]
        text = (get_text(doc_id) or "").strip()
        if not text:
            return d, None
        answer = _ask(client, READ_SYSTEM, f"Kérdés: {question}\n\nA dokumentum ({d['name']}) szövege:\n\n"
                                          f"{text[:MAX_DOC_CHARS]}", 1500, log)
        return d, answer

    with ThreadPoolExecutor(max_workers=MAX_READ) as pool:
        results = list(pool.map(read, chosen))

    found, empty = [], []
    for d, answer in results:
        label = f"{d['name']} (projekt: {names.get(d['project_id'], '?')})"
        if answer and answer.strip().upper().rstrip(".") != "NINCS":
            found.append(f"### {label}\n{answer.strip()}")
        else:
            empty.append(label)
    if not found:
        return ("A kereső átnézte ezeket, de a kért adatot nem találta bennük: " + "; ".join(empty)
                + ". Ha biztosan valamelyikben van, olvasd be a read_project_document eszközzel (csak az aktuális "
                  "projektét), vagy kérdezd meg a felhasználót.")
    out = ("A kereső kolléga eredménye (szó szerinti idézetek; a válaszban add meg, melyik dokumentumból "
           "származik az adat):\n\n" + "\n\n".join(found))
    if empty:
        out += "\n\nÁtnézve, de nem volt benne: " + "; ".join(empty)
    return out


# ---------------------------------------------------------------------------
# Webes kutatás: a kolléga keres a neten, és csak a releváns idézeteket adja vissza a forrással
# ---------------------------------------------------------------------------
WEB_TOOL = {
    "name": "research_web",
    "description": (
        "Internetes kutatás a kereső kollégával: ő keres a weben, átolvassa a találatokat, és csak a kérdéshez "
        "tartozó szó szerinti idézeteket adja vissza a forrás címével és linkjével. Aktuális adatokhoz "
        "(jogszabályok, adómértékek, határidők, díjak, cégadatok, hírek) ezt használd elsőként: olcsóbb, mint a "
        "saját keresésed. A saját web_search eszközödet csak akkor használd, ha a kolléga eredménye nem elég, "
        "ellentmondásos, vagy egy kulcsfontosságú adatot ellenőrizni kell."),
    "input_schema": {"type": "object", "properties": {
        "question": {"type": "string", "description": (
            "Mit keressen, pontosan és önállóan megfogalmazva, a szükséges részletekkel (ország, időpont, "
            "jogszabály, cég neve), pl. „A 2026-ban hatályos magyar kiberbiztonsági törvény végrehajtási rendelete: "
            "száma és a 3-as biztonsági osztály követelményei”.")}},
        "required": ["question"]},
}

WEB_SYSTEM = (
    "Egy cég főmunkatársának keresel adatot az interneten. Keress rá a kérdésre (több kereséssel is, ha kell), "
    "olvasd át a találatokat, és csak azt add vissza, ami a kérdésre válaszol: szó szerinti idézeteket (legfeljebb "
    "6, egyenként legfeljebb kb. 500 karakter), mindegyik alatt a forrás címével és pontos linkjével. "
    "A hivatalos és elsődleges forrásokat részesítsd előnyben (pl. njt.hu, magyarkozlony.hu, nav.gov.hu, "
    "e-cegjegyzek.hu, a hatóságok és a cég saját oldala); sajtót, blogot csak akkor, ha nincs jobb. Jelezd, ha egy "
    "forrás régi, vagy ha a források ellentmondanak egymásnak. Ne foglalj össze, ne értelmezz, ne egészíts ki. "
    "Ha nem találsz megbízható választ, írd meg röviden, mit kerestél és mit nem találtál.")


def web_research(client, question, log=None, max_searches=5, max_rounds=4):
    """A kolléga webes keresése (Haiku + szerveroldali web_search). Visszaad: Sarának szóló szöveg."""
    question = str(question or "").strip()
    if not question:
        raise ValueError("Add meg, mit keressen (question).")
    messages = [{"role": "user", "content": question}]
    tools = [{"type": "web_search_20250305", "name": "web_search", "max_uses": max_searches}]
    parts, sources = [], {}
    for _ in range(max_rounds):
        kwargs = dict(model=FINDER_MODEL, max_tokens=3000, system=WEB_SYSTEM, messages=messages, tools=tools)
        try:
            response = client.messages.create(thinking={"type": "disabled"}, **kwargs)
        except Exception as e:
            if getattr(e, "status_code", None) != 400 or "thinking" not in str(e).lower():
                raise
            response = client.messages.create(**kwargs)
        if log:
            log(FINDER_MODEL, getattr(response, "usage", None))
        for block in response.content:
            if getattr(block, "type", "") == "text":
                parts.append(block.text)
                for c in getattr(block, "citations", None) or []:
                    if getattr(c, "url", None):
                        sources[c.url] = getattr(c, "title", None) or c.url
        if getattr(response, "stop_reason", "") == "pause_turn":   # a keresés folytatódik
            messages = messages + [{"role": "assistant", "content": response.content}]
            continue
        break
    text = "".join(parts).strip()
    if not text:
        return "A kereső kolléga nem adott eredményt; keress rá magad a web_search eszközzel."
    out = "A kereső kolléga webes eredménye (idézetek a forrásokkal; ellenőrizd a kulcsadatokat):\n\n" + text
    if sources:
        out += "\n\nHivatkozott források:\n" + "\n".join(f"- {t}: {u}" for u, t in sources.items())
    return out
