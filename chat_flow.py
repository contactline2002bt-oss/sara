"""Egy chatkérdés teljes lebonyolítása: modellhívás, szerveroldali keresés folytatása (pause_turn)
és a dokumentumkészítő eszközök végrehajtása (tool_use). Streamlittől független, ezért tesztelhető."""
import export_tools as ex

MAX_FILE_BYTES = 10 * 1024 * 1024
HANDLERS_MAY_RETURN_FILES = True   # az app ezzel ellenőrzi, hogy a friss változat fut-e

# Ezt kapja Sara, ha a fájl tartalma nem fért bele egyetlen válaszba (a fájl ilyenkor nem készül el).
TRUNCATED_NOTE = (
    "[Az alkalmazás üzenete] A fájl NEM készült el: a tartalma túl hosszú volt egyetlen válaszhoz, ezért "
    "az eszközhívás megszakadt. Készítsd el újra úgy, hogy több, egyenként rövidebb fájlra bontod "
    "(pl. „… – 1. rész”, „… – 2. rész”), válaszonként egyetlen eszközhívással, vagy készíts tömörebb változatot."
)


# Gyorsítótár (prompt caching): a kérés ismétlődő eleje (eszközök, utasítások, korábbi üzenetek, csatolt
# dokumentumok) egy órára eltárolódik, és a következő hívásnál töredékáron megy. A tárolás határát az API
# a beszélgetés végéhez igazítja. Ha az API mégsem fogadná el, gyorsítótár nélkül megy tovább.
CACHE_CONTROL = {"type": "ephemeral", "ttl": "1h"}
_cache_ok = True


def _call_model(client, usage_cb=None, **kwargs):
    """Streamelve hívjuk a modellt: hosszú dokumentumnál a nem streamelt hívás időtúllépésbe futna,
    nagy max_tokens értéknél pedig az SDK el sem indítja."""
    global _cache_ok
    if _cache_ok:
        try:
            with client.messages.stream(extra_body={"cache_control": CACHE_CONTROL}, **kwargs) as stream:
                msg = stream.get_final_message()
        except Exception as e:
            if getattr(e, "status_code", None) != 400 or "cache" not in str(e).lower():
                raise
            _cache_ok = False   # ez a fiók / modell nem fogadja: gyorsítótár nélkül megyünk tovább
            return _call_model(client, usage_cb, **kwargs)
    else:
        with client.messages.stream(**kwargs) as stream:
            msg = stream.get_final_message()
    if usage_cb:
        try:
            usage_cb(kwargs.get("model", ""), msg.usage)
        except Exception:
            pass
    return msg


def _execute_tool(block, save_file, created, problems, handlers=None):
    """Egy tool_use blokk végrehajtása; mindig tool_result blokkot ad vissza."""
    def result(text, error=False):
        r = {"type": "tool_result", "tool_use_id": block.id, "content": text}
        if error:
            r["is_error"] = True
        return r

    produced = None
    if handlers and block.name in handlers:  # az alkalmazás saját eszközei (feladatok, dokumentumok olvasása, javítása)
        try:
            out = handlers[block.name](block.input or {})
        except Exception as e:
            return result(f"Hiba: {e}", error=True)
        if not (isinstance(out, tuple) and len(out) == 3 and isinstance(out[2], (bytes, bytearray))):
            return result(str(out))
        produced = out   # fájlt készítő eszköz (pl. Word-dokumentum javítása): ugyanúgy mentjük, mint a többit

    try:
        name, mime, data = produced or ex.run_tool(block.name, block.input)
        if len(data) > MAX_FILE_BYTES:
            raise ex.ToolError("A kész fájl túl nagy (legfeljebb 10 MB lehet). Készíts kisebbet vagy oszd több fájlra.")
        ok, info = save_file(name, mime, data)  # siker esetén info = a tárolt fájlnév, hiba esetén a hiba oka
        if not ok:
            problems.append(f"A(z) {name} fájl elkészült, de a mentése nem sikerült: {info}")
            return result(f"A fájl elkészült, de a mentése a dokumentumtárba nem sikerült: {info} "
                          "Ne próbáld újra; mondd meg a felhasználónak, hogy a mentés nem sikerült.", error=True)
        name = info or name
        created.append(name)
        return result(f"Elkészült és mentve a projekt dokumentumtárába: {name}. "
                      "A letöltőgombot az alkalmazás a válaszod alá teszi.")
    except ex.ToolError as e:
        return result(f"Hiba: {e} Javítsd, és hívd meg újra az eszközt.", error=True)
    except Exception as e:  # váratlan hiba ne állítsa meg az egész beszélgetést
        return result(f"Váratlan hiba a fájl készítésekor: {e}", error=True)


# Csak olvasó eszközök: egyszerre futhatnak (pl. a kereső kolléga a dokumentumtárban és a weben egy időben).
# A fájlt készítő és a feladatokat módosító eszközök sorban futnak, mert közös állapotot írnak.
PARALLEL_SAFE = {"find_information", "research_web", "read_project_document"}


# Körbefutás elleni védelem: egy kérdésen belül legfeljebb ennyi keresést kérhet Sara a kollégától, és ugyanazt
# kétszer nem. Utána nem keres tovább, hanem a felhasználótól kér pontosítást.
LOOKUP_TOOLS = {"find_information", "research_web"}
MAX_LOOKUPS = 3
LOOKUP_LIMIT_NOTE = (
    "Nem kerestem tovább: ebben a kérdésben elérted a keresések keretét ({n}). Ne fogalmazd át újra a keresést. "
    "Válaszolj azzal, amit eddig találtál, és kérd meg a felhasználót, hogy pontosítsa a kérdést vagy adjon több "
    "információt (például melyik dokumentumban, melyik cégről, melyik időszakról van szó). Röviden írd le, mit "
    "kerestél és mit nem találtál.")
LOOKUP_REPEAT_NOTE = (
    "Ezt a keresést ebben a kérdésben már lefuttattad, az eredménye fent van. Ne ismételd meg; ha nem elég, kérj "
    "pontosítást a felhasználótól.")


def _lookup_key(block):
    q = " ".join(str((block.input or {}).get("question", "")).lower().split())
    return f"{block.name}:{q}"


def _execute_all(blocks, save_file, created, problems, handlers, lookups=None):
    """Egy válasz összes eszközhívása; a párhuzamosan futtathatók egyszerre. Az eredmények sorrendje a hívásoké.
    lookups: a kérdésen belüli keresések nyilvántartása ({"count": int, "seen": set}) a körbefutás ellen."""
    from concurrent.futures import ThreadPoolExecutor
    results = [None] * len(blocks)
    if lookups is not None:
        for i, b in enumerate(blocks):
            if b.name not in LOOKUP_TOOLS:
                continue
            key = _lookup_key(b)
            if key in lookups["seen"]:
                note = LOOKUP_REPEAT_NOTE
            elif lookups["count"] >= MAX_LOOKUPS:
                note = LOOKUP_LIMIT_NOTE.format(n=MAX_LOOKUPS)
            else:
                lookups["count"] += 1
                lookups["seen"].add(key)
                continue
            results[i] = {"type": "tool_result", "tool_use_id": b.id, "content": note}
    parallel = [i for i, b in enumerate(blocks) if b.name in PARALLEL_SAFE and results[i] is None]
    if len(parallel) > 1:
        with ThreadPoolExecutor(max_workers=min(len(parallel), 4)) as pool:
            futures = {i: pool.submit(_execute_tool, blocks[i], save_file, created, problems, handlers)
                       for i in parallel}
            for i, fut in futures.items():
                results[i] = fut.result()
    for i, b in enumerate(blocks):
        if results[i] is None:
            results[i] = _execute_tool(b, save_file, created, problems, handlers)
    return results


def run_conversation(client, model, system, api_messages, tools, save_file, max_rounds=8, max_tokens=64000,
                     handlers=None, usage_cb=None):
    """Visszaad: (válasz szövege, források {url: cím}, elkészült fájlnevek, problémák listája).

    A problémák a felhasználónak szóló, érthető mondatok arról, ha egy fájl nem készült el vagy nem lett
    elmentve; így soha nem marad észrevétlen, ha Sara nem tudott dokumentumot leadni."""
    parts, sources, created, problems = [], {}, [], []
    lookups = {"count": 0, "seen": set()}   # a kollégától kért keresések ebben a kérdésben
    sep = ""
    retried_truncation = False
    for _ in range(max_rounds):
        try:
            response = _call_model(
                client, usage_cb, model=model, max_tokens=max_tokens, system=system, messages=api_messages, tools=tools
            )
        except Exception as e:  # hálózati vagy API-hiba: a már elkészült fájlokról így is beszámolunk
            if any(t.get("name") == "web_fetch" for t in tools) and "web_fetch" in str(e):
                tools = [t for t in tools if t.get("name") != "web_fetch"]   # a fiók nem fogadja: nélküle
                continue
            problems.append(f"Hiba történt a modell hívásakor: {e}")
            break

        # Az új modellek gondolkodási (thinking) és keresési blokkot is visszaadhatnak: csak a szöveget vesszük
        text = []
        for block in response.content:
            if block.type == "text":
                text.append(block.text)
                for citation in getattr(block, "citations", None) or []:
                    if getattr(citation, "url", None):
                        sources[citation.url] = citation.title or citation.url
        if text:
            parts.append(sep + "".join(text))

        if response.stop_reason == "pause_turn":  # a szerveroldali keresés megszakadt, folytatjuk
            api_messages.append({"role": "assistant", "content": response.content})
            sep = ""
            continue

        if response.stop_reason == "tool_use":  # dokumentumkészítő eszköz: végrehajtjuk, és visszaadjuk az eredményt
            blocks = [b for b in response.content if b.type == "tool_use"]
            results = _execute_all(blocks, save_file, created, problems, handlers, lookups)
            if not results:
                break
            api_messages.append({"role": "assistant", "content": response.content})
            api_messages.append({"role": "user", "content": results})
            sep = "\n\n"
            continue

        if response.stop_reason == "max_tokens":
            # A válasz elérte a hosszkorlátot. Ha épp fájlt készített, a félbemaradt eszközhívás használhatatlan:
            # egyszer megkérjük Sarát, hogy bontsa részekre; a csonka blokkot nem küldjük vissza.
            cut_tool = any(b.type == "tool_use" for b in response.content)
            if cut_tool and not retried_truncation:
                retried_truncation = True
                kept = "".join(text).strip()
                if kept:
                    api_messages.append({"role": "assistant", "content": kept})
                api_messages.append({"role": "user", "content": TRUNCATED_NOTE})
                sep = "\n\n"
                continue
            problems.append(
                "A fájl nem készült el, mert a tartalma túl hosszú volt egyetlen válaszhoz. "
                "Kérd Sarától részletekben (pl. fejezetenként külön fájlban)." if cut_tool else
                "A válasz elérte a hosszkorlátot, ezért a vége lemaradt. Kérd Sarától a folytatást."
            )
        break
    else:
        if not created:
            problems.append("Sara többszöri próbálkozásra sem tudta befejezni a feladatot. Próbáld meg újra, egyszerűbb kéréssel.")

    return "".join(parts).strip(), sources, created, list(dict.fromkeys(problems))


# ---------------------------------------------------------------------------
# Összefoglaló készítése az előzmények törlése előtt
# ---------------------------------------------------------------------------
SUMMARY_SEPARATOR = "====="

SUMMARY_INSTRUCTION = (
    "[Az alkalmazás kérése] Készíts összefoglalót a fenti beszélgetésről. A beszélgetés üzenetei és a "
    "hozzájuk csatolt munkaanyagok (képek, fájlok) ezután véglegesen törlődnek, és a jövőben kizárólag "
    "az összefoglalókból fogsz dolgozni. Ezért rögzíts minden olyan részletet, amire később szükség lehet; "
    "ami kimarad, az elvész. Ne általánosságokat írj, hanem konkrétumokat: neveket, összegeket, dátumokat, "
    "jogszabályhelyeket. Ne találj ki semmit, ami a beszélgetésben nem szerepelt.\n\n"
    "A válaszod pontosan ebben a formában legyen, más szöveg nélkül:\n"
    "LEÍRÁS: <2–3 mondat arról, miről szólt ez a beszélgetésszakasz és mi lett az eredménye>\n"
    f"{SUMMARY_SEPARATOR}\n"
    "<az összefoglaló Markdownban, az alábbi szakaszokkal; az üres szakaszt hagyd ki>\n"
    "## Tények és számok\n## Döntések és megállapítások\n## Kijavított tévedések\n"
    "## Nyitott kérdések és teendők\n## Elkészült fájlok\n## Csatolt munkaanyagok lényege\n\n"
    "A „Kijavított tévedések” szakaszba azt írd, amit a beszélgetés során rosszul mondtál és később "
    "helyesbítettél (a helyes változattal), hogy többé ne ismételd meg. A korábbi összefoglalók tartalmát "
    "ne ismételd meg, csak az azóta történteket foglald össze."
)


def parse_summary(text):
    """A modell válaszából (leírás, összefoglaló). Ha a formátum nem stimmel, a teljes szöveg az összefoglaló."""
    text = str(text or "").strip()
    head, sep, body = text.partition(SUMMARY_SEPARATOR)
    if sep and body.strip():
        description = head.strip()
        for prefix in ("LEÍRÁS:", "Leírás:", "**LEÍRÁS:**", "**Leírás:**"):
            if description.startswith(prefix):
                description = description[len(prefix):].strip()
        return " ".join(description.split()), body.strip().lstrip("=").strip()
    plain = " ".join(text.replace("#", " ").replace("*", " ").split())
    return plain[:300] + ("…" if len(plain) > 300 else ""), text


def summarize(client, model, system, api_messages, doc_names=(), fix_request="", previous_draft="", max_tokens=16000,
              usage_cb=None):
    """Összefoglaló a beszélgetésről. Visszaad: (rövid leírás, összefoglaló Markdownban).
    A fix_request a felhasználó javítási kérése az előző tervezethez (previous_draft)."""
    instruction = SUMMARY_INSTRUCTION
    if doc_names:
        instruction += ("\n\nA projekt dokumentumtárában lévő fájlok (ezek megmaradnak, a tartalmukat nem kell "
                        "leírnod, csak hivatkozz rájuk név szerint, ahol kell): " + ", ".join(doc_names))
    if previous_draft and fix_request:
        instruction += (f"\n\nAz előző tervezeted ez volt:\n<tervezet>\n{previous_draft}\n</tervezet>\n\n"
                        f"A felhasználó ezt kéri rajta javítani: {fix_request}\n"
                        "Írd meg újra a teljes összefoglalót a javítással, ugyanabban a formában.")
    messages = list(api_messages) + [{"role": "user", "content": instruction}]
    response = _call_model(client, usage_cb, model=model, max_tokens=max_tokens, system=system, messages=messages)
    text = "".join(b.text for b in response.content if b.type == "text")
    if not text.strip():
        raise RuntimeError("Sara nem adott vissza összefoglalót.")
    description, content = parse_summary(text)
    if response.stop_reason == "max_tokens":
        content += "\n\n*(Az összefoglaló elérte a hosszkorlátot, a vége hiányozhat. Egészítsd ki, mielőtt jóváhagyod.)*"
    return description, content
