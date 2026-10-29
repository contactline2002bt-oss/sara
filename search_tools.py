"""Szabadszavas keresés segédfüggvényei: ékezet- és kisbetű-független egyeztetés, egyszerű magyar
toldalékkezelés és találati szövegkörnyezet. Streamlittől és adatbázistól független."""
import html
import re
import unicodedata
from functools import lru_cache

MAX_SEARCH_TEXT = 300_000  # ennyi karaktert tárolunk egy dokumentum szövegéből


def _fold_char(ch):
    """Egy karakter kisbetűs, ékezet nélküli megfelelője; mindig pontosan egy karakter."""
    base = "".join(c for c in unicodedata.normalize("NFKD", ch) if not unicodedata.combining(c)).lower()
    if len(base) == 1:
        return base
    lower = ch.lower()
    return lower if len(lower) == 1 else ch


# Karakterenkénti csere: a normalizált szöveg ugyanolyan hosszú, mint az eredeti, ezért a találat
# pozíciója az eredeti szövegben is érvényes (ebből készül a kiemelt szövegkörnyezet).
_FOLD = {cp: _fold_char(chr(cp)) for cp in range(0x2000) if _fold_char(chr(cp)) != chr(cp)}


def normalize(text):
    return str(text or "").translate(_FOLD)


def clean_text(text):
    """Tárolás előtt: a NUL karaktert a PostgreSQL nem fogadja el; a túl hosszú szöveget levágjuk."""
    return str(text or "").replace("\x00", "")[:MAX_SEARCH_TEXT]


def build_search_text(title, text):
    """A kereshető (normalizált) szöveg: cím/fájlnév + tartalom."""
    return normalize(clean_text(f"{title}\n{text or ''}"))


# Gyakori magyar toldalékok ékezet nélküli alakban.
_SUFFIXES = ["kent", "ban", "ben", "bol", "rol", "tol", "nak", "nek", "val", "vel", "hoz", "hez", "nal", "nel",
             "ert", "ba", "be", "ra", "re", "on", "en", "ot", "et", "at", "ok", "ek", "ak", "ig", "t", "k", "n"]
_MIN_STEM = 3
_WORD = r"[^\W_]"


@lru_cache(maxsize=50_000)
def stem_candidates(word):
    """A szó lehetséges tövei (maga a szó is): minden alak, ami gyakori toldalékok levágásával előáll.
    Szótár nélkül nem dönthető el, hol a tő vége (házat = ház+at, de számlát = számla+t), ezért az összes
    lehetőséget megtartjuk; két szó akkor „ugyanaz”, ha van közös lehetséges tövük."""
    found, queue = {word}, [word]
    for _ in range(3):  # legfeljebb három toldalék egymás után (pl. szerzodes-ek-ben)
        next_queue = []
        for w in queue:
            shorter = []
            m = re.fullmatch(r"(.+?)([b-df-hj-np-tv-z])\2[ae]l", w)  # -val/-vel hasonulva: szerzodessel
            if m and len(m.group(1)) + 1 >= 4:
                shorter.append(m.group(1) + m.group(2))
            shorter += [w[: -len(suf)] for suf in _SUFFIXES if w.endswith(suf) and len(w) - len(suf) >= _MIN_STEM]
            for cand in shorter:
                if cand not in found:
                    found.add(cand)
                    next_queue.append(cand)
        queue = next_queue
    return frozenset(found)


def query_terms(query):
    """A keresőkifejezés szavai normalizálva (az egybetűs szavakat kihagyjuk)."""
    words = [w for w in re.findall(_WORD + "+", normalize(query)) if len(w) >= 2]
    return list(dict.fromkeys(words))


def _shortest_stem(word):
    return min(stem_candidates(word), key=len)


def filter_stems(terms):
    """Az adatbázis-oldali előszűréshez: ha egy szöveg találat, ezek a tövek biztosan szerepelnek benne."""
    return [_shortest_stem(w) for w in terms]


def _word_matches(text_word, word):
    """Egyezik-e a szöveg egy szava a keresőszóval:
    - tartalmazza a beírt szót (összetett és képzett szavak: adó -> vagyonadó, érték -> értékelés), vagy
    - van közös lehetséges tövük (ragozott alakok: szerződést -> szerződésben), vagy
    - összetett szó, amelynek a vége a keresett szó töve (csak hosszabb tőnél, a téves találatok ellen)."""
    if word in text_word:
        return True
    text_stems, word_stems = stem_candidates(text_word), stem_candidates(word)
    if text_stems & word_stems:
        return True
    return any(len(ws) >= 5 and ts.endswith(ws) for ws in word_stems for ts in text_stems)


def _term_spans(norm_text, word, limit=200):
    spans = []
    for m in re.finditer(f"{_WORD}*{re.escape(_shortest_stem(word))}{_WORD}*", norm_text):
        if _word_matches(m.group(0), word):
            spans.append((m.start(), m.end()))
            if len(spans) >= limit:
                break
    return spans


def _spans(norm_text, terms):
    """A találatok (kezdet, vég) párjai sorrendben; mindig a teljes szó van kijelölve."""
    spans = sorted(span for word in terms for span in _term_spans(norm_text, word))
    merged = []
    for start, end in spans:
        if merged and start <= merged[-1][1]:
            merged[-1] = (merged[-1][0], max(merged[-1][1], end))
        else:
            merged.append((start, end))
    return merged


def matches(text, terms):
    """Igaz, ha a szövegben minden keresőszóra van találat."""
    norm_text = normalize(text)
    return bool(terms) and all(_term_spans(norm_text, w, limit=1) for w in terms)


def count_hits(text, terms):
    return len(_spans(normalize(text), terms))


def highlight(text, terms):
    """A teljes szöveg HTML-ként, a találatok <mark>-kal kiemelve (rövid szöveghez, pl. címhez)."""
    text = str(text or "")
    out, pos = [], 0
    for start, end in _spans(normalize(text), terms):
        out.append(html.escape(text[pos:start]))
        out.append(f"<mark>{html.escape(text[start:end])}</mark>")
        pos = end
    out.append(html.escape(text[pos:]))
    return "".join(out)


def snippets(text, terms, width=110, max_snippets=2):
    """Legfeljebb max_snippets részlet a szövegből a találatok körül, HTML-ként, kiemeléssel."""
    text = str(text or "")
    result, covered_until = [], -1
    for start, end in _spans(normalize(text), terms):
        if start < covered_until:
            continue
        left, right = max(0, start - width), min(len(text), end + width)
        # szóhatárra igazítás, hogy ne vágjunk ketté szavakat
        if left > 0:
            cut = text.find(" ", left, start)
            left = cut + 1 if cut != -1 else left
        if right < len(text):
            cut = text.rfind(" ", end, right)
            right = cut if cut != -1 else right
        piece = re.sub(r"\s+", " ", text[left:right])
        result.append(("… " if left > 0 else "") + highlight(piece, terms) + (" …" if right < len(text) else ""))
        covered_until = right
        if len(result) >= max_snippets:
            break
    return result
