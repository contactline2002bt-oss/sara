"""Meglévő Word-dokumentum javítása az eredeti formázás megtartásával.

Lara nem írja újra a dokumentumot, hanem módosításokat ad (mit mire cseréljen, mit töröljön, hova szúrjon be),
és ezeket közvetlenül az eredeti .docx XML-jében hajtjuk végre. Minden más (betűtípus, stílusok, számozás,
élőfej, táblázatok, megjegyzések) változatlan marad. Kérésre a módosítások a Word „Változások követése”
jelöléseivel kerülnek be, így egyenként elfogadhatók vagy elvethetők.

Csak az lxml kell (a python-docx függősége).
"""
import copy
import datetime
import io
import re
import zipfile

from lxml import etree

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W = "{%s}" % W_NS
XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"
ACTIONS = ("replace", "delete", "insert_after", "insert_before", "comment", "format", "table_borders")
JC = {"left": "left", "center": "center", "right": "right", "justify": "both"}
COMMENTS_PART = "word/comments.xml"
COMMENTS_TYPE = "application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"
COMMENTS_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments"


class EditError(Exception):
    """Érthető hibaüzenet, amit Lara visszakap, és ki tud javítani."""


def _w(tag):
    return etree.Element(W + tag)


# ---------------------------------------------------------------------------
# Bekezdések és szövegük
# ---------------------------------------------------------------------------
def _in_deleted(el, stop):
    p = el.getparent()
    while p is not None and p is not stop:
        if p.tag in (W + "del", W + "moveFrom"):
            return True
        p = p.getparent()
    return False


def _runs(p):
    """A bekezdés látható szövegű futásai (a törölt, a mezőkód- és a beágyazott bekezdések nélkül)."""
    out = []
    for r in p.iter(W + "r"):
        if _in_deleted(r, p):
            continue
        anc = r.getparent()
        nested = False
        while anc is not None and anc is not p:   # szövegdobozban lévő bekezdés futása nem ehhez tartozik
            if anc.tag == W + "p":
                nested = True
                break
            anc = anc.getparent()
        if not nested:
            out.append(r)
    return out


def _piece_text(child):
    if child.tag == W + "t":
        return child.text or ""
    if child.tag == W + "tab":
        return "\t"
    if child.tag in (W + "br", W + "cr"):
        return "\n"
    if child.tag == W + "noBreakHyphen":
        return "-"
    return None


def _atomize(p):
    """Minden futást egyelemű futásokra bont (azonos formázással), hogy karakterpontosan vághassunk."""
    for r in _runs(p):
        content = [c for c in r if c.tag != W + "rPr"]
        if len(content) <= 1:
            continue
        rpr = r.find(W + "rPr")
        parent, idx = r.getparent(), r.getparent().index(r)
        parent.remove(r)
        for k, c in enumerate(content):
            nr = _w("r")
            for key, val in r.attrib.items():
                nr.set(key, val)
            if rpr is not None:
                nr.append(copy.deepcopy(rpr))
            nr.append(c)
            parent.insert(idx + k, nr)


def _pieces(p):
    """[(futás, szöveg)] – a futások a bekezdés sorrendjében; a nem szöveges futás szövege ''."""
    out = []
    for r in _runs(p):
        text = ""
        for c in r:
            t = _piece_text(c)
            if t:
                text += t
        out.append((r, text))
    return out


def _para_text(p):
    return "".join(t for _, t in _pieces(p))


def _norm_map(text):
    """Szóközök összevonása, nem törő szóköz és tipográfiai jelek egységesítése; visszaad: (szöveg, indextérkép)."""
    trans = {" ": " ", " ": " ", " ": " ", "\t": " ", "\n": " ",
             "„": '"', "”": '"', "“": '"', "’": "'", "‘": "'",
             "–": "-", "—": "-"}
    out, idx = [], []
    for i, ch in enumerate(text):
        ch = trans.get(ch, ch)
        if ch == " " and out and out[-1] == " ":
            continue
        out.append(ch)
        idx.append(i)
    return "".join(out), idx


def _find_all(text, needle):
    """A needle előfordulásai a szövegben [(kezdet, vég)]: előbb pontos, aztán egységesített egyezés."""
    hits, start = [], 0
    while needle:
        i = text.find(needle, start)
        if i < 0:
            break
        hits.append((i, i + len(needle)))
        start = i + 1
    if hits:
        return hits
    nt, idx = _norm_map(text)
    nn, _ = _norm_map(needle.strip())
    start = 0
    while nn:
        i = nt.find(nn, start)
        if i < 0:
            break
        hits.append((idx[i], idx[i + len(nn) - 1] + 1))
        start = i + 1
    return hits


_LABEL = re.compile(r"^\s*(?:\d+(?:\.\d+)*\.?|[a-zA-Z]\)|[ivxIVX]+\.|•)\s+")
_COMMENT_MARK = re.compile(r"\s*\[M\d+\]")


def _candidates(find):
    """A keresett szöveg változatai: ahogy Lara megadta, a megjegyzésjelölők és a sorszám nélkül
    (a kinyert szövegben ezek benne vannak, a dokumentum szövegében nincsenek)."""
    out = [find]
    no_marks = _COMMENT_MARK.sub("", find)
    out.append(no_marks)
    out.append(_LABEL.sub("", no_marks, count=1))
    return [c for c in dict.fromkeys(out) if c.strip()]


# ---------------------------------------------------------------------------
# Változáskövetés
# ---------------------------------------------------------------------------
class _Rev:
    def __init__(self, root, author, tracked):
        ids = [int(v) for v in root.xpath("//@w:id", namespaces={"w": W_NS}) if str(v).lstrip("-").isdigit()]
        self.next_id = max(ids + [0]) + 1000
        self.author = author
        self.tracked = tracked
        self.date = datetime.datetime.now(datetime.timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")

    def mark(self, tag):
        el = _w(tag)
        el.set(W + "id", str(self.next_id))
        el.set(W + "author", self.author)
        el.set(W + "date", self.date)
        self.next_id += 1
        return el


def _make_runs(text, rpr):
    runs = []
    for k, line in enumerate(str(text).split("\n")):
        if k:
            br = _w("r")
            if rpr is not None:
                br.append(copy.deepcopy(rpr))
            br.append(_w("br"))
            runs.append(br)
        for j, part in enumerate(line.split("\t")):
            if j:
                tr = _w("r")
                if rpr is not None:
                    tr.append(copy.deepcopy(rpr))
                tr.append(_w("tab"))
                runs.append(tr)
            if part:
                r = _w("r")
                if rpr is not None:
                    r.append(copy.deepcopy(rpr))
                t = _w("t")
                t.text = part
                t.set(XML_SPACE, "preserve")
                r.append(t)
                runs.append(r)
    return runs


def _clean_rpr(rpr):
    if rpr is None:
        return None
    rpr = copy.deepcopy(rpr)
    for tag in ("ins", "del", "moveFrom", "moveTo", "rPrChange"):
        for el in rpr.findall(W + tag):
            rpr.remove(el)
    return rpr


def _split_run(r, offset):
    """A szöveges futás kettévágása; visszaad: a második fele (az eredeti marad az első)."""
    t = r.find(W + "t")
    text = t.text or ""
    second = copy.deepcopy(r)
    t.text = text[:offset]
    t.set(XML_SPACE, "preserve")
    t2 = second.find(W + "t")
    t2.text = text[offset:]
    t2.set(XML_SPACE, "preserve")
    r.addnext(second)
    return second


def _isolate(p, start, end):
    """A [start, end) karaktertartomány futásai külön futásokban; visszaad: a futások listája."""
    _atomize(p)
    pos = 0
    for r, text in _pieces(p):
        n = len(text)
        if n and pos < start < pos + n and r.find(W + "t") is not None:
            _split_run(r, start - pos)
            break
        pos += n
    pos = 0
    for r, text in _pieces(p):
        n = len(text)
        if n and pos < end < pos + n and r.find(W + "t") is not None:
            _split_run(r, end - pos)
            break
        pos += n
    out, pos = [], 0
    for r, text in _pieces(p):
        n = len(text)
        if n and start <= pos and pos + n <= end:
            out.append(r)
        pos += n
    return out


def _delete_runs(runs, rev):
    for r in runs:
        if rev.tracked:
            for t in r.findall(W + "t"):
                t.tag = W + "delText"
            for t in r.findall(W + "instrText"):
                t.tag = W + "delInstrText"
            d = rev.mark("del")
            r.addprevious(d)
            d.append(r)
        else:
            r.getparent().remove(r)


def _insert_runs_after(anchor, new_runs, rev):
    """Új futások az anchor elem után (változáskövetéssel w:ins-be csomagolva)."""
    if not new_runs:
        return
    if rev.tracked:
        ins = rev.mark("ins")
        for r in new_runs:
            ins.append(r)
        anchor.addnext(ins)
    else:
        for r in reversed(new_runs):
            anchor.addnext(r)


def _first_rpr(runs):
    for r in runs:
        if r.find(W + "t") is not None:
            return _clean_rpr(r.find(W + "rPr"))
    return _clean_rpr(runs[0].find(W + "rPr")) if runs else None


def _mark_paragraph(p, tag, rev):
    """A bekezdésjel törlése vagy beszúrása változáskövetéssel (w:pPr/w:rPr/w:del|w:ins)."""
    ppr = p.find(W + "pPr")
    if ppr is None:
        ppr = _w("pPr")
        p.insert(0, ppr)
    rpr = ppr.find(W + "rPr")
    if rpr is None:
        rpr = _w("rPr")
        sect = ppr.find(W + "sectPr")
        if sect is not None:
            sect.addprevious(rpr)
        else:
            ppr.append(rpr)
    rpr.append(rev.mark(tag))


def _new_paragraph_like(anchor, text, rev):
    p = _w("p")
    ppr = anchor.find(W + "pPr")
    if ppr is not None:
        ppr = copy.deepcopy(ppr)
        for tag in ("sectPr", "pPrChange"):
            for el in ppr.findall(W + tag):
                ppr.remove(el)
        prpr = ppr.find(W + "rPr")
        if prpr is not None:
            for tag in ("ins", "del"):
                for el in prpr.findall(W + tag):
                    prpr.remove(el)
        p.append(ppr)
    rpr = _first_rpr(_runs(anchor))
    runs = _make_runs(text, rpr)
    if rev.tracked:
        _mark_paragraph(p, "ins", rev)
        if runs:
            ins = rev.mark("ins")
            for r in runs:
                ins.append(r)
            p.append(ins)
    else:
        for r in runs:
            p.append(r)
    return p


# ---------------------------------------------------------------------------
# Egy módosítás végrehajtása
# ---------------------------------------------------------------------------
MC_FALLBACK = "{http://schemas.openxmlformats.org/markup-compatibility/2006}Fallback"


def _in_fallback(p):
    """A szövegdobozok régi (VML) tartalék példánya: a Word nem ezt mutatja, ezért nem számít külön találatnak."""
    a = p.getparent()
    while a is not None:
        if a.tag == MC_FALLBACK:
            return True
        a = a.getparent()
    return False


def _paragraphs(body):
    return [p for p in body.iter(W + "p") if not _in_fallback(p)]


def _locate(body, find):
    """[(bekezdés, kezdet, vég)] a keresett szövegre; az első olyan változat, amelyik talál."""
    paras = _paragraphs(body)
    for cand in _candidates(find):
        hits = []
        for p in paras:
            text = _para_text(p)
            if not text:
                continue
            for s, e in _find_all(text, cand):
                hits.append((p, s, e))
        if hits:
            return hits
    return []


def _narrow(old, new, s, e):
    """Csak a ténylegesen változó rész cseréje (szóhatárra igazítva), hogy a követett változás ne az egész
    mondatot húzza át egy-két szó miatt. Visszaad: (új kezdet, új vég, beszúrandó szöveg)."""
    if old == new:
        return s, s, ""
    i = 0
    while i < min(len(old), len(new)) and old[i] == new[i]:
        i += 1
    while i > 0 and old[i - 1] not in " \t\n":   # vissza a szó elejére
        i -= 1
    j = 0
    while j < min(len(old), len(new)) - i and old[len(old) - 1 - j] == new[len(new) - 1 - j]:
        j += 1
    while j > 0 and old[len(old) - j] not in " \t\n.,;:)":   # előre a szó végére
        j -= 1
    return s + i, e - j, new[i:len(new) - j]


def _ppr(p):
    ppr = p.find(W + "pPr")
    if ppr is None:
        ppr = _w("pPr")
        p.insert(0, ppr)
    return ppr


def _ppr_child(ppr, tag, after=()):
    """A pPr gyermeke (létrehozza, ha nincs); a Word a sorrendre érzékeny, ezért a jc és a spacing a megfelelő
    helyre kerül (a pStyle, numPr stb. után, az rPr és a sectPr elé)."""
    el = ppr.find(W + tag)
    if el is not None:
        return el
    el = _w(tag)
    anchor = None
    for name in ("rPr", "sectPr", "pPrChange"):
        anchor = ppr.find(W + name)
        if anchor is not None:
            break
    if anchor is not None:
        anchor.addprevious(el)
    else:
        ppr.append(el)
    return el


def _format_paragraph(p, edit):
    """Igazítás, térköz, félkövér egy bekezdésen (a formázás nem követett változás, közvetlenül kerül be)."""
    changed = False
    ppr = _ppr(p)
    align = str(edit.get("align") or "").strip()
    if align:
        if align not in JC:
            raise EditError(f"ismeretlen igazítás: {align!r} (lehet: {', '.join(JC)})")
        _ppr_child(ppr, "jc").set(W + "val", JC[align])
        changed = True
    for key, attr in (("space_before", "before"), ("space_after", "after")):
        if edit.get(key) is not None:
            sp = _ppr_child(ppr, "spacing")
            sp.set(W + attr, str(int(float(edit[key]) * 20)))   # pont → huszadpont
            sp.set(W + ("beforeAutospacing" if attr == "before" else "afterAutospacing"), "0")
            changed = True
    if edit.get("bold") is not None:
        for r in _runs(p):
            rpr = r.find(W + "rPr")
            if rpr is None:
                rpr = _w("rPr")
                r.insert(0, rpr)
            for b in rpr.findall(W + "b"):
                rpr.remove(b)
            b = _w("b")
            if not edit.get("bold"):
                b.set(W + "val", "0")
            rpr.insert(0, b)
        changed = True
    return changed


def _set_table_borders(tbl, visible):
    tblpr = tbl.find(W + "tblPr")
    if tblpr is None:
        tblpr = _w("tblPr")
        tbl.insert(0, tblpr)
    for old in tblpr.findall(W + "tblBorders"):
        tblpr.remove(old)
    borders = _w("tblBorders")
    for side in ("top", "left", "bottom", "right", "insideH", "insideV"):
        el = _w(side)
        if visible:
            for k, v in (("val", "single"), ("sz", "4"), ("space", "0"), ("color", "auto")):
                el.set(W + k, v)
        else:
            el.set(W + "val", "nil")
        borders.append(el)
    # a tblBorders a tblW / jc után, a tblLayout / tblCellMar / tblLook előtt áll
    anchor = None
    for name in ("tblLayout", "tblCellMar", "tblLook", "tblCaption"):
        anchor = tblpr.find(W + name)
        if anchor is not None:
            break
    if anchor is not None:
        anchor.addprevious(borders)
    else:
        tblpr.append(borders)
    for tc in tbl.iter(W + "tc"):   # a cellák saját szegélyét is egységesítjük
        tcpr = tc.find(W + "tcPr")
        if tcpr is not None:
            for old in tcpr.findall(W + "tcBorders"):
                tcpr.remove(old)


def _apply(body, edit, rev):
    action = str(edit.get("action") or "").strip()
    find = str(edit.get("find") or "")
    new_text = edit.get("new_text")
    new_text = "" if new_text is None else str(new_text)
    if action not in ACTIONS:
        raise EditError(f"ismeretlen művelet: {action!r} (lehet: {', '.join(ACTIONS)})")
    if not find.strip():
        raise EditError("üres a find mező")
    if action in ("insert_after", "insert_before") and not new_text.strip():
        raise EditError("beszúrásnál a new_text nem lehet üres")
    if action == "format" and not any(edit.get(k) is not None and edit.get(k) != ""
                                      for k in ("align", "space_before", "space_after", "bold")):
        raise EditError("formázásnál adj meg legalább egyet: align, space_before, space_after, bold")
    if action == "comment" and not new_text.strip():
        raise EditError("megjegyzésnél a new_text (a megjegyzés szövege) nem lehet üres")
    hits = _locate(body, find)
    if not hits:
        raise EditError("a keresett szöveg nem található a dokumentumban (egyetlen bekezdésen/cellán belüli, "
                        "szó szerinti részletet adj meg)")
    if edit.get("paragraph"):   # csak az adott sorszámú bekezdésben (1-től, a dokumentum összes bekezdése szerint)
        try:
            target = _paragraphs(body)[int(edit["paragraph"]) - 1]
        except (ValueError, IndexError):
            raise EditError(f"nincs {edit.get('paragraph')}. bekezdés")
        hits = [h for h in hits if h[0] is target]
        if not hits:
            raise EditError(f"a keresett szöveg nincs a(z) {edit.get('paragraph')}. bekezdésben")
    if edit.get("occurrence"):   # az n-edik előfordulás (1-től): azonos szövegű helyek megkülönböztetésére
        try:
            n = int(edit["occurrence"])
        except ValueError:
            raise EditError("az occurrence egész szám legyen (1 = az első előfordulás)")
        if not 1 <= n <= len(hits):
            raise EditError(f"a keresett szöveg csak {len(hits)} helyen fordul elő, {n}. nincs")
        hits = [hits[n - 1]]
    if len(hits) > 1 and not edit.get("all"):
        raise EditError(f"a keresett szöveg {len(hits)} helyen fordul elő; adj meg hosszabb, egyedi részletet, "
                        "állítsd az all mezőt true-ra, vagy add meg, hányadik előfordulásról van szó (occurrence: 1, "
                        "2, …)")

    if action == "comment":
        for p, s, e in reversed(hits):
            runs = _isolate(p, s, e)
            if not runs:
                continue
            cid = rev.comments.add(new_text, rev)
            start, end = _w("commentRangeStart"), _w("commentRangeEnd")
            start.set(W + "id", cid)
            end.set(W + "id", cid)
            runs[0].addprevious(start)
            runs[-1].addnext(end)
            ref_run = _w("r")
            rpr = _w("rPr")
            style = _w("rStyle")
            style.set(W + "val", "CommentReference")
            rpr.append(style)
            ref_run.append(rpr)
            ref = _w("commentReference")
            ref.set(W + "id", cid)
            ref_run.append(ref)
            end.addnext(ref_run)
        return len(hits)

    if action == "format":
        seen = []
        for p, _, _ in hits:
            if not any(p is q for q in seen):
                seen.append(p)
                _format_paragraph(p, edit)
        return len(seen)

    if action == "table_borders":
        tables = []
        for p, _, _ in hits:
            t = p.getparent()
            while t is not None and t.tag != W + "tbl":
                t = t.getparent()
            if t is None:
                raise EditError("a keresett szöveg nem táblázatban van")
            if not any(t is q for q in tables):
                tables.append(t)
        visible = str(edit.get("borders") or "none").lower() in ("all", "true", "visible", "látható")
        for t in tables:
            _set_table_borders(t, visible)
        return len(tables)

    if action in ("insert_after", "insert_before"):
        seen = []
        for p, _, _ in hits:
            if any(p is q for q in seen):
                continue
            seen.append(p)
            lines = [line for line in new_text.split("\n\n")] if "\n\n" in new_text else new_text.split("\n")
            new_ps = [_new_paragraph_like(p, line, rev) for line in lines if line.strip()]
            for np in new_ps:   # a beszúrt bekezdés formázása is megadható (igazítás, térköz, félkövér)
                if any(edit.get(k) is not None for k in ("align", "space_before", "space_after", "bold")):
                    _format_paragraph(np, edit)
            if action == "insert_after":
                for np in reversed(new_ps):
                    p.addnext(np)
            else:
                for np in new_ps:
                    p.addprevious(np)
        return len(seen)

    # csere / törlés: hátulról előre, hogy a korábbi találatok helye ne csússzon el
    for p, s, e in reversed(hits):
        text = _para_text(p)
        whole = action == "delete" and not text[:s].strip() and not text[e:].strip()
        if action == "delete" and not whole and s > 0 and text[s - 1] == " " and (e == len(text) or text[e] in " .,;:)"):
            s -= 1   # a törölt részlet előtti szóköz se maradjon ott
        if action == "replace":
            s, e, new_text = _narrow(text[s:e], new_text, s, e)
            if s == e and not new_text:
                continue
            if s == e:   # tiszta beszúrás a bekezdés belsejébe: egy szomszédos karaktert cserélünk önmagára+újra
                if s > 0:
                    s, new_text = s - 1, text[s - 1] + new_text
                else:
                    e, new_text = e + 1, new_text + text[0]
        runs = _isolate(p, s, e)
        if not runs:
            continue
        rpr = _first_rpr(runs)
        last = runs[-1]
        new_runs = _make_runs(new_text, rpr) if action == "replace" and new_text else []
        if rev.tracked:
            _delete_runs(runs, rev)                      # a futások w:del-be kerülnek
            _insert_runs_after(last.getparent(), new_runs, rev)
        else:
            for r in reversed(new_runs):
                last.addnext(r)
            _delete_runs(runs, rev)
        if whole:
            if rev.tracked:
                _mark_paragraph(p, "del", rev)
            else:
                parent = p.getparent()
                if parent.tag == W + "tc" and len(parent.findall(W + "p")) == 1:
                    pass   # a táblázatcella utolsó bekezdése nem törölhető, üresen marad
                else:
                    parent.remove(p)
    return len(hits)


# ---------------------------------------------------------------------------
# Belépési pont
# ---------------------------------------------------------------------------
class _Comments:
    """A Word-megjegyzések része (word/comments.xml): a meglévőhöz fűzünk, vagy újat hozunk létre."""

    def __init__(self, zin):
        self.zin, self.added = zin, 0
        try:
            self.root = etree.fromstring(zin.read(COMMENTS_PART))
            self.exists = True
        except KeyError:
            self.root = etree.Element(W + "comments", nsmap={"w": W_NS})
            self.exists = False
        ids = [int(c.get(W + "id")) for c in self.root.findall(W + "comment") if str(c.get(W + "id", "")).isdigit()]
        self.next_id = max(ids + [-1]) + 1

    def add(self, text, rev):
        c = _w("comment")
        cid = str(self.next_id)
        self.next_id += 1
        c.set(W + "id", cid)
        c.set(W + "author", rev.author)
        c.set(W + "date", rev.date)
        c.set(W + "initials", rev.author[:2])
        for k, line in enumerate(str(text).split("\n")):
            p = _w("p")
            if k == 0:   # az első bekezdés elején a megjegyzés hivatkozási jele (Word így várja)
                ar = _w("r")
                ar.append(_w("annotationRef"))
                p.append(ar)
            for r in _make_runs(line, None):
                p.append(r)
            c.append(p)
        self.root.append(c)
        self.added += 1
        return cid

    def write(self, files):
        """files: {részfájl neve: bájtok} – a módosított részek kerülnek bele."""
        if not self.added:
            return
        files[COMMENTS_PART] = etree.tostring(self.root, xml_declaration=True, encoding="UTF-8", standalone=True)
        if self.exists:
            return
        ct = etree.fromstring(self.zin.read("[Content_Types].xml"))
        ns = ct.nsmap.get(None)
        ov = etree.SubElement(ct, "{%s}Override" % ns)
        ov.set("PartName", "/" + COMMENTS_PART)
        ov.set("ContentType", COMMENTS_TYPE)
        files["[Content_Types].xml"] = etree.tostring(ct, xml_declaration=True, encoding="UTF-8", standalone=True)
        rels_name = "word/_rels/document.xml.rels"
        rels = etree.fromstring(self.zin.read(rels_name))
        rns = rels.nsmap.get(None)
        used = {r.get("Id") for r in rels}
        n = 1
        while f"rIdLara{n}" in used:
            n += 1
        rel = etree.SubElement(rels, "{%s}Relationship" % rns)
        rel.set("Id", f"rIdLara{n}")
        rel.set("Type", COMMENTS_REL)
        rel.set("Target", "comments.xml")
        files[rels_name] = etree.tostring(rels, xml_declaration=True, encoding="UTF-8", standalone=True)


def apply_edits(data, edits, tracked=True, author="Sara"):
    """A módosítások végrehajtása a .docx bájtjain. Visszaad: (új bájtok, jelentés-sorok).
    Ha bármelyik módosítás nem hajtható végre, EditError-t ad az összes hibával, és nem készül fájl."""
    if not isinstance(edits, list) or not edits:
        raise EditError("Nincs megadva egyetlen módosítás sem (edits).")
    try:
        zin = zipfile.ZipFile(io.BytesIO(data))
        root = etree.fromstring(zin.read("word/document.xml"))
    except Exception:
        raise EditError("A fájl nem olvasható Word-dokumentumként (.docx).")
    body = root.find(W + "body")
    rev = _Rev(root, author, tracked)
    rev.comments = _Comments(zin)
    report, errors = [], []
    for n, edit in enumerate(edits, 1):
        if not isinstance(edit, dict):
            errors.append(f"{n}. módosítás: hibás formátum")
            continue
        try:
            count = _apply(body, edit, rev)
            report.append(f"{n}. {edit.get('action')}: {count} helyen")
        except EditError as e:
            errors.append(f"{n}. módosítás ({str(edit.get('find', ''))[:60]!r}): {e}")
    if errors:
        raise EditError("Nem készült fájl, mert ezek a módosítások nem hajthatók végre:\n" + "\n".join(errors)
                        + "\nJavítsd őket (a többit változatlanul hagyd), és hívd meg újra az eszközt az összes "
                        "módosítással.")
    changed = {"word/document.xml": etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)}
    rev.comments.write(changed)
    out = io.BytesIO()
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zout:
        names = set()
        for item in zin.infolist():
            names.add(item.filename)
            zout.writestr(item, changed.get(item.filename, zin.read(item.filename)))
        for name, payload in changed.items():   # új rész (pl. az első megjegyzésnél a comments.xml)
            if name not in names:
                zout.writestr(name, payload)
    return out.getvalue(), report
