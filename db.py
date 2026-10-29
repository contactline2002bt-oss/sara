"""Adatbázis-műveletek a Supabase REST felületén (PostgREST) keresztül. Streamlittől független.

Olvasásnál és írásnál is DbError keletkezik, ha a kérés nem sikerül: így a felület mindig meg tudja
mondani, mi a baj, és nem mutat üres listát hiba helyett."""
import time
import base64
from datetime import datetime, timezone

import requests

import search_tools as stx

_url = None
_headers = None
TIMEOUT = 60

# Ezeknek az oszlopoknak és tábláknak létezniük kell (schema.sql hozza létre őket)
REQUIRED = [
    ("projects", "id,name,status,created_at,sort_order,last_opened_at"),
    ("messages", "id,project_id,role,content,created_at"),
    ("project_documents", "id,project_id,name,mime_type,size_bytes,created_at,content_text,search_text,summary"),
    ("project_summaries", "id,project_id,description,content,search_text,created_at,updated_at"),
    ("message_attachments", "id,project_id,message_id,name,mime_type,size_bytes,created_at"),
]


class DbError(Exception):
    pass


# Jogosultsági szabály (RLS) mellett a tiltott módosítás vagy törlés nem ad hibát, csak nulla sort érint,
# ezért ezt külön ellenőrizzük, hogy az alkalmazás ne állítson olyat, ami nem történt meg.
NOT_APPLIED = ("az adatbázis nem hajtotta végre a műveletet (valószínűleg jogosultság hiányzik): ellenőrizd, hogy a "
               "SUPABASE_KEY a service_role (secret) kulcs-e, és futtasd le a schema.sql-t.")


def init(url, key):
    global _url, _headers
    _url = url.rstrip("/")
    _headers = {"apikey": key, "Authorization": f"Bearer {key}", "Content-Type": "application/json"}


def now_iso():
    return datetime.now(timezone.utc).isoformat()


def _explain(res):
    if "42501" in res.text or "row-level security" in res.text:
        return ("a Supabase jogosultsági szabálya elutasította a műveletet. Az alkalmazásnak a service_role "
                "(secret) kulcsot kell használnia: ellenőrizd a SUPABASE_KEY titkos beállítást, majd futtasd le a schema.sql-t.")
    return f"{res.status_code}: {res.text[:300]}"


def _request(method, table, params=None, json=None, prefer="return=representation"):
    try:
        res = requests.request(method, f"{_url}/rest/v1/{table}", headers={**_headers, "Prefer": prefer},
                               params=params, json=json, timeout=TIMEOUT)
    except requests.RequestException as e:
        raise DbError(f"Az adatbázis nem érhető el: {e}") from e
    if res.status_code not in (200, 201, 204):
        raise DbError(_explain(res))
    if prefer == "return=minimal" or not res.content:
        return []
    return res.json()


def _in(values):
    return "in.(" + ",".join(str(v) for v in values) + ")"


def _chunks(values, size=40):
    values = list(values)
    for i in range(0, len(values), size):
        yield values[i:i + size]


def check_schema():
    """Azoknak a tábláknak a listája, amelyekből hiányzik valami (üres lista = minden rendben)."""
    missing = []
    for table, columns in REQUIRED:
        try:
            _request("GET", table, {"select": columns, "limit": 1})
        except DbError as e:
            missing.append(f"{table} ({e})")
    return missing


# --- Projektek -------------------------------------------------------------
def get_projects():
    """Minden projekt a kézi sorrend szerint (ahol nincs sorrend, ott a legújabb elöl)."""
    return _request("GET", "projects", {"select": "*", "order": "sort_order.asc.nullslast,created_at.desc"})


def create_project(name, sort_order, **extra):
    """Új projekt; törzsaktánál extra: kind="torzs", record_type, confidential."""
    rows = _request("POST", "projects", json={"name": name.strip(), "status": "active", "sort_order": sort_order,
                                               "last_opened_at": now_iso(), **extra})
    return rows[0] if rows else None


_records_ok = None
_records_checked = 0.0


def has_records():
    """Megvannak-e a törzsaktákhoz kellő oszlopok és a beállítások táblája (schema.sql 1/e)."""
    global _records_ok, _records_checked
    if _records_ok is None or (_records_ok is False and time.time() - _records_checked > 60):
        _records_checked = time.time()
        try:
            _request("GET", "projects", {"select": "kind,record_type,links", "limit": 1})
            _request("GET", "app_settings", {"select": "key", "limit": 1})
            _records_ok = True
        except DbError:
            _records_ok = False
    return _records_ok


def get_setting(key, default=None):
    rows = _request("GET", "app_settings", {"key": f"eq.{key}", "select": "value"})
    return rows[0]["value"] if rows else default


def set_setting(key, value):
    _request("POST", "app_settings", json={"key": key, "value": value, "updated_at": now_iso()},
             prefer="resolution=merge-duplicates,return=representation")


def update_project(project_id, **fields):
    rows = _request("PATCH", "projects", {"id": f"eq.{project_id}"}, json=fields)
    if not rows:
        raise DbError(NOT_APPLIED)
    return rows[0]


def delete_project(project_id):
    """A projekt és minden hozzá tartozó adat végleges törlése."""
    for table in ("message_attachments", "messages", "project_documents", "project_summaries"):
        _request("DELETE", table, {"project_id": f"eq.{project_id}"}, prefer="return=minimal")
    try:  # a feladatok táblája újabb; ha még nincs létrehozva, a projekt attól még törölhető
        _request("DELETE", "project_tasks", {"project_id": f"eq.{project_id}"}, prefer="return=minimal")
    except DbError:
        pass
    _request("DELETE", "projects", {"id": f"eq.{project_id}"}, prefer="return=minimal")
    if _request("GET", "projects", {"id": f"eq.{project_id}", "select": "id"}):
        raise DbError(NOT_APPLIED)


# --- Üzenetek --------------------------------------------------------------
def get_messages(project_id):
    return _request("GET", "messages", {"project_id": f"eq.{project_id}", "order": "created_at.asc"})


def save_message(project_id, role, content):
    rows = _request("POST", "messages", json={"project_id": project_id, "role": role, "content": content})
    return rows[0] if rows else None


def delete_messages(project_id, message_ids):
    """A megadott üzenetek és a hozzájuk csatolt munkaanyagok törlése."""
    for chunk in _chunks(message_ids):
        _request("DELETE", "message_attachments", {"project_id": f"eq.{project_id}", "message_id": _in(chunk)},
                 prefer="return=minimal")
    for chunk in _chunks(message_ids):
        _request("DELETE", "messages", {"project_id": f"eq.{project_id}", "id": _in(chunk)}, prefer="return=minimal")
        if _request("GET", "messages", {"project_id": f"eq.{project_id}", "id": _in(chunk), "select": "id"}):
            raise DbError(NOT_APPLIED)


# --- Beszélgetésbe csatolt munkaanyagok ---------------------------------------
def get_attachments(project_id):
    """A projekt csatolt munkaanyagai tartalom nélkül."""
    return _request("GET", "message_attachments", {
        "project_id": f"eq.{project_id}", "select": "id,message_id,name,mime_type,size_bytes,created_at",
        "order": "id.asc"})


def save_attachment(project_id, message_id, name, mime_type, data):
    _request("POST", "message_attachments", json={
        "project_id": project_id, "message_id": message_id, "name": name, "mime_type": mime_type,
        "size_bytes": len(data), "content_b64": base64.b64encode(data).decode("ascii"),
    }, prefer="return=minimal")


def get_attachment_content(attachment_id):
    rows = _request("GET", "message_attachments", {"id": f"eq.{attachment_id}", "select": "content_b64"})
    if not rows:
        raise DbError("A csatolt fájl nem található.")
    return rows[0]["content_b64"]


# --- Dokumentumtár ---------------------------------------------------------
DOC_META = "id,project_id,name,mime_type,size_bytes,created_at,summary"
_doc_org_checked = 0.0
_doc_org = None   # van-e kategória és elavult oszlop (schema.sql 1/d); amíg nincs, a rendezés nem érhető el


def has_doc_org():
    """Megvan-e a dokumentumok rendezéséhez kellő két oszlop. Ha nincs (a schema.sql új része még nem futott),
    az alkalmazás ugyanúgy működik, csak rendezés nélkül."""
    global _doc_org, _doc_org_checked
    if _doc_org is None or (_doc_org is False and time.time() - _doc_org_checked > 60):   # SQL-futtatás után magától
        _doc_org_checked = time.time()
        try:
            _request("GET", "project_documents", {"select": "category,obsolete", "limit": 1})
            _doc_org = True
        except DbError:
            _doc_org = False
    return _doc_org


def doc_meta():
    return DOC_META + (",category,obsolete" if has_doc_org() else "")


def set_document_org(doc_id, **fields):
    """Kategória (category) és/vagy elavult jelölés (obsolete) módosítása."""
    fields = {k: v for k, v in fields.items() if k in ("category", "obsolete")}
    if not fields:
        return
    rows = _request("PATCH", "project_documents", {"id": f"eq.{doc_id}"}, json=fields)
    if not rows:
        raise DbError(NOT_APPLIED)


def get_documents(project_id):
    """Az adott projekt dokumentumainak listája (tartalom nélkül), a legújabb elöl."""
    return _request("GET", "project_documents", {
        "project_id": f"eq.{project_id}", "select": doc_meta(), "order": "created_at.desc"})


def get_documents_meta(project_ids):
    """Több projekt dokumentumainak listája kivonattal (tartalom nélkül) – a kereső kollégának."""
    docs = []
    for chunk in _chunks(project_ids):
        docs += _request("GET", "project_documents", {"project_id": _in(chunk), "select": doc_meta(),
                                                      "order": "created_at.desc"})
    return docs


def save_document(project_id, name, mime_type, data, text, summary=None):
    """Dokumentum mentése a kereséshez kinyert szöveggel és a kivonattal együtt (text lehet None, pl. képnél)."""
    _request("POST", "project_documents", json={
        "project_id": project_id, "name": name, "mime_type": mime_type, "size_bytes": len(data),
        "content_b64": base64.b64encode(data).decode("ascii"),
        "content_text": stx.clean_text(text), "search_text": stx.build_search_text(name, text),
        "summary": summary or None,
    }, prefer="return=minimal")


def set_document_summary(doc_id, summary):
    rows = _request("PATCH", "project_documents", {"id": f"eq.{doc_id}"}, json={"summary": summary})
    if not rows:
        raise DbError(NOT_APPLIED)


def get_document_content(doc_id):
    rows = _request("GET", "project_documents", {"id": f"eq.{doc_id}", "select": "content_b64"})
    if not rows:
        raise DbError("A dokumentum nem található.")
    return rows[0]["content_b64"]


def get_document_text(doc_id):
    rows = _request("GET", "project_documents", {"id": f"eq.{doc_id}", "select": "content_text"})
    return (rows[0].get("content_text") or "") if rows else ""


def delete_document(doc_id):
    _request("DELETE", "project_documents", {"id": f"eq.{doc_id}"}, prefer="return=minimal")


def unindexed_documents():
    """A kereséshez még fel nem dolgozott (régebben feltöltött) dokumentumok."""
    return _request("GET", "project_documents", {"search_text": "is.null", "select": "id,name,mime_type"})


def get_textless_documents(project_id):
    """A projekt szöveg nélküli PDF-jei és képei (szkennelt anyag): ezekhez a szövegfelismerés pótolható."""
    return _request("GET", "project_documents", {
        "project_id": f"eq.{project_id}", "select": "id,project_id,name,mime_type,size_bytes",
        "mime_type": "in.(application/pdf,image/png,image/jpeg)",
        "or": "(content_text.is.null,content_text.eq.)", "order": "created_at.desc"})


def set_document_text(doc_id, name, text):
    rows = _request("PATCH", "project_documents", {"id": f"eq.{doc_id}"}, json={
        "content_text": stx.clean_text(text), "search_text": stx.build_search_text(name, text)})
    if not rows:
        raise DbError(NOT_APPLIED)


# --- Összefoglalók ---------------------------------------------------------
def get_summaries(project_id):
    """A projekt összefoglalói időrendben (a legrégebbi elöl)."""
    return _request("GET", "project_summaries", {
        "project_id": f"eq.{project_id}", "select": "id,project_id,description,content,created_at,updated_at",
        "order": "created_at.asc"})


def add_summary(project_id, description, content):
    rows = _request("POST", "project_summaries", json={
        "project_id": project_id, "description": description, "content": content,
        "search_text": stx.build_search_text(description, content)})
    if not rows:
        raise DbError("az összefoglaló mentését az adatbázis nem igazolta vissza.")
    return rows[0]


def update_summary(summary_id, description, content):
    rows = _request("PATCH", "project_summaries", {"id": f"eq.{summary_id}"}, json={
        "description": description, "content": content, "updated_at": now_iso(),
        "search_text": stx.build_search_text(description, content)})
    if not rows:
        raise DbError(NOT_APPLIED)


def delete_summary(summary_id):
    _request("DELETE", "project_summaries", {"id": f"eq.{summary_id}"}, prefer="return=minimal")
    if _request("GET", "project_summaries", {"id": f"eq.{summary_id}", "select": "id"}):
        raise DbError(NOT_APPLIED)


# --- Feladatok -------------------------------------------------------------
TASK_STATUSES = ("open", "waiting", "done", "obsolete", "cancelled")


# --- Tokenfelhasználás -----------------------------------------------------
def log_usage(project_id, kind, model, usage):
    """A modellhívás tokenjeinek naplózása. Soha nem dob hibát: ha a tábla még nincs meg, csendben kimarad."""
    try:
        stu = getattr(usage, "server_tool_use", None)
        row = {
            "project_id": project_id, "kind": kind, "model": model,
            "input_tokens": int(getattr(usage, "input_tokens", 0) or 0),
            "cache_write_tokens": int(getattr(usage, "cache_creation_input_tokens", 0) or 0),
            "cache_read_tokens": int(getattr(usage, "cache_read_input_tokens", 0) or 0),
            "output_tokens": int(getattr(usage, "output_tokens", 0) or 0),
            "web_searches": int(getattr(stu, "web_search_requests", 0) or 0) if stu else 0,
        }
        _request("POST", "usage_log", json=row, prefer="return=minimal")
    except Exception:
        pass


def get_project_usage(project_id, since_iso):
    """A projekt chathívásai egy időpont óta (a beszélgetés méretének kijelzéséhez)."""
    return _request("GET", "usage_log", {"select": "*", "project_id": f"eq.{project_id}", "kind": "eq.chat",
                                         "created_at": f"gte.{since_iso}", "order": "created_at.asc"})


def get_usage(since_iso=None):
    params = {"select": "*", "order": "created_at.asc"}
    if since_iso:
        params["created_at"] = f"gte.{since_iso}"
    return _request("GET", "usage_log", params)


def get_tasks(project_id):
    """A projekt feladatai a felvétel sorrendjében. DbError, ha a tábla még nincs létrehozva."""
    return _request("GET", "project_tasks", {"project_id": f"eq.{project_id}", "select": "*",
                                             "order": "created_at.asc,id.asc"})


def add_task(project_id, topic, title, note=""):
    rows = _request("POST", "project_tasks", json={
        "project_id": project_id, "topic": topic.strip(), "title": title.strip(), "note": (note or "").strip(),
        "status": "open"})
    if not rows:
        raise DbError("a feladat mentését az adatbázis nem igazolta vissza.")
    return rows[0]


def update_task(task_id, **fields):
    if "status" in fields and fields["status"] not in TASK_STATUSES:
        raise DbError(f"ismeretlen állapot: {fields['status']}")
    rows = _request("PATCH", "project_tasks", {"id": f"eq.{task_id}"}, json={**fields, "updated_at": now_iso()})
    if not rows:
        raise DbError(NOT_APPLIED)
    return rows[0]


def delete_task(task_id):
    _request("DELETE", "project_tasks", {"id": f"eq.{task_id}"}, prefer="return=minimal")
    if _request("GET", "project_tasks", {"id": f"eq.{task_id}", "select": "id"}):
        raise DbError(NOT_APPLIED)


# --- Keresés ---------------------------------------------------------------
def search(project_ids, stems, limit=60):
    """Előszűrés: dokumentumok és összefoglalók, amelyek szövegében minden tő előfordul (a pontos, szavankénti
    egyeztetést a search_tools végzi). Visszaad: (dokumentumok, összefoglalók); a dokumentumoknál a kinyert
    szöveg is benne van."""
    if not project_ids or not stems:
        return [], []
    condition = "(" + ",".join(f"search_text.ilike.*{s}*" for s in stems) + ")"
    docs, sums = [], []
    for chunk in _chunks(project_ids):
        base = {"project_id": _in(chunk), "and": condition, "order": "created_at.desc", "limit": limit}
        docs += _request("GET", "project_documents", {**base, "select": doc_meta() + ",content_text"})
        sums += _request("GET", "project_summaries", {
            **base, "select": "id,project_id,description,content,created_at,updated_at"})
    return docs, sums


def rename_document(doc_id, name):
    rows = _request("PATCH", "project_documents", {"id": f"eq.{doc_id}"}, json={"name": name})
    if not rows:
        raise DbError(NOT_APPLIED)
