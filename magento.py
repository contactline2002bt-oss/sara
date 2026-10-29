"""BISA webshop (Magento 2) – csak olvasó kapcsolat Sara számára: termékek, készlet, rendelések.

A kapcsolat a Magento REST felületén megy, egy integrációs hozzáférési tokennel (Magento admin → System →
Extensions → Integrations). A token csak olvasási jogot kapjon (Catalog, Inventory, Sales); ez a modul
semmit nem ír és nem töröl a webshopban. Streamlittől független.

Titkos beállítások: MAGENTO_URL (pl. https://bisa.hu), MAGENTO_TOKEN, opcionálisan MAGENTO_STORE (alapból
"default") és MAGENTO_STOCK_ID (az eladható mennyiséghez, alapból 1).
"""
from datetime import datetime, timedelta, timezone

import requests

TIMEOUT = 30
MAX_ITEMS = 15


class MagentoError(Exception):
    pass


def configured(cfg):
    return bool(str(cfg.get("url") or "").strip() and str(cfg.get("token") or "").strip())


def _get(cfg, path, params=None):
    base = str(cfg["url"]).rstrip("/")
    store = str(cfg.get("store") or "default").strip() or "default"
    url = f"{base}/rest/{store}/V1/{path.lstrip('/')}"
    try:
        res = requests.get(url, params=params or {}, timeout=TIMEOUT,
                           headers={"Authorization": f"Bearer {cfg['token']}", "Accept": "application/json"})
    except requests.RequestException as e:
        raise MagentoError(f"A webshop nem érhető el: {e}") from e
    if res.status_code == 401:
        raise MagentoError("A webshop elutasította a tokent (401). Ellenőrizd a MAGENTO_TOKEN értékét, és hogy a "
                           "Magento adminban engedélyezve van-e: Stores → Configuration → Services → OAuth → "
                           "Consumer Settings → „Allow OAuth Access Tokens to be used as standalone Bearer tokens” = Yes.")
    if res.status_code == 403:
        raise MagentoError("A tokennek nincs jogosultsága ehhez az adathoz (403). Az integrációnál add meg a "
                           "Catalog, Inventory és Sales olvasási jogot.")
    if res.status_code == 404:
        return None
    if res.status_code >= 400:
        raise MagentoError(f"A webshop hibát jelzett ({res.status_code}): {res.text[:300]}")
    try:
        return res.json()
    except ValueError as e:
        raise MagentoError("A webshop válasza nem értelmezhető (nem JSON).") from e


def _criteria(groups, page_size=MAX_ITEMS, sort=None):
    """searchCriteria paraméterek. groups: [[(mező, érték, feltétel), ...], ...] – csoporton belül VAGY,
    csoportok között ÉS kapcsolat."""
    p = {"searchCriteria[pageSize]": page_size, "searchCriteria[currentPage]": 1}
    for gi, group in enumerate(groups):
        for fi, (field, value, cond) in enumerate(group):
            pre = f"searchCriteria[filter_groups][{gi}][filters][{fi}]"
            p[f"{pre}[field]"], p[f"{pre}[value]"], p[f"{pre}[condition_type]"] = field, value, cond
    if sort:
        p["searchCriteria[sortOrders][0][field]"], p["searchCriteria[sortOrders][0][direction]"] = sort
    return p


def _qty(cfg, sku):
    """(eladható mennyiség, raktári mennyiség, raktáron van-e) – ami nem kérdezhető le, az None."""
    salable = None
    stock_id = str(cfg.get("stock_id") or "1")
    try:   # MSI (Magento 2.3+): az eladható mennyiség a foglalásokkal együtt
        v = _get(cfg, f"inventory/get-product-salable-quantity/{requests.utils.quote(sku, safe='')}/{stock_id}")
        salable = float(v) if v is not None else None
    except (MagentoError, TypeError, ValueError):
        pass
    qty = in_stock = None
    try:
        item = _get(cfg, f"stockItems/{requests.utils.quote(sku, safe='')}") or {}
        qty, in_stock = item.get("qty"), item.get("is_in_stock")
    except MagentoError:
        pass
    return salable, qty, in_stock


def _num(v):
    try:
        f = float(v)
    except (TypeError, ValueError):
        return "–"
    return f"{f:,.0f}".replace(",", " ") if f == int(f) else f"{f:,.2f}".replace(",", " ").replace(".", ",")


def products(cfg, text, with_stock=True):
    q = str(text or "").strip()
    if not q:
        raise MagentoError("Adj meg terméknevet vagy cikkszámot (SKU).")
    data = _get(cfg, "products", {**_criteria([[("name", f"%{q}%", "like"), ("sku", f"%{q}%", "like")]]),
                                  "fields": "items[sku,name,price,status,type_id,visibility],total_count"}) or {}
    items = data.get("items") or []
    if not items:
        return f"Nincs „{q}” nevű vagy cikkszámú termék a BISA webshopban."
    lines = [f"BISA webshop – termékek „{q}” keresésre ({data.get('total_count', len(items))} találat"
             + (f", az első {len(items)}" if data.get("total_count", 0) > len(items) else "") + "):",
             "| Cikkszám | Név | Ár | Állapot | Eladható | Raktáron |", "|---|---|---:|---|---:|---:|"]
    for it in items:
        sku = str(it.get("sku") or "")
        salable, qty, in_stock = _qty(cfg, sku) if with_stock and it.get("type_id") in ("simple", "virtual") \
            else (None, None, None)
        status = "aktív" if str(it.get("status")) == "1" else "inaktív"
        stock = _num(qty) + ("" if in_stock in (None, True) else " (nincs készleten)") if qty is not None else "–"
        lines.append(f"| {sku} | {str(it.get('name') or '').replace('|', '/')} | {_num(it.get('price'))} | {status} | "
                     f"{_num(salable) if salable is not None else '–'} | {stock} |")
    lines.append("_Eladható: a raktári készlet a még nem teljesített rendelések foglalásával csökkentve. "
                 "Összetett (pl. konfigurálható) terméknél a készlet a változatainál látszik._")
    return "\n".join(lines)


STATUS_HU = {"pending": "függőben", "processing": "feldolgozás alatt", "complete": "teljesítve",
             "canceled": "törölve", "closed": "lezárva (visszatérítve)", "holded": "felfüggesztve",
             "pending_payment": "fizetésre vár", "payment_review": "fizetés ellenőrzése", "fraud": "gyanús"}


def orders(cfg, order_number="", days=7, status="", sku=""):
    order_number = str(order_number or "").strip().lstrip("#")
    groups = []
    if order_number:
        groups.append([("increment_id", order_number, "eq")])
    else:
        days = max(1, min(int(days or 7), 92))
        since = (datetime.now(timezone.utc) - timedelta(days=days)).strftime("%Y-%m-%d %H:%M:%S")
        groups.append([("created_at", since, "gteq")])
        if status:
            groups.append([("status", str(status).strip(), "eq")])
    fields = ("items[increment_id,created_at,status,grand_total,order_currency_code,total_qty_ordered,"
              "customer_firstname,customer_lastname,items[sku,name,qty_ordered,parent_item_id]],total_count")
    data = _get(cfg, "orders", {**_criteria(groups, page_size=50 if sku else MAX_ITEMS,
                                            sort=("created_at", "DESC")), "fields": fields}) or {}
    items = data.get("items") or []
    sku = str(sku or "").strip().lower()
    if sku:
        items = [o for o in items if any(sku in str(i.get("sku") or "").lower() for i in o.get("items") or [])]
    if not items:
        return ("Nincs ilyen rendelés." if order_number else
                f"Az elmúlt {days} napban nincs a feltételeknek megfelelő rendelés.")
    head = (f"BISA webshop – {order_number} számú rendelés:" if order_number else
            f"BISA webshop – rendelések az elmúlt {days} napból"
            + (f", állapot: {status}" if status else "") + (f", termék: {sku}" if sku else "")
            + f" ({len(items)} db, a legfrissebb elöl):")
    lines = [head, "| Rendelésszám | Dátum | Vevő | Állapot | Tételek | Végösszeg |", "|---|---|---|---|---|---:|"]
    for o in items[:MAX_ITEMS]:
        lines_items = "; ".join(f"{_num(i.get('qty_ordered'))} × {i.get('name')} ({i.get('sku')})"
                                for i in o.get("items") or [] if not i.get("parent_item_id"))
        buyer = " ".join(x for x in (o.get("customer_lastname"), o.get("customer_firstname")) if x) or "vendég"
        st = STATUS_HU.get(o.get("status"), o.get("status") or "–")
        lines.append(f"| {o.get('increment_id')} | {str(o.get('created_at') or '')[:16]} | {buyer} | {st} | "
                     f"{lines_items.replace('|', '/')} | {_num(o.get('grand_total'))} {o.get('order_currency_code') or ''} |")
    return "\n".join(lines)


TOOLS = [
    {"name": "bisa_termek_kereses",
     "description": (
         "A BISA webshop (Magento) termékeinek élő lekérdezése név- vagy cikkszám-részlet alapján: cikkszám, név, "
         "ár, aktív-e, eladható és raktári mennyiség. Készletre, árra, termékadatra vonatkozó kérdésnél használd. "
         "Csak olvas; a webshopban semmit nem módosít."),
     "input_schema": {"type": "object", "properties": {
         "kereses": {"type": "string", "description": "terméknév vagy cikkszám (részlet is lehet)"}},
         "required": ["kereses"]}},
    {"name": "bisa_rendelesek",
     "description": (
         "A BISA webshop (Magento) rendeléseinek élő lekérdezése: egy rendelés a száma alapján, vagy az elmúlt "
         "napok rendelései (legfeljebb 92 nap), szűrhetően állapotra (pending, processing, complete, canceled…) "
         "és termékre (cikkszám-részlet). Csak olvas. A vevők személyes adatait (név) csak a feladathoz "
         "szükséges mértékben idézd."),
     "input_schema": {"type": "object", "properties": {
         "rendelesszam": {"type": "string", "description": "pl. 000001234; ha megvan, a többi szűrő nem számít"},
         "napok": {"type": "integer", "description": "hány napra visszamenőleg (alapból 7)"},
         "allapot": {"type": "string", "description": "Magento állapotkód, pl. processing"},
         "cikkszam": {"type": "string", "description": "csak az ezt a terméket tartalmazó rendelések"}}}},
]


def handlers(cfg):
    return {
        "bisa_termek_kereses": lambda a: products(cfg, a.get("kereses")),
        "bisa_rendelesek": lambda a: orders(cfg, a.get("rendelesszam"), a.get("napok") or 7,
                                            a.get("allapot") or "", a.get("cikkszam") or ""),
    }
