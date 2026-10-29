"""Cég adószámának érvényessége a NAV Online Számla rendszer /queryTaxpayer műveletével (v3.0).

A NAV leírása szerint a taxpayerValidity akkor true, ha a lekérdezett adószám létezik és érvényes; nem érvényes
(pl. felfüggesztett, törölt) vagy nem létező adószámra false. Létező adószámnál a név, a teljes adószám, a
gazdasági forma és a címek is visszajönnek.

Hozzáférés: a cég Online Számla technikai felhasználója (onlineszamla.nav.gov.hu → Felhasználók → Technikai
felhasználó): NAV_LOGIN, NAV_PASSWORD, NAV_SIGNKEY és a cég adószámának első 8 jegye (NAV_TAXNUMBER) a
Streamlit titkos beállításaiban. A kérés aláírása a leírás 1.5.2. pontja szerint:
requestSignature = SHA3-512(requestId + timestamp[yyyyMMddHHmmss, UTC] + aláírókulcs), nagybetűsen.
"""
import hashlib
import re
import secrets
import urllib.request
from datetime import datetime, timezone
from xml.sax.saxutils import escape

from lxml import etree

PROD = "https://api.onlineszamla.nav.gov.hu/invoiceService/v3/queryTaxpayer"
TEST = "https://api-test.onlineszamla.nav.gov.hu/invoiceService/v3/queryTaxpayer"
NS = {"api": "http://schemas.nav.gov.hu/OSA/3.0/api", "common": "http://schemas.nav.gov.hu/NTCA/1.0/common",
      "base": "http://schemas.nav.gov.hu/OSA/3.0/base"}
INCORPORATION = {"ORGANIZATION": "gazdasági társaság", "SELF_EMPLOYED": "egyéni vállalkozó",
                 "TAXABLE_PERSON": "adószámos magánszemély"}

TOOL = {
    "name": "adoszam_ellenorzes",
    "description": (
        "Cég (vagy egyéni vállalkozó) adószámának élő ellenőrzése a NAV-nál (Online Számla adózó-lekérdezés): "
        "érvényes-e az adószám (nem felfüggesztett, nem törölt), mi a NAV szerinti név és székhely. Szerződés "
        "előtt, cégkivonat mellé használd. Csak az adószámot add meg (8 vagy 11 jegy)."),
    "input_schema": {"type": "object", "properties": {"adoszam": {"type": "string"}}, "required": ["adoszam"]},
}


class NavError(Exception):
    pass


def configured(cfg):
    return all(str(cfg.get(k) or "").strip() for k in ("login", "password", "signkey", "taxnumber"))


def build_request(cfg, tax_number, now=None, request_id=None):
    now = now or datetime.now(timezone.utc)
    request_id = request_id or ("SARA" + secrets.token_hex(8).upper())[:30]
    ts = now.strftime("%Y-%m-%dT%H:%M:%S.") + f"{now.microsecond // 1000:03d}Z"
    mask = now.strftime("%Y%m%d%H%M%S")
    pw = hashlib.sha512(str(cfg["password"]).encode("utf-8")).hexdigest().upper()
    sig = hashlib.sha3_512((request_id + mask + str(cfg["signkey"])).encode("utf-8")).hexdigest().upper()
    sw = cfg.get("software") or {}
    xml = f"""<?xml version="1.0" encoding="UTF-8"?>
<QueryTaxpayerRequest xmlns:common="{NS['common']}" xmlns="{NS['api']}">
 <common:header><common:requestId>{request_id}</common:requestId><common:timestamp>{ts}</common:timestamp>
  <common:requestVersion>3.0</common:requestVersion><common:headerVersion>1.0</common:headerVersion></common:header>
 <common:user><common:login>{escape(str(cfg['login']))}</common:login>
  <common:passwordHash cryptoType="SHA-512">{pw}</common:passwordHash>
  <common:taxNumber>{escape(str(cfg['taxnumber'])[:8])}</common:taxNumber>
  <common:requestSignature cryptoType="SHA3-512">{sig}</common:requestSignature></common:user>
 <software><softwareId>{escape(sw.get('id', 'HU' + str(cfg['taxnumber'])[:8] + 'SARA00001'))[:18]}</softwareId>
  <softwareName>Sara uzleti asszisztens</softwareName><softwareOperation>ONLINE_SERVICE</softwareOperation>
  <softwareMainVersion>1.0</softwareMainVersion><softwareDevName>{escape(sw.get('dev', 'Sara'))}</softwareDevName>
  <softwareDevContact>{escape(sw.get('contact', 'sara@example.hu'))}</softwareDevContact>
  <softwareDevCountryCode>HU</softwareDevCountryCode></software>
 <taxNumber>{tax_number}</taxNumber>
</QueryTaxpayerRequest>"""
    return xml.encode("utf-8")


def parse_response(data):
    root = etree.fromstring(data)
    func = root.findtext(".//common:result/common:funcCode", namespaces=NS)
    if func and func != "OK":
        msg = root.findtext(".//common:result/common:message", namespaces=NS) or ""
        code = root.findtext(".//common:result/common:errorCode", namespaces=NS) or ""
        raise NavError(f"A NAV hibát adott: {code} {msg}".strip())
    valid = root.findtext("api:taxpayerValidity", namespaces=NS)
    out = {"valid": None if valid is None else valid.strip().lower() == "true",
           "info_date": root.findtext("api:infoDate", namespaces=NS)}
    data_el = root.find("api:taxpayerData", NS)
    if data_el is not None:
        out["name"] = data_el.findtext("api:taxpayerName", namespaces=NS)
        out["short_name"] = data_el.findtext("api:taxpayerShortName", namespaces=NS)
        tn = data_el.find("api:taxNumberDetail", NS)
        if tn is not None:
            parts = [tn.findtext(f"base:{t}", namespaces=NS) or "" for t in ("taxpayerId", "vatCode", "countyCode")]
            out["tax_number"] = "-".join(p for p in parts if p)
        out["incorporation"] = INCORPORATION.get(data_el.findtext("api:incorporation", namespaces=NS) or "", "")
        addrs = []
        for item in data_el.iterfind(".//api:taxpayerAddressItem", NS):
            kind = item.findtext("api:taxpayerAddressType", namespaces=NS)
            a = item.find("api:taxpayerAddress", NS)
            if a is None:
                continue
            g = lambda t: a.findtext(f"base:{t}", namespaces=NS) or ""   # noqa: E731
            line = " ".join(x for x in (g("postalCode"), g("city") + ",", g("streetName"), g("publicPlaceCategory"),
                                        g("number")) if x.strip(" ,"))
            addrs.append(({"HQ": "székhely", "SITE": "telephely", "BRANCH": "fióktelep"}.get(kind, kind), line))
        out["addresses"] = addrs
    return out


def check_digit_ok(d8):
    """A törzsszám 8. jegye ellenőrző szám: az első 7 jegy 9-7-3-1 súlyozott összegének kiegészítése 10-re."""
    total = sum(int(d) * w for d, w in zip(d8[:7], (9, 7, 3, 1, 9, 7, 3)))
    return (10 - total % 10) % 10 == int(d8[7])


def query(cfg, tax_number, opener=None, test=False):
    full = re.sub(r"\D", "", str(tax_number or ""))
    if len(full) not in (8, 11):
        raise NavError("Az adószám 8 vagy 11 számjegyből áll (pl. 12345678 vagy 12345678-2-42).")
    digits = full[:8]
    if not check_digit_ok(digits):
        raise NavError(f"Ilyen adószám nem létezik: a(z) {digits} törzsszám ellenőrző számjegye hibás "
                       "(valószínűleg elírás).")
    if not configured(cfg):
        raise NavError("Az adószám-ellenőrzéshez a cég NAV Online Számla technikai felhasználója kell "
                       "(NAV_LOGIN, NAV_PASSWORD, NAV_SIGNKEY, NAV_TAXNUMBER a titkos beállításokban).")
    req = urllib.request.Request(TEST if test else PROD, data=build_request(cfg, digits),
                                 headers={"Content-Type": "application/xml", "Accept": "application/xml"})
    try:
        with (opener or urllib.request.urlopen)(req, timeout=30) as resp:
            body = resp.read()
    except urllib.error.HTTPError as e:
        body = e.read()
    r = parse_response(body)
    r["requested"] = full
    return digits, r


def report(digits, r):
    if r.get("valid") is True:
        head = f"✅ A(z) {r.get('tax_number') or digits} adószám a NAV szerint ÉRVÉNYES."
    elif r.get("valid") is False:
        head = (f"⛔ A(z) {digits} adószám a NAV szerint NEM ÉRVÉNYES (nem létezik, vagy felfüggesztett / törölt). "
                "Szerződéskötés előtt tisztázni kell.")
    else:
        head = f"A NAV nem adott egyértelmű választ a(z) {digits} adószámra."
    lines = [head]
    got = re.sub(r"\D", "", r.get("tax_number") or "")
    if r.get("valid") is True and len(r.get("requested") or "") == 11 and len(got) == 11 and got != r["requested"]:
        q = r["requested"]
        lines = [f"⚠️ A megadott {q[:8]}-{q[8]}-{q[9:]} adószám ebben a formában NEM LÉTEZIK: a törzsszám "
                 f"({digits}) érvényes, de a NAV szerint a teljes adószám {r['tax_number']}. "
                 "Ellenőrizni kell, melyik a helyes (elírás vagy régi adat)."]
    if r.get("name"):
        lines.append(f"NAV szerinti név: {r['name']}" + (f" ({r['short_name']})" if r.get("short_name") else ""))
    if r.get("incorporation"):
        lines.append(f"Forma: {r['incorporation']}")
    for kind, addr in r.get("addresses") or []:
        lines.append(f"{kind.capitalize()}: {addr}")
    if r.get("info_date"):
        lines.append(f"Adatok utolsó változása: {r['info_date'][:10].replace('-', '.')}.")
    try:
        from zoneinfo import ZoneInfo
        now = datetime.now(ZoneInfo("Europe/Budapest"))
    except Exception:
        now = datetime.now()
    lines.append(f"Lekérdezve a NAV Online Számla rendszeréből: {now:%Y.%m.%d. %H:%M}.")
    return "\n".join(lines)
