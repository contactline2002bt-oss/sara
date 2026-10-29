# Sara asszisztens

Üzleti asszisztens (Streamlit + Supabase + Claude).

- **SARA** – a fő asszisztens, alapos munka (Claude Sonnet 5.5)
- **JEAN** – Sara olcsóbb helyettese egyszerű kérdésekhez, adatgyűjtéshez (Claude Haiku 5.5)

Az oldalsávban a „Ki dolgozik?” kapcsolóval lehet váltani köztük. Jean napján a felület zsályazöld.

## Telepítés

1. **Adatbázis (Supabase):** hozz létre egy új projektet, majd SQL Editor → illeszd be a `schema.sql` teljes tartalmát → Run.
2. **Kód:** töltsd fel ezt a mappát egy (privát) GitHub-tárolóba.
3. **Futtatás (Streamlit Community Cloud):** New app → válaszd ki a tárolót, a fő fájl: `app.py`.
4. **Titkos beállítások** (Streamlit → App settings → Secrets):

```toml
SUPABASE_URL = "https://<projekt>.supabase.co"
SUPABASE_KEY = "<Supabase kulcs>"
ANTHROPIC_API_KEY = "sk-ant-..."
APP_PASSWORD = "<belépési jelszó>"

# nem kötelezők
# SARA_MODEL = "claude-sonnet-5-5"
# JEAN_MODEL = "claude-haiku-5-5"
# ASSZISZTENS_ALAP = "SARA"        # belépéskor ki dolgozik: "SARA" vagy "JEAN"
# SARA_SZEMELY = "..."             # személyes adatok Sarának (név, megszólítás)
# NAV_LOGIN / NAV_PASSWORD / NAV_SIGNKEY / NAV_TAXNUMBER   # NAV Online Számla adószám-ellenőrzés
# MAGENTO_URL = "https://bisa-webshop.hu"   # BISA webshop: termékek, készlet, rendelések (csak olvasás)
# MAGENTO_TOKEN = "..."                     # Magento integrációs hozzáférési token
# MAGENTO_STORE = "default"   MAGENTO_STOCK_ID = "1"
```

A `secrets.toml` soha ne kerüljön a tárolóba (a `.gitignore` kizárja).
