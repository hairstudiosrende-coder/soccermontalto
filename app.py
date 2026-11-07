# -*- coding: utf-8 -*-
"""
GESTIONE SQUADRA – prototipo
Streamlit + Google Sheets
Gestione per STAGIONE (es. 2026/2027): rosa e staff, contratti, pagamenti,
cassa, fornitori e fatture, sponsor, scadenze, ruoli e utenti.
"""

import hashlib
from html import escape
import io
import os
import uuid
from datetime import date, datetime, timedelta
from zoneinfo import ZoneInfo

import altair as alt
import gspread
import pandas as pd
import streamlit as st

TZ = ZoneInfo("Europe/Rome")

st.set_page_config(page_title="Gestione Squadra", page_icon="⚽", layout="wide")

# ════════════════════════════════════════════════════════════════════
# STRUTTURA DEI FOGLI
# ════════════════════════════════════════════════════════════════════
SHEETS = {
    "Impostazioni": ["chiave", "valore"],
    "Stagioni": ["stagione", "data_inizio", "data_fine", "attiva"],
    "Ruoli": ["ruolo", "pagine", "puo_modificare", "vede_importi"],
    "Utenti": ["username", "password_hash", "nome", "ruolo", "attivo"],
    "Categorie": ["tipo", "nome"],
    "Tesserati": ["id", "stagione", "tipo", "reparto", "cognome", "nome", "anno_nascita",
                  "provenienza", "telefono", "email", "iban", "compenso_stagione",
                  "stato", "note"],
    "Voci_Contratto": ["id", "stagione", "tesserato_id", "voce", "importo", "note"],
    "Pagamenti": ["id", "data", "stagione", "tesserato_id", "mese_rif", "voce", "importo",
                  "metodo", "note", "utente", "mov_id", "migrato"],   # archivio: dal 2° rilascio tutto è in Movimenti
    "Fornitori": ["id", "nome", "categoria", "piva", "telefono", "email", "iban", "note"],
    "Fatture": ["id", "stagione", "fornitore_id", "numero", "data", "scadenza", "categoria",
                "descrizione", "importo", "pagato", "stato", "note", "tesserato_id", "nel_compenso"],
    "Movimenti": ["id", "data", "stagione", "tipo", "categoria", "descrizione", "importo",
                  "metodo", "rif_tipo", "rif_id", "utente", "note", "tesserato_id", "nel_compenso",
                  "beneficiario", "mese_rif", "destinatario", "voce"],
    "Sponsor": ["id", "stagione", "nome", "referente", "telefono", "importo_concordato", "note"],
    "Scadenze": ["id", "stagione", "tipo", "descrizione", "tesserato_id", "data_scadenza",
                 "stato", "note"],
    "Log": ["timestamp", "utente", "azione", "foglio", "dettaglio"],
}
NUM_COLS = {"compenso_stagione", "importo", "pagato", "importo_concordato"}

PAGINE = ["Panoramica", "Rosa & Staff", "Cassa & Pagamenti", "Fornitori & Fatture",
          "Sponsor", "Scadenze", "Impostazioni"]
# nomi di pagina delle versioni precedenti (ruoli già salvati nel foglio)
PAGINE_ALIAS = {"Pagamenti": "Cassa & Pagamenti", "Cassa": "Cassa & Pagamenti"}
ICONE = {"Panoramica": "🏟️", "Rosa & Staff": "👥", "Cassa & Pagamenti": "💰",
         "Fornitori & Fatture": "🧾", "Sponsor": "🤝", "Scadenze": "📅", "Impostazioni": "⚙️"}

# Tipi di categoria (tutte modificabili da Impostazioni → Categorie)
TIPI_CATEGORIA = {
    "TESSERATO": "Tipo tesserato",
    "REPARTO": "Reparto / Area",
    "VOCE_CONTRATTO": "Voce contratto",
    "ENTRATA": "Categoria entrata",
    "USCITA": "Categoria uscita",
    "SCADENZA": "Tipo scadenza",
    "METODO": "Metodo di pagamento",
    "FORNITORE": "Categoria fornitore",
}
CATEGORIE_DEFAULT = {
    "TESSERATO": ["Giocatore", "Allenatore", "Vice allenatore", "Preparatore atletico",
                  "Preparatore portieri", "Fisioterapista / Massaggiatore", "Medico",
                  "Team manager", "Magazziniere", "Dirigente"],
    "REPARTO": ["Portieri", "Difensori", "Centrocampisti", "Attaccanti", "Staff tecnico",
                "Staff medico", "Dirigenza"],
    "VOCE_CONTRATTO": ["Compenso", "Alloggio", "Vitto", "Volo / Viaggi", "Procura", "Premio",
                       "Spese mediche", "Rimborso spese"],
    "ENTRATA": ["Sponsor", "Contributo Comune / Regione", "Contributo Federazione / Lega",
                "Quote soci", "Biglietteria / Abbonamenti", "Cessione giocatori",
                "Bar / Merchandising", "Eventi / Cene", "Donazioni", "Versamento presidente",
                "Altro"],
    "USCITA": ["Compensi giocatori", "Compensi staff", "Alloggi", "Vitto / Ristoranti",
               "Trasferte / Viaggi", "Voli", "Procuratori", "Materiale sportivo",
               "Abbigliamento", "Attrezzatura / Tecnologia", "Iscrizioni / Tesseramenti",
               "Arbitri / Ambulanza", "Affitto campo", "Utenze", "Spese mediche",
               "Assicurazioni", "Commercialista", "Altro"],
    "SCADENZA": ["Visita medica", "Tesseramento", "Iscrizione campionato", "Assicurazione",
                 "Scadenza contratto", "Pagamento", "Altro"],
    "METODO": ["Contanti", "Bonifico", "Carta / POS", "Assegno"],
    "FORNITORE": ["Materiale sportivo", "Ristorazione", "Alloggi", "Trasporti",
                  "Servizi medici", "Tecnologia", "Stampa / Grafica", "Altro"],
}
# Voce contratto → categoria di uscita in cassa
VOCE_TO_USCITA = {"Alloggio": "Alloggi", "Vitto": "Vitto / Ristoranti", "Volo / Viaggi": "Voli",
                  "Procura": "Procuratori", "Spese mediche": "Spese mediche"}

MESI = ["Gennaio", "Febbraio", "Marzo", "Aprile", "Maggio", "Giugno", "Luglio", "Agosto",
        "Settembre", "Ottobre", "Novembre", "Dicembre"]


# ════════════════════════════════════════════════════════════════════
# UTILITÀ
# ════════════════════════════════════════════════════════════════════
def oggi():
    return datetime.now(TZ).date()


def adesso():
    return datetime.now(TZ).strftime("%Y-%m-%d %H:%M:%S")


def nuovo_id(prefix):
    return f"{prefix}{uuid.uuid4().hex[:8]}"


def to_num(v):
    if v is None or v == "":
        return 0.0
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).replace("€", "").replace(" ", "").strip()
    if "," in s:
        s = s.replace(".", "").replace(",", ".")
    try:
        return float(s)
    except ValueError:
        return 0.0


def parse_date(v):
    if v in (None, ""):
        return None
    if isinstance(v, (int, float)):  # data seriale di Google Sheets
        return (datetime(1899, 12, 30) + timedelta(days=float(v))).date()
    s = str(v).strip()[:10]
    for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y"):
        try:
            return datetime.strptime(s, fmt).date()
        except ValueError:
            pass
    return None


def fmt_data(v):
    d = parse_date(v)
    return d.strftime("%d/%m/%Y") if d else ""


def euro(x):
    s = f"{to_num(x):,.2f}"
    return "€ " + s.replace(",", "X").replace(".", ",").replace("X", ".")


def hash_pw(pw, salt=None):
    salt = salt or uuid.uuid4().hex[:16]
    h = hashlib.pbkdf2_hmac("sha256", pw.encode(), salt.encode(), 120_000).hex()
    return f"{salt}${h}"


def check_pw(pw, stored):
    try:
        salt, _ = str(stored).split("$", 1)
    except ValueError:
        return False
    return hash_pw(pw, salt) == stored


def stagione_da_anno(anno):
    return f"{anno}/{anno + 1}"


def stagione_da_data(d):
    return stagione_da_anno(d.year if d.month >= 7 else d.year - 1)


def to_excel(dfs: dict):
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as xw:
        for nome, df in dfs.items():
            df.to_excel(xw, sheet_name=nome[:31], index=False)
    return buf.getvalue()


# ════════════════════════════════════════════════════════════════════
# GOOGLE SHEETS
# ════════════════════════════════════════════════════════════════════
def ripara_chiave(pk: str) -> str:
    """Ricostruisce la private_key qualunque sia il formato incollato nei Secrets
    (\\n letterali, a capo veri, spazi, virgolette, righe spezzate)."""
    import re
    import textwrap
    s = str(pk).strip().strip('"').strip("'")
    s = s.replace("\\\\n", "\n").replace("\\n", "\n").replace("\r", "")
    corpo = re.sub(r"-----(BEGIN|END)[A-Z ]*PRIVATE KEY-----", "", s)
    corpo = re.sub(r"[^A-Za-z0-9+/=]", "", corpo)
    return ("-----BEGIN PRIVATE KEY-----\n" + "\n".join(textwrap.wrap(corpo, 64))
            + "\n-----END PRIVATE KEY-----\n")


def credenziali():
    """Legge le credenziali: o il JSON intero (gcp_json) o il blocco [gcp_service_account]."""
    import json
    if "gcp_json" in st.secrets:
        creds = json.loads(str(st.secrets["gcp_json"]), strict=False)
    else:
        creds = dict(st.secrets["gcp_service_account"])
    creds["private_key"] = ripara_chiave(creds.get("private_key", ""))
    return creds


def diagnosi_chiave():
    """Informazioni sulla chiave senza mostrarla."""
    try:
        import base64
        pk = credenziali()["private_key"]
        corpo = "".join(pk.splitlines()[1:-1])
        righe = [f"Caratteri della chiave: **{len(corpo)}** (una chiave completa ne ha circa 1.620–1.630)",
                 f"Inizia con: `{corpo[:3]}…`  (deve iniziare con `MII`)",
                 f"Finisce con: `…{corpo[-3:]}`"]
        try:
            base64.b64decode(corpo, validate=True)
            righe.append("Formato base64: ✅ corretto")
        except Exception:
            righe.append("Formato base64: ❌ la chiave è tagliata o contiene caratteri in più")
        return "\n\n".join(righe)
    except Exception as e:
        return f"Impossibile leggere i Secrets: {e}"


def sheet_id():
    for k in ("SHEET_ID", "sheet_id", "gsheet_id", "GSHEET_ID"):
        if k in st.secrets:
            return str(st.secrets[k]).strip()
    return ""


@st.cache_resource(show_spinner=False)
def book():
    gc = gspread.service_account_from_dict(credenziali())
    sid = sheet_id()
    if sid.startswith("http"):
        return gc.open_by_url(sid)
    return gc.open_by_key(sid)


@st.cache_resource(show_spinner="Preparo il foglio Google…")
def struttura():
    """Controlla TUTTI i fogli con 2 sole richieste e crea solo ciò che manca.
    Restituisce (worksheets, intestazioni)."""
    sh = book()
    esistenti = {ws.title: ws for ws in sh.worksheets()}
    presenti = [n for n in SHEETS if n in esistenti]
    intest = {}
    if presenti:
        res = sh.values_batch_get(ranges=[f"'{n}'!1:1" for n in presenti])
        for n, vr in zip(presenti, res.get("valueRanges", [])):
            v = vr.get("values", [])
            intest[n] = [str(h).strip() for h in v[0]] if v else []
    for nome, cols in SHEETS.items():
        if nome not in esistenti:
            ws = sh.add_worksheet(title=nome, rows=200, cols=max(len(cols), 8))
            ws.update(range_name="A1", values=[cols], value_input_option="RAW")
            esistenti[nome], intest[nome] = ws, list(cols)
            continue
        header = intest.get(nome, [])
        mancanti = [c for c in cols if c not in header]
        if mancanti:
            nuovo = header + mancanti
            ws = esistenti[nome]
            if ws.col_count < len(nuovo):
                ws.add_cols(len(nuovo) - ws.col_count)
            ws.update(range_name="A1", values=[nuovo], value_input_option="RAW")
            intest[nome] = nuovo
    return esistenti, intest


def worksheets():
    return struttura()[0]


def intestazione(nome):
    return struttura()[1][nome]


@st.cache_data(ttl=300, show_spinner=False)
def leggi_tutto():
    """Una sola chiamata a Google per leggere tutti i fogli."""
    struttura()
    nomi = list(SHEETS.keys())
    res = book().values_batch_get(
        ranges=[f"'{n}'" for n in nomi],
        params={"valueRenderOption": "UNFORMATTED_VALUE"},
    )
    out = {}
    for nome, vr in zip(nomi, res.get("valueRanges", [])):
        vals = vr.get("values", [])
        cols = SHEETS[nome]
        if not vals:
            out[nome] = pd.DataFrame(columns=cols)
            continue
        header = [str(h).strip() for h in vals[0]]
        w = len(header)
        righe = [(list(r) + [""] * w)[:w] for r in vals[1:]
                 if any(str(c).strip() for c in r)]
        tab = pd.DataFrame(righe, columns=header)
        for c in cols:
            if c not in tab.columns:
                tab[c] = ""
        tab = tab[cols].copy()
        for c in cols:
            if c in NUM_COLS:
                tab[c] = tab[c].apply(to_num).astype(float)
            else:
                tab[c] = tab[c].apply(lambda x: "" if x is None else
                                      (str(int(x)) if isinstance(x, float) and x.is_integer()
                                       else str(x)).strip()).astype(object)
        out[nome] = tab
    return out


def df(nome):
    return leggi_tutto()[nome].copy()


def _riga_valori(nome, rec):
    return [("" if rec.get(c) is None else rec.get(c, "")) for c in intestazione(nome)]


def _log_riga(azione, foglio, dettaglio):
    return [adesso(), (st.session_state.get("utente") or {}).get("username", "-"),
            azione, foglio, str(dettaglio)[:300]]


def aggiungi(nome, rec, log=True):
    worksheets()[nome].append_row(_riga_valori(nome, rec), value_input_option="RAW",
                                  insert_data_option="INSERT_ROWS", table_range="A1")
    if log:
        scrivi_log("Aggiunto", nome, rec.get("id") or rec.get("nome") or rec.get("chiave", ""))
    leggi_tutto.clear()


def aggiungi_molti(nome, recs):
    if recs:
        worksheets()[nome].append_rows([_riga_valori(nome, r) for r in recs],
                                       value_input_option="RAW",
                                       insert_data_option="INSERT_ROWS", table_range="A1")
        leggi_tutto.clear()


def _trova_riga(nome, chiave_col, valore, extra=None):
    """Una sola richiesta: legge solo la colonna chiave (o tutto se servono più colonne)."""
    ws = worksheets()[nome]
    header = intestazione(nome)
    if extra:
        righe = ws.get_all_values()
        for i, riga in enumerate(righe[1:], start=2):
            riga = (riga + [""] * len(header))[:len(header)]
            if riga[header.index(chiave_col)].strip() == str(valore).strip() and all(
                    riga[header.index(k)].strip() == str(v).strip() for k, v in extra.items()):
                return i
        return None
    colonna = ws.col_values(header.index(chiave_col) + 1)
    for i, v in enumerate(colonna[1:], start=2):
        if str(v).strip() == str(valore).strip():
            return i
    return None


def aggiorna(nome, valore, campi: dict, chiave_col="id"):
    """Aggiorna solo le celle cambiate: 1 lettura + 1 scrittura."""
    r = _trova_riga(nome, chiave_col, valore)
    if r is None:
        return False
    header = intestazione(nome)
    blocchi = [{"range": gspread.utils.rowcol_to_a1(r, header.index(k) + 1),
                "values": [["" if v is None else v]]}
               for k, v in campi.items() if k in header]
    if blocchi:
        worksheets()[nome].batch_update(blocchi, value_input_option="RAW")
    scrivi_log("Modificato", nome, f"{valore}: {', '.join(campi.keys())}")
    leggi_tutto.clear()
    return True


def aggiorna_molti(nome, modifiche: dict, chiave_col="id"):
    """{valore_chiave: {campo: valore}} → 1 lettura + 1 scrittura per tutte le righe."""
    if not modifiche:
        return
    ws = worksheets()[nome]
    header = intestazione(nome)
    colonna = ws.col_values(header.index(chiave_col) + 1)
    riga_di = {str(v).strip(): i for i, v in enumerate(colonna, start=1) if i > 1}
    blocchi = []
    for chiave, campi in modifiche.items():
        r = riga_di.get(str(chiave).strip())
        if r is None:
            continue
        blocchi += [{"range": gspread.utils.rowcol_to_a1(r, header.index(k) + 1),
                     "values": [["" if v is None else v]]} for k, v in campi.items() if k in header]
    if blocchi:
        ws.batch_update(blocchi, value_input_option="RAW")
    scrivi_log("Modificato in blocco", nome, f"{len(modifiche)} righe")
    leggi_tutto.clear()


def elimina(nome, valore, chiave_col="id", extra=None):
    r = _trova_riga(nome, chiave_col, valore, extra)
    if r is None:
        return False
    worksheets()[nome].delete_rows(r)
    scrivi_log("Eliminato", nome, str(valore))
    leggi_tutto.clear()
    return True


def scrivi_log(azione, foglio, dettaglio):
    try:
        worksheets()["Log"].append_row(_log_riga(azione, foglio, dettaglio),
                                       value_input_option="RAW",
                                       insert_data_option="INSERT_ROWS", table_range="A1")
    except Exception:
        pass


def avviso(msg):
    """Messaggio che resta visibile anche dopo il ricaricamento della pagina."""
    st.session_state["_avviso"] = msg


def impostazione(chiave, default=""):
    imp = df("Impostazioni")
    r = imp[imp["chiave"] == chiave]
    return r.iloc[0]["valore"] if len(r) else default


def salva_impostazione(chiave, valore):
    imp = df("Impostazioni")
    if (imp["chiave"] == chiave).any():
        aggiorna("Impostazioni", chiave, {"valore": valore}, chiave_col="chiave")
    else:
        aggiungi("Impostazioni", {"chiave": chiave, "valore": valore}, log=False)


def categorie(tipo):
    c = df("Categorie")
    lista = c[c["tipo"] == tipo]["nome"].tolist()
    return lista or CATEGORIE_DEFAULT.get(tipo, ["Altro"])


# ════════════════════════════════════════════════════════════════════
# PERMESSI
# ════════════════════════════════════════════════════════════════════
def ruolo_info():
    u = st.session_state.get("utente")
    if not u:
        return {"pagine": [], "modifica": False, "importi": False}
    r = df("Ruoli")
    r = r[r["ruolo"] == u["ruolo"]]
    if not len(r):
        return {"pagine": ["Panoramica"], "modifica": False, "importi": False}
    r = r.iloc[0]
    pagine = []
    for p in str(r["pagine"]).split(","):
        p = PAGINE_ALIAS.get(p.strip(), p.strip())
        if p in PAGINE and p not in pagine:
            pagine.append(p)
    return {"pagine": pagine,
            "modifica": str(r["puo_modificare"]).upper() == "SI",
            "importi": str(r["vede_importi"]).upper() == "SI"}


def puo_modificare():
    return ruolo_info()["modifica"]


def vede_importi():
    return ruolo_info()["importi"]


# ════════════════════════════════════════════════════════════════════
# GRAFICA – stile campo da calcio
# ════════════════════════════════════════════════════════════════════
def _chiaro(hex_col):
    """True se il colore è chiaro (per scegliere testo scuro o chiaro sopra)."""
    h = str(hex_col).lstrip("#")
    try:
        r, g, b = int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16)
    except (ValueError, IndexError):
        return False
    return (0.299 * r + 0.587 * g + 0.114 * b) > 160


# Palette "campo da calcio futuristico"
NEON, CIANO, ORO, ROSSO = "#3dffa0", "#2ee6ff", "#ffd166", "#ff5d73"
TESTO, TENUE = "#e9fff4", "rgba(233,255,244,.62)"

# Linee del campo al neon (sfondo fisso dietro a tutto)
CAMPO_SVG = """<div class="campo-neon"><svg viewBox="0 0 1440 900" preserveAspectRatio="none">
<defs><filter id="glw"><feGaussianBlur stdDeviation="3" result="b"/>
<feMerge><feMergeNode in="b"/><feMergeNode in="SourceGraphic"/></feMerge></filter></defs>
<g stroke="#3dffa0" stroke-opacity=".2" fill="none" stroke-width="2" filter="url(#glw)">
<rect x="40" y="40" width="1360" height="820" rx="6"/><line x1="720" y1="40" x2="720" y2="860"/>
<circle cx="720" cy="450" r="125"/><circle cx="720" cy="450" r="4" fill="#3dffa0"/>
<rect x="40" y="250" width="170" height="400"/><rect x="40" y="350" width="64" height="200"/>
<rect x="1230" y="250" width="170" height="400"/><rect x="1336" y="350" width="64" height="200"/>
<path d="M210 382 A 92 92 0 0 1 210 518"/><path d="M1230 382 A 92 92 0 0 0 1230 518"/>
</g></svg></div>"""


@st.cache_data(show_spinner=False)
def _stemma_file():
    """Se nel repository c'è stemma.png (o .jpg/.svg/.webp) lo usa al posto dello scudo."""
    import base64
    from pathlib import Path
    try:
        cartella = Path(__file__).resolve().parent
    except NameError:
        cartella = Path.cwd()
    tipi = {".png": "image/png", ".jpg": "image/jpeg", ".jpeg": "image/jpeg",
            ".svg": "image/svg+xml", ".webp": "image/webp"}
    for est, mime in tipi.items():
        f = cartella / f"stemma{est}"
        if f.exists():
            return f"data:{mime};base64,{base64.b64encode(f.read_bytes()).decode()}"
    return ""


def applica_stile(c1="#1f7a3a", c2="#ffffff"):
    st.markdown(f"""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Orbitron:wght@700;800;900&family=Rajdhani:wght@500;600;700&family=JetBrains+Mono:wght@500;700;800&display=swap');
    :root {{
        --c1:{c1}; --c2:{c2};
        --neon:{NEON}; --ciano:{CIANO}; --oro:{ORO}; --rosso:{ROSSO};
        --bg:#03110d; --vetro:rgba(8,36,27,.62); --vetro2:rgba(3,20,14,.78);
        --bordo:rgba(61,255,160,.2); --testo:{TESTO}; --tenue:{TENUE};
        --f-titolo:'Orbitron','Rajdhani',system-ui,sans-serif;
        --f-testo:'Rajdhani','Inter',system-ui,sans-serif;
        --f-num:'JetBrains Mono','DejaVu Sans Mono',ui-monospace,monospace;
    }}

    /* ── SFONDO: notte allo stadio + strisce del prato ── */
    .stApp {{
        background:
          radial-gradient(900px 520px at 12% -8%, rgba(46,230,255,.16), transparent 60%),
          radial-gradient(900px 520px at 92% -8%, rgba(61,255,160,.15), transparent 60%),
          repeating-linear-gradient(90deg, rgba(4,26,16,.9) 0 110px, rgba(3,17,13,0) 110px 220px),
          linear-gradient(180deg, #041a12 0%, #020c08 100%) !important;
        background-attachment: fixed !important;
    }}
    .campo-neon {{ position:fixed; inset:0; pointer-events:none; z-index:0; opacity:.6; }}
    .campo-neon svg {{ width:100%; height:100%; }}
    [data-testid="stHeader"] {{ background:transparent !important; }}
    .block-container {{ padding-top:2rem; max-width:1320px; position:relative; z-index:1; }}

    /* ── TESTI ── */
    .stApp, .stApp p, .stApp span, .stApp li, .stApp label, .stApp h1, .stApp h2, .stApp h3,
    .stApp h4, .stApp h5, [data-testid="stWidgetLabel"] *, [data-testid="stMarkdownContainer"] *,
    [data-testid="stCaptionContainer"], [data-testid="stSidebar"] * {{ color:var(--testo) !important; }}
    .stApp {{ font-family:var(--f-testo); font-size:16.5px; }}
    [data-testid="stCaptionContainer"], [data-testid="stCaptionContainer"] * {{ color:var(--tenue) !important; }}
    [data-testid="stWidgetLabel"] p {{ font-size:12px !important; letter-spacing:1.6px; text-transform:uppercase;
        font-weight:700 !important; color:var(--tenue) !important; }}
    .stApp h3 {{ font-family:var(--f-testo); font-weight:700; letter-spacing:2.5px; text-transform:uppercase;
        font-size:17px !important; }}
    .stApp h4 {{ font-family:var(--f-testo); font-weight:700; letter-spacing:2.2px; text-transform:uppercase;
        font-size:15px !important; margin-top:18px; display:flex; align-items:center; gap:10px; }}
    .stApp h4:before {{ content:""; width:9px; height:9px; border-radius:2px; background:var(--neon);
        box-shadow:0 0 10px var(--neon); flex:0 0 9px; }}
    .stApp h5 {{ font-family:var(--f-testo); letter-spacing:1.5px; text-transform:uppercase; font-size:14px !important; }}

    /* ── BARRA LATERALE ── */
    [data-testid="stSidebar"] > div:first-child {{
        background:linear-gradient(180deg, rgba(3,20,14,.96), rgba(3,20,14,.86)) !important;
        border-right:1px solid var(--bordo); backdrop-filter:blur(10px);
    }}
    .stemma {{ display:flex; align-items:center; gap:14px; margin:2px 0 18px; }}
    .stemma .scudo {{ width:58px; height:66px; flex:0 0 58px; position:relative;
        filter:drop-shadow(0 0 10px rgba(61,255,160,.55)); }}
    .stemma .scudo .b, .stemma .scudo .s {{ position:absolute;
        clip-path:polygon(50% 100%, 0 74%, 0 0, 100% 0, 100% 74%); }}
    .stemma .scudo .b {{ inset:0; background:var(--neon); }}
    .stemma .scudo .s {{ inset:2px; background:linear-gradient(90deg, var(--c1) 50%, var(--c2) 50%); }}
    .stemma .scudo span {{ position:absolute; inset:0; display:flex; align-items:center; justify-content:center;
        padding-bottom:10px; font-family:var(--f-titolo); font-weight:900; font-size:14px; letter-spacing:.5px;
        color:#03110d !important; text-shadow:0 0 6px #fff; }}
    .stemma img {{ width:62px; height:62px; object-fit:contain; flex:0 0 62px;
        filter:drop-shadow(0 0 10px rgba(61,255,160,.5)); }}
    .stemma .nome {{ font-family:var(--f-testo); font-weight:700; font-size:19px; line-height:1.05; }}
    .stemma .citta {{ font-size:11px; letter-spacing:2.6px; text-transform:uppercase; color:var(--tenue) !important;
        margin-top:4px; }}
    [data-testid="stSidebar"] [data-baseweb="select"] > div {{
        background:rgba(61,255,160,.07) !important; border:1px solid var(--bordo) !important;
    }}
    [data-testid="stSidebar"] [data-baseweb="select"] * {{
        color:var(--neon) !important; -webkit-text-fill-color:var(--neon) !important; fill:var(--neon) !important;
        font-family:var(--f-num) !important; font-weight:700 !important;
    }}
    /* menu a pulsanti illuminati */
    [data-testid="stSidebar"] [role="radiogroup"] {{ gap:4px; }}
    [data-testid="stSidebar"] [role="radiogroup"] label {{
        width:100%; padding:10px 14px; border-radius:12px; margin:0; border:1px solid transparent;
        transition:all .15s;
    }}
    [data-testid="stSidebar"] [role="radiogroup"] label > div:first-child {{ display:none; }}
    [data-testid="stSidebar"] [role="radiogroup"] label p {{ font-size:16px !important; font-weight:600 !important;
        color:var(--tenue) !important; }}
    [data-testid="stSidebar"] [role="radiogroup"] label:hover {{ background:rgba(61,255,160,.07);
        border-color:var(--bordo); }}
    [data-testid="stSidebar"] [role="radiogroup"] label:has(input:checked) {{
        background:linear-gradient(90deg, var(--neon), #9dffcf); box-shadow:0 0 22px rgba(61,255,160,.45);
    }}
    [data-testid="stSidebar"] [role="radiogroup"] label:has(input:checked) p {{
        color:#03110d !important; font-weight:800 !important;
    }}
    [data-testid="stSidebar"] hr {{ border-color:var(--bordo) !important; }}

    /* ── INTESTAZIONE PAGINA ── */
    .banner {{
        position:relative; overflow:hidden; border-radius:22px; padding:24px 30px 22px; margin-bottom:22px;
        background:linear-gradient(115deg, color-mix(in srgb, var(--c1) 60%, transparent), rgba(4,26,18,.7) 62%);
        border:1px solid rgba(61,255,160,.38);
        box-shadow:0 0 0 1px rgba(46,230,255,.08) inset, 0 20px 50px rgba(0,0,0,.35);
    }}
    .banner:before {{ content:""; position:absolute; right:78px; top:0; bottom:0;
        border-left:2px solid rgba(61,255,160,.25); pointer-events:none; }}
    .banner:after {{ content:""; position:absolute; right:-56px; top:-56px; width:250px; height:250px;
        border-radius:50%; border:2px solid rgba(61,255,160,.35); box-shadow:0 0 30px rgba(61,255,160,.22);
        pointer-events:none; }}
    .banner .k {{ font-size:12px; letter-spacing:4px; text-transform:uppercase; font-weight:700;
        color:var(--neon) !important; }}
    .banner .t {{ font-family:var(--f-titolo); font-size:34px; font-weight:900; letter-spacing:1.5px;
        text-transform:uppercase; margin-top:6px; line-height:1.1; position:relative; z-index:1; }}
    .banner .s {{ color:var(--tenue) !important; font-size:16px; margin-top:4px; position:relative; z-index:1; }}
    .banner .dx {{ position:absolute; right:112px; top:26px; text-align:right; z-index:1; }}
    .banner .chip {{ display:inline-block; padding:7px 16px; border-radius:30px; font-weight:800; font-size:13px;
        letter-spacing:1.2px; background:var(--oro); box-shadow:0 0 20px rgba(255,209,102,.5);
        font-family:var(--f-num); }}
    .banner .chip, .banner .chip * {{ color:#03110d !important; }}
    .banner .live {{ font-size:12px; color:var(--tenue) !important; letter-spacing:1px; margin-top:10px; }}
    .banner .live i {{ display:inline-block; width:7px; height:7px; border-radius:50%; background:var(--neon);
        box-shadow:0 0 8px var(--neon); margin-right:6px; }}

    /* ── CARD NUMERI ── */
    .kpi {{
        position:relative; overflow:hidden; border-radius:18px; padding:15px 18px 14px; height:100%;
        min-height:104px; background:var(--vetro); border:1px solid var(--bordo); backdrop-filter:blur(8px);
        transition:transform .15s, box-shadow .15s;
    }}
    .kpi:hover {{ transform:translateY(-2px); box-shadow:0 10px 28px rgba(0,0,0,.35); }}
    .kpi:before {{ content:""; position:absolute; left:0; right:0; top:0; height:3px;
        background:linear-gradient(90deg, var(--neon), transparent); }}
    .kpi .l {{ font-size:11.5px; letter-spacing:2.2px; text-transform:uppercase; font-weight:700;
        color:var(--tenue) !important; }}
    .kpi .v {{ font-family:var(--f-num); font-size:25px; font-weight:800; margin-top:8px; letter-spacing:-.5px;
        white-space:nowrap; }}
    .kpi .n {{ font-size:13px; color:var(--tenue) !important; margin-top:3px; }}
    .kpi.pos .v {{ color:var(--neon) !important; text-shadow:0 0 12px rgba(61,255,160,.5); }}
    .kpi.neg:before {{ background:linear-gradient(90deg, var(--rosso), transparent); }}
    .kpi.neg .v {{ color:var(--rosso) !important; text-shadow:0 0 12px rgba(255,93,115,.45); }}
    .kpi.gold:before {{ background:linear-gradient(90deg, var(--oro), transparent); }}
    .kpi.gold .v {{ color:var(--oro) !important; text-shadow:0 0 12px rgba(255,209,102,.45); }}
    .kpi.ciano:before {{ background:linear-gradient(90deg, var(--ciano), transparent); }}
    .kpi.ciano .v {{ color:var(--ciano) !important; text-shadow:0 0 12px rgba(46,230,255,.45); }}
    /* tabellone tipo risultato partita */
    .kpi.score {{ display:flex; align-items:center; justify-content:space-between; gap:10px; }}
    .kpi.score .t {{ text-align:center; }}
    .kpi.score .v {{ font-size:23px; }}
    .kpi.score .vs {{ font-family:var(--f-titolo); font-weight:900; font-size:12px; letter-spacing:3px;
        color:var(--tenue) !important; border:1px solid var(--bordo); border-radius:8px; padding:6px 9px; }}
    .pill {{ font-family:var(--f-num); font-size:12px; font-weight:800; padding:4px 9px; border-radius:8px; }}
    .pill, .pill * {{ color:#03110d !important; }}
    .scad {{ display:flex; align-items:center; gap:12px; padding:10px 14px; border-radius:14px; margin-bottom:8px;
        background:var(--vetro2); border:1px solid var(--bordo); }}

    /* ── SCHEDE ── */
    .stTabs [data-baseweb="tab-list"] {{ gap:8px; border-bottom:none !important; flex-wrap:wrap; }}
    .stTabs [data-baseweb="tab"] {{
        background:var(--vetro2) !important; border:1px solid var(--bordo) !important; border-radius:12px !important;
        padding:8px 16px !important; height:auto !important;
    }}
    .stTabs [data-baseweb="tab"] p {{ font-weight:700 !important; font-size:15px !important; color:var(--tenue) !important; }}
    .stTabs [aria-selected="true"] {{ background:var(--neon) !important; border-color:var(--neon) !important;
        box-shadow:0 0 18px rgba(61,255,160,.45); }}
    .stTabs [aria-selected="true"] p {{ color:#03110d !important; font-weight:800 !important; }}
    .stTabs [data-baseweb="tab-highlight"], .stTabs [data-baseweb="tab-border"] {{ display:none !important; }}

    /* ── RIQUADRI, MODULI, TABELLE ── */
    [data-testid="stForm"] {{ background:var(--vetro) !important; border:1px solid var(--bordo) !important;
        border-radius:18px !important; backdrop-filter:blur(8px); }}
    [data-testid="stExpander"] details {{ background:var(--vetro2) !important; border:1px solid var(--bordo) !important;
        border-radius:14px !important; }}
    [data-testid="stExpander"] summary p {{ font-weight:700 !important; }}
    [data-testid="stAlert"] {{ background:rgba(46,230,255,.08) !important; border:1px solid rgba(46,230,255,.3) !important;
        border-radius:14px !important; }}
    [data-testid="stDataFrame"] {{ border:1px solid var(--bordo); border-radius:14px; overflow:hidden;
        box-shadow:0 0 0 1px rgba(46,230,255,.05) inset; }}

    /* ── CAMPI DA COMPILARE: chiari e leggibili, bordo neon quando li usi ── */
    .stApp input, .stApp textarea,
    [data-baseweb="input"], [data-baseweb="base-input"], [data-baseweb="textarea"],
    .stApp [data-baseweb="select"] > div, [data-testid="stNumberInput"] button {{
        background:#e9fff4 !important; color:#03110d !important; -webkit-text-fill-color:#03110d !important;
        caret-color:#03110d !important; border-color:transparent !important; border-radius:10px !important;
        font-family:var(--f-testo) !important; font-weight:600 !important; font-size:16px !important;
    }}
    [data-baseweb="input"]:focus-within, [data-baseweb="textarea"]:focus-within,
    .stApp [data-baseweb="select"] > div:focus-within {{
        border:2px solid var(--neon) !important; box-shadow:0 0 0 4px rgba(61,255,160,.18) !important;
    }}
    .stApp [data-baseweb="select"] *, .stApp [data-testid="stNumberInput"] button *,
    .stApp [data-baseweb="input"] *, .stApp [data-baseweb="base-input"] *, .stApp [data-baseweb="tag"] * {{
        color:#03110d !important; -webkit-text-fill-color:#03110d !important; fill:#03110d !important;
    }}
    .stApp [data-baseweb="tag"] {{ background:#b9ffd9 !important; }}
    .stApp input::placeholder, .stApp textarea::placeholder {{
        color:#6b8a7c !important; -webkit-text-fill-color:#6b8a7c !important;
    }}
    [data-baseweb="popover"], [data-baseweb="menu"], [role="listbox"], [data-baseweb="calendar"] {{
        background:#e9fff4 !important;
    }}
    [data-baseweb="popover"] *, [role="option"], [role="option"] *,
    .stApp [data-baseweb="popover"] *, .stApp [role="option"] * {{
        color:#03110d !important; -webkit-text-fill-color:#03110d !important;
    }}
    [role="option"]:hover, [aria-selected="true"][role="option"] {{ background:#b9ffd9 !important; }}
    /* scelte a pallino (radio) dentro le pagine: diventano "pillole" */
    [data-testid="stMain"] [role="radiogroup"], .main [role="radiogroup"] {{ gap:8px; flex-wrap:wrap; }}
    [data-testid="stMain"] [role="radiogroup"] label, .main [role="radiogroup"] label {{ border:1px solid var(--bordo); border-radius:30px; padding:6px 14px 6px 10px;
        background:var(--vetro2); margin:0 !important; }}
    [data-testid="stMain"] [role="radiogroup"] label:has(input:checked),
    .main [role="radiogroup"] label:has(input:checked) {{ border-color:var(--ciano);
        box-shadow:0 0 14px rgba(46,230,255,.35) inset; }}
    [data-testid="stMain"] [role="radiogroup"] label:has(input:checked) p,
    .main [role="radiogroup"] label:has(input:checked) p {{ color:var(--ciano) !important; font-weight:700 !important; }}

    /* ── PULSANTI ── */
    .stButton button, [data-testid="stFormSubmitButton"] button, [data-testid="stDownloadButton"] button {{
        background:linear-gradient(90deg, var(--neon), #9dffcf) !important; border:none !important;
        border-radius:12px !important; box-shadow:0 0 18px rgba(61,255,160,.35); transition:filter .15s, box-shadow .15s;
    }}
    .stButton button:hover, [data-testid="stFormSubmitButton"] button:hover,
    [data-testid="stDownloadButton"] button:hover {{ filter:brightness(1.08); box-shadow:0 0 28px rgba(61,255,160,.55); }}
    .stApp .stButton button *, .stApp [data-testid="stFormSubmitButton"] button *,
    .stApp [data-testid="stDownloadButton"] button * {{ color:#03110d !important; font-weight:800 !important; }}
    [data-testid="stSidebar"] .stButton button {{ background:var(--vetro2) !important; border:1px solid var(--bordo) !important;
        box-shadow:none; }}
    [data-testid="stSidebar"] .stButton button * {{ color:var(--testo) !important; }}

    /* ── TELEFONO ── */
    @media (max-width: 760px) {{
        .banner {{ padding:18px; }} .banner .t {{ font-size:22px; }}
        .banner .dx {{ position:static; text-align:left; margin-top:12px; }}
        .banner:before, .banner:after {{ display:none; }}
        .kpi .v {{ font-size:20px; }} .campo-neon {{ opacity:.35; }}
    }}
    </style>
    {CAMPO_SVG}
    """, unsafe_allow_html=True)


def stemma():
    nome = impostazione("nome_squadra", "La mia squadra")
    parole = [p for p in nome.replace(".", " ").split() if p[:1].isalpha()]
    sigla = "".join(p[0] for p in parole[:3]).upper() or "⚽"
    citta = escape(impostazione("citta", ""))
    img = _stemma_file()
    logo = (f'<img src="{img}" alt="stemma">' if img else
            f'<div class="scudo"><div class="b"></div><div class="s"></div><span>{escape(sigla)}</span></div>')
    st.markdown(f"""<div class="stemma">{logo}
        <div><div class="nome">{escape(nome)}</div><div class="citta">{citta}</div></div></div>""",
                unsafe_allow_html=True)


SOPRATITOLI = {"Panoramica": "Matchday finanziario", "Rosa & Staff": "La squadra",
               "Cassa & Pagamenti": "Registro unico", "Fornitori & Fatture": "Da pagare",
               "Sponsor": "Partner", "Scadenze": "Calendario", "Impostazioni": "Configurazione"}


def banner(titolo, sottotitolo=""):
    nome = escape(impostazione("nome_squadra", "La mia squadra"))
    stg = escape(st.session_state.get("stagione", ""))
    ora = datetime.now(TZ).strftime("%H:%M")
    st.markdown(f"""<div class="banner">
        <div class="k">● {SOPRATITOLI.get(titolo, nome)}</div>
        <div class="t">{escape(titolo)}</div>
        <div class="s">{nome}{' · ' + escape(sottotitolo) if sottotitolo else ''}</div>
        <div class="dx"><span class="chip">STAGIONE {stg}</span>
        <div class="live"><i></i>aggiornato alle {ora}</div></div></div>""", unsafe_allow_html=True)


def kpi(col, label, valore, nota="", tono=""):
    col.markdown(f"""<div class="kpi {tono}"><div class="l">{label}</div>
        <div class="v">{valore}</div><div class="n">{nota}</div></div>""",
                 unsafe_allow_html=True)


def kpi_score(col, entrate, uscite):
    """Tabellone 'Entrate VS Uscite' come il risultato di una partita."""
    num = lambda x: euro(x).replace("€ ", "").rsplit(",", 1)[0]
    col.markdown(f"""<div class="kpi score">
        <div class="t"><div class="l">Entrate</div><div class="v" style="color:{NEON};
            text-shadow:0 0 12px rgba(61,255,160,.5)">{num(entrate)}</div></div>
        <div class="vs">VS</div>
        <div class="t"><div class="l">Uscite</div><div class="v" style="color:{ROSSO};
            text-shadow:0 0 12px rgba(255,93,115,.45)">{num(uscite)}</div></div></div>""",
                 unsafe_allow_html=True)


def grafico_tema(ch):
    return (ch.configure(background="transparent", font="Rajdhani")
            .configure_axis(labelColor=TENUE, titleColor=TENUE, gridColor="rgba(233,255,244,.07)",
                            domainColor="rgba(233,255,244,.18)", tickColor="rgba(233,255,244,.18)",
                            labelFontSize=12, titleFontSize=12)
            .configure_legend(labelColor=TESTO, titleColor=TESTO, labelFontSize=13, symbolType="square")
            .configure_view(strokeWidth=0))


# ════════════════════════════════════════════════════════════════════
# DATI CALCOLATI
# ════════════════════════════════════════════════════════════════════
def stagione_corrente_df(nome):
    d = df(nome)
    return d[d["stagione"] == st.session_state["stagione"]].copy()


def mesi_stagione():
    s = df("Stagioni")
    r = s[s["stagione"] == st.session_state["stagione"]]
    anno = int(st.session_state["stagione"][:4])
    inizio = parse_date(r.iloc[0]["data_inizio"]) if len(r) else None
    fine = parse_date(r.iloc[0]["data_fine"]) if len(r) else None
    inizio = inizio or date(anno, 7, 1)
    fine = fine or date(anno + 1, 6, 30)
    out, d = [], date(inizio.year, inizio.month, 1)
    while d <= fine:
        out.append(f"{MESI[d.month - 1]} {d.year}")
        d = date(d.year + (d.month == 12), d.month % 12 + 1, 1)
    return out


def nome_completo(r):
    return f"{r['cognome']} {r['nome']}".strip()


def movimenti_tesserati():
    """Uscite collegate a un tesserato (registro unico Movimenti)."""
    m = stagione_corrente_df("Movimenti")
    return m[(m["tipo"] == "Uscita") & (m["tesserato_id"] != "")].copy()


def pagamenti_diretti():
    """Soldi dati direttamente al giocatore/staff (bonifico, contanti…)."""
    m = movimenti_tesserati()
    return m[m["destinatario"] == "Giocatore"].copy()


def spese_per_tesserati():
    """Pagamenti a strutture/terzi per conto di un tesserato (alloggio, visite, terapie…)."""
    m = movimenti_tesserati()
    return m[m["destinatario"] != "Giocatore"].copy()


def tesserati_con_saldi():
    t = stagione_corrente_df("Tesserati")
    if t.empty:
        return t.assign(extra=0.0, totale=0.0, pagato_diretto=0.0, pagato_terzi=0.0, pagato=0.0,
                        residuo=0.0, spese_extra=0.0, costo_totale=0.0, nominativo="")
    v = stagione_corrente_df("Voci_Contratto")
    sp = spese_per_tesserati()
    extra = v.groupby("tesserato_id")["importo"].sum()
    pag = pagamenti_diretti().groupby("tesserato_id")["importo"].sum()
    terzi = sp[sp["nel_compenso"] == "SI"].groupby("tesserato_id")["importo"].sum()
    fuori = sp[sp["nel_compenso"] != "SI"].groupby("tesserato_id")["importo"].sum()
    t["extra"] = t["id"].map(extra).fillna(0.0)
    t["totale"] = t["compenso_stagione"] + t["extra"]
    t["pagato_diretto"] = t["id"].map(pag).fillna(0.0)
    t["pagato_terzi"] = t["id"].map(terzi).fillna(0.0)          # scala dal compenso
    t["pagato"] = t["pagato_diretto"] + t["pagato_terzi"]
    t["residuo"] = t["totale"] - t["pagato"]
    t["spese_extra"] = t["id"].map(fuori).fillna(0.0)           # a carico società, fuori compenso
    t["costo_totale"] = t["pagato"] + t["spese_extra"]
    t["nominativo"] = t.apply(nome_completo, axis=1)
    return t


def ordina_reparti(t):
    ordine = {r: i for i, r in enumerate(categorie("REPARTO"))}
    t["_o"] = t["reparto"].map(ordine).fillna(99)
    return t.sort_values(["_o", "cognome", "nome"]).drop(columns="_o")


def mappa_tesserati():
    t = stagione_corrente_df("Tesserati")
    return {r["id"]: nome_completo(r) for _, r in t.iterrows()}


def registra_movimento(tipo, data_mov, categoria, descrizione, importo, metodo,
                       rif_tipo="", rif_id="", note="", tesserato_id="", nel_compenso="",
                       beneficiario="", mese_rif="", destinatario="", voce=""):
    """UNICO punto in cui si registrano soldi in entrata/uscita."""
    if tesserato_id and not destinatario:
        destinatario = "Struttura"
    if destinatario == "Giocatore":
        nel_compenso = "SI"
    mid = nuovo_id("M")
    aggiungi("Movimenti", {
        "id": mid, "data": data_mov.isoformat(), "stagione": st.session_state["stagione"],
        "tipo": tipo, "categoria": categoria, "descrizione": descrizione,
        "importo": float(importo), "metodo": metodo, "rif_tipo": rif_tipo, "rif_id": rif_id,
        "utente": st.session_state["utente"]["username"], "note": note,
        "tesserato_id": tesserato_id, "nel_compenso": nel_compenso if tesserato_id else "",
        "beneficiario": beneficiario, "mese_rif": mese_rif,
        "destinatario": destinatario if tesserato_id else "", "voce": voce})
    return mid


def migra_dati():
    """Una sola volta: porta i vecchi Pagamenti nel registro unico Movimenti."""
    if impostazione("versione_dati") == "2":
        return
    pag = df("Pagamenti")
    mov = df("Movimenti")
    ids_mov = set(mov["id"])
    nuovi = []
    with st.spinner("Aggiorno i dati al nuovo registro unico…"):
        mod_mov, mod_pag = {}, {}
        for _, p in pag[pag["migrato"] != "SI"].iterrows():
            campi = {"tesserato_id": p["tesserato_id"], "destinatario": "Giocatore", "nel_compenso": "SI",
                     "voce": p["voce"], "mese_rif": p["mese_rif"], "rif_tipo": "", "rif_id": ""}
            if p["mov_id"] and p["mov_id"] in ids_mov:
                mod_mov[p["mov_id"]] = campi
            else:
                nuovi.append({"id": nuovo_id("M"), "data": p["data"], "stagione": p["stagione"],
                              "tipo": "Uscita", "categoria": "Compensi giocatori",
                              "descrizione": f"{p['voce']} {p['mese_rif']}", "importo": p["importo"],
                              "metodo": p["metodo"], "utente": p["utente"], "note": p["note"], **campi})
            mod_pag[p["id"]] = {"migrato": "SI"}
        # vecchie spese a strutture: tesserato presente ma destinatario vuoto
        for _, m in mov.iterrows():
            if m["tesserato_id"] and not m["destinatario"] and m["id"] not in mod_mov:
                mod_mov[m["id"]] = {"destinatario": "Struttura"}
        aggiorna_molti("Movimenti", mod_mov)
        aggiungi_molti("Movimenti", nuovi)
        aggiorna_molti("Pagamenti", mod_pag)
        salva_impostazione("versione_dati", "2")


SCELTE_COMPENSO = ["NO – spesa extra a carico della società",
                   "SÌ – scala dal compenso del giocatore"]


def scelta_compenso(valore_radio):
    return "SI" if str(valore_radio).startswith("SÌ") else "NO"


def blocca_se_sola_lettura():
    if not puo_modificare():
        st.info("🔒 Il tuo ruolo è in sola lettura per questa sezione.")
        return True
    return False


# ════════════════════════════════════════════════════════════════════
# PRIMO ACCESSO
# ════════════════════════════════════════════════════════════════════
def pagina_setup():
    applica_stile()
    st.markdown("""<div class="banner"><div class="t">⚽ Benvenuto</div>
        <div class="s">Primo accesso: crea la tua squadra e l'utente amministratore</div>
        </div>""", unsafe_allow_html=True)
    anno = oggi().year if oggi().month >= 7 else oggi().year - 1
    with st.form("setup"):
        c1, c2 = st.columns(2)
        nome = c1.text_input("Nome squadra *")
        citta = c2.text_input("Città")
        c3, c4, c5 = st.columns(3)
        col1 = c3.color_picker("Colore sociale 1", "#1f7a3a")
        col2 = c4.color_picker("Colore sociale 2", "#ffffff")
        anni = list(range(anno - 2, anno + 3))
        a = c5.selectbox("Stagione iniziale", anni, index=2,
                         format_func=stagione_da_anno)
        st.markdown("**Amministratore** (ruolo *Presidente*, accesso completo)")
        c6, c7, c8 = st.columns(3)
        nome_admin = c6.text_input("Nome e cognome *")
        user = c7.text_input("Username *")
        pw = c8.text_input("Password * (min. 6 caratteri)", type="password")
        ok = st.form_submit_button("Crea squadra", type="primary", width="stretch")
    if ok:
        if not nome.strip() or not user.strip() or len(pw) < 6 or not nome_admin.strip():
            st.error("Compila tutti i campi obbligatori (password di almeno 6 caratteri).")
            return
        with st.spinner("Creo la squadra…"):
            stg = stagione_da_anno(a)
            if not (df("Stagioni")["stagione"] == stg).any():
                aggiungi_molti("Stagioni", [{"stagione": stg, "data_inizio": f"{a}-07-01",
                                             "data_fine": f"{a + 1}-06-30", "attiva": "SI"}])
            if not (df("Ruoli")["ruolo"] == "Presidente").any():
                aggiungi_molti("Ruoli", [{"ruolo": "Presidente", "pagine": ",".join(PAGINE),
                                          "puo_modificare": "SI", "vede_importi": "SI"}])
            if df("Categorie").empty:
                aggiungi_molti("Categorie", [{"tipo": t, "nome": n}
                                             for t, lista in CATEGORIE_DEFAULT.items() for n in lista])
            aggiungi_molti("Utenti", [{"username": user.strip().lower(), "password_hash": hash_pw(pw),
                                       "nome": nome_admin.strip(), "ruolo": "Presidente",
                                       "attivo": "SI"}])
            gia = set(df("Impostazioni")["chiave"])
            imp = {"nome_squadra": nome.strip(), "citta": citta.strip(), "colore_1": col1,
                   "colore_2": col2, "ruolo_admin": "Presidente", "versione_dati": "2",
                   "setup_completato": "SI"}
            for k in [k for k in imp if k in gia]:
                aggiorna("Impostazioni", k, {"valore": imp.pop(k)}, chiave_col="chiave")
            aggiungi_molti("Impostazioni", [{"chiave": k, "valore": v} for k, v in imp.items()])
            scrivi_log("Creata squadra", "Impostazioni", nome.strip())
        avviso("✅ Squadra creata! Ora accedi con il tuo utente.")
        st.rerun()


def pagina_login():
    applica_stile(impostazione("colore_1", "#1f7a3a"), impostazione("colore_2", "#ffffff"))
    nome = escape(impostazione("nome_squadra", "La mia squadra"))
    _, c, _ = st.columns([1, 1.3, 1])
    with c:
        msg = st.session_state.pop("_avviso", None)
        if msg:
            st.success(msg)
        st.markdown(f"""<div class="banner" style="text-align:center"><div class="t">⚽ {nome}</div>
            <div class="s">Area riservata – gestione società</div></div>""",
                    unsafe_allow_html=True)
        with st.form("login"):
            u = st.text_input("Username")
            p = st.text_input("Password", type="password")
            ok = st.form_submit_button("Entra", type="primary", width="stretch")
        if ok:
            ut = df("Utenti")
            r = ut[(ut["username"] == u.strip().lower()) & (ut["attivo"].str.upper() == "SI")]
            if len(r) and check_pw(p, r.iloc[0]["password_hash"]):
                st.session_state["utente"] = r.iloc[0].to_dict()
                scrivi_log("Accesso", "-", u)
                st.rerun()
            else:
                st.error("Credenziali non valide.")


# ════════════════════════════════════════════════════════════════════
# PANORAMICA
# ════════════════════════════════════════════════════════════════════
def pagina_panoramica():
    banner("Panoramica", "come va la stagione")
    t = tesserati_con_saldi()
    sc = stagione_corrente_df("Scadenze")
    giocatori = (t["tipo"] == "Giocatore").sum() if len(t) else 0

    if not vede_importi():
        c = st.columns(3)
        kpi(c[0], "Giocatori in rosa", giocatori)
        kpi(c[1], "Staff", len(t) - giocatori)
        kpi(c[2], "Scadenze aperte", (sc["stato"] != "Fatto").sum() if len(sc) else 0)
        sezione_scadenze_prossime(sc)
        return

    m = stagione_corrente_df("Movimenti")
    entrate = m[m["tipo"] == "Entrata"]["importo"].sum()
    uscite = m[m["tipo"] == "Uscita"]["importo"].sum()
    saldo = entrate - uscite
    contanti = (m[(m["metodo"] == "Contanti") & (m["tipo"] == "Entrata")]["importo"].sum()
                - m[(m["metodo"] == "Contanti") & (m["tipo"] == "Uscita")]["importo"].sum())
    residuo_t = t["residuo"].clip(lower=0).sum() if len(t) else 0
    f = stagione_corrente_df("Fatture")
    fatt_aperte = (f["importo"] - f["pagato"]).clip(lower=0).sum() if len(f) else 0
    sp = stagione_corrente_df("Sponsor")
    inc_sp = m[m["rif_tipo"] == "Sponsor"].groupby("rif_id")["importo"].sum()
    sp_da_incassare = ((sp["importo_concordato"] - sp["id"].map(inc_sp).fillna(0.0)).clip(lower=0).sum()
                       if len(sp) else 0)
    previsione = saldo - residuo_t - fatt_aperte + sp_da_incassare

    extra_soc = t["spese_extra"].sum() if len(t) else 0
    c = st.columns([1.25, 1, 1, 1])
    kpi_score(c[0], entrate, uscite)
    kpi(c[1], "Saldo cassa", euro(saldo), f"di cui contanti {euro(contanti)}",
        "pos" if saldo >= 0 else "neg")
    kpi(c[2], "Da pagare a tesserati", euro(residuo_t), f"{len(t)} tesserati · residuo contratti", "neg")
    kpi(c[3], "Previsione fine stagione", euro(previsione), "saldo − da pagare + sponsor", "gold")
    st.write("")
    c = st.columns(4)
    kpi(c[0], "Sponsor da incassare", euro(sp_da_incassare), f"{len(sp)} sponsor", "ciano")
    kpi(c[1], "Fatture da pagare", euro(fatt_aperte),
        f"{(f['stato'] != 'Pagata').sum() if len(f) else 0} aperte", "neg")
    kpi(c[2], "Spese extra società", euro(extra_soc), "strutture fuori compenso", "gold")
    kpi(c[3], "Rosa", f"{giocatori} {'giocatore' if giocatori == 1 else 'giocatori'}",
        f"{len(t) - giocatori} staff")

    st.write("")
    g1, g2 = st.columns([1.3, 1])
    with g1:
        st.markdown("#### Entrate e uscite per mese")
        if len(m):
            m["d"] = m["data"].apply(parse_date)
            m = m.dropna(subset=["d"])
            m["ord"] = m["d"].apply(lambda d: f"{d.year}-{d.month:02d}")
            m["mese"] = m["d"].apply(lambda d: f"{MESI[d.month - 1][:3].upper()} {str(d.year)[2:]}")
            agg = m.groupby(["ord", "mese", "tipo"])["importo"].sum().reset_index().sort_values("ord")
            ch = alt.Chart(agg).mark_bar(cornerRadiusTopLeft=6, cornerRadiusTopRight=6, size=18).encode(
                x=alt.X("mese:N", title=None, sort=agg["mese"].unique().tolist(), axis=alt.Axis(labelAngle=0)),
                xOffset="tipo:N",
                y=alt.Y("importo:Q", title=None),
                color=alt.Color("tipo:N", scale=alt.Scale(domain=["Entrata", "Uscita"], range=[NEON, ROSSO]),
                                legend=alt.Legend(orient="top", title=None)),
                tooltip=["mese", "tipo", alt.Tooltip("importo:Q", format=",.2f")])
            st.altair_chart(grafico_tema(ch.properties(height=300)), width="stretch")
        else:
            st.caption("Nessun movimento ancora registrato.")
    with g2:
        st.markdown("#### Dove vanno i soldi")
        u = m[m["tipo"] == "Uscita"] if len(m) else m
        if len(u):
            agg = u.groupby("categoria")["importo"].sum().sort_values(ascending=False)
            agg = agg.head(8).reset_index()
            ch = alt.Chart(agg).mark_bar(color=ORO, cornerRadiusEnd=6, size=14).encode(
                y=alt.Y("categoria:N", sort="-x", title=None),
                x=alt.X("importo:Q", title="€"),
                tooltip=["categoria", alt.Tooltip("importo:Q", format=",.2f")])
            st.altair_chart(grafico_tema(ch.properties(height=300)), width="stretch")
        else:
            st.caption("Nessuna uscita ancora registrata.")

    sezione_scadenze_prossime(sc)


def sezione_scadenze_prossime(sc):
    st.markdown("#### Scadenze nei prossimi 30 giorni")
    if sc.empty:
        st.caption("Nessuna scadenza inserita.")
        return
    sc = sc[sc["stato"] != "Fatto"].copy()
    sc["d"] = sc["data_scadenza"].apply(parse_date)
    sc = sc.dropna(subset=["d"])
    sc = sc[sc["d"] <= oggi() + timedelta(days=30)].sort_values("d")
    if sc.empty:
        st.caption("Niente in scadenza. 👍")
        return
    mt = mappa_tesserati()
    for _, r in sc.iterrows():
        scaduta = r["d"] < oggi()
        vicina = r["d"] <= oggi() + timedelta(days=7)
        col = ROSSO if scaduta else (ORO if vicina else NEON)
        chi = f" · {escape(mt.get(r['tesserato_id'], ''))}" if r["tesserato_id"] else ""
        st.markdown(f"<div class='scad'><span class='pill' style='background:{col}'>"
                    f"{'SCADUTA' if scaduta else r['d'].strftime('%d/%m')}</span>"
                    f"<b>{escape(r['tipo'])}</b> {escape(r['descrizione'])}{chi}</div>", unsafe_allow_html=True)


# ════════════════════════════════════════════════════════════════════
# ROSA & STAFF
# ════════════════════════════════════════════════════════════════════
def form_tesserato(key, r=None, mostra_importi=True):
    r = r or {}
    tipi, reparti = categorie("TESSERATO"), categorie("REPARTO")
    with st.form(key, clear_on_submit=r == {}):
        c = st.columns(3)
        cognome = c[0].text_input("Cognome *", r.get("cognome", ""))
        nome = c[1].text_input("Nome", r.get("nome", ""))
        tipo = c[2].selectbox("Tipo", tipi, index=tipi.index(r["tipo"]) if r.get("tipo") in tipi else 0)
        c = st.columns(3)
        reparto = c[0].selectbox("Reparto / Area", reparti,
                                 index=reparti.index(r["reparto"]) if r.get("reparto") in reparti else 0)
        anno = c[1].text_input("Anno di nascita", r.get("anno_nascita", ""))
        prov = c[2].text_input("Provenienza (squadra precedente)", r.get("provenienza", ""))
        c = st.columns(3)
        tel = c[0].text_input("Telefono", r.get("telefono", ""))
        mail = c[1].text_input("Email", r.get("email", ""))
        iban = c[2].text_input("IBAN", r.get("iban", ""))
        c = st.columns(3)
        comp = (c[0].number_input("Compenso stagione (€)", min_value=0.0, step=100.0,
                                  value=float(r.get("compenso_stagione", 0) or 0))
                if mostra_importi else float(r.get("compenso_stagione", 0) or 0))
        stati = ["Attivo", "Svincolato", "Ceduto", "Infortunato"]
        stato = c[1].selectbox("Stato", stati, index=stati.index(r["stato"]) if r.get("stato") in stati else 0)
        note = st.text_area("Note", r.get("note", ""), height=80)
        ok = st.form_submit_button("💾 Salva", type="primary")
    if ok:
        if not cognome.strip():
            st.error("Il cognome è obbligatorio.")
            return None
        return {"tipo": tipo, "reparto": reparto, "cognome": cognome.strip().upper(),
                "nome": nome.strip().upper(), "anno_nascita": anno.strip(),
                "provenienza": prov.strip(), "telefono": tel.strip(), "email": mail.strip(),
                "iban": iban.strip().replace(" ", "").upper(), "compenso_stagione": float(comp),
                "stato": stato, "note": note.strip()}
    return None


def pagina_rosa():
    banner("Rosa & Staff", "giocatori, mister, preparatori e contratti")
    tab_rosa, tab_nuovo, tab_scheda = st.tabs(["📋 Rosa", "➕ Nuovo tesserato", "🗂️ Scheda e contratto"])
    t = tesserati_con_saldi()
    imp = vede_importi()
    anno_stg = int(st.session_state["stagione"][:4])

    with tab_rosa:
        if t.empty:
            st.info("Nessun tesserato in questa stagione. Aggiungilo da **➕ Nuovo tesserato**.")
        else:
            c = st.columns([2, 1])
            cerca = c[0].text_input("🔍 Cerca", placeholder="cognome, nome, provenienza…")
            solo = c[1].selectbox("Mostra", ["Tutti", "Solo giocatori", "Solo staff"])
            v = ordina_reparti(t)
            if cerca:
                v = v[v.apply(lambda r: cerca.lower() in " ".join(map(str, r.values)).lower(), axis=1)]
            if solo == "Solo giocatori":
                v = v[v["tipo"] == "Giocatore"]
            elif solo == "Solo staff":
                v = v[v["tipo"] != "Giocatore"]
            for rep in v["reparto"].unique():
                g = v[v["reparto"] == rep].copy()
                g["Età"] = g["anno_nascita"].apply(
                    lambda a: anno_stg - int(a) if str(a).isdigit() else "")
                cols = {"nominativo": "Nominativo", "tipo": "Tipo", "anno_nascita": "Anno",
                        "Età": "Età", "provenienza": "Provenienza", "stato": "Stato"}
                if imp:
                    cols.update({"compenso_stagione": "Compenso", "extra": "Extra",
                                 "totale": "Totale", "pagato": "Pagato", "residuo": "Residuo",
                                 "costo_totale": "Totale speso"})
                st.markdown(f"#### {rep}  <span style='opacity:.6;font-size:14px'>({len(g)})</span>",
                            unsafe_allow_html=True)
                st.dataframe(g[list(cols)].rename(columns=cols), hide_index=True,
                             width="stretch",
                             column_config={k: st.column_config.NumberColumn(format="€ %.2f")
                                            for k in ["Compenso", "Extra", "Totale", "Pagato", "Residuo",
                                                      "Totale speso"]})
            if imp:
                c = st.columns(3)
                kpi(c[0], "Totale contratti", euro(v["totale"].sum()))
                kpi(c[1], "Pagato", euro(v["pagato"].sum()), tono="pos")
                kpi(c[2], "Residuo", euro(v["residuo"].sum()), tono="gold")
                st.write("")
                st.download_button("⬇️ Scarica rosa in Excel", to_excel({"Rosa": v.drop(columns=["id"])}),
                                   f"rosa_{st.session_state['stagione'].replace('/', '-')}.xlsx")

    with tab_nuovo:
        if not blocca_se_sola_lettura():
            rec = form_tesserato("nuovo_tess", mostra_importi=imp)
            if rec:
                rec.update({"id": nuovo_id("T"), "stagione": st.session_state["stagione"]})
                if not imp:
                    rec["compenso_stagione"] = 0.0
                aggiungi("Tesserati", rec)
                avviso(f"✅ {rec['cognome']} {rec['nome']} aggiunto alla rosa.")
                st.rerun()

    with tab_scheda:
        if t.empty:
            st.caption("Nessun tesserato.")
            return
        t = ordina_reparti(t)
        sel = st.selectbox("Seleziona tesserato", t["id"].tolist(),
                           format_func=lambda i: f"{t.set_index('id').loc[i, 'nominativo']} – "
                                                 f"{t.set_index('id').loc[i, 'tipo']}")
        r = t[t["id"] == sel].iloc[0].to_dict()
        if imp:
            st.markdown("#### 📑 Contratto")
            c = st.columns(4)
            kpi(c[0], "Compenso", euro(r["compenso_stagione"]))
            kpi(c[1], "Voci extra a contratto", euro(r["extra"]), "alloggio, vitto, premi… pattuiti")
            kpi(c[2], "Totale contratto", euro(r["totale"]))
            kpi(c[3], "Residuo da pagare", euro(r["residuo"]), "sul contratto", "gold")
            st.markdown("#### 💸 Speso finora per lui")
            c = st.columns(4)
            kpi(c[0], "Pagato al giocatore", euro(r["pagato_diretto"]), "bonifico / contanti", "pos")
            kpi(c[1], "Strutture nel compenso", euro(r["pagato_terzi"]), "scalati dal contratto", "pos")
            kpi(c[2], "Extra fuori contratto", euro(r["spese_extra"]), "strutture a carico società", "neg")
            kpi(c[3], "TOTALE SPESO", euro(r["costo_totale"]), "tutto ciò che è costato", "gold")
            st.write("")

        s1, s2, s3 = st.tabs(["✏️ Dati", "📑 Voci contratto", "💶 Pagamenti e spese"])
        with s1:
            if puo_modificare():
                rec = form_tesserato(f"mod_{sel}", r, mostra_importi=imp)
                if rec:
                    if not imp:
                        rec.pop("compenso_stagione")
                    aggiorna("Tesserati", sel, rec)
                    avviso("Dati aggiornati.")
                    st.rerun()
                with st.expander("🔁 Conferma per un'altra stagione"):
                    stg = [s for s in df("Stagioni")["stagione"].tolist()
                           if s != st.session_state["stagione"]]
                    if not stg:
                        st.caption("Crea prima la nuova stagione da Impostazioni → Stagioni.")
                    else:
                        c = st.columns(2)
                        dest = c[0].selectbox("Stagione", stg, key=f"dst_{sel}")
                        nuovo_comp = c[1].number_input("Nuovo compenso (€)", min_value=0.0,
                                                       step=100.0, value=float(r["compenso_stagione"]),
                                                       key=f"nc_{sel}")
                        if st.button("Conferma tesserato", key=f"conf_{sel}"):
                            nuovo = {k: r[k] for k in SHEETS["Tesserati"]}
                            nuovo.update({"id": nuovo_id("T"), "stagione": dest,
                                          "compenso_stagione": float(nuovo_comp), "stato": "Attivo"})
                            aggiungi("Tesserati", nuovo)
                            st.success(f"Confermato per la stagione {dest}.")
                with st.expander("🗑️ Elimina tesserato"):
                    st.warning("Elimina il tesserato e le sue voci contratto. I movimenti di cassa già "
                               "registrati restano (i soldi sono usciti davvero) ma senza nome collegato.")
                    if st.checkbox("Confermo l'eliminazione", key=f"ck_{sel}") and \
                            st.button("Elimina definitivamente", key=f"del_{sel}"):
                        for vid in stagione_corrente_df("Voci_Contratto").query("tesserato_id == @sel")["id"]:
                            elimina("Voci_Contratto", vid)
                        elimina("Tesserati", sel)
                        st.rerun()
            else:
                for k in ["tipo", "reparto", "anno_nascita", "provenienza", "telefono", "email", "note"]:
                    st.write(f"**{k.replace('_', ' ').capitalize()}:** {r[k]}")

        with s2:
            if not imp:
                st.info("🔒 Il tuo ruolo non vede gli importi.")
            else:
                v = stagione_corrente_df("Voci_Contratto")
                v = v[v["tesserato_id"] == sel]
                if len(v):
                    st.dataframe(v[["voce", "importo", "note"]].rename(
                        columns={"voce": "Voce", "importo": "Importo", "note": "Note"}),
                        hide_index=True, width="stretch",
                        column_config={"Importo": st.column_config.NumberColumn(format="€ %.2f")})
                else:
                    st.caption("Solo compenso base. Aggiungi qui alloggio, vitto, volo, procura, premi…")
                ex = spese_per_tesserati()
                ex = ex[(ex["tesserato_id"] == sel) & (ex["nel_compenso"] != "SI")].copy()
                st.markdown("#### ➕ Extra fuori contratto (pagati a strutture)")
                if ex.empty:
                    st.caption("Nessuna spesa extra: alloggi, visite o terapie pagate alle strutture "
                               "con la scelta *NO – spesa extra* compariranno qui.")
                else:
                    ex["Data"] = ex["data"].apply(fmt_data)
                    st.dataframe(ex[["Data", "beneficiario", "categoria", "descrizione", "importo"]].rename(
                        columns={"beneficiario": "Struttura", "categoria": "Tipo",
                                 "descrizione": "Descrizione", "importo": "Importo"}),
                        hide_index=True, width="stretch",
                        column_config={"Importo": st.column_config.NumberColumn(format="€ %.2f")})
                c = st.columns(3)
                kpi(c[0], "Totale contratto", euro(r["totale"]), "compenso + voci a contratto")
                kpi(c[1], "Extra fuori contratto", euro(r["spese_extra"]), tono="neg")
                kpi(c[2], "Contratto + extra", euro(r["totale"] + r["spese_extra"]), tono="gold")
                st.write("")
                if puo_modificare():
                    st.markdown("##### Aggiungi voce a contratto")
                    with st.form(f"voce_{sel}", clear_on_submit=True):
                        c = st.columns([1.2, 1, 2])
                        voce = c[0].selectbox("Voce", [x for x in categorie("VOCE_CONTRATTO") if x != "Compenso"]
                                              or categorie("VOCE_CONTRATTO"))
                        imp_v = c[1].number_input("Importo (€)", min_value=0.0, step=50.0)
                        nt = c[2].text_input("Note")
                        if st.form_submit_button("➕ Aggiungi voce"):
                            aggiungi("Voci_Contratto", {"id": nuovo_id("V"), "stagione": st.session_state["stagione"],
                                                        "tesserato_id": sel, "voce": voce,
                                                        "importo": float(imp_v), "note": nt})
                            st.rerun()
                    if len(v):
                        dv = st.selectbox("Elimina voce", [""] + v["id"].tolist(),
                                          format_func=lambda i: "—" if not i else
                                          f"{v.set_index('id').loc[i, 'voce']} {euro(v.set_index('id').loc[i, 'importo'])}")
                        if dv and st.button("🗑️ Elimina voce"):
                            elimina("Voci_Contratto", dv)
                            st.rerun()
        with s3:
            if not imp:
                st.info("🔒 Il tuo ruolo non vede gli importi.")
            else:
                st.markdown("#### 💶 Pagati direttamente a lui")
                p = pagamenti_diretti()
                p = p[p["tesserato_id"] == sel].copy()
                if p.empty:
                    st.caption("Nessun pagamento registrato.")
                else:
                    p["Data"] = p["data"].apply(fmt_data)
                    st.dataframe(p[["Data", "mese_rif", "voce", "importo", "metodo", "note"]].rename(
                        columns={"mese_rif": "Mese", "voce": "Voce", "importo": "Importo",
                                 "metodo": "Metodo", "note": "Note"}),
                        hide_index=True, width="stretch",
                        column_config={"Importo": st.column_config.NumberColumn(format="€ %.2f")})
                st.markdown("#### 🏨 Pagati a strutture per lui")
                sp = spese_per_tesserati()
                sp = sp[sp["tesserato_id"] == sel].copy()
                if sp.empty:
                    st.caption("Nessun pagamento a strutture (alloggio, visite, terapie…).")
                else:
                    sp["Data"] = sp["data"].apply(fmt_data)
                    sp["Nel compenso"] = sp["nel_compenso"].map({"SI": "✅ Sì"}).fillna("➖ No (extra)")
                    st.dataframe(sp[["Data", "beneficiario", "categoria", "descrizione", "importo",
                                     "Nel compenso"]].rename(
                        columns={"beneficiario": "Struttura", "categoria": "Tipo",
                                 "descrizione": "Descrizione", "importo": "Importo"}),
                        hide_index=True, width="stretch",
                        column_config={"Importo": st.column_config.NumberColumn(format="€ %.2f")})
                    c = st.columns(3)
                    kpi(c[0], "Strutture nel compenso", euro(sp[sp["nel_compenso"] == "SI"]["importo"].sum()),
                        "scalati dal contratto", "pos")
                    kpi(c[1], "Strutture extra", euro(sp[sp["nel_compenso"] != "SI"]["importo"].sum()),
                        "fuori contratto", "neg")
                    kpi(c[2], "Totale strutture", euro(sp["importo"].sum()), tono="gold")
                    st.caption("Per cambiare la scelta compenso: Cassa & Pagamenti → 🏨 Pagati a strutture.")
                st.markdown("#### 🧮 Totale speso per lui in questa stagione")
                c = st.columns(3)
                kpi(c[0], "Pagato al giocatore", euro(r["pagato_diretto"]), tono="pos")
                kpi(c[1], "Pagato a strutture", euro(r["pagato_terzi"] + r["spese_extra"]))
                kpi(c[2], "TOTALE SPESO", euro(r["costo_totale"]), tono="gold")


# ════════════════════════════════════════════════════════════════════
# CASSA & PAGAMENTI  (registro unico: foglio Movimenti)
# ════════════════════════════════════════════════════════════════════
TIPI_REGISTRAZIONE = ["💶 Pagamento a giocatore / staff", "🏨 Pagamento a struttura per un giocatore",
                      "📤 Altra uscita", "📥 Entrata"]


def pagina_cassa_pagamenti():
    banner("Cassa & Pagamenti", "un solo registro per tutti i soldi della stagione")
    m = stagione_corrente_df("Movimenti")
    ent = m[m["tipo"] == "Entrata"]["importo"].sum()
    usc = m[m["tipo"] == "Uscita"]["importo"].sum()
    segno = m["importo"].where(m["tipo"] == "Entrata", -m["importo"]) if len(m) else pd.Series(dtype=float)
    contanti = segno[m["metodo"] == "Contanti"].sum() if len(m) else 0.0
    if vede_importi():
        c = st.columns(4)
        kpi(c[0], "Entrate", euro(ent), tono="pos")
        kpi(c[1], "Uscite", euro(usc), tono="neg")
        kpi(c[2], "Saldo", euro(ent - usc), tono="pos" if ent >= usc else "neg")
        kpi(c[3], "Saldo contanti", euro(contanti), f"banca/altro {euro(ent - usc - contanti)}", "gold")
        st.write("")
    else:
        st.info("🔒 Il tuo ruolo non vede gli importi.")
        return

    t = ordina_reparti(tesserati_con_saldi())
    nomi = t.set_index("id")["nominativo"].to_dict() if len(t) else {}
    mesi = mesi_stagione()
    mese_corr = f"{MESI[oggi().month - 1]} {oggi().year}"
    schede = st.tabs(["➕ Registra", "📒 Movimenti", "📊 Residui giocatori", "🏨 Pagati a strutture",
                      "🧮 Totale per giocatore", "🗓️ Elenco del mese"])
    with schede[0]:
        sezione_registra(t, nomi, mesi, mese_corr)
    with schede[1]:
        sezione_movimenti(m, nomi)
    with schede[2]:
        sezione_residui(t)
    with schede[3]:
        sezione_pagati_strutture(nomi)
    with schede[4]:
        sezione_totale_per_giocatore(t)
    with schede[5]:
        sezione_elenco_mese(t, nomi, mesi, mese_corr)


def _indice(lista, valore):
    return lista.index(valore) if valore in lista else 0


def sezione_registra(t, nomi, mesi, mese_corr):
    if blocca_se_sola_lettura():
        return
    cosa = st.radio("Cosa vuoi registrare?", TIPI_REGISTRAZIONE, horizontal=True, key="reg_cosa")
    usc, ent, met = categorie("USCITA"), categorie("ENTRATA"), categorie("METODO")
    ids = t["id"].tolist() if len(t) else []
    if cosa in TIPI_REGISTRAZIONE[:2] and not ids:
        st.info("Prima inserisci i tesserati in **Rosa & Staff**.")
        return

    # ── 1. pagamento diretto al giocatore ──
    if cosa == TIPI_REGISTRAZIONE[0]:
        st.caption("Soldi dati **direttamente** al giocatore o allo staff (bonifico, contanti…): "
                   "scalano dal suo compenso e finiscono nell'**Elenco del mese**.")
        with st.form("reg_gioc", clear_on_submit=True):
            c = st.columns([2, 1, 1])
            tid = c[0].selectbox("Tesserato", ids, format_func=lambda i: nomi[i])
            mese = c[1].selectbox("Mese di riferimento", mesi, index=_indice(mesi, mese_corr))
            dt = c[2].date_input("Data pagamento", oggi(), format="DD/MM/YYYY")
            c = st.columns(3)
            voce = c[0].selectbox("Voce", categorie("VOCE_CONTRATTO"))
            imp = c[1].number_input("Importo (€)", min_value=0.0, step=50.0)
            metodo = c[2].selectbox("Metodo", met)
            note = st.text_input("Note")
            ok = st.form_submit_button("💾 Registra pagamento", type="primary")
        if ok:
            if imp <= 0:
                st.error("Inserisci un importo maggiore di zero.")
                return
            r = t[t["id"] == tid].iloc[0]
            cat = VOCE_TO_USCITA.get(voce) or ("Compensi giocatori" if r["tipo"] == "Giocatore"
                                               else "Compensi staff")
            if cat not in usc:
                cat = usc[0]
            registra_movimento("Uscita", dt, cat, f"{voce} {mese} – {nomi[tid]}", imp, metodo, note=note,
                               tesserato_id=tid, destinatario="Giocatore", mese_rif=mese, voce=voce,
                               beneficiario=nomi[tid])
            avviso(f"✅ Pagamento di {euro(imp)} a {nomi[tid]} registrato.")
            st.rerun()

    # ── 2. pagamento a struttura per un giocatore ──
    elif cosa == TIPI_REGISTRAZIONE[1]:
        st.caption("Alloggio, visita medica, terapia… pagati **direttamente alla struttura** per un giocatore. "
                   "Scegli se l'importo **scala dal suo compenso** o è una **spesa extra della società**.")
        forn = df("Fornitori")
        with st.form("reg_strutt", clear_on_submit=True):
            c = st.columns([2, 2, 1])
            tid = c[0].selectbox("Giocatore / tesserato che ne ha usufruito", ids, format_func=lambda i: nomi[i])
            fsel = c[1].selectbox("Struttura (dai fornitori)", [""] + forn["nome"].tolist(),
                                  format_func=lambda x: x or "— scrivi a destra se non è in elenco —")
            altro = c[2].text_input("…oppure nome struttura")
            c = st.columns([1.4, 1, 1, 1])
            cat = c[0].selectbox("Tipo di spesa", usc, index=_indice(usc, "Alloggi"))
            imp = c[1].number_input("Importo (€)", min_value=0.0, step=10.0)
            dt = c[2].date_input("Data", oggi(), format="DD/MM/YYYY")
            metodo = c[3].selectbox("Metodo", met)
            c = st.columns([1, 2])
            mese = c[0].selectbox("Mese di riferimento", mesi, index=_indice(mesi, mese_corr))
            desc = c[1].text_input("Descrizione", placeholder="es. Hotel settembre, visita ortopedica, 5 sedute fisioterapia")
            compenso = st.radio("Va nel compenso del giocatore?", SCELTE_COMPENSO, horizontal=True)
            note = st.text_input("Note")
            ok = st.form_submit_button("💾 Registra pagamento alla struttura", type="primary")
        if ok:
            struttura = (altro.strip() or fsel).strip()
            if imp <= 0 or not struttura:
                st.error("Servono la struttura e un importo maggiore di zero.")
                return
            nc = scelta_compenso(compenso)
            registra_movimento("Uscita", dt, cat, f"{desc.strip() or cat} – {struttura} (per {nomi[tid]})",
                               imp, metodo, note=note, tesserato_id=tid, nel_compenso=nc,
                               destinatario="Struttura", beneficiario=struttura, mese_rif=mese)
            avviso(f"✅ {euro(imp)} a {struttura} per {nomi[tid]} – "
                   f"{'scalato dal compenso' if nc == 'SI' else 'spesa extra società'}.")
            st.rerun()

    # ── 3/4. altra uscita o entrata ──
    else:
        uscita = cosa == TIPI_REGISTRAZIONE[2]
        with st.form("reg_mov", clear_on_submit=True):
            c = st.columns([1, 1.4, 1, 1])
            dt = c[0].date_input("Data", oggi(), format="DD/MM/YYYY")
            cat = c[1].selectbox("Categoria", usc if uscita else ent)
            imp = c[2].number_input("Importo (€)", min_value=0.0, step=10.0)
            metodo = c[3].selectbox("Metodo", met)
            desc = st.text_input("Descrizione *", placeholder="es. Pranzo trasferta Cotronei" if uscita
                                 else "es. Contributo Comune")
            note = st.text_input("Note")
            ok = st.form_submit_button("💾 Registra", type="primary")
        if ok:
            if imp <= 0 or not desc.strip():
                st.error("Servono descrizione e importo.")
                return
            registra_movimento("Uscita" if uscita else "Entrata", dt, cat, desc.strip(), imp, metodo, note=note)
            avviso("✅ Movimento registrato.")
            st.rerun()


def _annulla_pagamento_fattura(mov):
    """Se elimino il pagamento di una fattura, la fattura torna (in parte) da pagare."""
    f = df("Fatture")
    r = f[f["id"] == mov["rif_id"]]
    if len(r):
        r = r.iloc[0]
        pagato = max(r["pagato"] - mov["importo"], 0.0)
        stato = "Pagata" if pagato >= r["importo"] - 0.005 else ("Parziale" if pagato > 0 else "Da pagare")
        aggiorna("Fatture", r["id"], {"pagato": float(pagato), "stato": stato})


DESTINATARI = {"": "— nessun giocatore collegato —", "Giocatore": "👤 Pagato direttamente al giocatore",
               "Struttura": "🏨 Pagato a una struttura per il giocatore"}


def modifica_movimento(v, nomi):
    """Modulo precompilato per correggere un movimento già registrato."""
    mid = st.selectbox("Movimento da modificare", v["id"].tolist(), key="mv_edit", format_func=lambda i:
                       f"{v.set_index('id').loc[i, 'Data']} – {v.set_index('id').loc[i, 'descrizione']} – "
                       f"{euro(v.set_index('id').loc[i, 'importo'])}")
    r = v.set_index("id").loc[mid]
    uscita = r["tipo"] == "Uscita"
    cats = categorie("USCITA" if uscita else "ENTRATA")
    if r["categoria"] and r["categoria"] not in cats:
        cats = [r["categoria"]] + cats
    metodi = categorie("METODO")
    if r["metodo"] and r["metodo"] not in metodi:
        metodi = [r["metodo"]] + metodi
    mesi = [""] + mesi_stagione()
    if r["mese_rif"] and r["mese_rif"] not in mesi:
        mesi.append(r["mese_rif"])
    voci = [""] + categorie("VOCE_CONTRATTO")
    if r["voce"] and r["voce"] not in voci:
        voci.append(r["voce"])
    ids_t = [""] + list(nomi)
    if r["tesserato_id"] and r["tesserato_id"] not in ids_t:
        ids_t.append(r["tesserato_id"])
    if r["rif_tipo"] == "Fattura":
        st.caption("🧾 Pagamento di una fattura: se cambi l'importo aggiorno anche quanto resta da pagare.")
    elif r["rif_tipo"] == "Sponsor":
        st.caption("🤝 Incasso sponsor: l'importo aggiornato vale anche nella pagina Sponsor.")

    with st.form(f"mv_form_{mid}"):
        st.markdown(f"**{'📤 Uscita' if uscita else '📥 Entrata'}**")
        c = st.columns([1, 1.4, 1, 1])
        d0 = parse_date(r["data"]) or oggi()
        dt = c[0].date_input("Data", d0, format="DD/MM/YYYY")
        cat = c[1].selectbox("Categoria", cats, index=_indice(cats, r["categoria"]))
        imp = c[2].number_input("Importo (€)", min_value=0.0, step=10.0, value=float(r["importo"]))
        met = c[3].selectbox("Metodo", metodi, index=_indice(metodi, r["metodo"]))
        desc = st.text_input("Descrizione", r["descrizione"])
        note = st.text_input("Note", r["note"])
        dest, tid, benef, nc, mese, voce = "", "", r["beneficiario"], "", r["mese_rif"], r["voce"]
        if uscita:
            st.markdown("**Collegamento al giocatore**")
            c = st.columns([1.3, 1.6])
            tid = c[0].selectbox("Giocatore", ids_t, index=_indice(ids_t, r["tesserato_id"]),
                                 format_func=lambda i: nomi.get(i, "— nessuno —" if not i else "(tesserato eliminato)"))
            dkeys = list(DESTINATARI)
            dest = c[1].selectbox("Tipo di pagamento", dkeys,
                                  index=_indice(dkeys, r["destinatario"] if r["tesserato_id"] else ""),
                                  format_func=lambda k: DESTINATARI[k])
            c = st.columns([1.3, 1, 1, 1.6])
            benef = c[0].text_input("Struttura / a chi", r["beneficiario"])
            mese = c[1].selectbox("Mese di riferimento", mesi, index=_indice(mesi, r["mese_rif"]),
                                  format_func=lambda x: x or "—")
            voce = c[2].selectbox("Voce", voci, index=_indice(voci, r["voce"]), format_func=lambda x: x or "—")
            nc_radio = c[3].radio("Se è una struttura: va nel compenso?", SCELTE_COMPENSO,
                                  index=1 if r["nel_compenso"] == "SI" else 0)
            nc = scelta_compenso(nc_radio)
        ok = st.form_submit_button("💾 Salva modifiche", type="primary")
    if not ok:
        return
    if imp <= 0 or not desc.strip():
        st.error("Servono descrizione e importo maggiore di zero.")
        return
    if uscita:
        if tid and not dest:
            dest = "Struttura"       # giocatore scelto ma tipo non indicato
        if not tid:
            dest, nc = "", ""
        if dest == "Giocatore":
            nc = "SI"
    campi = {"data": dt.isoformat(), "categoria": cat, "importo": float(imp), "metodo": met,
             "descrizione": desc.strip(), "note": note, "tesserato_id": tid, "destinatario": dest,
             "nel_compenso": nc, "beneficiario": benef.strip(), "mese_rif": mese, "voce": voce,
             "stagione": stagione_da_data(dt) if stagione_da_data(dt) in set(df("Stagioni")["stagione"])
             else r["stagione"]}
    if r["rif_tipo"] == "Fattura" and abs(float(imp) - float(r["importo"])) > 0.004:
        f = df("Fatture")
        fr = f[f["id"] == r["rif_id"]]
        if len(fr):
            fr = fr.iloc[0]
            pagato = max(fr["pagato"] - float(r["importo"]) + float(imp), 0.0)
            stato = "Pagata" if pagato >= fr["importo"] - 0.005 else ("Parziale" if pagato > 0 else "Da pagare")
            aggiorna("Fatture", fr["id"], {"pagato": float(pagato), "stato": stato})
    aggiorna("Movimenti", mid, campi)
    avviso("✅ Movimento aggiornato.")
    st.rerun()


def sezione_movimenti(m, nomi):
    if m.empty:
        st.caption("Nessun movimento.")
        return
    m = m.copy()
    m["d"] = m["data"].apply(parse_date)
    m["Mese"] = m["d"].apply(lambda d: f"{MESI[d.month - 1]} {d.year}" if d else "")
    m["Giocatore"] = m["tesserato_id"].map(nomi).fillna("")
    c = st.columns([1, 1.3, 1, 1.3, 1.4])
    f_tipo = c[0].selectbox("Tipo", ["Tutti", "Entrata", "Uscita"], key="mv_tipo")
    f_cat = c[1].selectbox("Categoria", ["Tutte"] + sorted(m["categoria"].unique().tolist()), key="mv_cat")
    f_mese = c[2].selectbox("Mese", ["Tutti"] + [x for x in mesi_stagione() if x in m["Mese"].values], key="mv_mese")
    gioc = sorted([g for g in m["Giocatore"].unique() if g])
    f_gioc = c[3].selectbox("Giocatore", ["Tutti"] + gioc, key="mv_gioc")
    f_txt = c[4].text_input("🔍 Cerca", key="mv_txt")
    v = m
    if f_tipo != "Tutti":
        v = v[v["tipo"] == f_tipo]
    if f_cat != "Tutte":
        v = v[v["categoria"] == f_cat]
    if f_mese != "Tutti":
        v = v[v["Mese"] == f_mese]
    if f_gioc != "Tutti":
        v = v[v["Giocatore"] == f_gioc]
    if f_txt:
        v = v[v["descrizione"].str.contains(f_txt, case=False, na=False)]
    v = v.sort_values("d", ascending=False).copy()
    v["Data"] = v["data"].apply(fmt_data)
    v["Entrata"] = v["importo"].where(v["tipo"] == "Entrata")
    v["Uscita"] = v["importo"].where(v["tipo"] == "Uscita")
    v["A chi"] = v.apply(lambda r: "👤 Giocatore" if r["destinatario"] == "Giocatore" else
                         (f"🏨 {r['beneficiario']}" if r["destinatario"] == "Struttura" else
                          (r["beneficiario"] or "")), axis=1)
    v["Nel compenso"] = v.apply(lambda r: "" if not r["tesserato_id"] else
                                ("✅ Sì" if r["nel_compenso"] == "SI" else "➖ No"), axis=1)
    out = v[["Data", "descrizione", "categoria", "Entrata", "Uscita", "metodo", "Giocatore", "A chi",
             "Nel compenso", "note"]].rename(columns={"descrizione": "Descrizione", "categoria": "Categoria",
                                                       "metodo": "Metodo", "note": "Note"})
    st.dataframe(out, hide_index=True, width="stretch",
                 column_config={k: st.column_config.NumberColumn(format="€ %.2f") for k in ["Entrata", "Uscita"]})
    st.caption(f"{len(v)} movimenti · entrate {euro(v['Entrata'].sum())} · uscite {euro(v['Uscita'].sum())}")
    st.download_button("⬇️ Scarica movimenti (Excel)", to_excel({"Movimenti": out}), "movimenti.xlsx")
    if puo_modificare() and len(v):
        with st.expander("✏️ Modifica un movimento"):
            modifica_movimento(v, nomi)
        with st.expander("🗑️ Elimina un movimento"):
            mid = st.selectbox("Movimento", v["id"].tolist(), key="mv_del", format_func=lambda i:
                               f"{v.set_index('id').loc[i, 'Data']} – {v.set_index('id').loc[i, 'descrizione']} – "
                               f"{euro(v.set_index('id').loc[i, 'importo'])}")
            mv = v.set_index("id").loc[mid]
            if mv["rif_tipo"] == "Fattura":
                st.caption("È il pagamento di una fattura: eliminandolo la fattura torna da pagare per questo importo.")
            if st.checkbox("Confermo l'eliminazione", key="mv_ck") and st.button("Elimina movimento", key="mv_btn"):
                if mv["rif_tipo"] == "Fattura":
                    _annulla_pagamento_fattura(mv)
                elimina("Movimenti", mid)
                avviso("🗑️ Movimento eliminato.")
                st.rerun()


def sezione_residui(t):
    if t.empty:
        st.caption("Nessun tesserato.")
        return
    colonne = {"nominativo": "Nominativo", "tipo": "Tipo", "reparto": "Reparto",
               "compenso_stagione": "Compenso", "extra": "Extra", "totale": "Totale contratto",
               "pagato_diretto": "Pagato al giocatore", "pagato_terzi": "Pagato a strutture (nel compenso)",
               "residuo": "Residuo", "spese_extra": "Spese extra società", "costo_totale": "Costo totale"}
    v = t[list(colonne)].rename(columns=colonne)
    st.dataframe(v, hide_index=True, width="stretch",
                 column_config={k: st.column_config.NumberColumn(format="€ %.2f")
                                for k in list(colonne.values())[3:]})
    st.caption("**Residuo** = totale contratto − pagato al giocatore − pagato a strutture *nel compenso*. "
               "Le **spese extra** non toccano il residuo ma entrano nel **costo totale** del giocatore.")
    c = st.columns(4)
    kpi(c[0], "Totale contratti", euro(t["totale"].sum()))
    kpi(c[1], "Pagato (diretto + strutture)", euro(t["pagato"].sum()), tono="pos")
    kpi(c[2], "Residuo", euro(t["residuo"].sum()), tono="gold")
    kpi(c[3], "Spese extra società", euro(t["spese_extra"].sum()), "fuori compenso", "neg")
    st.write("")
    st.download_button("⬇️ Scarica residui", to_excel({"Residui": v}), "residui.xlsx")


def sezione_elenco_mese(t, nomi, mesi, mese_corr):
    mese = st.selectbox("Mese", mesi, index=_indice(mesi, mese_corr), key="elenco_mese")
    p = pagamenti_diretti()
    p = p[p["mese_rif"] == mese].copy()
    if p.empty:
        st.caption("Nessun pagamento diretto ai tesserati per questo mese.")
        return
    p["metodo2"] = p["metodo"].apply(lambda x: "CONTANTI" if x == "Contanti" else "BONIFICO")
    pv = p.pivot_table(index="tesserato_id", columns="metodo2", values="importo", aggfunc="sum", fill_value=0.0)
    for col in ["BONIFICO", "CONTANTI"]:
        if col not in pv:
            pv[col] = 0.0
    pv["TOTALE"] = pv["BONIFICO"] + pv["CONTANTI"]
    pv = pv.reset_index()
    pv.insert(0, "GIOCATORE", pv["tesserato_id"].map(nomi).fillna("—"))
    pv["IBAN"] = pv["tesserato_id"].map(t.set_index("id")["iban"] if len(t) else {}).fillna("")
    pv = pv[["GIOCATORE", "BONIFICO", "CONTANTI", "TOTALE", "IBAN"]]
    tot = pd.DataFrame([{"GIOCATORE": "TOTALE", "BONIFICO": pv["BONIFICO"].sum(),
                         "CONTANTI": pv["CONTANTI"].sum(), "TOTALE": pv["TOTALE"].sum(), "IBAN": ""}])
    pv = pd.concat([pv, tot], ignore_index=True)
    st.dataframe(pv, hide_index=True, width="stretch",
                 column_config={k: st.column_config.NumberColumn(format="€ %.2f")
                                for k in ["BONIFICO", "CONTANTI", "TOTALE"]})
    st.download_button("⬇️ Scarica elenco pagamenti", to_excel({mese: pv}),
                       f"pagamenti_{mese.replace(' ', '_')}.xlsx")


def sezione_totale_per_giocatore(t):
    """Riepilogo di fine stagione: quanto è costato ogni tesserato, con dettaglio per tipo di spesa."""
    st.caption(f"Tutto quello che la società ha speso per ogni tesserato nella stagione "
               f"**{st.session_state['stagione']}**: soldi dati a lui + strutture (nel compenso ed extra).")
    solo = st.radio("Mostra", ["Tutti", "Solo giocatori", "Solo staff"], horizontal=True, key="tot_solo")
    v = t.copy()
    if solo == "Solo giocatori":
        v = v[v["tipo"] == "Giocatore"]
    elif solo == "Solo staff":
        v = v[v["tipo"] != "Giocatore"]
    if v.empty:
        st.caption("Nessun tesserato.")
        return
    # dettaglio strutture per tipo di spesa
    sp = spese_per_tesserati()
    det = (sp.pivot_table(index="tesserato_id", columns="categoria", values="importo",
                          aggfunc="sum", fill_value=0.0) if len(sp) else pd.DataFrame())
    out = v[["id", "nominativo", "tipo", "reparto", "totale", "pagato_diretto", "pagato_terzi",
             "spese_extra", "costo_totale", "residuo"]].copy()
    cat_cols = []
    if len(det):
        det.columns = [f"di cui {c}" for c in det.columns]
        cat_cols = list(det.columns)
        out = out.merge(det, left_on="id", right_index=True, how="left").fillna({c: 0.0 for c in cat_cols})
    out = out.sort_values("costo_totale", ascending=False).drop(columns="id")
    nomi_col = {"nominativo": "Nominativo", "tipo": "Tipo", "reparto": "Reparto",
                "totale": "Totale contratto", "pagato_diretto": "Pagato al giocatore",
                "pagato_terzi": "Strutture nel compenso", "spese_extra": "Extra fuori contratto",
                "costo_totale": "TOTALE SPESO", "residuo": "Residuo da pagare"}
    out = out.rename(columns=nomi_col)
    num = [c for c in out.columns if c not in ("Nominativo", "Tipo", "Reparto")]
    tot = {c: out[c].sum() for c in num}
    tot.update({"Nominativo": "TOTALE", "Tipo": "", "Reparto": ""})
    out_tot = pd.concat([out, pd.DataFrame([tot])], ignore_index=True)

    c = st.columns(4)
    kpi(c[0], "Totale speso", euro(tot["TOTALE SPESO"]), f"{len(out)} tesserati", "gold")
    kpi(c[1], "Dato ai tesserati", euro(tot["Pagato al giocatore"]), tono="pos")
    kpi(c[2], "Pagato a strutture", euro(tot["Strutture nel compenso"] + tot["Extra fuori contratto"]))
    media = tot["TOTALE SPESO"] / len(out) if len(out) else 0
    kpi(c[3], "Media per tesserato", euro(media))
    st.write("")
    st.dataframe(out_tot, hide_index=True, width="stretch",
                 column_config={k: st.column_config.NumberColumn(format="€ %.2f") for k in num})

    graf = out[out["TOTALE SPESO"] > 0].head(15)
    if len(graf):
        st.markdown("#### Chi è costato di più")
        lunga = graf.melt(id_vars="Nominativo",
                          value_vars=["Pagato al giocatore", "Strutture nel compenso", "Extra fuori contratto"],
                          var_name="Voce", value_name="Euro")
        ch = alt.Chart(lunga).mark_bar(cornerRadiusEnd=3).encode(
            y=alt.Y("Nominativo:N", sort=graf["Nominativo"].tolist(), title=None),
            x=alt.X("sum(Euro):Q", title="€", stack="zero"),
            color=alt.Color("Voce:N", scale=alt.Scale(
                domain=["Pagato al giocatore", "Strutture nel compenso", "Extra fuori contratto"],
                range=[NEON, ORO, ROSSO]), legend=alt.Legend(orient="top", title=None)),
            tooltip=["Nominativo", "Voce", alt.Tooltip("Euro:Q", format=",.2f")])
        st.altair_chart(grafico_tema(ch.properties(height=max(220, 28 * len(graf)))), width="stretch")
    st.download_button("⬇️ Scarica totale per giocatore (Excel)", to_excel({"Totale per giocatore": out_tot}),
                       f"totale_per_giocatore_{st.session_state['stagione'].replace('/', '-')}.xlsx")


def sezione_pagati_strutture(nomi):
    st.caption("Alloggi, visite, terapie… pagati alle strutture per conto dei giocatori. "
               "Per registrarne uno nuovo: **➕ Registra → 🏨 Pagamento a struttura**.")
    sp = spese_per_tesserati()
    if sp.empty:
        st.caption("Nessun pagamento a strutture registrato.")
        return
    sp["d"] = sp["data"].apply(parse_date)
    sp = sp.sort_values("d", ascending=False)
    sp["Giocatore"] = sp["tesserato_id"].map(nomi).fillna("—")
    sp["Data"] = sp["data"].apply(fmt_data)
    sp["Nel compenso"] = sp["nel_compenso"].map({"SI": "✅ Sì"}).fillna("➖ No (extra)")
    f = st.selectbox("Filtra giocatore", ["Tutti"] + sorted(sp["Giocatore"].unique().tolist()), key="filtro_strutt")
    v = sp if f == "Tutti" else sp[sp["Giocatore"] == f]
    st.dataframe(v[["Data", "Giocatore", "beneficiario", "categoria", "descrizione", "importo", "metodo",
                    "mese_rif", "Nel compenso"]].rename(
        columns={"beneficiario": "Struttura", "categoria": "Tipo", "descrizione": "Descrizione",
                 "importo": "Importo", "metodo": "Metodo", "mese_rif": "Mese"}),
        hide_index=True, width="stretch",
        column_config={"Importo": st.column_config.NumberColumn(format="€ %.2f")})
    c = st.columns(3)
    kpi(c[0], "Nel compenso", euro(v[v["nel_compenso"] == "SI"]["importo"].sum()), "scala dal residuo", "gold")
    kpi(c[1], "Spese extra società", euro(v[v["nel_compenso"] != "SI"]["importo"].sum()), "fuori compenso", "neg")
    kpi(c[2], "Totale strutture", euro(v["importo"].sum()))
    if puo_modificare():
        st.write("")
        with st.expander("🔁 Cambia scelta compenso"):
            mid = st.selectbox("Pagamento", v["id"].tolist(), key="sel_strutt", format_func=lambda i:
                               f"{v.set_index('id').loc[i, 'Data']} – {v.set_index('id').loc[i, 'Giocatore']} – "
                               f"{v.set_index('id').loc[i, 'beneficiario']} – {euro(v.set_index('id').loc[i, 'importo'])}")
            attuale = v.set_index("id").loc[mid, "nel_compenso"]
            etichetta = ("➖ Togli dal compenso (diventa spesa extra)" if attuale == "SI"
                         else "✅ Metti nel compenso (scala dal residuo)")
            if st.button(etichetta, key="toggle_strutt"):
                aggiorna("Movimenti", mid, {"nel_compenso": "NO" if attuale == "SI" else "SI"})
                avviso("✅ Scelta aggiornata.")
                st.rerun()
            st.caption("Per eliminarlo: scheda **📒 Movimenti → Elimina un movimento**.")


# ════════════════════════════════════════════════════════════════════
# FORNITORI & FATTURE
# ════════════════════════════════════════════════════════════════════
def pagina_fatture():
    banner("Fornitori & Fatture", "cosa dobbiamo pagare")
    forn = df("Fornitori")
    nomi_f = forn.set_index("id")["nome"].to_dict()
    f = stagione_corrente_df("Fatture")
    tab1, tab2, tab3 = st.tabs(["🧾 Fatture", "➕ Nuova fattura", "🏢 Fornitori"])

    with tab1:
        if f.empty:
            st.caption("Nessuna fattura in questa stagione.")
        else:
            f["residuo"] = f["importo"] - f["pagato"]
            f["Fornitore"] = f["fornitore_id"].map(nomi_f).fillna("—")
            f["Giocatore"] = f["tesserato_id"].map(mappa_tesserati()).fillna("")
            f["d_scad"] = f["scadenza"].apply(parse_date)
            c = st.columns(3)
            kpi(c[0], "Totale fatture", euro(f["importo"].sum()))
            kpi(c[1], "Pagato", euro(f["pagato"].sum()), tono="pos")
            scad = f[(f["stato"] != "Pagata") & f["d_scad"].notna()]
            scad = scad[scad["d_scad"] < oggi()]
            kpi(c[2], "Da pagare", euro(f["residuo"].clip(lower=0).sum()),
                f"{len(scad)} scadute", "neg" if len(scad) else "gold")
            st.write("")
            fs = st.selectbox("Stato", ["Aperte", "Tutte", "Pagate"])
            v = f if fs == "Tutte" else (f[f["stato"] == "Pagata"] if fs == "Pagate" else f[f["stato"] != "Pagata"])
            v = v.sort_values("d_scad", na_position="last")
            out = v.assign(Data=v["data"].apply(fmt_data), Scadenza=v["scadenza"].apply(fmt_data))[
                ["Fornitore", "numero", "Data", "Scadenza", "categoria", "descrizione", "Giocatore", "importo",
                 "pagato", "residuo", "stato"]].rename(columns={"numero": "N.", "categoria": "Categoria",
                                                      "descrizione": "Descrizione", "importo": "Importo",
                                                      "pagato": "Pagato", "residuo": "Residuo", "stato": "Stato"})
            st.dataframe(out, hide_index=True, width="stretch",
                         column_config={k: st.column_config.NumberColumn(format="€ %.2f")
                                        for k in ["Importo", "Pagato", "Residuo"]})
            aperte = f[f["stato"] != "Pagata"]
            if puo_modificare() and len(aperte):
                st.markdown("##### 💳 Registra pagamento fattura")
                with st.form("pag_fatt", clear_on_submit=True):
                    c = st.columns([2, 1, 1, 1])
                    fid = c[0].selectbox("Fattura", aperte["id"].tolist(), format_func=lambda i:
                                         f"{aperte.set_index('id').loc[i, 'Fornitore']} n.{aperte.set_index('id').loc[i, 'numero']} "
                                         f"– residuo {euro(aperte.set_index('id').loc[i, 'residuo'])}")
                    imp = c[1].number_input("Importo (€) – 0 = saldo", min_value=0.0, step=10.0)
                    dt = c[2].date_input("Data", oggi(), format="DD/MM/YYYY")
                    met = c[3].selectbox("Metodo", categorie("METODO"))
                    if st.form_submit_button("Registra pagamento", type="primary"):
                        r = aperte.set_index("id").loc[fid]
                        imp = min(imp or r["residuo"], r["residuo"])
                        pagato = r["pagato"] + imp
                        stato = "Pagata" if pagato >= r["importo"] - 0.005 else "Parziale"
                        aggiorna("Fatture", fid, {"pagato": float(pagato), "stato": stato})
                        registra_movimento("Uscita", dt, r["categoria"],
                                           f"Fattura n.{r['numero']} – {r['Fornitore']}", imp, met, "Fattura", fid,
                                           tesserato_id=r["tesserato_id"], nel_compenso=r["nel_compenso"],
                                           beneficiario=r["Fornitore"])
                        avviso("Pagamento registrato anche in Cassa.")
                        st.rerun()
            if puo_modificare():
                with st.expander("🗑️ Elimina fattura"):
                    did = st.selectbox("Fattura da eliminare", f["id"].tolist(), format_func=lambda i:
                                       f"{f.set_index('id').loc[i, 'Fornitore']} n.{f.set_index('id').loc[i, 'numero']}")
                    if f.set_index("id").loc[did, "pagato"] > 0:
                        st.caption("Questa fattura ha pagamenti registrati: i movimenti in Cassa restano.")
                    if st.button("Elimina fattura"):
                        elimina("Fatture", did)
                        st.rerun()

    with tab2:
        if not blocca_se_sola_lettura():
            if forn.empty:
                st.info("Aggiungi prima un fornitore nella scheda **🏢 Fornitori**.")
            else:
                with st.form("fatt", clear_on_submit=True):
                    c = st.columns([2, 1, 1, 1])
                    fid = c[0].selectbox("Fornitore", forn["id"].tolist(), format_func=lambda i: nomi_f[i])
                    num = c[1].text_input("Numero fattura")
                    dt = c[2].date_input("Data", oggi(), format="DD/MM/YYYY")
                    sc = c[3].date_input("Scadenza", oggi() + timedelta(days=30), format="DD/MM/YYYY")
                    c = st.columns([1.3, 2, 1])
                    cu = categorie("USCITA")
                    cat = c[0].selectbox("Categoria", cu, index=cu.index("Materiale sportivo")
                                         if "Materiale sportivo" in cu else 0)
                    desc = c[1].text_input("Descrizione")
                    imp = c[2].number_input("Importo totale (€)", min_value=0.0, step=10.0)
                    note = st.text_input("Note")
                    mt = mappa_tesserati()
                    c = st.columns([1, 1.6])
                    tid = c[0].selectbox("Giocatore di riferimento (facoltativo)", [""] + list(mt),
                                         format_func=lambda i: mt.get(i, "— nessuno —"), key="fatt_tess")
                    compenso = c[1].radio("Se c'è un giocatore: va nel suo compenso?", SCELTE_COMPENSO,
                                          key="fatt_comp")
                    nc = scelta_compenso(compenso) if tid else ""
                    gia = st.checkbox("Già pagata (registra subito in cassa)")
                    met = st.selectbox("Metodo (se già pagata)", categorie("METODO"))
                    if st.form_submit_button("💾 Salva fattura", type="primary"):
                        if imp <= 0:
                            st.error("Inserisci l'importo.")
                        else:
                            nid = nuovo_id("F")
                            aggiungi("Fatture", {"id": nid, "stagione": st.session_state["stagione"],
                                                 "fornitore_id": fid, "numero": num, "data": dt.isoformat(),
                                                 "scadenza": sc.isoformat(), "categoria": cat, "descrizione": desc,
                                                 "importo": float(imp), "pagato": float(imp) if gia else 0.0,
                                                 "stato": "Pagata" if gia else "Da pagare", "note": note,
                                                 "tesserato_id": tid, "nel_compenso": nc})
                            if gia:
                                registra_movimento("Uscita", dt, cat, f"Fattura n.{num} – {nomi_f[fid]}",
                                                   imp, met, "Fattura", nid, tesserato_id=tid,
                                                   nel_compenso=nc, beneficiario=nomi_f[fid])
                            avviso("Fattura salvata.")
                            st.rerun()

    with tab3:
        if puo_modificare():
            with st.form("forn", clear_on_submit=True):
                c = st.columns([2, 1, 1])
                nome = c[0].text_input("Ragione sociale *")
                cat = c[1].selectbox("Categoria", categorie("FORNITORE"))
                piva = c[2].text_input("P.IVA / C.F.")
                c = st.columns(3)
                tel = c[0].text_input("Telefono")
                mail = c[1].text_input("Email")
                iban = c[2].text_input("IBAN")
                note = st.text_input("Note")
                if st.form_submit_button("➕ Aggiungi fornitore"):
                    if nome.strip():
                        aggiungi("Fornitori", {"id": nuovo_id("FO"), "nome": nome.strip(), "categoria": cat,
                                               "piva": piva, "telefono": tel, "email": mail,
                                               "iban": iban.replace(" ", "").upper(), "note": note})
                        st.rerun()
                    else:
                        st.error("Inserisci il nome.")
        if len(forn):
            tutte = df("Fatture")
            tutte["res"] = tutte["importo"] - tutte["pagato"]
            forn["Da pagare"] = forn["id"].map(tutte.groupby("fornitore_id")["res"].sum()).fillna(0.0)
            st.dataframe(forn.drop(columns=["id"]).rename(columns=str.capitalize), hide_index=True,
                         width="stretch",
                         column_config={"Da pagare": st.column_config.NumberColumn(format="€ %.2f")})
            if puo_modificare():
                with st.expander("🗑️ Elimina fornitore"):
                    did = st.selectbox("Fornitore", forn["id"].tolist(), format_func=lambda i: nomi_f[i])
                    if st.button("Elimina fornitore"):
                        if (df("Fatture")["fornitore_id"] == did).any():
                            st.error("Questo fornitore ha fatture registrate: elimina prima le fatture.")
                        else:
                            elimina("Fornitori", did)
                            st.rerun()
        st.caption("Il fornitore è comune a tutte le stagioni. Allegati PDF: in arrivo in una prossima versione.")


# ════════════════════════════════════════════════════════════════════
# SPONSOR
# ════════════════════════════════════════════════════════════════════
def pagina_sponsor():
    banner("Sponsor", "accordi e incassi")
    sp = stagione_corrente_df("Sponsor")
    m = stagione_corrente_df("Movimenti")
    inc = m[m["rif_tipo"] == "Sponsor"].groupby("rif_id")["importo"].sum()
    if len(sp):
        sp["incassato"] = sp["id"].map(inc).fillna(0.0)
        sp["da_incassare"] = sp["importo_concordato"] - sp["incassato"]
        c = st.columns(3)
        kpi(c[0], "Concordato", euro(sp["importo_concordato"].sum()))
        kpi(c[1], "Incassato", euro(sp["incassato"].sum()), tono="pos")
        kpi(c[2], "Da incassare", euro(sp["da_incassare"].clip(lower=0).sum()), tono="gold")
        st.write("")
        st.dataframe(sp[["nome", "referente", "telefono", "importo_concordato", "incassato", "da_incassare", "note"]]
                     .rename(columns={"nome": "Sponsor", "referente": "Referente", "telefono": "Telefono",
                                      "importo_concordato": "Concordato", "incassato": "Incassato",
                                      "da_incassare": "Da incassare", "note": "Note"}),
                     hide_index=True, width="stretch",
                     column_config={k: st.column_config.NumberColumn(format="€ %.2f")
                                    for k in ["Concordato", "Incassato", "Da incassare"]})
    else:
        st.caption("Nessuno sponsor in questa stagione.")
    if not puo_modificare():
        return
    c1, c2 = st.columns(2)
    with c1:
        st.markdown("##### ➕ Nuovo sponsor")
        with st.form("sp", clear_on_submit=True):
            nome = st.text_input("Nome sponsor *")
            a, b = st.columns(2)
            ref = a.text_input("Referente")
            tel = b.text_input("Telefono")
            imp = st.number_input("Importo concordato (€)", min_value=0.0, step=100.0)
            note = st.text_input("Note (es. maglia, cartellone, social)")
            salva_sp = st.form_submit_button("Salva sponsor", type="primary")
            if salva_sp and not nome.strip():
                st.error("Scrivi il nome dello sponsor.")
            if salva_sp and nome.strip():
                aggiungi("Sponsor", {"id": nuovo_id("S"), "stagione": st.session_state["stagione"],
                                     "nome": nome.strip(), "referente": ref, "telefono": tel,
                                     "importo_concordato": float(imp), "note": note})
                st.rerun()
    with c2:
        st.markdown("##### 💶 Registra incasso")
        if sp.empty:
            st.caption("Aggiungi prima uno sponsor.")
        else:
            with st.form("inc", clear_on_submit=True):
                sid = st.selectbox("Sponsor", sp["id"].tolist(),
                                   format_func=lambda i: sp.set_index("id").loc[i, "nome"])
                a, b, c = st.columns(3)
                imp = a.number_input("Importo (€)", min_value=0.0, step=100.0)
                dt = b.date_input("Data", oggi(), format="DD/MM/YYYY")
                met = c.selectbox("Metodo", categorie("METODO"))
                salva_inc = st.form_submit_button("Registra incasso", type="primary")
                if salva_inc and imp <= 0:
                    st.error("Inserisci un importo maggiore di zero.")
                if salva_inc and imp > 0:
                    cat = "Sponsor" if "Sponsor" in categorie("ENTRATA") else categorie("ENTRATA")[0]
                    registra_movimento("Entrata", dt, cat, f"Sponsor – {sp.set_index('id').loc[sid, 'nome']}",
                                       imp, met, "Sponsor", sid)
                    st.rerun()
        if len(sp):
            with st.expander("🗑️ Elimina sponsor"):
                did = st.selectbox("Sponsor da eliminare", sp["id"].tolist(),
                                   format_func=lambda i: sp.set_index("id").loc[i, "nome"], key="delsp")
                if st.button("Elimina sponsor"):
                    elimina("Sponsor", did)
                    st.rerun()


# ════════════════════════════════════════════════════════════════════
# SCADENZE
# ════════════════════════════════════════════════════════════════════
def pagina_scadenze():
    banner("Scadenze", "visite mediche, tesseramenti, iscrizioni")
    sc = stagione_corrente_df("Scadenze")
    mt = mappa_tesserati()
    if puo_modificare():
        with st.expander("➕ Nuova scadenza", expanded=sc.empty):
            with st.form("sc", clear_on_submit=True):
                c = st.columns([1, 2, 1.5, 1])
                tipo = c[0].selectbox("Tipo", categorie("SCADENZA"))
                desc = c[1].text_input("Descrizione *")
                tid = c[2].selectbox("Tesserato (facoltativo)", [""] + list(mt),
                                     format_func=lambda i: mt.get(i, "—"))
                dt = c[3].date_input("Scadenza", oggi() + timedelta(days=30), format="DD/MM/YYYY")
                note = st.text_input("Note")
                salva_sc = st.form_submit_button("Salva", type="primary")
                if salva_sc and not desc.strip():
                    st.error("Scrivi una descrizione.")
                if salva_sc and desc.strip():
                    aggiungi("Scadenze", {"id": nuovo_id("SC"), "stagione": st.session_state["stagione"],
                                          "tipo": tipo, "descrizione": desc.strip(), "tesserato_id": tid,
                                          "data_scadenza": dt.isoformat(), "stato": "Da fare", "note": note})
                    st.rerun()
    if sc.empty:
        st.caption("Nessuna scadenza inserita.")
        return
    sc["d"] = sc["data_scadenza"].apply(parse_date)
    sc = sc.sort_values("d", na_position="last")

    def semaforo(r):
        if r["stato"] == "Fatto":
            return "✅ Fatto"
        if r["d"] and r["d"] < oggi():
            return "🔴 Scaduta"
        if r["d"] and r["d"] <= oggi() + timedelta(days=30):
            return "🟡 Entro 30 gg"
        return "🟢 In regola"
    sc["Stato"] = sc.apply(semaforo, axis=1)
    sc["Tesserato"] = sc["tesserato_id"].map(mt).fillna("")
    sc["Data"] = sc["data_scadenza"].apply(fmt_data)
    filtro = st.radio("Mostra", ["Da fare", "Tutte"], horizontal=True)
    v = sc[sc["stato"] != "Fatto"] if filtro == "Da fare" else sc
    st.dataframe(v[["Stato", "Data", "tipo", "descrizione", "Tesserato", "note"]].rename(
        columns={"tipo": "Tipo", "descrizione": "Descrizione", "note": "Note"}),
        hide_index=True, width="stretch")
    if puo_modificare() and len(v):
        c = st.columns([3, 1, 1])
        sid = c[0].selectbox("Scadenza", v["id"].tolist(), format_func=lambda i:
                             f"{v.set_index('id').loc[i, 'Data']} – {v.set_index('id').loc[i, 'descrizione']}")
        if c[1].button("✅ Segna come fatta"):
            aggiorna("Scadenze", sid, {"stato": "Fatto"})
            st.rerun()
        if c[2].button("🗑️ Elimina"):
            elimina("Scadenze", sid)
            st.rerun()


# ════════════════════════════════════════════════════════════════════
# IMPOSTAZIONI
# ════════════════════════════════════════════════════════════════════
def pagina_impostazioni():
    banner("Impostazioni", "squadra, stagioni, categorie, ruoli e utenti")
    if not puo_modificare():
        st.info("🔒 Solo lettura.")
        return
    t1, t2, t3, t4, t5, t6 = st.tabs(["⚽ Squadra", "📆 Stagioni", "🏷️ Categorie", "🛡️ Ruoli", "👤 Utenti", "📜 Storico"])

    with t1:
        with st.form("squadra"):
            c = st.columns(2)
            nome = c[0].text_input("Nome squadra", impostazione("nome_squadra"))
            citta = c[1].text_input("Città", impostazione("citta"))
            c = st.columns(2)
            c1 = c[0].color_picker("Colore sociale 1", impostazione("colore_1", "#1f7a3a"))
            c2 = c[1].color_picker("Colore sociale 2", impostazione("colore_2", "#ffffff"))
            if st.form_submit_button("💾 Salva", type="primary"):
                for k, v in {"nome_squadra": nome, "citta": citta, "colore_1": c1, "colore_2": c2}.items():
                    salva_impostazione(k, v)
                st.rerun()

    with t2:
        s = df("Stagioni")
        st.dataframe(s.assign(data_inizio=s["data_inizio"].apply(fmt_data),
                              data_fine=s["data_fine"].apply(fmt_data))
                     .rename(columns={"stagione": "Stagione", "data_inizio": "Inizio", "data_fine": "Fine",
                                      "attiva": "Attiva"}), hide_index=True, width="stretch")
        st.markdown("##### ➕ Nuova stagione")
        ultimo = max([int(x[:4]) for x in s["stagione"] if str(x)[:4].isdigit()] or [oggi().year - 1])
        with st.form("stg"):
            c = st.columns(3)
            a = c[0].number_input("Anno di inizio", value=ultimo + 1, step=1, format="%d")
            di = c[1].date_input("Inizio", date(int(a), 7, 1), format="DD/MM/YYYY")
            dfin = c[2].date_input("Fine", date(int(a) + 1, 6, 30), format="DD/MM/YYYY")
            att = st.checkbox("Rendila la stagione attiva")
            if st.form_submit_button("Crea stagione", type="primary"):
                nome = stagione_da_anno(int(a))
                if (s["stagione"] == nome).any():
                    st.error("Esiste già.")
                else:
                    if att:
                        aggiorna_molti("Stagioni", {x: {"attiva": "NO"} for x in s["stagione"]},
                                       chiave_col="stagione")
                    aggiungi("Stagioni", {"stagione": nome, "data_inizio": di.isoformat(),
                                          "data_fine": dfin.isoformat(), "attiva": "SI" if att else "NO"})
                    avviso(f"Stagione {nome} creata.")
                    st.rerun()
        if len(s) > 1:
            c = st.columns([2, 1])
            nuova = c[0].selectbox("Imposta stagione attiva (quella che si apre di default)", s["stagione"].tolist())
            if c[1].button("Imposta come attiva"):
                aggiorna_molti("Stagioni", {x: {"attiva": "SI" if x == nuova else "NO"} for x in s["stagione"]},
                               chiave_col="stagione")
                st.rerun()

    with t3:
        st.caption("Tutto quello che compare nei menu a tendina si decide qui.")
        tipo = st.selectbox("Gruppo", list(TIPI_CATEGORIA), format_func=lambda k: TIPI_CATEGORIA[k])
        cat = df("Categorie")
        lista = cat[cat["tipo"] == tipo]["nome"].tolist()
        st.write(" · ".join(f"`{x}`" for x in lista) or "_vuoto: uso i valori di default_")
        c = st.columns([2, 1])
        nuova = c[0].text_input("Nuova voce", key=f"nv_{tipo}")
        if c[1].button("➕ Aggiungi", key=f"add_{tipo}") and nuova.strip():
            if nuova.strip() in lista:
                st.warning("Esiste già.")
            else:
                aggiungi("Categorie", {"tipo": tipo, "nome": nuova.strip()})
                st.rerun()
        if lista:
            c = st.columns([2, 1])
            via = c[0].selectbox("Rimuovi voce", lista, key=f"rv_{tipo}")
            if c[1].button("🗑️ Rimuovi", key=f"del_{tipo}"):
                elimina("Categorie", via, chiave_col="nome", extra={"tipo": tipo})
                st.rerun()

    with t4:
        r = df("Ruoli")
        st.dataframe(r.rename(columns={"ruolo": "Ruolo", "pagine": "Pagine visibili",
                                       "puo_modificare": "Può modificare", "vede_importi": "Vede importi"}),
                     hide_index=True, width="stretch")
        admin = impostazione("ruolo_admin", "Presidente")
        scelta = st.selectbox("Crea o modifica ruolo", ["➕ Nuovo ruolo"] + r["ruolo"].tolist())
        esist = r[r["ruolo"] == scelta].iloc[0].to_dict() if scelta in r["ruolo"].values else {}
        with st.form(f"ruolo_{scelta}"):
            nome = st.text_input("Nome ruolo (es. Segreteria, Mister, Direttore sportivo)",
                                 esist.get("ruolo", ""), disabled=bool(esist))
            pag = st.multiselect("Pagine che può vedere", PAGINE,
                                 default=list(dict.fromkeys(
                                     PAGINE_ALIAS.get(p.strip(), p.strip())
                                     for p in str(esist.get("pagine", "Panoramica")).split(",")
                                     if PAGINE_ALIAS.get(p.strip(), p.strip()) in PAGINE)))
            c = st.columns(2)
            mod = c[0].checkbox("Può inserire / modificare", esist.get("puo_modificare", "NO") == "SI")
            imp = c[1].checkbox("Vede gli importi (€)", esist.get("vede_importi", "NO") == "SI")
            if st.form_submit_button("💾 Salva ruolo", type="primary"):
                rec = {"pagine": ",".join(pag), "puo_modificare": "SI" if mod else "NO",
                       "vede_importi": "SI" if imp else "NO"}
                if esist:
                    if esist["ruolo"] == admin and ("Impostazioni" not in pag or not mod):
                        st.error(f"Il ruolo {admin} deve mantenere Impostazioni e modifica.")
                    else:
                        aggiorna("Ruoli", esist["ruolo"], rec, chiave_col="ruolo")
                        st.rerun()
                elif nome.strip() and not (r["ruolo"] == nome.strip()).any():
                    aggiungi("Ruoli", {"ruolo": nome.strip(), **rec})
                    st.rerun()
                else:
                    st.error("Nome mancante o già esistente.")
        if esist and esist["ruolo"] != admin:
            if st.button(f"🗑️ Elimina ruolo {esist['ruolo']}"):
                if (df("Utenti")["ruolo"] == esist["ruolo"]).any():
                    st.error("Ci sono utenti con questo ruolo: cambiali prima.")
                else:
                    elimina("Ruoli", esist["ruolo"], chiave_col="ruolo")
                    st.rerun()

    with t5:
        u = df("Utenti")
        st.dataframe(u[["username", "nome", "ruolo", "attivo"]].rename(columns=str.capitalize),
                     hide_index=True, width="stretch")
        ruoli = df("Ruoli")["ruolo"].tolist()
        st.markdown("##### ➕ Nuovo utente")
        with st.form("utente", clear_on_submit=True):
            c = st.columns(4)
            un = c[0].text_input("Username")
            nm = c[1].text_input("Nome")
            pw = c[2].text_input("Password", type="password")
            rl = c[3].selectbox("Ruolo", ruoli)
            if st.form_submit_button("Crea utente", type="primary"):
                un = un.strip().lower()
                if not un or len(pw) < 6:
                    st.error("Username e password (min. 6) obbligatori.")
                elif (u["username"] == un).any():
                    st.error("Username già usato.")
                else:
                    aggiungi("Utenti", {"username": un, "password_hash": hash_pw(pw), "nome": nm,
                                        "ruolo": rl, "attivo": "SI"})
                    st.rerun()
        st.markdown("##### ✏️ Modifica utente")
        c = st.columns(4)
        sel = c[0].selectbox("Utente", u["username"].tolist())
        if sel:
            r = u[u["username"] == sel].iloc[0]
            nr = c[1].selectbox("Ruolo", ruoli, index=ruoli.index(r["ruolo"]) if r["ruolo"] in ruoli else 0,
                                key=f"mod_ruolo_{sel}")
            npw = c[2].text_input("Nuova password (vuoto = invariata)", type="password", key=f"mod_pw_{sel}")
            att = c[3].selectbox("Attivo", ["SI", "NO"], index=0 if r["attivo"] == "SI" else 1,
                                 key=f"mod_att_{sel}")
            io_stesso = sel == st.session_state["utente"]["username"]
            if st.button("💾 Salva utente"):
                if io_stesso and (att == "NO" or nr != r["ruolo"]):
                    st.error("Non puoi disattivare o cambiare il ruolo a te stesso.")
                else:
                    campi = {"ruolo": nr, "attivo": att}
                    if npw:
                        if len(npw) < 6:
                            st.error("Password troppo corta.")
                            st.stop()
                        campi["password_hash"] = hash_pw(npw)
                    aggiorna("Utenti", sel, campi, chiave_col="username")
                    avviso("Utente aggiornato.")
                    st.rerun()

    with t6:
        lg = df("Log")
        st.dataframe(lg.iloc[::-1].head(300).rename(columns=str.capitalize), hide_index=True,
                     width="stretch")
        st.markdown("##### 💾 Backup")
        st.download_button("⬇️ Scarica backup completo (Excel)",
                           to_excel({n: df(n).drop(columns=["password_hash"], errors="ignore") for n in SHEETS}),
                           f"backup_squadra_{oggi().isoformat()}.xlsx")


# ════════════════════════════════════════════════════════════════════
# AVVIO
# ════════════════════════════════════════════════════════════════════
def main():
    try:
        leggi_tutto()
    except Exception as e:
        applica_stile()
        st.error("Non riesco a collegarmi al Google Sheet. Controlla i *Secrets* "
                 "(gcp_service_account e SHEET_ID) e che il foglio sia condiviso con l'email del service account.")
        if "PEM" in str(e) or "private" in str(e).lower():
            st.warning("Il problema è la **private_key** nei Secrets. Diagnosi:")
            st.markdown(diagnosi_chiave())
        st.exception(e)
        st.stop()

    if impostazione("setup_completato") != "SI" or df("Utenti").empty:
        pagina_setup()
        return
    migra_dati()
    if "utente" not in st.session_state:
        pagina_login()
        return

    applica_stile(impostazione("colore_1", "#1f7a3a"), impostazione("colore_2", "#ffffff"))
    msg = st.session_state.pop("_avviso", None)
    if msg:
        st.toast(msg)
    stagioni = df("Stagioni")
    elenco = sorted(stagioni["stagione"].tolist(), reverse=True) or [stagione_da_data(oggi())]
    attiva = stagioni[stagioni["attiva"] == "SI"]["stagione"].tolist()
    if st.session_state.get("stagione") not in elenco:
        st.session_state["stagione"] = attiva[0] if attiva else elenco[0]

    info = ruolo_info()
    with st.sidebar:
        stemma()
        st.selectbox("Stagione", elenco, key="stagione")
        pagine = [p for p in PAGINE if p in info["pagine"]] or ["Panoramica"]
        pagina = st.radio("Menu", pagine, format_func=lambda p: f"{ICONE[p]}  {p}", label_visibility="collapsed")
        st.divider()
        u = st.session_state["utente"]
        st.caption(f"👤 {u['nome']} · {u['ruolo']}")
        st.caption(f"🕒 {datetime.now(TZ).strftime('%d/%m/%Y %H:%M')}")
        c = st.columns(2)
        if c[0].button("🔄 Aggiorna", width="stretch"):
            leggi_tutto.clear()
            st.rerun()
        if c[1].button("🚪 Esci", width="stretch"):
            for k in ["utente", "stagione"]:
                st.session_state.pop(k, None)
            st.rerun()

    {"Panoramica": pagina_panoramica, "Rosa & Staff": pagina_rosa, "Cassa & Pagamenti": pagina_cassa_pagamenti,
     "Fornitori & Fatture": pagina_fatture, "Sponsor": pagina_sponsor,
     "Scadenze": pagina_scadenze, "Impostazioni": pagina_impostazioni}[pagina]()


main()
