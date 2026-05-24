"""
Compilatore file SDA - versione Streamlit
Avvio: streamlit run compilatore_sda_streamlit.py
"""

import streamlit as st
from datetime import datetime
import io
import csv


# ─────────────────────────────────────────────
#  LOGICA CORE
# ─────────────────────────────────────────────

def oggi():
    return datetime.now().strftime("%d%m%Y")

def ora_now():
    return datetime.now().strftime("%H:%M")

def build_riga13(op, giro, ldv, data, orario):
    prefix = " " + op[:3] + giro.upper().ljust(4)[:4]
    body = prefix + ldv.ljust(13)[:13] + "001"
    return body.ljust(46) + data + orario

def build_riga22(op, giro, ldv, data, orario):
    prefix = " " + op[:3] + giro.upper().ljust(4)[:4]
    parte1 = ldv[:9]
    parte2 = ldv[9:22]
    has_x = len(ldv) > 2 and ldv[1] == "U" and ldv[2] == "W"
    x_str = "X" if has_x else " "
    line_a = (prefix + parte1 + "    " + "001" + parte2).ljust(46) + data + orario
    line_b = (prefix + parte2 + "001" + "             " + x_str).ljust(46) + data + orario
    return line_a + "\n" + line_b

def genera_file(op, data, orario, filiale, progressivo, ldv_list):
    errors = []
    righe = [
        "1INIZIOFILE  25.07.05-SNAPSHOT                      AP"
        + filiale + progressivo
    ]
    for item in ldv_list:
        ldv = item["ldv"].strip().upper()
        giro = item["giro"].strip().upper().ljust(4)[:4]
        if not ldv:
            continue
        if len(ldv) == 13:
            righe.append(build_riga13(op, giro, ldv, data, orario))
        elif len(ldv) == 22:
            righe.append(build_riga22(op, giro, ldv, data, orario))
        else:
            errors.append(f"LDV '{ldv}': {len(ldv)} caratteri (attesi 13 o 22)")
    righe.append(" FINE FILE")
    return "\n".join(righe), errors

def nome_file(data, orario, op):
    ora = orario.replace(":", "")
    return f"SDA_{data}{ora}00_{op}.txt"


def parse_testo_lista(testo, giro_default):
    """Una LDV per riga. Separatori supportati: TAB, ; ,"""
    risultati = []
    for line in testo.splitlines():
        line = line.strip()
        if not line:
            continue
        for sep in ["\t", ";", ","]:
            if sep in line:
                parts = line.split(sep, 1)
                ldv = parts[0].strip()
                giro = parts[1].strip() if len(parts) > 1 and parts[1].strip() else giro_default
                risultati.append({"ldv": ldv, "giro": giro})
                break
        else:
            risultati.append({"ldv": line, "giro": giro_default})
    return risultati


def parse_csv_upload(file_bytes, giro_default):
    """
    Legge CSV o Excel (se openpyxl installato).
    Colonne cercate: LDV / LETTERA / BARCODE / CODICE  +  GIRO (opzionale).
    """
    text = file_bytes.decode("utf-8-sig", errors="replace")
    sep = ";"
    first = text.splitlines()[0] if text.splitlines() else ""
    if first.count(",") > first.count(";"):
        sep = ","
    reader = csv.DictReader(io.StringIO(text), delimiter=sep)
    risultati = []
    ldv_col = giro_col = None
    for row in reader:
        if ldv_col is None:
            keys_upper = {k.strip().upper(): k for k in row.keys()}
            for c in ["LDV", "LETTERA", "LETTERA DI VETTURA", "BARCODE", "CODICE"]:
                if c in keys_upper:
                    ldv_col = keys_upper[c]; break
            for c in ["GIRO", "GIRO PARALLELO", "PERCORSO"]:
                if c in keys_upper:
                    giro_col = keys_upper[c]; break
            if ldv_col is None:
                cols = list(row.keys())
                ldv_col = cols[0]
                if len(cols) > 1:
                    giro_col = cols[1]
        ldv = row.get(ldv_col, "").strip()
        if not ldv:
            continue
        giro = row.get(giro_col, "").strip() if giro_col else ""
        risultati.append({"ldv": ldv, "giro": giro or giro_default})
    return risultati


def parse_xlsx_upload(file_bytes, giro_default):
    import openpyxl
    wb = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=True)
    ws = wb.active
    rows = list(ws.iter_rows(values_only=True))
    if not rows:
        return []
    header = [str(c).strip().upper() if c else "" for c in rows[0]]
    ldv_idx = giro_idx = None
    for c in ["LDV", "LETTERA", "LETTERA DI VETTURA", "BARCODE", "CODICE"]:
        if c in header:
            ldv_idx = header.index(c); break
    for c in ["GIRO", "GIRO PARALLELO", "PERCORSO"]:
        if c in header:
            giro_idx = header.index(c); break
    if ldv_idx is None:
        ldv_idx = 0
        if len(header) > 1:
            giro_idx = 1
    risultati = []
    for row in rows[1:]:
        ldv = str(row[ldv_idx]).strip() if row[ldv_idx] else ""
        if not ldv or ldv == "None":
            continue
        giro = str(row[giro_idx]).strip() if giro_idx is not None and row[giro_idx] else ""
        risultati.append({"ldv": ldv, "giro": giro or giro_default})
    return risultati


# ─────────────────────────────────────────────
#  STREAMLIT UI
# ─────────────────────────────────────────────

st.set_page_config(page_title="Compilatore SDA", page_icon="📦", layout="wide")
st.title("📦 Compilatore file SDA")
st.caption("Genera file .txt compatibile con lettore TC77")

# ── Intestazione ──────────────────────────────
with st.expander("⚙️ Intestazione file", expanded=True):
    c1, c2, c3, c4, c5 = st.columns([1, 1.4, 1, 1.2, 1.4])
    operatore  = c1.text_input("Operatore",        value="059",      max_chars=3)
    data       = c2.text_input("Data (DDMMYYYY)",  value=oggi(),     max_chars=8)
    orario     = c3.text_input("Orario (HH:MM)",   value=ora_now(),  max_chars=5)
    filiale    = c4.text_input("Filiale AP",        value="230520",   max_chars=6)
    progressivo= c5.text_input("Progressivo",       value="69305811", max_chars=8)

# ── Sessione LDV ─────────────────────────────
if "ldv_rows" not in st.session_state:
    st.session_state.ldv_rows = [
        {"ldv": "281412J080140", "giro": "G101"},
        {"ldv": "281412J079774", "giro": "G131"},
    ]

st.subheader("Lettere di vettura")

# ── Tabs: manuale / import lista ─────────────
tab_manual, tab_testo, tab_file = st.tabs(
    ["✏️ Inserimento manuale", "📋 Incolla lista testo", "📂 Importa CSV / Excel"]
)

# ── Tab 1: manuale ───────────────────────────
with tab_manual:
    col_h1, col_h2, col_h3 = st.columns([5, 1.5, 0.5])
    col_h1.markdown("<small style='color:gray'>LDV (13 o 22 caratteri)</small>", unsafe_allow_html=True)
    col_h2.markdown("<small style='color:gray'>Giro (G101, RUS, 021…)</small>", unsafe_allow_html=True)

    to_remove = None
    for i, row in enumerate(st.session_state.ldv_rows):
        c1, c2, c3 = st.columns([5, 1.5, 0.5])
        st.session_state.ldv_rows[i]["ldv"] = c1.text_input(
            f"ldv_{i}", value=row["ldv"], key=f"ldv_{i}",
            label_visibility="collapsed", max_chars=22, placeholder="LDV 13 o 22 caratteri"
        )
        st.session_state.ldv_rows[i]["giro"] = c2.text_input(
            f"giro_{i}", value=row["giro"], key=f"giro_{i}",
            label_visibility="collapsed", max_chars=4, placeholder="G101/RUS/021"
        )
        if c3.button("✕", key=f"rm_{i}"):
            to_remove = i

    if to_remove is not None:
        st.session_state.ldv_rows.pop(to_remove)
        st.rerun()

    col_a, col_b = st.columns([1, 5])
    if col_a.button("➕ Aggiungi riga"):
        st.session_state.ldv_rows.append({"ldv": "", "giro": "G101"})
        st.rerun()
    if col_b.button("🗑 Svuota lista", type="secondary"):
        st.session_state.ldv_rows = []
        st.rerun()

# ── Tab 2: incolla testo ─────────────────────
with tab_testo:
    st.markdown(
        "Incolla una LDV per riga. Per specificare il giro usa uno di questi separatori: "
        "`LDV TAB GIRO` · `LDV;GIRO` · `LDV,GIRO`  \n"
        "Se non specifichi il giro, verrà usato il **giro di default** qui sotto."
    )
    giro_def_t = st.text_input("Giro di default (es. G101, RUS, 021)", value="G101", max_chars=4, key="giro_def_testo")
    testo_lista = st.text_area(
        "Lista LDV",
        height=200,
        placeholder="281412J080140\n281412J079774\t G122\n3UW1TZY000016;G112",
        key="testo_lista"
    )
    if st.button("📥 Aggiungi alla lista", key="btn_import_testo"):
        if testo_lista.strip():
            nuove = parse_testo_lista(testo_lista, giro_def_t.strip().upper() or "G101")
            st.session_state.ldv_rows.extend(nuove)
            st.success(f"✓ Aggiunte {len(nuove)} LDV alla lista")
            st.rerun()
        else:
            st.warning("Nessun testo inserito.")

# ── Tab 3: file CSV/Excel ─────────────────────
with tab_file:
    st.markdown(
        "Carica un file **CSV** o **Excel (.xlsx)**. "
        "Il file deve avere una colonna `LDV` (o `BARCODE`, `CODICE`) "
        "e opzionalmente una colonna `GIRO`."
    )
    giro_def_f = st.text_input("Giro di default (es. G101, RUS, 021)", value="G101", max_chars=4, key="giro_def_file")
    uploaded = st.file_uploader("Scegli file", type=["csv", "txt", "xlsx", "xls"], key="uploader")
    if uploaded:
        ext = uploaded.name.rsplit(".", 1)[-1].lower()
        file_bytes = uploaded.read()
        try:
            if ext in ("xlsx", "xls"):
                try:
                    nuove = parse_xlsx_upload(file_bytes, giro_def_f.strip().upper() or "G101")
                except ImportError:
                    st.error("Per leggere .xlsx installa openpyxl:  `pip install openpyxl`")
                    nuove = []
            else:
                nuove = parse_csv_upload(file_bytes, giro_def_f.strip().upper() or "G101")

            if nuove:
                st.success(f"✓ Trovate {len(nuove)} LDV nel file — premi il pulsante per aggiungerle")
                st.dataframe(
                    [{"LDV": r["ldv"], "Giro": r["giro"]} for r in nuove],
                    use_container_width=True, height=180
                )
                if st.button("📥 Aggiungi alla lista", key="btn_import_file"):
                    st.session_state.ldv_rows.extend(nuove)
                    st.success(f"✓ Aggiunte {len(nuove)} LDV alla lista")
                    st.rerun()
            else:
                st.warning("Nessuna LDV trovata nel file.")
        except Exception as e:
            st.error(f"Errore lettura file: {e}")

# ── Riepilogo e generazione ───────────────────
st.divider()
n_ldv = sum(1 for r in st.session_state.ldv_rows if r["ldv"].strip())
st.markdown(f"**Lista attuale: {n_ldv} LDV**")

op = operatore.strip().zfill(3)[-3:]

if st.button("⚙️ Genera file SDA", type="primary"):
    testo, errors = genera_file(
        op, data.strip(), orario.strip(),
        filiale.strip(), progressivo.strip(),
        st.session_state.ldv_rows
    )
    st.session_state["output"] = testo
    st.session_state["errors"] = errors
    st.session_state["filename"] = nome_file(data.strip(), orario.strip(), op)

if "output" in st.session_state:
    for e in st.session_state.get("errors", []):
        st.error(f"⚠️ {e}")
    if not st.session_state.get("errors"):
        st.success(f"✓ File generato — {n_ldv} LDV")

    st.subheader("Anteprima")
    st.code(st.session_state["output"], language=None)
    st.download_button(
        label="💾 Scarica .txt",
        data=st.session_state["output"].encode("utf-8"),
        file_name=st.session_state.get("filename", "output.txt"),
        mime="text/plain"
    )

st.divider()
st.caption("LDV 13 car. → 1 riga · LDV 22 car. → 2 righe · X automatica se UW in pos 2-3 · Giro: qualsiasi codice fino a 4 caratteri (G101, RUS, 021, 029…)")
