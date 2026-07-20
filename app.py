# -*- coding: utf-8 -*-
"""
App de consulta y actualización de stock — Horizonte SRL + Años Luz SRL.
Corre LOCAL en cada PC, pero la base es una planilla de Google Sheets compartida
(dos pestañas: HZ y AZ), así todos ven el mismo stock.

Necesita 2 archivos en esta carpeta:
  - credenciales.json  → la llave de la cuenta de servicio de Google (NO se comparte).
  - planilla.txt       → una línea con la URL de la planilla de Google Sheets.

- Buscar (todos): por código, descripción, adicional o proveedor + filtro de depósito.
- Actualizar: subir el Excel de stock; reemplaza esa pestaña (HZ o AZ) en la nube.
"""
import os
import pandas as pd
import streamlit as st
import gspread
from gspread_dataframe import set_with_dataframe, get_as_dataframe

st.set_page_config(page_title="Stock HZ + AZ", page_icon="📦", layout="centered")
AQUI = os.path.dirname(os.path.abspath(__file__))
CRED = os.path.join(AQUI, "credenciales.json")
URL_FILE = os.path.join(AQUI, "planilla.txt")

DEPOSITOS = {"1": "HERAS", "9": "ALTO", "15": "PERICO", "7": "Pulmón rollos",
             "8": "Central", "2": "URDI", "5": "NECO"}
def _dep(cod):
    n = DEPOSITOS.get(str(cod))
    return f"{cod} · {n}" if n else str(cod)

COLS = ["codigo", "descripcion", "adicional", "deposito", "empresa", "stock",
        "cod_prov", "proveedor"]


def _hay_secrets():
    """En la NUBE (Streamlit Cloud) las credenciales vienen de st.secrets;
    en LOCAL vienen de los archivos credenciales.json + planilla.txt."""
    try:
        return "gcp_service_account" in st.secrets
    except Exception:
        return False


def _faltan_archivos():
    if _hay_secrets():
        return []   # en la nube: nada que falte, viene de secrets
    faltan = []
    if not os.path.exists(CRED):
        faltan.append("**credenciales.json** (la llave de Google)")
    if not os.path.exists(URL_FILE):
        faltan.append("**planilla.txt** (con la URL de tu Google Sheet)")
    return faltan


@st.cache_resource(show_spinner=False)
def _abrir_planilla():
    if _hay_secrets():
        gc = gspread.service_account_from_dict(dict(st.secrets["gcp_service_account"]))
        url = st.secrets["planilla_url"]
    else:
        gc = gspread.service_account(filename=CRED)
        url = open(URL_FILE, encoding="utf-8").read().strip()
    return gc.open_by_url(url)


def _limpiar(df):
    for c in COLS:
        if c not in df.columns:
            df[c] = ""
    df["codigo"] = df["codigo"].astype(str).str.replace(r"\.0$", "", regex=True).str.strip()
    df["stock"] = pd.to_numeric(df["stock"], errors="coerce").fillna(0)
    for c in ("descripcion", "adicional", "proveedor", "deposito", "empresa", "cod_prov"):
        df[c] = df[c].fillna("").astype(str)
    return df[COLS]


@st.cache_data(ttl=300, show_spinner="Leyendo stock…")
def cargar_consulta():
    """Solo lo que tiene stock (~25k) → rápido. Es el modo por defecto."""
    sh = _abrir_planilla()
    try:
        df = get_as_dataframe(sh.worksheet("CONSULTA"), header=0).dropna(how="all")
    except Exception:
        return pd.DataFrame(columns=COLS)
    return _limpiar(df)


@st.cache_data(ttl=600, show_spinner="Leyendo TODO el stock (incluye los que están en 0)…")
def cargar_todo():
    """TODOS los artículos y depósitos, incluidos los que están en 0 (~136k).
    Más lento la primera vez, pero permite ver el stock en cero."""
    sh = _abrir_planilla()
    partes = []
    for hoja in ("HZ", "AZ"):
        try:
            d = get_as_dataframe(sh.worksheet(hoja), header=0).dropna(how="all")
            if not d.empty:
                partes.append(d)
        except Exception:
            pass
    if not partes:
        return pd.DataFrame(columns=COLS)
    return _limpiar(pd.concat(partes, ignore_index=True))


@st.cache_data(ttl=3600, show_spinner=False)
def mapa_proveedores():
    """Mapa artículo→proveedor. En la nube vive en la pestaña PROVEEDORES de la
    planilla (privada); en local, si existe, usa proveedores.csv."""
    p = os.path.join(AQUI, "proveedores.csv")
    if os.path.exists(p):
        m = pd.read_csv(p, dtype=str).fillna("")
        m["codigo"] = m["codigo"].str.strip()
        return m
    try:
        sh = _abrir_planilla()
        m = get_as_dataframe(sh.worksheet("PROVEEDORES"), header=0).dropna(how="all")
        for c in ("codigo", "cod_prov", "proveedor"):
            m[c] = m[c].fillna("").astype(str).str.strip() if c in m.columns else ""
        return m[["codigo", "cod_prov", "proveedor"]]
    except Exception:
        return pd.DataFrame(columns=["codigo", "cod_prov", "proveedor"])


def _escribir_hoja(sh, nombre, df):
    try:
        ws = sh.worksheet(nombre)
    except gspread.WorksheetNotFound:
        ws = sh.add_worksheet(title=nombre, rows=len(df) + 10, cols=len(COLS))
    ws.clear()
    # RAW = guardar tal cual (si no, Google Sheets 'interpreta' los códigos como
    # números y les come los ceros de adelante: 0010100400 → 10100400).
    dfx = df.copy()
    for c in ("codigo", "deposito", "cod_prov"):
        if c in dfx.columns:
            dfx[c] = dfx[c].astype(str)
    dfx = dfx.fillna("")
    valores = [list(dfx.columns)] + dfx.astype(object).values.tolist()
    ws.resize(rows=len(valores) + 5, cols=len(dfx.columns))
    ws.update(valores, value_input_option="RAW")


def guardar_en_sheets(empresa, df):
    """Escribe la pestaña de la empresa (HZ/AZ) con TODO su stock, y reconstruye
    la pestaña CONSULTA (solo stock != 0, ambas empresas) que lee la app Android
    → el celular queda siempre al día sin manejar 136k filas."""
    sh = _abrir_planilla()
    _escribir_hoja(sh, empresa, df)
    # la otra empresa: la leo de la planilla para combinar
    otra = "AZ" if empresa == "HZ" else "HZ"
    try:
        od = get_as_dataframe(sh.worksheet(otra), header=0).dropna(how="all")
        od["stock"] = pd.to_numeric(od.get("stock", 0), errors="coerce").fillna(0)
        od = od[[c for c in COLS if c in od.columns]]
    except Exception:
        od = pd.DataFrame(columns=COLS)
    combinado = pd.concat([df, od], ignore_index=True)
    consulta = combinado[combinado["stock"] != 0].copy()
    _escribir_hoja(sh, "CONSULTA", consulta)


# ─────────────────────────── APP ───────────────────────────────────
st.markdown("""
<style>
#MainMenu, footer {visibility:hidden;}
.block-container {padding-top:1.2rem; max-width:720px;}
.pc{background:#fff;border:1px solid #e2e8f0;border-radius:12px;padding:11px 14px;
    margin-bottom:10px;box-shadow:0 1px 3px rgba(13,27,42,.06);}
.pc-top{display:flex;justify-content:space-between;align-items:baseline;gap:8px}
.pc .cod{font-family:ui-monospace,Consolas,monospace;font-weight:700;font-size:16px;color:#0d1b2a}
.pc .tot{font-weight:800;font-size:14px;white-space:nowrap}
.pc .tot.hay{color:#14805c}.pc .tot.no{color:#94a3b8}
.pc .desc{font-size:14px;color:#0d1b2a;margin:1px 0;line-height:1.3}
.pc .prov{font-size:12px;color:#64748b;margin:2px 0 7px}
.pc .deps{display:flex;flex-wrap:wrap;gap:6px}
.pc .chip{font-size:12.5px;font-weight:600;border-radius:20px;padding:3px 11px;white-space:nowrap}
.pc .chip.ok{background:#dff3ea;color:#14805c}
.pc .chip.zero{background:#eef2f6;color:#94a3b8}
</style>
""", unsafe_allow_html=True)
st.title("📦 Stock HZ + AZ")

# ─────────────────────────── LOGIN ─────────────────────────────────
def _usuarios():
    try:
        return {str(k).lower(): str(v) for k, v in dict(st.secrets.get("usuarios", {})).items()}
    except Exception:
        return {}

_USERS = _usuarios()
if _USERS:  # si hay usuarios configurados (en la nube), se pide login
    if not st.session_state.get("auth_ok"):
        st.subheader("🔒 Iniciar sesión")
        with st.form("login"):
            u = st.text_input("Usuario")
            p = st.text_input("Contraseña", type="password")
            entrar = st.form_submit_button("Entrar", type="primary")
        if entrar:
            if _USERS.get(u.strip().lower()) == p:
                st.session_state["auth_ok"] = True
                st.session_state["user"] = u.strip().lower().capitalize()
                st.rerun()
            else:
                st.error("Usuario o contraseña incorrectos.")
        st.stop()
    else:
        cabe = st.columns([4, 1])
        cabe[0].caption(f"👤 {st.session_state.get('user', '')}")
        if cabe[1].button("Salir", use_container_width=True):
            st.session_state.clear()
            st.rerun()

faltan = _faltan_archivos()
if faltan:
    st.error("Faltan archivos de configuración en la carpeta de la app:")
    for f in faltan:
        st.markdown(f"- {f}")
    st.info("Ponelos en la misma carpeta que `app.py` y recargá la página.")
    st.stop()

modo = st.radio("¿Qué querés hacer?", ["🔍 Buscar", "⬆️ Actualizar (subir Excel)"],
                horizontal=True, label_visibility="collapsed")

# ─────────────────────────── BUSCAR ────────────────────────────────
def _num(v):
    v = float(v)
    return f"{v:,.0f}" if v == int(v) else f"{v:,.1f}"

def _depname(cod):
    return DEPOSITOS.get(str(cod), f"Dep {cod}")

if modo == "🔍 Buscar":
    incluir_ceros = st.toggle("Incluir productos sin stock (más lento)", value=False,
                              help="Mostrá también los artículos y depósitos que están en 0.")
    try:
        stock = cargar_todo() if incluir_ceros else cargar_consulta()
    except Exception as e:
        st.error(f"No pude leer la planilla. Revisá las credenciales y que compartiste la "
                 f"planilla con el mail del robot.\n\n{e}")
        st.stop()
    if stock.empty:
        st.warning("La planilla está vacía. Andá a «Actualizar» y subí el Excel de stock.")
        st.stop()

    q = st.text_input("Buscar", placeholder="Código, descripción, adicional o proveedor…")
    deps = ["Todos"] + sorted(stock["deposito"].unique().tolist())
    dep_sel = st.selectbox("Depósito", deps,
                           format_func=lambda x: "Todos" if x == "Todos" else _dep(x))

    if not q.strip():
        st.info("Escribí un código, descripción, adicional o proveedor para buscar.")
        st.stop()
    t = q.strip().lower()
    m = (stock["codigo"].str.lower().str.contains(t, na=False) |
         stock["descripcion"].str.lower().str.contains(t, na=False) |
         stock["adicional"].str.lower().str.contains(t, na=False) |
         stock["proveedor"].str.lower().str.contains(t, na=False))
    d = stock[m]
    if dep_sel != "Todos":
        d = d[d["deposito"] == dep_sel]

    if d.empty:
        st.warning("Sin resultados. Si el producto puede estar en 0, activá arriba "
                   "«Incluir productos sin stock».")
        st.stop()

    # una tarjeta por artículo; los de más stock, primero
    grupos = list(d.groupby(["empresa", "codigo"], sort=False))
    grupos.sort(key=lambda kv: kv[1]["stock"].sum(), reverse=True)
    st.caption(f"{len(grupos):,} producto(s)" + (" · mostrando los primeros 150" if len(grupos) > 150 else ""))

    cards = []
    for (emp, cod), g in grupos[:150]:
        g0 = g.iloc[0]
        tot = g["stock"].sum()
        tot_cls = "hay" if tot > 0 else "no"
        chips = ""
        for _, r in g.sort_values("stock", ascending=False).iterrows():
            cls = "ok" if r["stock"] > 0 else "zero"
            chips += f"<span class='chip {cls}'>{_depname(r['deposito'])}: {_num(r['stock'])}</span>"
        adic = f" · {g0['adicional']}" if g0["adicional"] else ""
        cards.append(
            f"<div class='pc'><div class='pc-top'><span class='cod'>{cod}</span>"
            f"<span class='tot {tot_cls}'>{_num(tot)} u</span></div>"
            f"<div class='desc'>{g0['descripcion']}{adic}</div>"
            f"<div class='prov'>{g0['proveedor']} · {emp}</div>"
            f"<div class='deps'>{chips}</div></div>")
    st.markdown("".join(cards), unsafe_allow_html=True)

# ─────────────────────────── ACTUALIZAR ────────────────────────────
else:
    st.subheader("⬆️ Actualizar stock")
    st.write("Subí el **Excel de stock** exportado de Tango. Reemplaza los datos de esa "
             "empresa en la planilla (la otra empresa no se toca). El último que sube, queda.")
    emp = st.radio("¿De qué empresa es este archivo?", ["HZ (Horizonte)", "AZ (Años Luz)"],
                   horizontal=True)
    emp_cod = "HZ" if emp.startswith("HZ") else "AZ"
    col_saldo_pref = "Saldo control stock" if emp_cod == "HZ" else "Saldo stock"

    archivo = st.file_uploader("Excel de stock (.xlsx)", type=["xlsx"])
    if archivo is not None:
        try:
            raw = pd.read_excel(archivo, sheet_name="Datos")
        except Exception:
            raw = pd.read_excel(archivo)
        if "Cód. Artículo" not in raw.columns or "Cód. Depósito" not in raw.columns:
            st.error("El Excel no tiene las columnas esperadas. ¿Es el export de STOCK POR DEPÓSITO?")
            st.stop()
        col_saldo = col_saldo_pref if col_saldo_pref in raw.columns else next(
            (c for c in raw.columns if str(c).lower().startswith("saldo")), None)
        if not col_saldo:
            st.error("No encontré la columna de saldo/stock en el Excel.")
            st.stop()

        nuevo = pd.DataFrame({
            "codigo": raw["Cód. Artículo"].astype(str).str.strip(),
            "descripcion": raw.get("Descripción", "").fillna("").astype(str).str.strip(),
            "adicional": raw.get("Desc. Adicional", "").fillna("").astype(str).str.strip(),
            "deposito": raw["Cód. Depósito"].astype(str).str.strip(),
            "empresa": emp_cod,
            "stock": pd.to_numeric(raw[col_saldo], errors="coerce").fillna(0),
        })
        m = mapa_proveedores()
        nuevo = nuevo.merge(m, on="codigo", how="left")
        nuevo["cod_prov"] = nuevo["cod_prov"].fillna("")
        nuevo["proveedor"] = nuevo["proveedor"].fillna("Sin proveedor")
        nuevo = nuevo[COLS]

        st.success(f"Archivo leído: {len(nuevo):,} filas · {nuevo['stock'].ne(0).sum():,} con stock.")
        st.dataframe(nuevo.head(15), hide_index=True, use_container_width=True)

        if st.button(f"✅ Guardar en la nube (pestaña {emp_cod})", type="primary"):
            with st.spinner("Subiendo a Google Sheets… puede tardar (son muchas filas)."):
                guardar_en_sheets(emp_cod, nuevo)
            cargar_consulta.clear()
            cargar_todo.clear()
            st.success(f"¡Listo! Se actualizó el stock de {emp_cod} en la nube. Todos lo ven ya.")
            st.balloons()
