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
import time
from datetime import datetime, timezone, timedelta
import pandas as pd
import streamlit as st
import gspread
from gspread_dataframe import set_with_dataframe, get_as_dataframe
from streamlit_autorefresh import st_autorefresh

_TZ_ARG = timezone(timedelta(hours=-3))   # hora de Argentina
_TIMEOUT_SEG = 5 * 60                       # cierre por inactividad: 5 minutos

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


def _rebuild_consulta(sh):
    """Reconstruye la pestaña CONSULTA (solo stock != 0, ambas empresas) leyendo
    HZ y AZ de la planilla. Es lo que lee la app en el celular (rápido, ~25k)."""
    partes = []
    for hoja in ("HZ", "AZ"):
        try:
            d = get_as_dataframe(sh.worksheet(hoja), header=0).dropna(how="all")
            d["stock"] = pd.to_numeric(d.get("stock", 0), errors="coerce").fillna(0)
            partes.append(d[[c for c in COLS if c in d.columns]])
        except Exception:
            pass
    comb = pd.concat(partes, ignore_index=True) if partes else pd.DataFrame(columns=COLS)
    _escribir_hoja(sh, "CONSULTA", comb[comb["stock"] != 0].copy())


def guardar_empresas(dfs_por_empresa, usuario):
    """dfs_por_empresa: dict {'HZ': df, 'AZ': df} (una o las dos). Escribe cada
    pestaña, reconstruye CONSULTA UNA sola vez y registra fecha/hora + usuario."""
    sh = _abrir_planilla()
    for emp, df in dfs_por_empresa.items():
        _escribir_hoja(sh, emp, df)
    _rebuild_consulta(sh)
    detalle = " + ".join(f"{e} ({len(d):,} filas)" for e, d in dfs_por_empresa.items())
    return _escribir_meta(sh, usuario, detalle)


def _escribir_meta(sh, usuario, detalle):
    ahora = datetime.now(_TZ_ARG).strftime("%d/%m/%Y %H:%M")
    try:
        ws = sh.worksheet("META")
    except gspread.WorksheetNotFound:
        ws = sh.add_worksheet(title="META", rows=5, cols=3)
    ws.clear()
    ws.update([["fecha", "usuario", "detalle"], [ahora, usuario, detalle]],
              value_input_option="RAW")
    return ahora


@st.cache_data(ttl=60, show_spinner=False)
def leer_meta():
    """Última actualización: (fecha, usuario, detalle) o None."""
    try:
        sh = _abrir_planilla()
        vals = sh.worksheet("META").get_all_values()
        if len(vals) >= 2 and any(vals[1]):
            f, u, d = (vals[1] + ["", "", ""])[:3]
            return f, u, d
    except Exception:
        pass
    return None


def _leer_stock_excel(archivo, emp_cod):
    """Lee un Excel de STOCK POR DEPÓSITO y devuelve el df normalizado con
    proveedor. Lanza ValueError con mensaje claro si el archivo no sirve."""
    col_pref = "Saldo control stock" if emp_cod == "HZ" else "Saldo stock"
    try:
        raw = pd.read_excel(archivo, sheet_name="Datos")
    except Exception:
        raw = pd.read_excel(archivo)
    if "Cód. Artículo" not in raw.columns or "Cód. Depósito" not in raw.columns:
        raise ValueError(f"{emp_cod}: al Excel le faltan columnas. ¿Es el export de STOCK POR DEPÓSITO?")
    col_saldo = col_pref if col_pref in raw.columns else next(
        (c for c in raw.columns if str(c).lower().startswith("saldo")), None)
    if not col_saldo:
        raise ValueError(f"{emp_cod}: no encontré la columna de saldo/stock.")
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
    return nuevo[COLS]


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
        if st.session_state.pop("_msg_timeout", False):
            st.info("Tu sesión se cerró por inactividad (5 minutos). Volvé a entrar.")
        st.subheader("🔒 Iniciar sesión")
        with st.form("login"):
            u = st.text_input("Usuario")
            p = st.text_input("Contraseña", type="password")
            entrar = st.form_submit_button("Entrar", type="primary")
        if entrar:
            if _USERS.get(u.strip().lower()) == p:
                st.session_state["auth_ok"] = True
                st.session_state["user"] = u.strip().lower().capitalize()
                st.session_state["_ultima_act"] = time.time()
                st.rerun()
            else:
                st.error("Usuario o contraseña incorrectos.")
        st.stop()
    else:
        # ── temporizador de inactividad: se cierra sola tras 5 min sin usarla,
        # pero cada interacción (buscar, actualizar, etc.) reinicia el contador ──
        tick = st_autorefresh(interval=60_000, key="idle_tick")  # revisa cada 1 min
        ahora = time.time()
        if st.session_state.get("_tick") == tick:      # el rerun lo causó el usuario
            st.session_state["_ultima_act"] = ahora     #   → hubo actividad
        st.session_state["_tick"] = tick
        if ahora - st.session_state.get("_ultima_act", ahora) > _TIMEOUT_SEG:
            st.session_state.clear()
            st.session_state["_msg_timeout"] = True
            st.rerun()
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

_meta = leer_meta()
if _meta:
    st.caption(f"🕒 Última actualización: **{_meta[0]}** · por **{_meta[1].capitalize()}** ({_meta[2]})")
else:
    st.caption("🕒 Sin actualizaciones registradas todavía.")

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
    st.write("Subí los **Excel de stock** de Tango. Podés subir **los dos a la vez** "
             "(Horizonte y Años Luz) o solo uno. El último que sube, queda.")
    c1, c2 = st.columns(2)
    arch_hz = c1.file_uploader("📄 Horizonte (HZ)", type=["xlsx"], key="up_hz")
    arch_az = c2.file_uploader("📄 Años Luz (AZ)", type=["xlsx"], key="up_az")

    # leer y previsualizar lo que se haya subido
    a_guardar, errores = {}, []
    for arch, emp_cod in [(arch_hz, "HZ"), (arch_az, "AZ")]:
        if arch is None:
            continue
        try:
            df = _leer_stock_excel(arch, emp_cod)
            a_guardar[emp_cod] = df
            st.success(f"{emp_cod}: {len(df):,} filas · {df['stock'].ne(0).sum():,} con stock.")
        except Exception as e:
            errores.append(str(e))
    for e in errores:
        st.error(e)

    puede = bool(a_guardar) and not errores
    if st.button("✅ Guardar en la nube", type="primary", disabled=not puede):
        usuario = st.session_state.get("user", "—")
        with st.spinner("Subiendo a Google Sheets… puede tardar (son muchas filas)."):
            ts = guardar_empresas(a_guardar, usuario)
        cargar_consulta.clear(); cargar_todo.clear(); leer_meta.clear()
        st.success(f"¡Listo! Actualizado: {', '.join(a_guardar)} · {ts} · por {usuario}. "
                   "Todos lo ven ya.")
        st.balloons()
