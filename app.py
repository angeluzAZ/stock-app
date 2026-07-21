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
import hmac
import hashlib
from datetime import datetime, timezone, timedelta
import pandas as pd
import streamlit as st
import gspread
from gspread_dataframe import set_with_dataframe, get_as_dataframe
from streamlit_autorefresh import st_autorefresh
from streamlit_cookies_controller import CookieController

_TZ_ARG = timezone(timedelta(hours=-3))   # hora de Argentina
_TIMEOUT_SEG = 5 * 60                       # cierre por inactividad: 5 minutos

st.set_page_config(page_title="Stock HZ + AZ", page_icon="📦", layout="centered")
AQUI = os.path.dirname(os.path.abspath(__file__))
CRED = os.path.join(AQUI, "credenciales.json")
URL_FILE = os.path.join(AQUI, "planilla.txt")

DEPOSITOS = {"1": "HERAS", "9": "ALTO", "15": "PERICO", "7": "Pulmón rollos",
             "8": "Central", "2": "URDI", "5": "NECO"}
# Qué depósitos se muestran de cada empresa (los demás se ocultan).
DEP_POR_EMPRESA = {"HZ": {"1", "9", "15", "8", "7"}, "AZ": {"2", "5"}}
def _dep(cod):
    n = DEPOSITOS.get(str(cod))
    return f"{cod} · {n}" if n else str(cod)

COLS = ["codigo", "descripcion", "adicional", "deposito", "empresa", "stock",
        "cod_prov", "proveedor"]
# Columnas que baja la app OFFLINE (sin proveedor, para no exponerlo).
COLS_OFF = ["codigo", "descripcion", "adicional", "deposito", "empresa", "stock"]


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
    con = comb[comb["stock"] != 0].copy()
    _escribir_hoja(sh, "CONSULTA", con)
    # CONSULTA_OFF: lo que baja la app OFFLINE del celular. SIN proveedor (privado)
    # y solo los depósitos que se muestran de cada empresa (HZ:1,9,15,8,7 · AZ:2,5).
    off = con[con.apply(
        lambda r: str(r["deposito"]).strip() in DEP_POR_EMPRESA.get(r["empresa"], set()),
        axis=1)]
    _escribir_hoja(sh, "CONSULTA_OFF", off[COLS_OFF].copy())


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
# Tamaño de letra (accesibilidad): multiplicador --tam aplicado a las tarjetas.
_TAMS = {"A": 1.0, "A+": 1.25, "A++": 1.55}
_tam = _TAMS.get(st.session_state.get("_tam_sel", "A"), 1.0)

st.markdown(f"""
<style>
:root{{
  --tam:{_tam};
  --bg:#eef1f8; --surface:#ffffff; --surface-2:#f6f8fc;
  --ink:#0f172a; --ink-soft:#586074; --line:#e2e7f0;
  --acc:#2f6bff; --acc-2:#5b8cff; --acc-soft:#e8f0ff;
  --verde:#067a4e; --verde-bg:#dcf5e9; --verde-bar:#18b877;
  --rojo:#c23b22; --rojo-bg:#fbe6e1;
  --sombra:0 1px 2px rgba(15,23,42,.06), 0 14px 34px -22px rgba(15,23,42,.28);
}}
@media (prefers-color-scheme:dark){{:root{{
  --bg:#090d18; --surface:#141a2b; --surface-2:#1b2236;
  --ink:#eef2fb; --ink-soft:#9aa6c2; --line:#28324a;
  --acc:#5b8cff; --acc-2:#7aa2ff; --acc-soft:#1a2540;
  --verde:#38e08e; --verde-bg:#123024; --verde-bar:#34d399;
  --rojo:#ff7a63; --rojo-bg:#3a1a14;
  --sombra:0 1px 2px rgba(0,0,0,.4), 0 18px 40px -22px rgba(0,0,0,.7);
}}}}
#MainMenu, footer {{visibility:hidden;}}
.stApp{{background:var(--bg);}}
.block-container {{padding-top:1.1rem; max-width:640px;}}

/* buscador grande y cómodo */
div[data-testid="stTextInput"] input{{
  font-size:calc(20px*var(--tam)) !important; font-weight:600 !important;
  padding:16px 18px !important; height:auto !important;
  border-radius:16px !important; box-shadow:var(--sombra);
}}
div[data-testid="stTextInput"] label p{{font-size:calc(15px*var(--tam)) !important; font-weight:700;}}

/* tarjeta de producto (diseño moderno) */
.card{{background:var(--surface);border:1px solid var(--line);border-radius:20px;
  padding:calc(15px*var(--tam)) 17px;margin-bottom:14px;box-shadow:var(--sombra);}}
.card .head{{display:flex;align-items:center;justify-content:space-between;gap:10px;}}
.card .cod{{font-family:ui-monospace,Consolas,monospace;font-weight:800;
  font-size:calc(14px*var(--tam));letter-spacing:1.2px;color:var(--ink-soft);}}
.card .badge{{display:flex;align-items:baseline;gap:5px;padding:7px 13px;border-radius:14px;font-weight:900;white-space:nowrap;}}
.card .badge .n{{font-size:calc(23px*var(--tam));}}
.card .badge .u{{font-size:calc(11px*var(--tam));font-weight:800;text-transform:uppercase;letter-spacing:.05em;}}
.card .badge.hay{{background:var(--verde-bg);color:var(--verde);}}
.card .badge.no{{background:var(--rojo-bg);color:var(--rojo);}}
.card .desc{{font-size:calc(19px*var(--tam));font-weight:700;color:var(--ink);margin:9px 0 3px;line-height:1.25;}}
.card .prov{{font-size:calc(14px*var(--tam));color:var(--ink-soft);font-weight:600;
  margin-bottom:11px;display:flex;align-items:center;gap:7px;}}
.card .prov .dot{{width:8px;height:8px;border-radius:50%;background:var(--acc);flex:none;}}
.card .deps{{display:flex;flex-direction:column;gap:8px;}}
.card .dep{{display:grid;grid-template-columns:auto 1fr auto;align-items:center;gap:11px;
  padding:calc(10px*var(--tam)) 13px;border-radius:12px;background:var(--surface-2);}}
.card .dep .mk{{font-size:calc(18px*var(--tam));width:calc(22px*var(--tam));text-align:center;}}
.card .dep.si .mk{{color:var(--verde);}} .card .dep.no .mk{{color:var(--ink-soft);opacity:.6;}}
.card .dep .lugar{{font-weight:700;font-size:calc(16px*var(--tam));color:var(--ink);}}
.card .dep .right{{display:flex;align-items:center;gap:11px;justify-content:flex-end;}}
.card .bar{{width:calc(70px*var(--tam));height:9px;border-radius:6px;background:var(--line);overflow:hidden;}}
.card .bar > i{{display:block;height:100%;border-radius:6px;background:var(--verde-bar);}}
.card .dep .val{{font-family:ui-monospace,Consolas,monospace;font-weight:900;
  font-size:calc(19px*var(--tam));min-width:calc(42px*var(--tam));text-align:right;}}
.card .dep.si .val{{color:var(--verde);}} .card .dep.no .val{{color:var(--ink-soft);}}
</style>
""", unsafe_allow_html=True)
st.title("📦 Stock HZ + AZ")

# ─────────────────────────── LOGIN ─────────────────────────────────
def _usuarios():
    try:
        return {str(k).lower(): str(v) for k, v in dict(st.secrets.get("usuarios", {})).items()}
    except Exception:
        return {}

def _cookie_secret():
    try:
        return str(st.secrets["gcp_service_account"]["private_key_id"]) or "x"
    except Exception:
        return "local-dev"

def _firma(user):
    return hmac.new(_cookie_secret().encode(), user.encode(), hashlib.sha256).hexdigest()[:20]

def _token(user):
    return f"{user}|{_firma(user)}"

def _valida_token(tok):
    try:
        user, sig = str(tok).split("|", 1)
        if hmac.compare_digest(sig, _firma(user)):
            return user
    except Exception:
        pass
    return None

_USERS = _usuarios()
if _USERS:  # si hay usuarios configurados (en la nube), se pide login
    cookies = CookieController()
    # ── recordar sesión: si al recargar no hay sesión pero hay cookie válida,
    # se restaura (queda logueado aunque refresques la página) ──
    if not st.session_state.get("auth_ok"):
        _u_ck = _valida_token(cookies.get("stk")) if cookies.get("stk") else None
        if _u_ck and _u_ck.lower() in _USERS:
            st.session_state["auth_ok"] = True
            st.session_state["user"] = _u_ck.capitalize()
            st.session_state["_ultima_act"] = time.time()

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
                cookies.set("stk", _token(u.strip().lower()),
                            expires=datetime.now(_TZ_ARG) + timedelta(hours=8))
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
            cookies.remove("stk")
            st.session_state.clear()
            st.session_state["_msg_timeout"] = True
            st.rerun()
        cabe = st.columns([4, 1])
        cabe[0].caption(f"👤 {st.session_state.get('user', '')}")
        if cabe[1].button("Salir", use_container_width=True):
            cookies.remove("stk")
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

cm1, cm2 = st.columns([3, 2])
with cm2:
    st.radio(
        "Tamaño de letra", list(_TAMS.keys()), key="_tam_sel",
        horizontal=True, label_visibility="collapsed",
        help="Agrandá el texto de las tarjetas (A+ / A++).")
with cm1:
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

    # solo los depósitos que corresponden a cada empresa (HZ: 1,9,15,8,7 · AZ: 2,5)
    stock = stock[stock.apply(
        lambda r: str(r["deposito"]).strip() in DEP_POR_EMPRESA.get(r["empresa"], set()),
        axis=1)]

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

    # una tarjeta por código (aunque esté en las dos empresas).
    # Orden: primero los códigos que EMPIEZAN con lo buscado; luego, más stock primero.
    grupos = list(d.groupby("codigo", sort=False))
    grupos.sort(key=lambda kv: (not kv[0].lower().startswith(t), -kv[1]["stock"].sum()))
    st.caption(f"{len(grupos):,} producto(s)" + (" · mostrando los primeros 150" if len(grupos) > 150 else ""))

    cards = []
    for cod, g in grupos[:150]:
        g0 = g.iloc[0]
        emps = sorted(g["empresa"].unique())
        emp = " / ".join(emps)
        tot = g["stock"].sum()
        badge_cls = "hay" if tot > 0 else "no"
        # stock que ya existe por depósito (los que están en 0 no vienen en la planilla)
        present = {str(r["deposito"]).strip(): float(r["stock"]) for _, r in g.iterrows()}
        # armar TODOS los depósitos permitidos de la(s) empresa(s), con 0 precargado
        expected = []
        for e in emps:
            for dcod in sorted(DEP_POR_EMPRESA.get(e, set()), key=lambda x: int(x)):
                if dep_sel != "Todos" and dcod != dep_sel:
                    continue
                expected.append((dcod, present.get(dcod, 0.0)))
        expected.sort(key=lambda x: x[1], reverse=True)     # más stock primero
        mx = max(1.0, max((s for _, s in expected), default=0.0))
        filas = ""
        for dcod, stk in expected:
            si = stk > 0
            cls = "si" if si else "no"
            mk = "✔" if si else "—"
            w = int(round(stk / mx * 100)) if si else 0
            filas += (f"<div class='dep {cls}'><span class='mk'>{mk}</span>"
                      f"<span class='lugar'>{_depname(dcod)}</span>"
                      f"<span class='right'><span class='bar'><i style='width:{w}%'></i></span>"
                      f"<span class='val'>{_num(stk)}</span></span></div>")
        adic = f" · {g0['adicional']}" if g0["adicional"] else ""
        cards.append(
            f"<div class='card'><div class='head'><span class='cod'>{cod}</span>"
            f"<span class='badge {badge_cls}'><span class='n'>{_num(tot)}</span><span class='u'>u</span></span></div>"
            f"<div class='desc'>{g0['descripcion']}{adic}</div>"
            f"<div class='prov'><span class='dot'></span>{g0['proveedor']} · {emp}</div>"
            f"<div class='deps'>{filas}</div></div>")
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
