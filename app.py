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


@st.cache_data(ttl=120, show_spinner="Leyendo stock…")
def cargar_stock():
    """Lee la pestaña CONSULTA (solo lo que tiene stock, ~25k → rápido en el
    celular). El Streamlit la mantiene al día cada vez que alguien sube un Excel."""
    sh = _abrir_planilla()
    try:
        df = get_as_dataframe(sh.worksheet("CONSULTA"), header=0).dropna(how="all")
    except Exception:
        return pd.DataFrame(columns=COLS)
    for c in COLS:
        if c not in df.columns:
            df[c] = ""
    df["codigo"] = df["codigo"].astype(str).str.replace(r"\.0$", "", regex=True).str.strip()
    df["stock"] = pd.to_numeric(df["stock"], errors="coerce").fillna(0)
    for c in ("descripcion", "adicional", "proveedor", "deposito", "empresa", "cod_prov"):
        df[c] = df[c].fillna("").astype(str)
    return df[COLS]


@st.cache_data
def mapa_proveedores():
    p = os.path.join(AQUI, "proveedores.csv")
    if os.path.exists(p):
        m = pd.read_csv(p, dtype=str).fillna("")
        m["codigo"] = m["codigo"].str.strip()
        return m
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
st.title("📦 Stock HZ + AZ")

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
if modo == "🔍 Buscar":
    try:
        stock = cargar_stock()
    except Exception as e:
        st.error(f"No pude leer la planilla. Revisá que credenciales.json y planilla.txt "
                 f"sean correctos y que compartiste la planilla con el mail del robot.\n\n{e}")
        st.stop()
    if stock.empty:
        st.warning("La planilla está vacía. Andá a «Actualizar» y subí el Excel de stock.")
        st.stop()

    q = st.text_input("Buscar", placeholder="Código, descripción, adicional o proveedor…")
    c1, c2 = st.columns(2)
    deps = ["Todos"] + sorted(stock["deposito"].unique().tolist())
    dep_sel = c1.selectbox("Depósito", deps,
                           format_func=lambda x: "Todos" if x == "Todos" else _dep(x))
    solo_con = c2.toggle("Solo con stock", value=True)

    d = stock
    if dep_sel != "Todos":
        d = d[d["deposito"] == dep_sel]
    if solo_con:
        d = d[d["stock"] != 0]
    if not q.strip():
        st.info("Escribí algo para buscar.")
        st.stop()
    t = q.strip().lower()
    m = (d["codigo"].str.lower().str.contains(t, na=False) |
         d["descripcion"].str.lower().str.contains(t, na=False) |
         d["adicional"].str.lower().str.contains(t, na=False) |
         d["proveedor"].str.lower().str.contains(t, na=False))
    d = d[m]

    st.caption(f"{len(d):,} resultado(s)")
    if d.empty:
        st.warning("Sin resultados. Si el producto existe pero está en 0, destildá «Solo con stock».")
    else:
        vista = pd.DataFrame({
            "Código": d["codigo"], "Descripción": d["descripcion"],
            "Adicional": d["adicional"], "Proveedor": d["proveedor"],
            "Depósito": d["deposito"].map(_dep), "Empresa": d["empresa"], "Stock": d["stock"],
        }).sort_values(["Código", "Depósito"])
        st.dataframe(vista, hide_index=True, use_container_width=True,
                     column_config={"Stock": st.column_config.NumberColumn(format="%.0f")})

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
            cargar_stock.clear()
            st.success(f"¡Listo! Se actualizó el stock de {emp_cod} en la nube. Todos lo ven ya.")
            st.balloons()
