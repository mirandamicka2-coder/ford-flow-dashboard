import streamlit as st
import pandas as pd
from datetime import datetime
import datetime as dt
import os
import base64

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

st.set_page_config(
    page_title="FLOW — Ford Logistics",
    page_icon="🚛",
    layout="wide",
    initial_sidebar_state="collapsed"
)

RECEPCIONES = ["GENERAL", "INTERPLANTA", "JIT", "RANGER"]
DARSENAS    = {"GENERAL": 5, "INTERPLANTA": 2, "JIT": 2, "RANGER": 4}
ICONOS      = {"GENERAL": "📦", "INTERPLANTA": "🔄", "JIT": "⚡", "RANGER": "🛻"}

# ── ESTILOS ───────────────────────────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=IBM+Plex+Sans:wght@300;400;500;600;700&family=IBM+Plex+Mono:wght@400;600&display=swap');

html, body, [class*="css"] {
    font-family: 'IBM Plex Sans', sans-serif;
    font-size: 16px;
}

/* FONDO BLANCO */
.main { background-color: #FFFFFF !important; }
[data-testid="stAppViewContainer"] { background-color: #FFFFFF !important; }
[data-testid="stHeader"] { background-color: #FFFFFF !important; box-shadow: none !important; }
.block-container {
    padding-top: 4rem !important;
    padding-bottom: 1rem;
    background-color: #FFFFFF;
    max-width: 100% !important;
}

/* HEADER sin corte */
.flow-header {
    background: linear-gradient(135deg, #001f4d 0%, #003478 60%, #0050b3 100%);
    padding: 14px 32px;
    margin-bottom: 20px;
    margin-top: 18px;
    margin-left: -1rem;
    margin-right: -1rem;
    display: flex;
    align-items: center;
    justify-content: space-between;
    box-shadow: 0 4px 20px rgba(0,52,120,0.25);
    width: calc(100% + 2rem);
    box-sizing: border-box;
}

/* TABS */
.stTabs [data-baseweb="tab-list"] { gap: 4px; border-bottom: 2px solid #dde3ec; }
.stTabs [data-baseweb="tab"] {
    background: #f8fafc;
    border-radius: 8px 8px 0 0;
    color: #374151;
    font-weight: 600;
    font-size: 15px;
    padding: 10px 22px;
    border: 1px solid #e2e8f0;
}
.stTabs [aria-selected="true"] {
    background: #003478 !important;
    color: white !important;
    border-color: #003478 !important;
}

/* BOTONES */
.stButton > button {
    width: 100% !important;
    height: 72px !important;

    font-size: 24px !important;
    font-weight: 900 !important;

    border-radius: 12px !important;

    letter-spacing: 0.08em !important;

    text-transform: uppercase !important;

    border: none !important;

    box-shadow: 0 4px 12px rgba(0,0,0,0.18) !important;
}

/* MÉTRICAS */
div[data-testid="stMetricValue"] { color: #003478 !important; font-size: 28px !important; }
div[data-testid="stMetricLabel"] { font-size: 15px !important; font-weight: 600 !important; }

/* TEXTOS generales */
p, div, span, label { color: #111827; }

</style>
""", unsafe_allow_html=True)


# ── UTILS ─────────────────────────────────────────────────────────────────────
def img_to_b64(path):
    try:
        with open(path, "rb") as f:
            return base64.b64encode(f.read()).decode()
    except:
        return None


def get_ipd_color(ipd):
    if ipd >= 56: return "#ef4444"
    if ipd >= 26: return "#f59e0b"
    return "#22c55e"


def get_criticidad_color(criticidad):
    c = str(criticidad).strip().upper()
    if c == "ALTA":  return "#ef4444"
    if c == "MEDIA": return "#f59e0b"
    return "#22c55e"


def registrar_accion(patente, accion, detalle=""):
    st.session_state.registro.append({
        "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "patente":   patente,
        "accion":    accion,
        "detalle":   detalle,
    })


# ── CARGA DE DATOS ────────────────────────────────────────────────────────────
def init_data():
    if "registro" not in st.session_state:
        st.session_state.registro = []

    if "camiones" not in st.session_state:
        ruta_excel = os.path.join(BASE_DIR, "reducido_prioridades_1.xlsx")

        if not os.path.exists(ruta_excel):
            st.error(f"❌ No se encontró 'reducido_prioridades_1.xlsx' en: {BASE_DIR}")
            st.session_state.camiones = []
            return

        try:
            df = pd.read_excel(ruta_excel, engine="openpyxl",
                               sheet_name="DEFINITIVO POWER BI")
            df = df.loc[:, ~df.columns.str.contains("^Unnamed")]
            df = df.dropna(subset=["PATENTE", "REMITO"])

            camiones = []
            for i, row in df.iterrows():
                hora_pago_raw = row.get("HORA PAGO")

                if pd.notna(hora_pago_raw):
                    try:
                        if isinstance(hora_pago_raw, dt.datetime):
                            ventana_str = hora_pago_raw.strftime("%H:%M")
                        elif isinstance(hora_pago_raw, dt.time):
                            ventana_str = hora_pago_raw.strftime("%H:%M")
                        else:
                            ventana_str = str(hora_pago_raw)[:5]
                    except:
                        ventana_str = "—"
                else:
                    ventana_str = "—"

                rec_raw = str(row.get("RECEPCIÓN", "GENERAL")).strip().upper()
                if rec_raw not in RECEPCIONES:
                    rec_raw = "GENERAL"

                ipd_excel        = int(row["IPD"])        if pd.notna(row.get("IPD"))        else 0
                criticidad_excel = str(row.get("CRITICIDAD", "BAJA")).strip().upper()
                pp    = int(row["CANTIDAD MÍNIMA"]   if pd.notna(row.get("CANTIDAD MÍNIMA"))   else 0)
                stock = int(row["CANTIDAD EN LÍNEA"] if pd.notna(row.get("CANTIDAD EN LÍNEA")) else 0)
                tipo_raw = str(row.get("TIPO", "CALL")).strip().upper()

                camiones.append({
                    "id":            f"CAM-{i+1:03}",
                    "patente":       str(row.get("PATENTE",      "—")).strip(),
                    "codigo":        str(row.get("RUTA",         "—")).strip(),
                    "descripcion":   str(row.get("DESCRIPCION ", "—")).strip(),
                    "recepcion":     rec_raw,
                    "proveedor":     str(row.get("PEDIDO A",     "—")).strip(),
                    "tipo":          tipo_raw,
                    "material":      str(row.get("PIEZA",        "—")).strip(),
                    "ventana":       ventana_str,
                    "hora_pago_raw": hora_pago_raw,
                    "llegada":       None,
                    "ipd":           ipd_excel,
                    "criticidad":    criticidad_excel,
                    "s_stock":       int(row.get("S.STOCK",  0) if pd.notna(row.get("S.STOCK"))  else 0),
                    "s_tiempo":      int(row.get("S.TIEMPO", 0) if pd.notna(row.get("S.TIEMPO")) else 0),
                    "s_tipo":        int(row.get("S.TIPO",   0) if pd.notna(row.get("S.TIPO"))   else 0),
                    "punto_pedido":  pp,
                    "stock_actual":  stock,
                    "cantidad_lote": int(row["CANTIDAD ENTREGADA"] if pd.notna(row.get("CANTIDAD ENTREGADA")) else 0),
                    "remito":        str(row.get("REMITO", "")).strip(),
                    "darsena":       None,
                    "descargado":    False,
                    "hora_descarga": None,
                })

            st.session_state.camiones = camiones

        except Exception as e:
            st.error(f"❌ Error al leer el Excel: {e}")
            st.session_state.camiones = []


# ── HEADER ────────────────────────────────────────────────────────────────────
def render_header():
    now      = datetime.now()
    ford_b64 = img_to_b64(os.path.join(BASE_DIR, "logo.png"))
    flow_b64 = img_to_b64(os.path.join(BASE_DIR, "flowlogo1.png"))

    ford_img = (f'<img src="data:image/png;base64,{ford_b64}" style="height:62px; object-fit:contain;">'
                if ford_b64 else "")
    flow_img = (f'<img src="data:image/png;base64,{flow_b64}" style="height:72px; object-fit:contain;">'
                if flow_b64 else
                "<span style='color:white;font-size:28px;font-weight:800;letter-spacing:2px;'>FLOW</span>")

    st.markdown(f"""
    <div class="flow-header">
        <div style="display:flex; align-items:center; gap:22px;">
            {ford_img}
            <div style="width:1px; height:58px; background:rgba(255,255,255,0.3);"></div>
            {flow_img}
            <div style="width:1px; height:58px; background:rgba(255,255,255,0.2);"></div>
            <div style="color:rgba(255,255,255,0.65); font-size:18px;
                        letter-spacing:0.20em; text-transform:uppercase; font-weight:500;">
                Ford Logistics Operations Window
            </div>
        </div>
        <div style="text-align:right;">
            <div style="color:white; font-size:42px; font-weight:800;
                        font-family:'IBM Plex Mono',monospace; line-height:1;">
                {now.strftime('%H:%M')}
            </div>
            <div style="color:rgba(255,255,255,0.75); font-size:15px; margin-top:4px;">
                {now.strftime('%d/%m/%Y')}
            </div>
            <div style="color:#7ec8f7; font-size:18px; font-weight:700; margin-top:6px;">
                PLANTA MONTAJE
            </div>
        </div>
    </div>
    """, unsafe_allow_html=True)


# ── TRUCK CARD (100% Streamlit nativo) ───────────────────────────────────────
def render_truck_card(c, show_buttons=True):
    ipd        = c.get("ipd", 0)
    criticidad = c.get("criticidad", "BAJA")
    descargado = c.get("descargado", False)
    ipd_color  = get_ipd_color(ipd)
    crit_color = get_criticidad_color(criticidad)

    # Fondo completo según criticidad
    if descargado:
        border_color = "#64748b"
        bg_color     = "#e2e8f0"
    
    elif criticidad == "ALTA":
        border_color = "#dc2626"
        bg_color     = "#fee2e2"
        
    elif criticidad == "MEDIA":
        border_color = "#d97706"
        bg_color     = "#fef3c7"
        
    else:
        border_color = "#16a34a"
        bg_color     = "#dcfce7"

    # Contenedor con borde izquierdo de color
    st.markdown(f"""
    <div style="
        background:{bg_color};
        border: 3px solid {border_color};
        border-radius: 12px;
        padding: 24px 28px 18px 28px;
        margin-bottom: 4px;
        box-shadow: 0 2px 8px rgba(0,0,0,0.07);
    ">
    """, unsafe_allow_html=True)

    # ── Fila superior: Patente + badges + IPD ────────────────────────────────
    col_info, col_ipd = st.columns([5, 1])

    with col_info:
        # Badge criticidad
        if descargado:
            badge_color = "#64748b"; badge_bg = "#f1f5f9"; badge_txt = "✅ DESCARGADO"
        elif criticidad == "ALTA":
            badge_color = "#b91c1c"; badge_bg = "#fee2e2"; badge_txt = "🔴 ALTA"
        elif criticidad == "MEDIA":
            badge_color = "#a16207"; badge_bg = "#fef9c3"; badge_txt = "🟡 MEDIA"
        else:
            badge_color = "#15803d"; badge_bg = "#dcfce7"; badge_txt = "🟢 BAJA"

        # Badge tipo
        tipo = c.get("tipo", "CALL")
        if "SECUENCIADO" in tipo:
            tipo_bg = "#dbeafe"; tipo_color = "#1d4ed8"; tipo_txt = "SECUENCIADO"
        elif "CARD" in tipo:
            tipo_bg = "#f1f5f9"; tipo_color = "#374151"; tipo_txt = "CARD"
        else:
            tipo_bg = "#f1f5f9"; tipo_color = "#374151"; tipo_txt = "CALL"

        st.markdown(f"""
        <div style="margin-bottom:8px;">
            <span style="font-size:20px; font-weight:800; color:#111827;
                         font-family:'IBM Plex Mono',monospace;">
                {c['patente']}
            </span>
            <span style="margin-left:10px; font-size:13px; color:#6b7280;">
                {c['codigo']}
            </span>
            <span style="display:inline-block; margin-left:10px; padding:3px 12px;
                         border-radius:20px; font-size:13px; font-weight:700;
                         background:{badge_bg}; color:{badge_color};">
                {badge_txt}
            </span>
            <span style="display:inline-block; margin-left:6px; padding:3px 12px;
                         border-radius:20px; font-size:13px; font-weight:700;
                         background:{tipo_bg}; color:{tipo_color};">
                {tipo_txt}
            </span>
            <div style="font-size:16px; color:#374151; margin-top:5px; font-weight:500;">
                {c.get('descripcion','—')}
            </div>
        </div>
        """, unsafe_allow_html=True)

    with col_ipd:
        st.markdown(f"""
        <div style="text-align:right; padding-right:4px;">
            <div style="font-size:36px; font-weight:800; color:{ipd_color};
                        font-family:'IBM Plex Mono',monospace; line-height:1;">
                {ipd}
            </div>
            <div style="font-size:11px; color:#9ca3af; text-transform:uppercase;
                        letter-spacing:0.14em; margin-top:2px;">IPD</div>
            <div style="font-size:12px; color:#6b7280; margin-top:6px; line-height:1.8;">
                Stock: <b style="color:#374151;">{c.get('s_stock',0)}</b><br>
                Tiempo: <b style="color:#374151;">{c.get('s_tiempo',0)}</b><br>
                Tipo: <b style="color:#374151;">{c.get('s_tipo',0)}</b>
            </div>
        </div>
        """, unsafe_allow_html=True)

    # ── Grilla de datos (100% nativa Streamlit) ──────────────────────────────
    g1, g2, g3, g4 = st.columns(4)

    with g1:
        st.markdown("<p style='font-size:12px;color:#6b7280;font-weight:600;"
                    "text-transform:uppercase;letter-spacing:0.06em;margin-bottom:2px;'>"
                    "Proveedor</p>", unsafe_allow_html=True)
        st.markdown(f"<p style='font-size:18px;font-weight:700;color:#111827;margin-top:0;'>"
                    f"{c['proveedor']}</p>", unsafe_allow_html=True)

    with g2:
        st.markdown("<p style='font-size:12px;color:#6b7280;font-weight:600;"
                    "text-transform:uppercase;letter-spacing:0.06em;margin-bottom:2px;'>"
                    "Material</p>", unsafe_allow_html=True)
        st.markdown(f"<p style='font-size:18px;font-weight:700;color:#111827;margin-top:0;'>"
                    f"{c['material']}</p>", unsafe_allow_html=True)

    with g3:
        st.markdown("<p style='font-size:12px;color:#6b7280;font-weight:600;"
                    "text-transform:uppercase;letter-spacing:0.06em;margin-bottom:2px;'>"
                    "Ventana</p>", unsafe_allow_html=True)
        st.markdown(f"<p style='font-size:18px;font-weight:700;color:#111827;margin-top:0;'>"
                    f"{c['ventana']}</p>", unsafe_allow_html=True)

    with g4:
        st.markdown("<p style='font-size:12px;color:#6b7280;font-weight:600;"
                    "text-transform:uppercase;letter-spacing:0.06em;margin-bottom:2px;'>"
                    "Llegada</p>", unsafe_allow_html=True)
        llegada = c.get('llegada') or '—'
        st.markdown(f"<p style='font-size:18px;font-weight:700;color:#111827;margin-top:0;'>"
                    f"{llegada}</p>", unsafe_allow_html=True)

    g5, g6, g7, g8 = st.columns(4)

    with g5:
        st.markdown("<p style='font-size:12px;color:#6b7280;font-weight:600;"
                    "text-transform:uppercase;letter-spacing:0.06em;margin-bottom:2px;'>"
                    "Remito</p>", unsafe_allow_html=True)
        st.markdown(f"<p style='font-size:18px;font-weight:600;color:#111827;"
                    f"font-family:IBM Plex Mono,monospace;margin-top:0;'>"
                    f"{c.get('remito','—')}</p>", unsafe_allow_html=True)

    with g6:
        st.markdown("<p style='font-size:12px;color:#6b7280;font-weight:600;"
                    "text-transform:uppercase;letter-spacing:0.06em;margin-bottom:2px;'>"
                    "Dársena</p>", unsafe_allow_html=True)
        darsena = f"Dársena {c['darsena']}" if c.get('darsena') else 'No asignada'
        st.markdown(f"<p style='font-size:18px;font-weight:700;color:#111827;margin-top:0;'>"
                    f"{darsena}</p>", unsafe_allow_html=True)

    with g7:
        st.markdown("<p style='font-size:12px;color:#6b7280;font-weight:600;"
                    "text-transform:uppercase;letter-spacing:0.06em;margin-bottom:2px;'>"
                    "Stock en línea</p>", unsafe_allow_html=True)
        st.markdown(f"<p style='font-size:18px;font-weight:700;color:#111827;margin-top:0;'>"
                    f"{c.get('stock_actual','—')} u.</p>", unsafe_allow_html=True)

    with g8:
        st.markdown("<p style='font-size:12px;color:#6b7280;font-weight:600;"
                    "text-transform:uppercase;letter-spacing:0.06em;margin-bottom:2px;'>"
                    "Punto de pedido</p>", unsafe_allow_html=True)
        st.markdown(f"<p style='font-size:18px;font-weight:700;color:#111827;margin-top:0;'>"
                    f"{c.get('punto_pedido','—')} u.</p>", unsafe_allow_html=True)

    # ── Alerta ───────────────────────────────────────────────────────────────
    if descargado:
        st.success(f"✅ Descargado a las {c.get('hora_descarga','—')}")
    elif criticidad == "ALTA":
        st.error("⚠️ CRITICIDAD ALTA — Riesgo de parada de línea")
    elif criticidad == "MEDIA":
        st.warning("⏱ CRITICIDAD MEDIA — Monitorear stock")

    st.markdown("</div>", unsafe_allow_html=True)
    st.markdown("<div style='margin-bottom:14px;'></div>", unsafe_allow_html=True)

    # ── Botón único: Descargado ───────────────────────────────────────────────
    if show_buttons and not descargado:
        col_btn = st.container()
        with col_btn:
            if st.button(
                "✅  DESCARGADO",
                key=f"desc_{c['id']}",
                type="primary",
                use_container_width=True
            ):
                hora_ahora = datetime.now().strftime("%H:%M")
                for cam in st.session_state.camiones:
                    if cam["id"] == c["id"]:
                        cam["descargado"]    = True
                        cam["hora_descarga"] = hora_ahora
                registrar_accion(c["patente"], "DESCARGA_CONFIRMADA",
                                 f"Dársena {c.get('darsena','?')} a las {hora_ahora}")
                st.rerun()


# ── RECEPCIÓN ─────────────────────────────────────────────────────────────────
def render_recepcion(recepcion):
    camiones_rec    = [c for c in st.session_state.camiones if c["recepcion"] == recepcion]
    camiones_sorted = sorted(
        camiones_rec,
        key=lambda c: (-c.get("ipd", 0), c.get("ventana", "99:99"))
    )

    criticos     = [c for c in camiones_rec if not c["descargado"] and c.get("criticidad") == "ALTA"]
    header_color = "#c0392b" if criticos else "#003478"

    st.markdown(f"""
    <div style="background:{header_color}; color:white; padding:14px 22px;
                border-radius:10px; margin-bottom:18px;
                display:flex; align-items:center; gap:12px;">
        <span style="font-size:22px;">{ICONOS[recepcion]}</span>
        <span style="font-size:20px; font-weight:800;">Recepción {recepcion}</span>
        <span style="font-size:16px; opacity:0.8; margin-left:auto;">
            {DARSENAS[recepcion]} dársenas
        </span>
    </div>
    """, unsafe_allow_html=True)

    activos     = [c for c in camiones_sorted if not c["descargado"]]
    descargados = [c for c in camiones_sorted if c["descargado"]]

    if not activos:
        st.info("✅ Sin camiones pendientes en esta recepción.")
    else:
        st.markdown(
            f"<p style='font-size:15px;font-weight:700;color:#374151;"
            f"text-transform:uppercase;letter-spacing:0.08em;margin-bottom:12px;'>"
            f"Cola de descarga — {len(activos)} camiones activos</p>",
            unsafe_allow_html=True
        )
        for i, c in enumerate(activos):
            col_num, col_card = st.columns([0.06, 0.94])
            with col_num:
                ipd_col = get_ipd_color(c.get("ipd", 0))
                st.markdown(
                    f"<div style='font-size:24px;font-weight:800;text-align:center;"
                    f"padding-top:22px;font-family:IBM Plex Mono,monospace;color:{ipd_col};'>"
                    f"#{i+1}</div>",
                    unsafe_allow_html=True
                )
            with col_card:
                render_truck_card(c)

    if descargados:
        with st.expander(f"✅ Historial descargados ({len(descargados)})"):
            for c in descargados:
                render_truck_card(c, show_buttons=False)


# ── TABLERO GENERAL ───────────────────────────────────────────────────────────
def render_tablero_general():
    todos = st.session_state.camiones
    cols  = st.columns(4)

    for i, rec in enumerate(RECEPCIONES):
        with cols[i]:
            camiones_rec = [c for c in todos if c["recepcion"] == rec]
            activos_rec  = sorted(
                [c for c in camiones_rec if not c["descargado"]],
                key=lambda c: (-c.get("ipd", 0), c.get("ventana", "99:99"))
            )
            criticos_rec = [c for c in activos_rec if c.get("criticidad") == "ALTA"]
            header_col   = "#c0392b" if criticos_rec else "#003478"

            st.markdown(f"""
            <div style="background:{header_col}; color:white; padding:12px 16px;
                        border-radius:10px 10px 0 0; text-align:center;
                        font-weight:800; font-size:18px;">
                {ICONOS[rec]} {rec}
                <br><span style="font-size:13px; opacity:0.8; font-weight:500;">
                    {len(activos_rec)} activos · {DARSENAS[rec]} dársenas
                </span>
            </div>
            """, unsafe_allow_html=True)

            if not activos_rec:
                st.markdown(
                    "<div style='background:white; border:1px solid #e2e8f0; "
                    "border-top:none; border-radius:0 0 10px 10px; padding:22px; "
                    "text-align:center; color:#6b7280; font-size:16px;'>"
                    "Sin pendientes</div>",
                    unsafe_allow_html=True
                )
            else:
                for j, c in enumerate(activos_rec):
                    ipd    = c.get("ipd", 0)
                    crit   = c.get("criticidad", "BAJA")
                    color  = get_ipd_color(ipd)
                    bg     = "#fff8f8" if crit == "ALTA" else "#fffdf0" if crit == "MEDIA" else "#f9fffc"
                    border = "#fecaca" if crit == "ALTA" else "#fde68a" if crit == "MEDIA" else "#bbf7d0"
                    icon   = "🔴" if crit == "ALTA" else "🟡" if crit == "MEDIA" else "🟢"
                    radius = "border-radius:0 0 10px 10px;" if j == len(activos_rec) - 1 else ""
                    mat    = c['material']
                    mat_s  = mat[:28] + "…" if len(mat) > 28 else mat

                    st.markdown(f"""
                    <div style="background:{bg}; border:1px solid {border};
                                border-top:none; padding:12px 16px; {radius}">
                        <div style="display:flex; justify-content:space-between; align-items:center;">
                            <div>
                                <div style="font-weight:800; font-size:15px; color:#111827;">
                                    #{j+1} {icon} {c['patente']}
                                </div>
                                <div style="font-size:13px; color:#374151; margin-top:3px; font-weight:500;">
                                    {mat_s}
                                </div>
                                <div style="font-size:12px; color:#6b7280; margin-top:2px;">
                                    {c['proveedor']} · {c['tipo']} · ⏰ {c['ventana']}
                                </div>
                            </div>
                            <div style="text-align:right;">
                                <div style="font-size:26px; font-weight:800; color:{color};
                                            font-family:'IBM Plex Mono',monospace;">{ipd}</div>
                                <div style="font-size:11px; color:#6b7280;">IPD</div>
                            </div>
                        </div>
                    </div>
                    """, unsafe_allow_html=True)


# ── MAIN ──────────────────────────────────────────────────────────────────────
def main():
    init_data()
    render_header()

    tabs = st.tabs([
        "📊 Tablero General",
        "📦 GENERAL",
        "🔄 INTERPLANTA",
        "⚡ JIT",
        "🛻 RANGER",
    ])

    with tabs[0]: render_tablero_general()
    with tabs[1]: render_recepcion("GENERAL")
    with tabs[2]: render_recepcion("INTERPLANTA")
    with tabs[3]: render_recepcion("JIT")
    with tabs[4]: render_recepcion("RANGER")


if __name__ == "__main__":
    main()