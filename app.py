import streamlit as st
import pandas as pd
import math
import json
import io
from datetime import date
from calculos import (
    calcular_caida_tension,
    seleccionar_conductor,
    calcular_conduit_fill,
    calcular_cdt_calibre,
    calcular_proteccion,
    factor_correccion_temp,
    temp_terminales_auto,
    ampacidad_base,
    formato_proteccion,
    calibre_tierra,
    TABLA_CONDUCTORES,
    TABLA_9_NOM,
    TABLA_CONDUIT,
    ORDEN_CALIBRES,
    OCPD_ESTANDARES,
    CANALIZACIONES,
    AISLAMIENTOS_T,
    POLOS_POR_CONFIG,
)
from reporte import generar_reporte_pdf
from reporte_docx import generar_reporte_docx
import auth

# ─────────────────────────────────────────────
# HELPERS — SESIÓN JSON Y EXPORTACIÓN EXCEL
# ─────────────────────────────────────────────
def _aplicar_seleccion_manual(calibre, cdt_nuevo, amp_nuevo, I_diseño_cond,
                                config_code, tierra_mat, tierra_auto):
    """Actualiza session_state["resultado_conductor"] con un calibre seleccionado
       manualmente, RECALCULANDO también la protección (OCPD), su formato comercial
       y el conductor de tierra para que todo quede consistente."""
    rc = st.session_state.get("resultado_conductor")
    if not rc:
        return
    new_ocpd_A, new_ocpd_st = calcular_proteccion(I_diseño_cond, amp_nuevo)
    new_ocpd_fmt = formato_proteccion(new_ocpd_A, config_code)
    new_tierra = (
        calibre_tierra(new_ocpd_A, tierra_mat)
        if tierra_auto else rc.get("tierra_calibre", "—")
    )
    rc["conductor"]         = calibre
    rc["cdt"]               = cdt_nuevo
    rc["ampacity_corr"]     = amp_nuevo
    rc["proteccion_A"]      = new_ocpd_A
    rc["proteccion_status"] = new_ocpd_st
    rc["proteccion_fmt"]    = new_ocpd_fmt
    rc["tierra_calibre"]    = new_tierra
    rc["seleccion_manual"]  = True


def _sanitize_json(obj):
    """Convierte recursivamente DataFrames y otros tipos no-JSON a estructuras serializables."""
    if isinstance(obj, pd.DataFrame):
        return obj.to_dict(orient="records")
    if isinstance(obj, dict):
        return {k: _sanitize_json(v) for k, v in obj.items()}
    if isinstance(obj, (list, tuple)):
        return [_sanitize_json(v) for v in obj]
    if hasattr(obj, "isoformat"):  # date / datetime
        return obj.isoformat()
    return obj


def exportar_sesion_json(circuitos: list, datos_proy: dict) -> str:
    return json.dumps(
        _sanitize_json({"version": "1.2", "proyecto": datos_proy, "circuitos": circuitos}),
        ensure_ascii=False, indent=2, default=str,
    )

def crear_excel_circuitos(circuitos: list) -> bytes:
    rows = []
    for c in circuitos:
        rec = (c.get("conduit_info") or {}).get("recomendada") or {}
        cf = c.get("cf", 1)
        cdt_val = c.get("cdt")
        cdt_max_val = c.get("cdt_max", 3.0)
        if cdt_val is None:
            estado = "SIN DATOS"
        elif cdt_val <= cdt_max_val:
            estado = "CUMPLE"
        else:
            estado = "NO CUMPLE"
        rows.append({
            "No.": c.get("id"),
            "Circuito": c.get("nombre", ""),
            "Potencia (W)": c.get("potencia"),
            "FP": c.get("fp"),
            "Voltaje (V)": c.get("voltaje"),
            "Configuración": c.get("configuracion", ""),
            "I real (A)": round(c.get("corriente", 0) or 0, 2),
            "I diseño (A)": round(c.get("corriente_diseño", 0) or 0, 2),
            "Longitud (m)": c.get("longitud"),
            "CF": cf,
            "Material": c.get("material"),
            "Conductor AWG": c.get("conductor"),
            "Amp. corr. (A)": round(c.get("ampacity_corr", 0) or 0, 1),
            "C.d.T. (%)": round(cdt_val, 2) if cdt_val is not None else "—",
            "C.d.T. máx. (%)": cdt_max_val,
            "OCPD (A)": c.get("proteccion_A", ""),
            "OCPD estado": c.get("proteccion_status", ""),
            "Canalización": c.get("canalizacion_label", c.get("canalizacion", "")),
            "T° term. (°C)": c.get("temp_term", ""),
            "Aislamiento (°C)": c.get("temp_aislamiento", ""),
            "Tubería": rec.get("tubo", ""),
            "Relleno (%)": rec.get("fill_pct", ""),
            "Estado": estado,
        })
    df = pd.DataFrame(rows)
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Circuitos")
        ws = writer.sheets["Circuitos"]
        for col in ws.columns:
            w = max(len(str(cell.value or "")) for cell in col)
            ws.column_dimensions[col[0].column_letter].width = max(12, min(w + 2, 35))
    return buf.getvalue()

# ─────────────────────────────────────────────
# CONFIGURACIÓN DE PÁGINA
# ─────────────────────────────────────────────
st.set_page_config(
    page_title="eSuit · Cálculo eléctrico",
    page_icon="⚡",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ─────────────────────────────────────────────
# CSS PERSONALIZADO
# ─────────────────────────────────────────────
st.markdown("""
<style>
    /* ── Tipografía sistema (iOS / macOS) ─────────────────────── */
    html, body, [class*="css"], .stApp {
        font-family: -apple-system, BlinkMacSystemFont, "SF Pro Display",
                     "SF Pro Text", "Helvetica Neue", Helvetica, Arial, sans-serif;
        -webkit-font-smoothing: antialiased;
        letter-spacing: -0.01em;
        background: #f5f5f7;
    }

    /* ── Encabezado principal ─────────────────────────────────── */
    .main-title {
        font-size: 2.1rem;
        font-weight: 700;
        color: #1d1d1f;
        margin: 4px 0 2px 0;
        letter-spacing: -0.025em;
    }
    .sub-title {
        color: #86868b;
        font-size: 0.95rem;
        font-weight: 400;
        margin-bottom: 24px;
    }

    /* ── Bloques de resultado (estilo iOS Notifications) ──────── */
    .resultado-ok, .resultado-warn, .resultado-error {
        border-radius: 14px;
        padding: 14px 18px;
        margin: 10px 0;
        font-weight: 500;
        backdrop-filter: blur(20px);
    }
    .resultado-ok      { background: rgba(52,199,89,0.12);   color: #1d6e3a; border: 1px solid rgba(52,199,89,0.3); }
    .resultado-warn    { background: rgba(255,159,10,0.13);  color: #8a5a00; border: 1px solid rgba(255,159,10,0.32); }
    .resultado-error   { background: rgba(255,59,48,0.12);   color: #a31621; border: 1px solid rgba(255,59,48,0.3); }

    /* ── Metric cards (estilo widget iOS) ─────────────────────── */
    .metric-card {
        background: #ffffff;
        border: 1px solid rgba(0,0,0,0.06);
        border-radius: 16px;
        padding: 16px;
        text-align: center;
        box-shadow: 0 1px 3px rgba(0,0,0,0.04);
        transition: transform 0.15s ease, box-shadow 0.15s ease;
    }
    .metric-card:hover {
        transform: translateY(-1px);
        box-shadow: 0 4px 12px rgba(0,0,0,0.08);
    }
    .metric-label {
        font-size: 0.72rem;
        color: #86868b;
        text-transform: uppercase;
        letter-spacing: 0.06em;
        font-weight: 600;
    }
    .metric-value {
        font-size: 1.65rem;
        font-weight: 600;
        color: #1d1d1f;
        margin-top: 4px;
        letter-spacing: -0.02em;
    }
    .metric-unit {
        font-size: 0.78rem;
        color: #86868b;
        margin-top: 2px;
    }

    /* ── Section header (estilo macOS) ────────────────────────── */
    .section-header {
        font-size: 1.1rem;
        font-weight: 600;
        color: #1d1d1f;
        margin: 24px 0 14px 0;
        padding-bottom: 8px;
        border-bottom: 1px solid rgba(0,0,0,0.08);
        letter-spacing: -0.01em;
    }

    /* ── Sidebar (iOS dark glass) ─────────────────────────────── */
    div[data-testid="stSidebar"] {
        background: #1c1c1e;
        border-right: 1px solid rgba(255,255,255,0.06);
    }
    div[data-testid="stSidebar"] > div { padding-top: 1rem; }
    div[data-testid="stSidebar"] * { color: #f5f5f7 !important; }
    div[data-testid="stSidebar"] input,
    div[data-testid="stSidebar"] textarea,
    div[data-testid="stSidebar"] select {
        background: rgba(255,255,255,0.08) !important;
        border-radius: 10px !important;
        border: 1px solid rgba(255,255,255,0.1) !important;
        color: #f5f5f7 !important;
    }
    div[data-testid="stSidebar"] label { font-size: 0.82rem !important; opacity: 0.85; }

    /* ── Tabs (estilo iOS segmented control) ──────────────────── */
    .stTabs [data-baseweb="tab-list"] {
        gap: 4px;
        background: rgba(0,0,0,0.04);
        padding: 4px;
        border-radius: 12px;
    }
    .stTabs [data-baseweb="tab"] {
        font-weight: 500;
        border-radius: 8px;
        padding: 8px 16px;
    }
    .stTabs [aria-selected="true"] {
        background: #ffffff !important;
        box-shadow: 0 1px 3px rgba(0,0,0,0.08);
    }

    /* ── Inputs (iOS) ────────────────────────────────────────── */
    .stTextInput input, .stNumberInput input, .stTextArea textarea {
        border-radius: 10px !important;
        border: 1px solid rgba(0,0,0,0.1) !important;
    }
    .stSelectbox > div > div {
        border-radius: 10px !important;
    }

    /* ── Buttons (iOS pill) ──────────────────────────────────── */
    .stButton > button {
        border-radius: 980px;
        font-weight: 500;
        padding: 6px 18px;
        border: none;
        transition: all 0.15s ease;
    }
    .stButton > button[kind="primary"] {
        background: #0071e3;
        color: white;
    }
    .stButton > button[kind="primary"]:hover {
        background: #0077ed;
        transform: translateY(-1px);
    }
    .stButton > button[kind="secondary"] {
        background: rgba(0,0,0,0.05);
        color: #1d1d1f;
    }

    /* ── Expander (iOS group) ─────────────────────────────────── */
    div[data-testid="stExpander"] {
        background: #ffffff;
        border-radius: 14px;
        border: 1px solid rgba(0,0,0,0.06);
        box-shadow: 0 1px 2px rgba(0,0,0,0.03);
        margin: 8px 0;
    }

    /* ── Dataframe (cleaner) ──────────────────────────────────── */
    .stDataFrame {
        border-radius: 12px;
        overflow: hidden;
    }

    /* ── Captions ────────────────────────────────────────────── */
    .stCaption { color: #86868b !important; }

    /* ── Info box for the formula display ─────────────────────── */
    .formula-box {
        background: #ffffff;
        border: 1px solid rgba(0,0,0,0.06);
        border-radius: 12px;
        padding: 12px 16px;
        margin: 6px 0 18px 0;
        font-family: "SF Mono", "Menlo", Consolas, monospace;
        font-size: 0.88rem;
        color: #1d1d1f;
        box-shadow: 0 1px 2px rgba(0,0,0,0.03);
    }

    /* ── Project header card (truncate-safe) ──────────────────── */
    .project-header-card {
        background: #ffffff;
        border: 1px solid rgba(0,0,0,0.06);
        border-radius: 14px;
        padding: 12px 16px;
        margin-bottom: 16px;
        box-shadow: 0 1px 2px rgba(0,0,0,0.03);
        overflow: hidden;
    }
    .project-header-row {
        display: flex;
        flex-wrap: wrap;
        gap: 8px 22px;
        font-size: 0.88rem;
    }
    .project-header-row span b { color: #1d1d1f; }
    .project-header-row span    { color: #6e6e73; word-break: break-word; }

    /* ── Footer ──────────────────────────────────────────────── */
    .stApp footer { display: none; }
</style>
""", unsafe_allow_html=True)

# ═════════════════════════════════════════════════════════
# SPLASH SCREEN + LOGIN GATE  (eSuit)
# ═════════════════════════════════════════════════════════
def _mostrar_splash_login():
    """Pantalla de bienvenida con logo y formulario de login.
       Bloquea el resto de la app hasta que el usuario se autentique."""
    # CSS específico de splash (centra todo)
    st.markdown("""
    <style>
        /* Oculta el sidebar mientras no haya login */
        section[data-testid="stSidebar"] { display: none; }
        div[data-testid="collapsedControl"] { display: none; }
        .splash-container {
            max-width: 440px;
            margin: 2rem auto 1rem auto;
            text-align: center;
        }
        .splash-logo {
            display: flex; justify-content: center;
            margin-bottom: 18px;
            animation: zoomIn 0.6s ease-out;
        }
        .splash-title {
            font-size: 3rem; font-weight: 700;
            letter-spacing: -0.04em; color: #1d1d1f;
            margin: 0;
            animation: fadeUp 0.6s ease-out 0.2s both;
        }
        .splash-subtitle {
            font-size: 1.05rem; color: #86868b;
            margin: 4px 0 28px 0; font-weight: 400;
            animation: fadeUp 0.6s ease-out 0.35s both;
        }
        .splash-card {
            background: #ffffff;
            border: 1px solid rgba(0,0,0,0.06);
            border-radius: 18px;
            padding: 28px 32px;
            box-shadow: 0 8px 28px rgba(0,0,0,0.06);
            animation: fadeUp 0.6s ease-out 0.5s both;
        }
        .splash-foot {
            color: #86868b; font-size: 0.8rem;
            margin-top: 1.5rem;
            animation: fadeUp 0.6s ease-out 0.65s both;
        }
        @keyframes zoomIn {
            from { opacity: 0; transform: scale(0.6); }
            to   { opacity: 1; transform: scale(1); }
        }
        @keyframes fadeUp {
            from { opacity: 0; transform: translateY(12px); }
            to   { opacity: 1; transform: translateY(0); }
        }
    </style>
    """, unsafe_allow_html=True)

    st.markdown(
        f"""<div class="splash-container">
            <div class="splash-logo">{auth.LOGO_SVG}</div>
            <h1 class="splash-title">eSuit</h1>
            <p class="splash-subtitle">
                Cálculo eléctrico profesional · NOM-001-SEDE-2012<br>
                Caída de tensión · Conductores · Protecciones · Puesta a tierra
            </p>
        </div>""",
        unsafe_allow_html=True,
    )

    # Si nunca se ha visto el splash, mostrarlo unos segundos antes del login.
    # (Botón "Continuar" para que el usuario decida cuándo pasar al login.)
    if not st.session_state.get("splash_shown", False):
        col_a, col_b, col_c = st.columns([1, 1, 1])
        with col_b:
            if st.button("Continuar →", type="primary", use_container_width=True):
                st.session_state["splash_shown"] = True
                st.rerun()
        st.markdown(
            '<div class="splash-foot">v2.0 · Build interno</div>',
            unsafe_allow_html=True
        )
        st.stop()

    # ── Formulario de login ──
    _, col_login, _ = st.columns([1, 2, 1])
    with col_login:
        with st.form("form_login", clear_on_submit=False):
            st.markdown("### Iniciar sesión")
            user = st.text_input("Usuario", placeholder="admin",
                                  autocomplete="username")
            pwd = st.text_input("Contraseña", type="password",
                                 placeholder="••••••", autocomplete="current-password")
            submit = st.form_submit_button("Iniciar sesión",
                                            type="primary", use_container_width=True)
            if submit:
                resultado = auth.verificar(user.strip(), pwd)
                if resultado:
                    st.session_state["auth_user"] = resultado
                    st.rerun()
                else:
                    st.error("Usuario o contraseña incorrectos.")

        # Pista del usuario por defecto, solo si nunca se ha cambiado
        try:
            usuarios = auth.cargar_usuarios()
            admin_default = (
                "admin" in usuarios
                and usuarios["admin"].get("default_pwd", False)
            )
            if admin_default:
                st.info(
                    "🔑 **Primer arranque** — usa `admin` / `admin123` para "
                    "iniciar sesión. Cambia esta contraseña desde el panel "
                    "de administración."
                )
        except Exception:
            pass

        st.markdown(
            '<div class="splash-foot">eSuit v2.0 · '
            '© Cálculo eléctrico profesional</div>',
            unsafe_allow_html=True
        )
    st.stop()


# Si no hay usuario autenticado, mostrar splash/login y bloquear el resto.
if "auth_user" not in st.session_state:
    _mostrar_splash_login()


# ─────────────────────────────────────────────
# SESSION STATE — INIT
# ─────────────────────────────────────────────
if "circuitos" not in st.session_state:
    st.session_state["circuitos"] = []

# ─────────────────────────────────────────────
# SIDEBAR — DATOS DEL PROYECTO
# ─────────────────────────────────────────────
with st.sidebar:
    # ── Tarjeta de sesión (usuario autenticado) ──
    _usr = st.session_state["auth_user"]
    _is_admin = (_usr.get("rol") == "admin")
    st.markdown(
        f"""<div style="display:flex;align-items:center;gap:10px;
            padding:10px 12px;margin-bottom:8px;
            background:rgba(255,255,255,0.06);border-radius:12px;
            border:1px solid rgba(255,255,255,0.08);">
            <div style="flex:0 0 auto;">{auth.LOGO_SVG_SMALL}</div>
            <div style="flex:1 1 auto;line-height:1.2;">
                <div style="font-weight:600;font-size:0.95rem;color:#fff;">
                    {_usr.get('nombre','—')}
                </div>
                <div style="font-size:0.75rem;opacity:0.65;color:#fff;">
                    {_usr.get('user','—')} · {_usr.get('rol','—')}
                </div>
            </div>
        </div>""",
        unsafe_allow_html=True,
    )
    if st.button("🚪 Cerrar sesión", key="btn_logout", use_container_width=True):
        for k in list(st.session_state.keys()):
            del st.session_state[k]
        st.rerun()

    # Aviso si admin sigue con contraseña por defecto
    if _is_admin and _usr.get("default_pwd"):
        st.warning(
            "⚠️ Estás usando la **contraseña por defecto** (`admin123`). "
            "Ve al panel de administración para cambiarla."
        )

    st.markdown("---")
    st.markdown("### ⚡ Datos del proyecto")

    with st.expander("📋 Información general", expanded=True):
        proyecto = st.text_area(
            "Nombre del proyecto",
            "Instalación Eléctrica Comercial",
            height=68,
            help="Acepta textos largos. Aparecerá en portada del PDF.",
        )
        cliente = st.text_area("Cliente / Propietario", "", height=68)
        ubicacion = st.text_area("Ubicación", "Salamanca, Gto.", height=68)
        descripcion_proy = st.text_area(
            "Descripción del proyecto",
            "Memoria técnica para la instalación eléctrica del inmueble, "
            "abarcando alimentación principal, circuitos derivados, puesta a "
            "tierra y protecciones contra sobrecorriente.",
            height=110,
            help="Texto descriptivo que aparecerá en la sección 'Descripción del proyecto' del reporte.",
        )

    with st.expander("👷 Responsable técnico", expanded=False):
        responsable = st.text_input("Nombre completo", "")
        cedula = st.text_input("Cédula profesional", "")
        fecha_proj = st.date_input("Fecha", value=date.today())

    with st.expander("📐 Normatividad y sistema", expanded=False):
        norma = st.selectbox(
            "Norma principal",
            ["NOM-001-SEDE-2012", "NEC 2023", "NOM-001-SEDE-2005"]
        )
        normas_extra = st.multiselect(
            "Normas complementarias",
            [
                "CFE DCDIAMT (Instalaciones aéreas MT)",
                "CFE DCDIASMT (Instalaciones subterráneas MT)",
                "NOM-007-ENER-2014 (Eficiencia energética)",
                "NMX-J-098-ANCE (Tensiones eléctricas estándar)",
                "NMX-J-235-ANCE (Conductores eléctricos)",
                "IEEE Std 141 (Red Book)",
                "IEC 60364 (Instalaciones eléctricas en edificios)",
            ],
            default=[],
            help="Se incluirán en la sección 'Normatividad aplicable' del reporte.",
        )
        tipo_sistema = st.selectbox(
            "Sistema eléctrico",
            ["Monofásico 120V", "Monofásico 240V",
             "Bifásico 240V (Mono 3H)", "Trifásico 220V", "Trifásico 440V"]
        )
        temp_ambiente = st.number_input(
            "Temperatura ambiente (°C)",
            value=35, min_value=-10, max_value=80, step=1,
            help="Temperatura del entorno donde corre la canalización. "
                 "Afecta el factor de corrección de ampacidad.",
        )

    fc_temp = factor_correccion_temp(temp_ambiente)
    if temp_ambiente > 30:
        st.info(f"FC temperatura: **{fc_temp:.2f}** · NOM-001 Tabla 310-15(B)(2)(a)")
    else:
        st.success(f"FC temperatura: **{fc_temp:.2f}** (≤ 30°C, sin derateo)")

    # Validación rápida en el sidebar
    if not proyecto.strip():
        st.warning("⚠️ El nombre del proyecto está vacío.")
    if not responsable.strip():
        st.caption("ℹ️ Tip: completa el responsable técnico para la firma del PDF.")

# ─────────────────────────────────────────────
# ENCABEZADO PRINCIPAL
# ─────────────────────────────────────────────
st.markdown('<div class="main-title">⚡ eSuit · Cálculo eléctrico profesional</div>', unsafe_allow_html=True)
st.markdown(
    '<div class="sub-title">Cálculo profesional según NOM-001-SEDE-2012 · '
    'caída de tensión, ampacidad, protecciones y puesta a tierra</div>',
    unsafe_allow_html=True
)

# Project header card (handles long texts gracefully)
def _esc(s):
    return (s or "").replace("<", "&lt;").replace(">", "&gt;")

st.markdown(
    f'''<div class="project-header-card">
        <div class="project-header-row">
            <span><b>Proyecto:</b> {_esc(proyecto.strip() or "—")}</span>
            <span><b>Cliente:</b> {_esc(cliente.strip() or "—")}</span>
            <span><b>Sistema:</b> {_esc(tipo_sistema)}</span>
            <span><b>Norma:</b> {_esc(norma)}</span>
        </div>
    </div>''',
    unsafe_allow_html=True
)

# ─────────────────────────────────────────────
# TABS PRINCIPALES
# ─────────────────────────────────────────────
# Pestañas: incluye gestión de usuarios solo para admin
_tab_labels = [
    "⚡ Caída de Tensión & Conductor",
    "🔧 Tubería (Conduit Fill)",
    "📊 Resumen General",
    "📄 Reporte PDF",
]
if _is_admin:
    _tab_labels.append("👥 Usuarios")
    tab1, tab2, tab3, tab4, tab_users = st.tabs(_tab_labels)
else:
    tab1, tab2, tab3, tab4 = st.tabs(_tab_labels)
    tab_users = None

# ══════════════════════════════════════════════
# TAB 1 — CAÍDA DE TENSIÓN Y SELECCIÓN DE CONDUCTOR
# ══════════════════════════════════════════════
with tab1:
    st.markdown('<div class="section-header">Método de Cálculo</div>', unsafe_allow_html=True)

    col_m1, col_m2 = st.columns(2)
    with col_m1:
        metodo_label = st.selectbox(
            "Método de caída de tensión",
            ["Por Impedancia (R·cosφ + X·senφ)", "Por Sección Transversal", "Aéreo (NEUTRANEL / ACSR)"],
        )
    with col_m2:
        config_label = st.selectbox(
            "Configuración del circuito",
            ["Monofásico 2H (Fase + Neutro)", "Monofásico 3H (Fase + Fase + Neutro)", "Trifásico"],
        )

    metodo_map = {
        "Por Impedancia (R·cosφ + X·senφ)": "impedancia",
        "Por Sección Transversal": "seccion",
        "Aéreo (NEUTRANEL / ACSR)": "aereo",
    }
    config_map = {
        "Monofásico 2H (Fase + Neutro)": "mono2h",
        "Monofásico 3H (Fase + Fase + Neutro)": "mono3h",
        "Trifásico": "trifasico",
    }
    metodo = metodo_map[metodo_label]
    config = config_map[config_label]

    formulas = {
        ("impedancia", "mono2h"):    "%VD = (200 × I × L × (R·cosφ + X·senφ)) / V",
        ("impedancia", "mono3h"):    "%VD = (200 × I × L × (R·cosφ + X·senφ)) / V",
        ("impedancia", "trifasico"): "%VD = (√3 × I × L × (R·cosφ + X·senφ) × 100) / V",
        ("seccion",    "mono2h"):    "%VD = (4 × L[m] × I) / (120 × S)",
        ("seccion",    "mono3h"):    "%VD = (2 × L[m] × I) / (120 × S)",
        ("seccion",    "trifasico"): "%VD = (2 × √3 × L[m] × I) / (V × S)",
        ("aereo",      "mono2h"):    "VD[V] = (2 × L[m] × I × ρ) / S",
        ("aereo",      "mono3h"):    "VD[V] = (2 × L[m] × I × ρ) / S",
        ("aereo",      "trifasico"): "VD[V] = (√3 × L[m] × I × ρ) / S",
    }
    st.markdown(
        f'<div class="formula-box"><b>Fórmula activa:</b>&nbsp;&nbsp;{formulas.get((metodo, config), "—")}</div>',
        unsafe_allow_html=True
    )

    st.markdown('<div class="section-header">Parámetros del Circuito</div>', unsafe_allow_html=True)

    voltaje_default = {"mono2h": 120, "mono3h": 240, "trifasico": 220}
    voltaje_nominal = st.number_input(
        "Voltaje nominal (V)", min_value=12, max_value=15000,
        value=voltaje_default[config], step=10
    )

    col1, col2, col3 = st.columns(3)
    with col1:
        potencia = st.number_input(
            "Potencia de la carga (W)",
            min_value=1, max_value=10_000_000, value=5000, step=100,
            help="Potencia activa de la carga en Watts. Debe ser > 0.",
        )
        fp = st.number_input(
            "Factor de potencia (cosφ)",
            min_value=0.1, max_value=1.0, value=0.90, step=0.01,
            help="Coseno del ángulo entre tensión y corriente. Inductivo típicamente 0.85-0.95.",
        )
        sen_phi = round(math.sqrt(max(0.0, 1 - fp**2)), 4)
        st.caption(f"senφ = √(1 − cosφ²) = **{sen_phi}**")
    with col2:
        longitud = st.number_input(
            "Longitud del circuito (m)",
            min_value=1, max_value=5000, value=40, step=1,
            help="Longitud de un solo conductor (ida). Las fórmulas multiplican por 2 para ida+vuelta en monofásico.",
        )
        material = st.selectbox("Material del conductor", ["Cobre (Cu)", "Aluminio (Al)"])
        cf = st.number_input(
            "Conductores por fase (CF)",
            min_value=1, max_value=10, value=1, step=1,
            help="Conductores en paralelo por fase. La corriente por conductor = I / CF",
        )
    with col3:
        cdt_max = st.number_input(
            "C.d.T. máxima permitida (%)",
            min_value=0.5, max_value=10.0, value=3.0, step=0.5,
            help="NOM-001: alimentadores 2%, ramales 3%, total 5%.",
        )
        conductores_activos = st.selectbox(
            "Conductores en tubo (agrupamiento)",
            [1, 2, 3, 4, 5, 6], index=2,
            help="3 conductores portadores = 1.00; 4-6 reducen ampacidad.",
        )
        es_motor = st.checkbox(
            "Es circuito derivado de motor (NEC 430)",
            value=False,
            help="Activa para usar 75°C en terminales (independiente de la corriente).",
        )

    # ── Carga continua: nota desplegable explicativa ─────────────
    with st.expander("💡 ¿Cuándo se aplica el factor 1.25 (carga continua)?", expanded=False):
        st.markdown("""
**Carga continua** (NEC 210.19 / 210.20 · NOM-001 Art. 210):

> Toda carga que opere durante **3 horas o más de forma ininterrumpida** se considera continua.
> En ese caso el conductor y la protección deben dimensionarse para **125 %** de la corriente real:

```
I_diseño = I_real × 1.25
```

**Aplica para:**
- 💡 Alumbrado comercial y de oficinas (operación todo el día)
- 🏭 Cargas de procesos industriales que corren turnos largos
- ❄️ Sistemas HVAC y bombas que arrancan y permanecen encendidos
- 🔌 Tableros principales que alimentan cargas continuas

**NO aplica para:**
- 🍳 Cargas intermitentes (electrodomésticos, contactos de hogar)
- 🚪 Cargas de muy corta duración (puertas eléctricas, etc.)
- ⚡ Cargas con tiempo de uso inferior a 3 horas
- 🛠️ Motores (estos llevan **1.25 fijo del motor mayor + 100 % del resto** por NEC 430.24)

**Recomendación:** Si tienes duda, deja marcado el checkbox — el dimensionamiento es conservador.
        """)

    carga_continua = st.checkbox(
        "Aplicar factor de carga continua ×1.25 (NEC 210.20)",
        value=True,
        help="Para cargas que operan ≥ 3 horas continuas. Ver expander arriba.",
    )

    # ── Canalización, temperatura de terminales y aislamiento (Tabla 9 NOM / Art. 110-14) ──
    st.markdown('<div class="section-header">Canalización y Aislamiento</div>', unsafe_allow_html=True)
    col4, col5, col6 = st.columns(3)
    with col4:
        canal_label = st.selectbox(
            "Tipo de canalización (Tabla 9 NOM)",
            list(CANALIZACIONES.keys()),
            index=2,  # default Acero
            help="Afecta R y X. Acero (magnética) tiene mayor reactancia que PVC/Aluminio.",
        )
    with col5:
        term_label = st.selectbox(
            "Temperatura terminales (Art. 110-14)",
            ["Auto (según In)", "60 °C", "75 °C", "90 °C"],
            help="Auto: 60°C si In ≤ 100 A, 75°C si In > 100 A. Motores ≠ Clase A usan 75°C.",
        )
    with col6:
        aisl_label = st.selectbox(
            "Aislamiento del conductor (ampacidad)",
            list(AISLAMIENTOS_T.keys()),
            index=1,  # default 75°C (THW)
            help="THW=75°C, THHW-2/XHHW-2=90°C. La ampacidad se selecciona de Tabla 310-15(B)(16).",
        )

    canalizacion = CANALIZACIONES[canal_label]
    temp_aislamiento = AISLAMIENTOS_T[aisl_label]
    temp_term_override = None if term_label.startswith("Auto") else int(term_label.split()[0])

    fc_agrup_val = {1: 1.0, 2: 0.88, 3: 0.82, 4: 0.75, 5: 0.67, 6: 0.58}[conductores_activos]
    factor_demanda = 1.25 if carga_continua else 1.0

    # ── Conductor de puesta a tierra (cable desnudo) ─────────────
    st.markdown('<div class="section-header">Conductor de puesta a tierra (cable desnudo)</div>', unsafe_allow_html=True)
    col_t1, col_t2, col_t3 = st.columns([1, 1, 1])
    with col_t1:
        tierra_auto = st.checkbox(
            "Auto (NOM-001 Tabla 250-122)",
            value=True,
            help="Selecciona automáticamente el calibre según la capacidad de la protección.",
        )
    with col_t2:
        tierra_material = st.selectbox(
            "Material de tierra",
            ["Cu (cobre desnudo)", "Al (aluminio)"],
            index=0,
        )
    with col_t3:
        if tierra_auto:
            st.text_input(
                "Calibre de tierra (auto)",
                value="— se calcula con la protección —",
                disabled=True,
            )
            tierra_manual = None
        else:
            tierra_manual = st.selectbox(
                "Calibre de tierra manual",
                ORDEN_CALIBRES,
                index=ORDEN_CALIBRES.index("8"),
            )

    # ── Conduit integrado para este circuito (opcional) ──────────
    with st.expander("🔧 Calcular tubería (conduit) para este circuito", expanded=False):
        st.caption("Activa esta opción para calcular automáticamente la tubería que aloja los conductores del circuito.")
        incluir_conduit_calc = st.checkbox(
            "Incluir cálculo de tubería en este circuito",
            value=False,
            key="chk_incluir_conduit_inline",
        )
        if incluir_conduit_calc:
            cc1, cc2 = st.columns(2)
            with cc1:
                tipo_tubo_inline = st.selectbox(
                    "Tipo de tubería",
                    list(TABLA_CONDUIT.keys()),
                    key="tubo_inline",
                )
                # Por defecto: 3 fases + 1 neutro + 1 tierra para trifásico
                num_default = {"mono2h": 3, "mono3h": 4, "trifasico": 5}.get(config, 3)
                num_conds_inline = st.number_input(
                    "Número total de conductores (forrados + tierra desnuda)",
                    min_value=1, max_value=40,
                    value=num_default, step=1,
                    key="ncond_inline",
                    help=f"Por defecto: {num_default} (fases + neutro + tierra).",
                )
            with cc2:
                tipo_aisl_inline = st.selectbox(
                    "Tipo de aislamiento",
                    ["THW (75°C)", "THHW (75/90°C)", "THWN-2 (90°C)", "XHHW-2 (90°C)"],
                    key="aisl_inline",
                )
                st.caption(
                    "ℹ️ El cálculo usa el OD del calibre principal del circuito. "
                    "Si tienes calibres mixtos (ej. tierra menor), el resultado es conservador."
                )
        else:
            tipo_tubo_inline = None
            num_conds_inline = None
            tipo_aisl_inline = None

    st.markdown('<div class="section-header">Resultados del Cálculo</div>', unsafe_allow_html=True)

    # ── Validación de entradas antes de calcular ─────────────────
    def _validar_entradas():
        errs = []
        if potencia <= 0:
            errs.append("Potencia debe ser > 0 W")
        if voltaje_nominal <= 0:
            errs.append("Voltaje nominal debe ser > 0 V")
        if longitud <= 0:
            errs.append("Longitud debe ser > 0 m")
        if not (0 < fp <= 1):
            errs.append("Factor de potencia debe estar en (0, 1]")
        if cdt_max <= 0:
            errs.append("C.d.T. máxima debe ser > 0 %")
        if cf < 1:
            errs.append("CF debe ser ≥ 1")
        if not proyecto.strip():
            errs.append("El nombre del proyecto está vacío (sidebar)")
        return errs

    if st.button("🔍 Calcular", key="btn_calc_conductor", type="primary"):
        errores_input = _validar_entradas()
        if errores_input:
            st.error("❌ No se puede calcular:\n\n" + "\n".join(f"• {e}" for e in errores_input))
            st.stop()
        mat = "Cu" if "Cobre" in material else "Al"
        corriente, conductor_min, cdt_real, conductor_verificado, ampacity_corr = calcular_caida_tension(
            potencia, fp, voltaje_nominal, config, longitud, mat, cdt_max, fc_temp, fc_agrup_val, metodo,
            factor_demanda=factor_demanda, cf=cf,
            canalizacion=canalizacion, temp_term=temp_term_override,
            temp_aislamiento=temp_aislamiento, es_motor=es_motor,
        )
        corriente_diseño = corriente * factor_demanda
        corriente_diseño_cond = corriente_diseño / cf
        ocpd_A, ocpd_status = calcular_proteccion(corriente_diseño_cond, ampacity_corr)
        temp_term_efectiva = temp_term_override or temp_terminales_auto(corriente, es_motor=es_motor)

        df_tabla = seleccionar_conductor(
            corriente, longitud, voltaje_nominal, config, mat,
            cdt_max, fc_temp, fc_agrup_val, fp, metodo,
            factor_demanda=factor_demanda, cf=cf,
            canalizacion=canalizacion, temp_term=temp_term_override,
            temp_aislamiento=temp_aislamiento, es_motor=es_motor,
        )

        # ── Cálculo de conductor de tierra (NOM-001 250-122) ─────
        mat_tierra = "Cu" if "Cu" in tierra_material else "Al"
        if tierra_auto:
            tierra_calibre = calibre_tierra(ocpd_A, mat_tierra)
        else:
            tierra_calibre = tierra_manual or "—"

        # ── Cálculo de conduit (si toggle activo) ────────────────
        conduit_inline_result = None
        if incluir_conduit_calc and tipo_tubo_inline:
            todos_tubos = calcular_conduit_fill(
                conductor_verificado, num_conds_inline, tipo_tubo_inline
            )
            conduit_inline_result = {
                "params": {
                    "calibre": conductor_verificado,
                    "num_conds": num_conds_inline,
                    "tipo_tubo": tipo_tubo_inline,
                    "aislamiento": tipo_aisl_inline,
                },
                "resultados": todos_tubos,
                "resultados_todos": todos_tubos,
            }
            recom = next(
                (x for x in todos_tubos if x["cumple"]),
                todos_tubos[-1] if todos_tubos else None,
            )
            conduit_inline_result["recomendada"] = recom

        # Protección con formato comercial (NxA)
        proteccion_fmt = formato_proteccion(ocpd_A, config)

        st.session_state["resultado_conductor"] = {
            "corriente": corriente,
            "corriente_diseño": corriente_diseño,
            "corriente_diseño_cond": corriente_diseño_cond,
            "factor_demanda": factor_demanda,
            "cf": cf,
            "conductor": conductor_verificado,
            "conductor_min": conductor_min,
            "cdt": cdt_real,
            "cdt_max": cdt_max,
            "potencia": potencia,
            "longitud": longitud,
            "voltaje": voltaje_nominal,
            "configuracion": config_label,
            "config_code": config,
            "metodo": metodo_label,
            "metodo_code": metodo,
            "material": mat,
            "fp": fp,
            "sen_phi": sen_phi,
            "ampacity_corr": ampacity_corr,
            "proteccion_A": ocpd_A,
            "proteccion_status": ocpd_status,
            "proteccion_fmt": proteccion_fmt,
            "polos": POLOS_POR_CONFIG.get(config, 1),
            "seleccion_manual": False,
            "canalizacion": canalizacion,
            "canalizacion_label": canal_label,
            "temp_term": temp_term_efectiva,
            "temp_term_auto": (temp_term_override is None),
            "temp_aislamiento": temp_aislamiento,
            "es_motor": es_motor,
            "tierra_material": mat_tierra,
            "tierra_calibre": tierra_calibre,
            "tierra_auto": tierra_auto,
            "conduit_inline": conduit_inline_result,
        }
        st.session_state["tabla_conductor"] = df_tabla
        st.success(
            f"✅ Cálculo completado · Conductor {conductor_verificado} AWG {mat} · "
            f"Protección {proteccion_fmt} A · Tierra {tierra_calibre} AWG {mat_tierra} "
            f"desnudo"
        )

    # ── Render persistente de resultados (sobrevive al rerun) ──
    if "resultado_conductor" in st.session_state:
        r = st.session_state["resultado_conductor"]
        cf_r       = r.get("cf", 1)
        corriente  = r.get("corriente", 0)
        corriente_diseño = r.get("corriente_diseño", 0)
        corriente_diseño_cond = r.get("corriente_diseño_cond", 0)
        cdt_real   = r.get("cdt", 0)
        cdt_max_r  = r.get("cdt_max", 3.0)
        ampacity_corr = r.get("ampacity_corr", 0)
        ocpd_A     = r.get("proteccion_A", "—")
        ocpd_status = r.get("proteccion_status", "—")
        conductor_min = r.get("conductor_min", "—")
        conductor_verificado = r.get("conductor", "—")
        mat = r.get("material", "Cu")

        c1, c2, c3, c4, c5 = st.columns(5)
        with c1:
            subtexto_i = f"A real · diseño {corriente_diseño:.1f} A"
            if cf_r > 1:
                subtexto_i += f" · {corriente_diseño_cond:.1f} A/cond"
            st.markdown(f"""<div class="metric-card">
                <div class="metric-label">Corriente de carga</div>
                <div class="metric-value">{corriente:.1f}</div>
                <div class="metric-unit">{subtexto_i}</div></div>""", unsafe_allow_html=True)
        with c2:
            st.markdown(f"""<div class="metric-card">
                <div class="metric-label">Calibre mín. (ampacidad)</div>
                <div class="metric-value">{conductor_min}</div>
                <div class="metric-unit">AWG / kcmil</div></div>""", unsafe_allow_html=True)
        with c3:
            st.markdown(f"""<div class="metric-card">
                <div class="metric-label">Calibre seleccionado</div>
                <div class="metric-value">{conductor_verificado}</div>
                <div class="metric-unit">AWG / kcmil</div></div>""", unsafe_allow_html=True)
        with c4:
            color_cdt = "#10b981" if cdt_real <= cdt_max_r else "#ef4444"
            st.markdown(f"""<div class="metric-card">
                <div class="metric-label">C.d.T. real</div>
                <div class="metric-value" style="color:{color_cdt}">{cdt_real:.2f}</div>
                <div class="metric-unit">% (máx {cdt_max_r}%)</div></div>""", unsafe_allow_html=True)
        with c5:
            color_ocpd = {
                "OK": "#34c759", "REVISAR": "#ff9f0a", "INSUFICIENTE": "#ff3b30"
            }.get(ocpd_status, "#86868b")
            ocpd_fmt = r.get("proteccion_fmt", f"{r.get('polos', 1)}x{ocpd_A}")
            st.markdown(f"""<div class="metric-card">
                <div class="metric-label">Protección</div>
                <div class="metric-value" style="color:{color_ocpd}">{ocpd_fmt}</div>
                <div class="metric-unit">{ocpd_A} A · {ocpd_status}</div></div>""", unsafe_allow_html=True)

        # ── Segunda fila de métricas: tierra y conduit ─────────
        tierra_cal = r.get("tierra_calibre", "—")
        tierra_mat = r.get("tierra_material", "Cu")
        conduit_data = r.get("conduit_inline")

        if tierra_cal != "—" or conduit_data:
            ct1, ct2 = st.columns(2)
            with ct1:
                tierra_origen = "auto (Tabla 250-122)" if r.get("tierra_auto") else "manual"
                st.markdown(f"""<div class="metric-card">
                    <div class="metric-label">Conductor de tierra (desnudo)</div>
                    <div class="metric-value">{tierra_cal}</div>
                    <div class="metric-unit">AWG {tierra_mat} · {tierra_origen}</div></div>""",
                    unsafe_allow_html=True)
            with ct2:
                if conduit_data:
                    rec = conduit_data["recomendada"]
                    cumple_t = rec.get("cumple", False)
                    color_t = "#34c759" if cumple_t else "#ff9f0a"
                    st.markdown(f"""<div class="metric-card">
                        <div class="metric-label">Tubería recomendada</div>
                        <div class="metric-value" style="color:{color_t}">{rec['tubo']}</div>
                        <div class="metric-unit">Relleno {rec['fill_pct']:.1f}% (máx 40%)</div></div>""",
                        unsafe_allow_html=True)
                else:
                    st.markdown(f"""<div class="metric-card">
                        <div class="metric-label">Tubería</div>
                        <div class="metric-value" style="color:#86868b">—</div>
                        <div class="metric-unit">No calculada en este circuito</div></div>""",
                        unsafe_allow_html=True)

        st.markdown("")
        cf_label = f"{cf_r} × " if cf_r > 1 else ""
        if cdt_real <= cdt_max_r:
            st.markdown(f"""<div class="resultado-ok">
                ✅ CUMPLE — {cf_label}Conductor {conductor_verificado} AWG {mat}
                · C.d.T. = {cdt_real:.2f}% ≤ {cdt_max_r}%
                · Amp. corregida: {ampacity_corr:.1f} A ≥ {corriente_diseño_cond:.1f} A/cond
                · Método: {r.get('metodo', '').split('(')[0].strip()} · {r.get('configuracion', '')}
            </div>""", unsafe_allow_html=True)
        else:
            st.markdown(f"""<div class="resultado-error">
                ❌ NO CUMPLE — C.d.T. = {cdt_real:.2f}% supera {cdt_max_r}% máximo permitido
            </div>""", unsafe_allow_html=True)

        st.caption(
            f"📐 Tabla 9 NOM · canalización **{r.get('canalizacion_label', '—').split('(')[0].strip()}** · "
            f"terminales **{r.get('temp_term', '—')}°C** · aislamiento **{r.get('temp_aislamiento', 75)}°C**"
        )

        if "tabla_conductor" in st.session_state:
            st.markdown('<div class="section-header">Tabla de verificación de conductores candidatos</div>', unsafe_allow_html=True)
            st.dataframe(st.session_state["tabla_conductor"], use_container_width=True, hide_index=True)

    # ──────────────────────────────────────────
    # SELECCIÓN MANUAL DE CONDUCTOR
    # ──────────────────────────────────────────
    if "resultado_conductor" in st.session_state:
        r = st.session_state["resultado_conductor"]
        mat_m = r["material"]
        taisl_m = r.get("temp_aislamiento", 75)
        # Filtrar a calibres con ampacidad > 0 para este material/aislamiento
        calibres_validos = [
            c for c in ORDEN_CALIBRES
            if ampacidad_base(c, mat_m, taisl_m) > 0
        ] or ORDEN_CALIBRES

        with st.expander("🔧 Selección manual de conductor", expanded=False):
            etiqueta_auto = (
                f"Conductor automático: **{r['conductor']} AWG {mat_m}** "
                f"· C.d.T. = {r['cdt']:.2f}%"
                + (" *(manual aplicado)*" if r.get("seleccion_manual") else "")
            )
            st.info(etiqueta_auto)

            default_idx = (
                calibres_validos.index(r["conductor"])
                if r["conductor"] in calibres_validos else 0
            )
            # Key incluye material para resetear el widget al cambiar material
            calibre_manual = st.selectbox(
                f"Elige el calibre ({mat_m})",
                calibres_validos,
                index=default_idx,
                key=f"calibre_manual_{mat_m}_{taisl_m}",
            )

            cf_r = r.get("cf", 1)
            canal_m = r.get("canalizacion", "AC")
            tterm_m = r.get("temp_term", 75)
            L_km_m = r["longitud"] / 1000
            calc_ok = True
            try:
                cdt_m = calcular_cdt_calibre(
                    calibre_manual, r["corriente"], L_km_m, mat_m,
                    r["fp"], r["voltaje"], r["config_code"], r["metodo_code"],
                    cf=cf_r, canalizacion=canal_m, temp_term=tterm_m,
                )
                amp_m = ampacidad_base(calibre_manual, mat_m, taisl_m) * fc_temp * fc_agrup_val
            except ValueError as ex:
                st.error(f"⚠️ {ex}")
                calc_ok = False
                cdt_m, amp_m = 0.0, 0.0
            I_diseño_cond_m = r["corriente"] * r.get("factor_demanda", 1.25) / cf_r

            if calc_ok:
                cm1, cm2, cm3 = st.columns(3)
                with cm1:
                    st.markdown(f"""<div class="metric-card">
                        <div class="metric-label">Calibre manual</div>
                        <div class="metric-value">{calibre_manual}</div>
                        <div class="metric-unit">AWG / kcmil · {mat_m}</div></div>""",
                        unsafe_allow_html=True)
                with cm2:
                    color_amp_m = "#10b981" if amp_m >= I_diseño_cond_m else "#ef4444"
                    req_label = f"{I_diseño_cond_m:.1f} A/cond" + (f" (CF={cf_r})" if cf_r > 1 else "")
                    st.markdown(f"""<div class="metric-card">
                        <div class="metric-label">Ampacidad corregida</div>
                        <div class="metric-value" style="color:{color_amp_m}">{amp_m:.1f}</div>
                        <div class="metric-unit">A req: {req_label}</div></div>""",
                        unsafe_allow_html=True)
                with cm3:
                    color_cdt_m = "#10b981" if cdt_m <= r["cdt_max"] else "#ef4444"
                    st.markdown(f"""<div class="metric-card">
                        <div class="metric-label">C.d.T. manual</div>
                        <div class="metric-value" style="color:{color_cdt_m}">{cdt_m:.2f}</div>
                        <div class="metric-unit">% (máx {r['cdt_max']}%)</div></div>""",
                        unsafe_allow_html=True)

                st.markdown("")
                cumple_m = cdt_m <= r["cdt_max"] and amp_m >= I_diseño_cond_m
                if cumple_m:
                    st.markdown(f"""<div class="resultado-ok">
                        ✅ {calibre_manual} AWG {mat_m} cumple ambos criterios —
                        Amp. {amp_m:.1f} A ≥ {I_diseño_cond_m:.1f} A/cond · C.d.T. {cdt_m:.2f}% ≤ {r['cdt_max']}%
                    </div>""", unsafe_allow_html=True)
                else:
                    problemas = []
                    if amp_m < I_diseño_cond_m:
                        problemas.append(f"ampacidad insuficiente ({amp_m:.1f} A < {I_diseño_cond_m:.1f} A/cond)")
                    if cdt_m > r["cdt_max"]:
                        problemas.append(f"C.d.T. excede límite ({cdt_m:.2f}% > {r['cdt_max']}%)")
                    st.markdown(f"""<div class="resultado-warn">
                        ⚠️ {calibre_manual} AWG {mat_m}: {" · ".join(problemas)}
                    </div>""", unsafe_allow_html=True)

                # ── Aviso visible si la selección manual NO está aplicada ──
                manual_pendiente = (calibre_manual != r.get("conductor"))
                if manual_pendiente:
                    st.markdown(f"""<div class="resultado-warn" style="
                        background: rgba(255,159,10,0.18); color:#8a5a00;
                        border: 1.5px solid rgba(255,159,10,0.5); font-size: 0.95rem;">
                        <b>⚠ Selección manual NO aplicada todavía.</b><br>
                        Ahora mismo el reporte usará <b>{r.get('conductor','—')} AWG</b>
                        (calculado automáticamente). Para usar
                        <b>{calibre_manual} AWG</b>, presiona el botón de abajo.
                    </div>""", unsafe_allow_html=True)

                btn_label = (
                    f"✔️ Aplicar selección manual ({calibre_manual} AWG)"
                    if manual_pendiente
                    else "✔️ Selección manual aplicada"
                )
                btn_disabled = not manual_pendiente
                if st.button(btn_label, key="btn_aplicar_manual",
                              type="primary" if manual_pendiente else "secondary",
                              disabled=btn_disabled, use_container_width=True):
                    _aplicar_seleccion_manual(
                        calibre_manual, cdt_m, amp_m, I_diseño_cond_m,
                        r.get("config_code", "mono2h"),
                        r.get("tierra_material", "Cu"),
                        r.get("tierra_auto", True),
                    )
                    st.success(f"✅ Conductor actualizado a **{calibre_manual} AWG {mat_m}** · C.d.T. = {cdt_m:.2f}%")
                    st.rerun()

    # ──────────────────────────────────────────
    # AGREGAR AL PROYECTO  (UX mejorada)
    # ──────────────────────────────────────────
    if "resultado_conductor" in st.session_state:
        r = st.session_state["resultado_conductor"]
        st.markdown('<div class="section-header">📌 Guardar este circuito en el proyecto</div>', unsafe_allow_html=True)

        # ── Detección de selección manual pendiente ──
        # El selectbox del expander manual está guardado en session_state con
        # key f"calibre_manual_{mat}_{aisl}". Si su valor difiere del actual,
        # avisamos al usuario que la selección manual NO está aplicada.
        mat_check = r.get("material", "Cu")
        taisl_check = r.get("temp_aislamiento", 75)
        key_manual = f"calibre_manual_{mat_check}_{taisl_check}"
        calibre_en_dropdown = st.session_state.get(key_manual)
        seleccion_pendiente = (
            calibre_en_dropdown is not None
            and calibre_en_dropdown != r.get("conductor")
        )
        if seleccion_pendiente:
            st.markdown(f"""<div class="resultado-warn" style="
                background: rgba(255,59,48,0.13); color:#a31621;
                border: 1.5px solid rgba(255,59,48,0.45); font-size: 0.95rem;">
                <b>⚠ ATENCIÓN — Tienes una selección manual SIN aplicar:</b><br>
                En el panel "🔧 Selección manual de conductor" elegiste
                <b>{calibre_en_dropdown} AWG</b>, pero esa selección
                <b>NO se ha guardado</b>. Si agregas el circuito ahora se
                guardará con el calibre automático: <b>{r.get('conductor')} AWG</b>.<br>
                <i>Sube al panel manual y presiona el botón ✔️ para aplicar tu selección.</i>
            </div>""", unsafe_allow_html=True)

        # Tarjeta-resumen de lo que se guardará
        rec_t = r.get("tierra_calibre", "—")
        ci = r.get("conduit_inline")
        conduit_resumen = (
            f"{ci['recomendada']['tubo']} · {ci['recomendada']['fill_pct']:.1f}% relleno"
            if ci else "— no calculada —"
        )
        st.markdown(
            f"""<div class="resultado-ok" style="background:rgba(0,113,227,0.08);
                color:#1d1d1f;border:1px solid rgba(0,113,227,0.25);">
                <b>Resumen del cálculo actual:</b><br>
                ⚡ {r.get('potencia',0):,} W · {r.get('voltaje','—')} V · {r.get('configuracion','—')}<br>
                🧵 Conductor: <b>{r.get('cf',1)}× {r.get('conductor','—')} AWG {r.get('material','—')}</b>
                · C.d.T. {r.get('cdt',0):.2f}%
                · Amp. {r.get('ampacity_corr',0):.1f} A<br>
                🛡️ Protección: <b>{r.get('proteccion_fmt','—')} A</b>
                · Tierra: <b>{rec_t} AWG {r.get('tierra_material','Cu')}</b> desnudo<br>
                🔧 Tubería: <b>{conduit_resumen}</b>
            </div>""",
            unsafe_allow_html=True,
        )

        # Diálogo de guardado
        n_circ = len(st.session_state["circuitos"]) + 1
        col_add1, col_add2 = st.columns([3, 1])
        with col_add1:
            nombre_circ = st.text_input(
                "Nombre / Identificación del circuito",
                value=f"Circuito {n_circ:02d}",
                key="nombre_circ_input",
                help="Ej: 'Iluminación PB', 'Tablero A-1', 'Alimentador motor #2'",
                placeholder="Ej. Iluminación Planta Baja",
            )
        with col_add2:
            st.markdown("<br>", unsafe_allow_html=True)
            agregar_clicked = st.button(
                "➕ Agregar al proyecto",
                key="btn_add_circ", type="primary",
                use_container_width=True,
            )

        if agregar_clicked:
            if not nombre_circ.strip():
                st.error("⚠️ Asigna un nombre al circuito antes de agregarlo.")
            else:
                # Conduit: usa el inline si existe, si no, fallback al de Tab 2
                conduit_info = None
                if ci:
                    conduit_info = {
                        "params": ci["params"],
                        "recomendada": ci["recomendada"],
                        "resultados_todos": ci.get("resultados_todos", ci.get("resultados", [])),
                    }
                elif "resultado_conduit" in st.session_state:
                    rc_c = st.session_state["resultado_conduit"]
                    recom = next((x for x in rc_c if x["cumple"]), rc_c[-1])
                    conduit_info = {
                        "params": st.session_state.get("conduit_params"),
                        "recomendada": recom,
                        "resultados_todos": rc_c,
                    }
                tabla_df = st.session_state.get("tabla_conductor")
                nuevo = {
                    "id": len(st.session_state["circuitos"]) + 1,
                    "nombre": nombre_circ.strip(),
                    **r,
                    "conduit_info": conduit_info,
                    "tabla_conductor_df": tabla_df,
                }
                st.session_state["circuitos"].append(nuevo)
                total = len(st.session_state["circuitos"])
                st.success(
                    f"✅ **{nombre_circ}** se agregó al proyecto. "
                    f"Ahora tienes **{total} circuito(s)**. "
                    f"Ve a la pestaña **📊 Resumen General** para revisar el tablero."
                )

        # Lista de circuitos ya guardados (preview claro)
        if st.session_state["circuitos"]:
            with st.expander(
                f"📂 Circuitos en el proyecto ({len(st.session_state['circuitos'])})",
                expanded=False,
            ):
                for c in st.session_state["circuitos"]:
                    cdt_v = c.get("cdt", 0) or 0
                    cdt_mx = c.get("cdt_max", 3.0)
                    icon = "✅" if cdt_v <= cdt_mx else "❌"
                    st.markdown(
                        f"- {icon} **#{c.get('id','?')} {c.get('nombre','')}** · "
                        f"{c.get('potencia','—')} W · {c.get('cf',1)}×{c.get('conductor','—')} {c.get('material','')} · "
                        f"OCPD {c.get('proteccion_fmt', str(c.get('proteccion_A','—'))+' A')}"
                    )

# ══════════════════════════════════════════════
# TAB 2 — CONDUIT FILL
# ══════════════════════════════════════════════
with tab2:
    st.markdown('<div class="section-header">Cálculo de Relleno de Tubería (Conduit Fill)</div>', unsafe_allow_html=True)
    st.caption("Según NOM-001-SEDE-2012 · Tabla 1 del Capítulo 3 (máx. 40% para 3+ conductores)")

    col1, col2 = st.columns(2)
    with col1:
        tipo_tubo = st.selectbox("Tipo de tubería", ["EMT (Conduit metálico ligero)", "IMC (Conduit metálico intermedio)", "RGS (Rígido galvanizado)"])
        num_conductores = st.number_input("Número total de conductores", min_value=1, max_value=20, value=3)
    with col2:
        calibre_conduit = st.selectbox("Calibre del conductor en la tubería", list(TABLA_CONDUCTORES.keys()))
        tipo_aislamiento = st.selectbox("Tipo de aislamiento", ["THW (75°C)", "THHW (75/90°C)", "THWN-2 (90°C)", "XHHW-2 (90°C)"])

    if st.button("🔍 Calcular tubería", key="btn_conduit", type="primary"):
        resultado_conduit = calcular_conduit_fill(calibre_conduit, num_conductores, tipo_tubo)

        recomendada = next((r for r in resultado_conduit if r.get("recomendada")), None)
        if recomendada:
            st.markdown(f"""<div class="resultado-ok">
                ★ Tubería recomendada: <b>{recomendada['tubo']}</b> ·
                {num_conductores} conductores {calibre_conduit} AWG ·
                Relleno: <b>{recomendada['fill_pct']:.1f}%</b> ≤ {recomendada['fill_max']}% (máx).
            </div>""", unsafe_allow_html=True)
        else:
            st.markdown(f"""<div class="resultado-error">
                ✗ Ningún tamaño cumple. Considera tubería de mayor familia o reducir N conductores.
            </div>""", unsafe_allow_html=True)

        st.markdown('<div class="section-header">Tabla completa de medidas</div>', unsafe_allow_html=True)
        df_conduit = pd.DataFrame([
            {
                "Tubería": r["tubo"],
                "Ø nominal": r.get("diametro", "—"),
                "Área int. (mm²)": r["area_tubo"],
                "Área conds. (mm²)": r["area_conds"],
                "Relleno (%)": r["fill_pct"],
                "Máx. (%)": r["fill_max"],
                "Cumple": "✓" if r["cumple"] else "✗",
                "Selección": "★" if r.get("recomendada") else "",
            }
            for r in resultado_conduit
        ])
        st.dataframe(df_conduit, use_container_width=True, hide_index=True)

        st.session_state["resultado_conduit"] = resultado_conduit
        st.session_state["conduit_params"] = {
            "calibre": calibre_conduit,
            "num_conds": num_conductores,
            "tipo_tubo": tipo_tubo,
            "aislamiento": tipo_aislamiento,
        }

# ══════════════════════════════════════════════
# TAB 3 — PANEL DE CIRCUITOS
# ══════════════════════════════════════════════
with tab3:
    st.markdown('<div class="section-header">Panel de Circuitos del Proyecto</div>', unsafe_allow_html=True)

    circuitos = st.session_state["circuitos"]

    if not circuitos:
        st.info("⚡ No hay circuitos en el proyecto. Ve a **Caída de Tensión**, calcula y usa **Agregar al proyecto**.")
    else:
        # ── Métricas resumen ────────────────────────────────
        total_kw = sum(c.get("potencia", 0) for c in circuitos) / 1000

        def _cumple(c):
            cdt = c.get("cdt")
            cdt_max = c.get("cdt_max", 3.0)
            return (cdt is not None) and (cdt <= cdt_max)

        todos_cumplen = all(_cumple(c) for c in circuitos)
        sm1, sm2, sm3, sm4 = st.columns(4)
        with sm1:
            st.metric("Circuitos", len(circuitos))
        with sm2:
            st.metric("Carga total", f"{total_kw:.2f} kW")
        with sm3:
            st.metric("Estado general", "✅ Cumplen" if todos_cumplen else "⚠️ Revisar")
        with sm4:
            st.metric("Temp. ambiente", f"{temp_ambiente} °C")

        st.markdown("")

        # ── Tablero de circuitos ─────────────────────────────
        st.markdown("**Tablero de circuitos**")
        filas_tb = []
        for i, c in enumerate(circuitos):
            cf = c.get("cf", 1)
            cond = c.get("conductor", "—")
            cond_str = f"{cf}×{cond}" if cf > 1 else cond
            rec = (c.get("conduit_info") or {}).get("recomendada") or {}
            tubo_str = rec.get("tubo", "—")
            ocpd_st = c.get("proteccion_status", "")
            ocpd_ico = "✅" if ocpd_st == "OK" else ("⚠️" if ocpd_st == "REVISAR" else "—")
            cdt = c.get("cdt")
            cdt_max = c.get("cdt_max", 3.0)
            corriente = c.get("corriente", 0) or 0
            corriente_d = c.get("corriente_diseño", 0) or 0
            amp_corr = c.get("ampacity_corr", 0) or 0
            filas_tb.append({
                "No.": c.get("id", i + 1),
                "Circuito": c.get("nombre", f"Circuito {i+1}"),
                "P (W)": c.get("potencia", "—"),
                "V (V)": c.get("voltaje", "—"),
                "Config.": (c.get("configuracion", "—") or "—").split("(")[0].strip(),
                "I (A)": f"{corriente:.1f}",
                "I diseño (A)": f"{corriente_d:.1f}",
                "Conductor": f"{cond_str} AWG {c.get('material', '—')}",
                "Amp (A)": f"{amp_corr:.1f}",
                "C.d.T. %": f"{cdt:.2f}" if cdt is not None else "—",
                "Máx %": cdt_max,
                "Protección": f"{c.get('proteccion_fmt', c.get('proteccion_A','—'))} {ocpd_ico}",
                "Tierra": f"{c.get('tierra_calibre','—')} {c.get('tierra_material','')}".strip(),
                "Tubería": tubo_str,
                "Estado": "✅" if _cumple(c) else "❌",
            })
        st.dataframe(pd.DataFrame(filas_tb), use_container_width=True, hide_index=True)

        # ── Acciones: Excel / JSON / Limpiar ────────────────
        st.markdown("")
        ca1, ca2, ca3, ca4 = st.columns(4)

        datos_proy_export = {
            "proyecto": proyecto, "cliente": cliente, "ubicacion": ubicacion,
            "responsable": responsable, "cedula": cedula, "fecha": str(fecha_proj),
            "norma": norma, "sistema": tipo_sistema, "temp_ambiente": temp_ambiente,
        }

        with ca1:
            try:
                excel_bytes = crear_excel_circuitos(circuitos)
                st.download_button(
                    "📊 Exportar Excel",
                    excel_bytes,
                    f"Circuitos_{proyecto.replace(' ','_')}.xlsx",
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
            except Exception as e:
                st.error(f"Excel: {e}")

        with ca2:
            json_bytes = exportar_sesion_json(circuitos, datos_proy_export).encode("utf-8")
            st.download_button(
                "💾 Guardar sesión",
                json_bytes,
                f"sesion_{proyecto.replace(' ','_')}.json",
                "application/json",
            )

        with ca3:
            if st.button("🗑️ Limpiar todos", key="btn_limpiar_todos"):
                st.session_state["circuitos"] = []
                st.rerun()

        with ca4:
            if len(circuitos) > 1:
                opciones = [f"#{c.get('id', i+1)} — {c.get('nombre', '')}"
                            for i, c in enumerate(circuitos)]
                idx_borrar = st.selectbox(
                    "Eliminar circuito",
                    range(len(opciones)),
                    format_func=lambda i: opciones[i],
                    key="sel_borrar",
                    label_visibility="collapsed",
                )
                if st.button("🗑️ Eliminar", key="btn_borrar_uno"):
                    st.session_state["circuitos"].pop(idx_borrar)
                    for j, c in enumerate(st.session_state["circuitos"]):
                        c["id"] = j + 1
                    st.rerun()

    # ── Importar sesión ──────────────────────────────────────
    st.markdown("")
    with st.expander("📂 Importar sesión guardada", expanded=False):
        uploaded = st.file_uploader("Selecciona un archivo .json", type="json", key="json_import")
        if uploaded:
            try:
                datos_imp = json.loads(uploaded.read().decode("utf-8"))
                circs_imp = datos_imp.get("circuitos", [])
                proy_imp = datos_imp.get("proyecto", {})
                st.info(
                    f"**{len(circs_imp)} circuito(s)** en el archivo"
                    + (f" — Proyecto: {proy_imp.get('proyecto','?')}" if proy_imp else "")
                )
                ci1, ci2 = st.columns(2)
                with ci1:
                    if st.button("Reemplazar circuitos actuales", key="btn_imp_replace"):
                        st.session_state["circuitos"] = circs_imp
                        st.success(f"{len(circs_imp)} circuito(s) cargados")
                        st.rerun()
                with ci2:
                    if st.button("Agregar a los existentes", key="btn_imp_append"):
                        base = len(st.session_state["circuitos"])
                        for j, c in enumerate(circs_imp):
                            c["id"] = base + j + 1
                        st.session_state["circuitos"] += circs_imp
                        st.success(f"{len(circs_imp)} circuito(s) añadidos")
                        st.rerun()
            except Exception as e:
                st.error(f"Error al leer el archivo: {e}")

# ══════════════════════════════════════════════
# TAB 4 — REPORTE PDF
# ══════════════════════════════════════════════
with tab4:
    st.markdown('<div class="section-header">Generación de Reporte PDF Profesional</div>', unsafe_allow_html=True)

    sin_resultado = "resultado_conductor" not in st.session_state
    sin_circuitos = not st.session_state["circuitos"]

    if sin_resultado and sin_circuitos:
        st.warning("⚠️ Calcula al menos un circuito antes de generar el reporte.")
    else:
        datos_proyecto = {
            "proyecto": proyecto, "cliente": cliente, "ubicacion": ubicacion,
            "responsable": responsable, "cedula": cedula, "fecha": str(fecha_proj),
            "norma": norma, "sistema": tipo_sistema, "temp_ambiente": temp_ambiente,
        }

        # ── Modo de reporte ─────────────────────────────────
        modos_disponibles = []
        if not sin_resultado:
            modos_disponibles.append("Circuito actual (memoria detallada)")
        if not sin_circuitos:
            modos_disponibles.append(f"Proyecto completo ({len(st.session_state['circuitos'])} circuito(s))")

        modo = st.radio("Modo de reporte", modos_disponibles, horizontal=True)

        if "Proyecto completo" in modo:
            st.markdown("**El reporte de proyecto incluirá:**")
            c1, c2 = st.columns(2)
            with c1:
                st.markdown("- 📋 Portada con datos del proyecto\n- 📐 Criterios de diseño\n- 📊 Tabla resumen de todos los circuitos")
            with c2:
                st.markdown("- ⚡ Memoria condensada por circuito\n- 🔧 Conduit (si está vinculado)\n- ✅ Conclusiones y firma")

            formato_lbl = st.radio(
                "Formato de salida",
                ["📄 PDF (no editable, profesional)",
                 "📝 Word .docx (editable, para complementar)"],
                horizontal=True, key="fmt_multi",
            )
            es_word = "Word" in formato_lbl

            if st.button(
                f"{'📝' if es_word else '📄'} Generar Reporte de Proyecto",
                type="primary", key="btn_doc_multi",
            ):
                with st.spinner(f"Generando reporte ({'Word' if es_word else 'PDF'})..."):
                    rc_default = st.session_state.get(
                        "resultado_conductor", st.session_state["circuitos"][0]
                    )
                    if es_word:
                        doc_bytes = generar_reporte_docx(
                            datos_proyecto, rc_default,
                            circuitos=st.session_state["circuitos"],
                        )
                        ext, mime = "docx", (
                            "application/vnd.openxmlformats-officedocument."
                            "wordprocessingml.document"
                        )
                    else:
                        doc_bytes = generar_reporte_pdf(
                            datos_proyecto, rc_default,
                            circuitos=st.session_state["circuitos"],
                        )
                        ext, mime = "pdf", "application/pdf"
                nombre = f"MemoriaTecnica_{proyecto.replace(' ','_')}_Proyecto.{ext}"
                st.download_button(
                    f"⬇️ Descargar {nombre}",
                    doc_bytes, nombre, mime, type="primary",
                )
                st.success(f"✅ Reporte generado: **{nombre}**")

        else:
            st.markdown("**El reporte del circuito actual incluirá:**")
            c1, c2 = st.columns(2)
            with c1:
                st.markdown("- 📋 Portada con datos del proyecto\n- 📐 Criterios de diseño y normativa\n- ⚡ Memoria de cálculo paso a paso\n- 📊 Tabla de verificación de conductores")
            with c2:
                st.markdown("- 🔧 Cálculo de relleno de tubería\n- ✅ Conclusiones y cumplimiento\n- 📝 Firma del responsable técnico")

            incluir_tabla = st.checkbox("Incluir tabla completa de conductores", value=True)
            incluir_conduit = st.checkbox("Incluir cálculo de tubería", value="resultado_conduit" in st.session_state)

            formato_lbl = st.radio(
                "Formato de salida",
                ["📄 PDF (no editable, profesional)",
                 "📝 Word .docx (editable, para complementar)"],
                horizontal=True, key="fmt_single",
            )
            es_word = "Word" in formato_lbl

            if st.button(
                f"{'📝' if es_word else '📄'} Generar Reporte",
                type="primary", key="btn_doc_single",
            ):
                tabla_conductor = st.session_state.get("tabla_conductor") if incluir_tabla else None
                resultado_conduit = st.session_state.get("resultado_conduit") if incluir_conduit else None
                conduit_params = st.session_state.get("conduit_params") if incluir_conduit else None

                with st.spinner(f"Generando reporte ({'Word' if es_word else 'PDF'})..."):
                    if es_word:
                        doc_bytes = generar_reporte_docx(
                            datos_proyecto,
                            st.session_state["resultado_conductor"],
                            tabla_conductor, resultado_conduit, conduit_params,
                        )
                        ext, mime = "docx", (
                            "application/vnd.openxmlformats-officedocument."
                            "wordprocessingml.document"
                        )
                    else:
                        doc_bytes = generar_reporte_pdf(
                            datos_proyecto,
                            st.session_state["resultado_conductor"],
                            tabla_conductor, resultado_conduit, conduit_params,
                        )
                        ext, mime = "pdf", "application/pdf"
                nombre = f"MemoriaTecnica_{proyecto.replace(' ','_')}.{ext}"
                st.download_button(
                    f"⬇️ Descargar {nombre}",
                    doc_bytes, nombre, mime, type="primary",
                )
                st.success(f"✅ Reporte generado: **{nombre}**")

# ══════════════════════════════════════════════
# TAB 5 — GESTIÓN DE USUARIOS  (solo admin)
# ══════════════════════════════════════════════
if tab_users is not None:
    with tab_users:
        st.markdown('<div class="section-header">👥 Gestión de Usuarios</div>',
                    unsafe_allow_html=True)
        st.caption(
            "Como administrador puedes crear, editar y eliminar usuarios. "
            "Los datos se almacenan en `data/users.json` con hash SHA-256 + salt."
        )

        usuarios = auth.listar_usuarios()
        df_users = pd.DataFrame([
            {
                "Usuario": u["user"],
                "Nombre": u["nombre"],
                "Rol": u["rol"],
                "Contraseña por defecto": "⚠️ sí" if u["default_pwd"] else "—",
            }
            for u in usuarios
        ])
        st.dataframe(df_users, use_container_width=True, hide_index=True)

        st.markdown('<div class="section-header">Crear nuevo usuario</div>',
                    unsafe_allow_html=True)
        with st.form("form_nuevo_usuario", clear_on_submit=True):
            cu1, cu2 = st.columns(2)
            with cu1:
                nuevo_user = st.text_input("Usuario (login)", placeholder="jperez")
                nuevo_nombre = st.text_input("Nombre completo", placeholder="Juan Pérez")
            with cu2:
                nuevo_pwd = st.text_input("Contraseña", type="password",
                                           placeholder="mínimo 4 caracteres")
                nuevo_rol = st.selectbox("Rol", ["usuario", "admin"])
            if st.form_submit_button("➕ Crear usuario", type="primary"):
                ok, msg = auth.agregar_usuario(
                    nuevo_user, nuevo_pwd, nuevo_rol, nuevo_nombre
                )
                if ok:
                    st.success(msg)
                    st.rerun()
                else:
                    st.error(msg)

        st.markdown('<div class="section-header">Cambiar contraseña</div>',
                    unsafe_allow_html=True)
        with st.form("form_cambiar_pwd", clear_on_submit=True):
            cp1, cp2 = st.columns(2)
            with cp1:
                u_sel = st.selectbox(
                    "Usuario",
                    [u["user"] for u in usuarios],
                    index=([u["user"] for u in usuarios].index(_usr["user"])
                           if _usr["user"] in [u["user"] for u in usuarios] else 0),
                )
            with cp2:
                pwd_nueva = st.text_input("Nueva contraseña", type="password")
            if st.form_submit_button("🔑 Cambiar contraseña", type="primary"):
                ok, msg = auth.cambiar_password(u_sel, pwd_nueva)
                if ok:
                    st.success(msg)
                    if u_sel == _usr["user"]:
                        # actualizar flag en sesión actual
                        st.session_state["auth_user"]["default_pwd"] = False
                    st.rerun()
                else:
                    st.error(msg)

        st.markdown('<div class="section-header">Cambiar rol</div>',
                    unsafe_allow_html=True)
        with st.form("form_cambiar_rol", clear_on_submit=True):
            cr1, cr2 = st.columns(2)
            with cr1:
                u_rol_sel = st.selectbox(
                    "Usuario",
                    [u["user"] for u in usuarios],
                    key="sel_user_rol",
                )
            with cr2:
                nuevo_rol_sel = st.selectbox("Nuevo rol", ["usuario", "admin"],
                                              key="sel_rol_nuevo")
            if st.form_submit_button("Actualizar rol", type="primary"):
                ok, msg = auth.cambiar_rol(u_rol_sel, nuevo_rol_sel)
                (st.success if ok else st.error)(msg)
                if ok:
                    st.rerun()

        st.markdown('<div class="section-header">Eliminar usuario</div>',
                    unsafe_allow_html=True)
        eliminables = [u["user"] for u in usuarios if u["user"] != _usr["user"]]
        if eliminables:
            with st.form("form_eliminar_usuario", clear_on_submit=True):
                u_del = st.selectbox("Usuario a eliminar", eliminables)
                confirmar = st.checkbox(
                    f"Confirmo que quiero eliminar permanentemente al usuario seleccionado",
                )
                if st.form_submit_button("🗑️ Eliminar", type="secondary"):
                    if not confirmar:
                        st.warning("Marca la casilla de confirmación para continuar.")
                    else:
                        ok, msg = auth.eliminar_usuario(u_del)
                        (st.success if ok else st.error)(msg)
                        if ok:
                            st.rerun()
        else:
            st.info("No hay otros usuarios para eliminar.")
