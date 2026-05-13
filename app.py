# eSuit · Cálculo eléctrico profesional · NOM-001-SEDE-2012
# Repo: https://github.com/JADA2825/eSuit
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
    TABLA_CONDUIT,
    ORDEN_CALIBRES,
    CANALIZACIONES,
    AISLAMIENTOS_T,
    POLOS_POR_CONFIG,
)
from reporte import generar_reporte_pdf
from reporte_docx import generar_reporte_docx
from cuadro_cargas import (
    asignar_fases, resumen_balanceo, generar_excel_cuadro_cargas,
)
import auth

# ─────────────────────────────────────────────
# HELPERS — SESIÓN JSON Y EXPORTACIÓN EXCEL
# ─────────────────────────────────────────────
def _aplicar_seleccion_manual(calibre, snap, cdt_nuevo, amp_nuevo, I_diseño_cond,
                                config_code, tierra_mat, tierra_auto):
    """Guarda un override de calibre asociado a un snapshot de los inputs actuales.
       El override se mantiene mientras los inputs no cambien y se invalida solo
       cuando difieren. Esto evita que el cálculo en vivo de Tab 1 lo sobreescriba.
    """
    st.session_state["manual_override"] = {
        "calibre": calibre,
        "snap": snap,
        "cdt": cdt_nuevo,
        "amp": amp_nuevo,
        "I_d_cond": I_diseño_cond,
        "config_code": config_code,
        "tierra_mat": tierra_mat,
        "tierra_auto": tierra_auto,
    }
    # También sincronizar el dict actual para que la UI lo refleje inmediatamente
    rc = st.session_state.get("resultado_conductor")
    if rc:
        new_ocpd_A, new_ocpd_st = calcular_proteccion(I_diseño_cond, amp_nuevo)
        rc["conductor"]         = calibre
        rc["cdt"]               = cdt_nuevo
        rc["ampacity_corr"]     = amp_nuevo
        rc["proteccion_A"]      = new_ocpd_A
        rc["proteccion_status"] = new_ocpd_st
        rc["proteccion_fmt"]    = formato_proteccion(new_ocpd_A, config_code)
        if tierra_auto:
            rc["tierra_calibre"] = calibre_tierra(new_ocpd_A, tierra_mat)
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
def _get_theme():
    """Devuelve 'light', 'dark' o 'system' (default)."""
    return st.session_state.get("ui_theme", "system")


_THEME = _get_theme()

# Bloques CSS con las variables. Light y Dark son explícitos.
# Sistema usa @media (prefers-color-scheme: dark) sobre defaults claros.
_VARS_LIGHT = """
    --bg-app:        linear-gradient(180deg, #fafafa 0%, #f0f2f5 100%);
    --bg-blob-1:     rgba(0,113,227,0.07);
    --bg-blob-2:     rgba(255,159,10,0.06);
    --bg-blob-3:     rgba(52,199,89,0.05);
    --bg-card:       rgba(255,255,255,0.92);
    --bg-glass:      rgba(255,255,255,0.78);
    --bg-input:      rgba(255,255,255,0.96);
    --bg-formula:    rgba(241,245,249,0.92);
    --bg-tab-list:   rgba(255,255,255,0.7);
    --bg-tab-active: #ffffff;
    --bg-pill-recom: rgba(0,113,227,0.08);
    --text-main:     #1d1d1f;
    --text-sub:      #3a3a3c;
    --text-muted:    #6e6e73;
    --text-tabactive:#0071e3;
    --border-light:  rgba(0,0,0,0.08);
    --border-input:  rgba(0,0,0,0.18);
    --border-tab:    rgba(0,0,0,0.10);
    --shadow-card:   0 1px 2px rgba(0,0,0,0.04), 0 8px 24px rgba(0,0,0,0.05);
    --shadow-strong: 0 1px 2px rgba(0,0,0,0.05), 0 12px 30px rgba(0,0,0,0.06);
    --accent-blue:   #0071e3;
    --accent-blue-h: #0077ed;
    --accent-orange: #ff9500;   /* Naranja eléctrico iOS */
    --accent-orange-h:#ffaa1a;
    --ok-bg:    rgba(34,197,94,0.14);
    --ok-bg2:   rgba(220,252,231,0.92);
    --ok-bg3:   rgba(240,253,244,0.92);
    --ok-bd:    rgba(34,197,94,0.40);
    --ok-text:  #15803d;
    --ok-text2: #14532d;
    --warn-bg:  rgba(245,158,11,0.16);
    --warn-bg2: rgba(254,243,199,0.92);
    --warn-bg3: rgba(255,251,235,0.92);
    --warn-bd:  rgba(245,158,11,0.42);
    --warn-text: #b45309;
    --warn-text2:#78350f;
    --err-bg:   rgba(239,68,68,0.14);
    --err-bg2:  rgba(254,226,226,0.92);
    --err-bg3:  rgba(254,242,242,0.92);
    --err-bd:   rgba(239,68,68,0.40);
    --err-text: #b91c1c;
    --err-text2:#7f1d1d;
"""

_VARS_DARK = """
    --bg-app:        linear-gradient(180deg, #0a0a0c 0%, #14161b 100%);
    --bg-blob-1:     rgba(10,132,255,0.16);
    --bg-blob-2:     rgba(255,159,10,0.10);
    --bg-blob-3:     rgba(52,199,89,0.10);
    --bg-card:       rgba(36,38,46,0.85);
    --bg-glass:      rgba(36,38,46,0.72);
    --bg-input:      rgba(255,255,255,0.08);
    --bg-formula:    rgba(28,32,46,0.92);
    --bg-tab-list:   rgba(255,255,255,0.08);
    --bg-tab-active: rgba(255,255,255,0.16);
    --bg-pill-recom: rgba(10,132,255,0.20);
    --text-main:     #f5f5f7;
    --text-sub:      #e5e5e7;
    --text-muted:    #a1a1a6;
    --text-tabactive:#0a84ff;
    --border-light:  rgba(255,255,255,0.12);
    --border-input:  rgba(255,255,255,0.20);
    --border-tab:    rgba(255,255,255,0.14);
    --shadow-card:   0 1px 2px rgba(0,0,0,0.4), 0 8px 24px rgba(0,0,0,0.5);
    --shadow-strong: 0 1px 2px rgba(0,0,0,0.45), 0 12px 30px rgba(0,0,0,0.55);
    --accent-blue:   #0a84ff;
    --accent-blue-h: #1a8eff;
    --accent-orange: #ff9f0a;   /* Naranja eléctrico iOS dark */
    --accent-orange-h:#ffaf2a;
    --ok-bg:    rgba(48,209,88,0.22);
    --ok-bg2:   rgba(20,80,40,0.65);
    --ok-bg3:   rgba(15,60,30,0.65);
    --ok-bd:    rgba(48,209,88,0.50);
    --ok-text:  #6ee7b7;
    --ok-text2: #a7f3d0;
    --warn-bg:  rgba(255,159,10,0.22);
    --warn-bg2: rgba(110,70,10,0.65);
    --warn-bg3: rgba(80,55,10,0.65);
    --warn-bd:  rgba(255,159,10,0.55);
    --warn-text:#fcd34d;
    --warn-text2:#fde68a;
    --err-bg:   rgba(255,69,58,0.22);
    --err-bg2:  rgba(110,30,30,0.65);
    --err-bg3:  rgba(80,25,25,0.65);
    --err-bd:   rgba(255,69,58,0.55);
    --err-text: #fca5a5;
    --err-text2:#fecaca;
"""

# Construir el bloque :root según el tema elegido por el usuario
if _THEME == "dark":
    _root_css = f":root {{{_VARS_DARK}}}"
elif _THEME == "light":
    _root_css = f":root {{{_VARS_LIGHT}}}"
else:  # system
    _root_css = (
        f":root {{{_VARS_LIGHT}}}\n"
        f"@media (prefers-color-scheme: dark) {{\n"
        f"  :root {{{_VARS_DARK}}}\n"
        f"}}"
    )

st.markdown(f"<style>\n{_root_css}\n</style>", unsafe_allow_html=True)

st.markdown("""
<style>
    /* ═══ BASE ═══════════════════════════════════════════════ */
    html, body, [class*="css"], .stApp {
        font-family: -apple-system, BlinkMacSystemFont, "SF Pro Display",
                     "SF Pro Text", "Helvetica Neue", Helvetica, Arial, sans-serif;
        -webkit-font-smoothing: antialiased;
        text-rendering: optimizeLegibility;
        letter-spacing: -0.01em;
        color: var(--text-main);
    }
    .stApp {
        background:
            radial-gradient(1100px 700px at 0% 0%, var(--bg-blob-1), transparent 50%),
            radial-gradient(900px 600px at 100% 100%, var(--bg-blob-2), transparent 55%),
            radial-gradient(700px 500px at 100% 0%, var(--bg-blob-3), transparent 60%),
            var(--bg-app);
        background-attachment: fixed;
    }
    .block-container { padding-top: 1.4rem; padding-bottom: 3rem; max-width: 1500px; }
    .stApp footer, .stApp #MainMenu { display: none; }

    /* Contraste universal — texto principal SIEMPRE legible sobre el fondo */
    h1, h2, h3, h4, h5, h6, p, span, label, div { color: var(--text-main); }
    .stMarkdown, .stMarkdown p, .stMarkdown li,
    .stCheckbox label, .stRadio label { color: var(--text-main) !important; }
    .stCaption, .stCaption *, small { color: var(--text-muted) !important; }

    /* ═══ ENCABEZADO PRINCIPAL ════════════════════════════════ */
    .main-title {
        font-size: 1.85rem; font-weight: 700;
        color: var(--text-main);
        margin: 4px 0 2px 0; letter-spacing: -0.025em;
    }
    .sub-title {
        color: var(--text-muted); font-size: 0.92rem; font-weight: 400;
        margin-bottom: 16px;
    }

    /* ═══ SECTION HEADER ═════════════════════════════════════ */
    .section-header {
        font-size: 1.05rem; font-weight: 600;
        color: var(--text-main);
        margin: 18px 0 10px 0;
        padding-bottom: 6px;
        border-bottom: 1px solid var(--border-light);
        letter-spacing: -0.01em;
    }
    .subsection-header {
        font-size: 0.78rem; font-weight: 700;
        color: var(--text-muted);
        text-transform: uppercase; letter-spacing: 0.08em;
        margin: 14px 0 6px 0;
    }

    /* ═══ LIQUID GLASS PANELS (iOS 26 style) ═══════════════ */
    .glass {
        position: relative;
        background: var(--bg-glass);
        backdrop-filter: saturate(200%) blur(28px);
        -webkit-backdrop-filter: saturate(200%) blur(28px);
        border: 1px solid var(--border-light);
        border-radius: 18px;
        padding: 18px 20px;
        box-shadow:
            inset 0 1px 0 rgba(255,255,255,0.22),
            var(--shadow-card);
        color: var(--text-main);
    }
    .glass-strong {
        position: relative;
        background: var(--bg-card);
        backdrop-filter: saturate(200%) blur(32px);
        -webkit-backdrop-filter: saturate(200%) blur(32px);
        border: 1px solid var(--border-light);
        border-radius: 20px;
        padding: 20px;
        box-shadow:
            inset 0 1px 0 rgba(255,255,255,0.25),
            var(--shadow-strong);
        color: var(--text-main);
    }
    .panel-results {
        background: var(--bg-card);
        backdrop-filter: saturate(200%) blur(32px);
        -webkit-backdrop-filter: saturate(200%) blur(32px);
        border: 1px solid var(--border-light);
        border-radius: 22px;
        padding: 20px;
        box-shadow:
            inset 0 1px 0 rgba(255,255,255,0.25),
            var(--shadow-strong);
        color: var(--text-main);
    }

    /* ═══ RESULT CARDS — TODAS LA MISMA ALTURA ══════════════ */
    .result-card {
        position: relative;
        display: flex;
        flex-direction: column;
        justify-content: space-between;
        border-radius: 18px;
        padding: 18px 20px;
        margin: 0 0 12px 0;
        min-height: 130px;
        border: 1px solid var(--border-light);
        background: var(--bg-card);
        backdrop-filter: saturate(200%) blur(28px);
        -webkit-backdrop-filter: saturate(200%) blur(28px);
        box-shadow:
            inset 0 1px 0 rgba(255,255,255,0.22),
            var(--shadow-card);
        color: var(--text-main);
        transition: transform 0.18s ease, box-shadow 0.18s ease;
        overflow: hidden;
    }
    /* Brillo superior tipo cristal (iOS 26) */
    .result-card::before {
        content: '';
        position: absolute;
        top: 0; left: 0; right: 0; height: 50%;
        background: linear-gradient(180deg,
            rgba(255,255,255,0.08), rgba(255,255,255,0));
        pointer-events: none;
        border-radius: 18px 18px 0 0;
    }
    .result-card:hover {
        transform: translateY(-2px);
        box-shadow:
            inset 0 1px 0 rgba(255,255,255,0.28),
            var(--shadow-strong);
    }
    .result-card.ok    { background: linear-gradient(135deg, var(--ok-bg2),   var(--ok-bg3));   border-color: var(--ok-bd); }
    .result-card.warn  { background: linear-gradient(135deg, var(--warn-bg2), var(--warn-bg3)); border-color: var(--warn-bd); }
    .result-card.error { background: linear-gradient(135deg, var(--err-bg2),  var(--err-bg3));  border-color: var(--err-bd); }

    /* Card "destacada" (la grande del conductor recomendado) */
    .result-card.featured { min-height: 160px; padding: 22px 24px; }
    .result-card.featured .rc-value { font-size: 2.6rem; }

    .rc-row { display: flex; align-items: center; gap: 12px; margin-bottom: 4px; }
    .rc-icon {
        width: 40px; height: 40px;
        display: flex; align-items: center; justify-content: center;
        font-size: 18px;
        border-radius: 12px;
        background: var(--bg-glass);
        border: 1px solid var(--border-light);
        flex: 0 0 auto;
        color: var(--text-main);
        box-shadow: inset 0 1px 0 rgba(255,255,255,0.20);
    }
    .rc-label {
        font-size: 0.7rem; font-weight: 700;
        color: var(--text-muted);
        text-transform: uppercase; letter-spacing: 0.08em;
        line-height: 1.2;
        margin-bottom: 6px;
    }
    .rc-value {
        font-size: 1.9rem; font-weight: 700;
        color: var(--text-main);
        line-height: 1.05; letter-spacing: -0.03em;
        margin: 0 0 4px 0;
    }
    .rc-value.ok    { color: var(--ok-text); }
    .rc-value.warn  { color: var(--warn-text); }
    .rc-value.error { color: var(--err-text); }
    .rc-unit {
        font-size: 0.8rem;
        color: var(--text-muted);
        margin-top: auto;
        font-weight: 500;
        line-height: 1.3;
    }
    .rc-pill {
        display: inline-flex; align-items: center; gap: 4px;
        padding: 4px 10px;
        border-radius: 999px;
        font-size: 0.72rem; font-weight: 700; letter-spacing: 0.04em;
        margin-top: 8px;
        align-self: flex-start;
    }
    .rc-pill.ok    { background: var(--ok-bg);   color: var(--ok-text2); }
    .rc-pill.warn  { background: var(--warn-bg); color: var(--warn-text2); }
    .rc-pill.error { background: var(--err-bg);  color: var(--err-text2); }

    /* Compact metric (para grids) */
    .metric-card {
        display: flex;
        flex-direction: column;
        justify-content: center;
        background: var(--bg-card);
        backdrop-filter: saturate(200%) blur(24px);
        -webkit-backdrop-filter: saturate(200%) blur(24px);
        border: 1px solid var(--border-light);
        border-radius: 14px;
        padding: 14px 16px;
        min-height: 96px;
        text-align: center;
        box-shadow:
            inset 0 1px 0 rgba(255,255,255,0.20),
            var(--shadow-card);
        color: var(--text-main);
        margin-bottom: 12px;
    }
    .metric-label {
        font-size: 0.7rem; font-weight: 700;
        color: var(--text-muted);
        text-transform: uppercase; letter-spacing: 0.06em;
        line-height: 1.2;
        margin-bottom: 6px;
    }
    .metric-value {
        font-size: 1.5rem; font-weight: 700;
        color: var(--text-main);
        margin: 0; letter-spacing: -0.02em;
        line-height: 1.15;
    }
    .metric-unit {
        font-size: 0.75rem; color: var(--text-muted);
        margin-top: 6px; line-height: 1.3;
    }

    /* ═══ BLOQUES DE MENSAJE  ═══════════════════════════════ */
    .resultado-ok, .resultado-warn, .resultado-error {
        border-radius: 14px; padding: 12px 16px; margin: 10px 0;
        font-weight: 500; backdrop-filter: blur(18px);
    }
    .resultado-ok    { background: var(--ok-bg);   color: var(--ok-text2);   border: 1px solid var(--ok-bd); }
    .resultado-warn  { background: var(--warn-bg); color: var(--warn-text2); border: 1px solid var(--warn-bd); }
    .resultado-error { background: var(--err-bg);  color: var(--err-text2);  border: 1px solid var(--err-bd); }

    /* ═══ SIDEBAR (siempre dark glass — independiente de tema global) ═ */
    section[data-testid="stSidebar"] {
        background:
            linear-gradient(180deg, rgba(20,22,30,0.94), rgba(28,30,40,0.97)),
            radial-gradient(600px 400px at 50% 0%, rgba(0,122,255,0.18), transparent 60%) !important;
        backdrop-filter: blur(30px);
        -webkit-backdrop-filter: blur(30px);
        border-right: 1px solid rgba(255,255,255,0.06);
    }
    section[data-testid="stSidebar"] > div { padding-top: 1rem; }

    /* Cualquier texto dentro del sidebar — fuerza blanco */
    section[data-testid="stSidebar"],
    section[data-testid="stSidebar"] *:not(button):not([class*="resultado"]):not(.stAlert) {
        color: #f5f5f7 !important;
    }
    section[data-testid="stSidebar"] small,
    section[data-testid="stSidebar"] .stCaption,
    section[data-testid="stSidebar"] [data-testid="stCaptionContainer"],
    section[data-testid="stSidebar"] [data-testid="stCaptionContainer"] * {
        color: #a1a1a6 !important;
    }

    /* Inputs: text_input, number_input, text_area */
    section[data-testid="stSidebar"] input,
    section[data-testid="stSidebar"] textarea,
    section[data-testid="stSidebar"] [data-baseweb="input"] input,
    section[data-testid="stSidebar"] [data-baseweb="textarea"] textarea,
    section[data-testid="stSidebar"] [data-testid="stTextInput"] input,
    section[data-testid="stSidebar"] [data-testid="stTextArea"] textarea,
    section[data-testid="stSidebar"] [data-testid="stNumberInput"] input,
    section[data-testid="stSidebar"] [data-testid="stDateInput"] input {
        background: rgba(255,255,255,0.10) !important;
        border-radius: 10px !important;
        border: 1px solid rgba(255,255,255,0.18) !important;
        color: #f5f5f7 !important;
        -webkit-text-fill-color: #f5f5f7 !important;
        caret-color: #f5f5f7 !important;
    }
    section[data-testid="stSidebar"] input::placeholder,
    section[data-testid="stSidebar"] textarea::placeholder {
        color: #a1a1a6 !important;
        opacity: 0.7;
    }
    section[data-testid="stSidebar"] input:focus,
    section[data-testid="stSidebar"] textarea:focus {
        border-color: #0a84ff !important;
        box-shadow: 0 0 0 3px rgba(10,132,255,0.25) !important;
    }

    /* Selectbox y multiselect en sidebar */
    section[data-testid="stSidebar"] [data-baseweb="select"] > div,
    section[data-testid="stSidebar"] [data-baseweb="select"] [role="combobox"],
    section[data-testid="stSidebar"] [data-baseweb="select"] [role="option"] {
        background: rgba(255,255,255,0.10) !important;
        border-radius: 10px !important;
        border-color: rgba(255,255,255,0.18) !important;
        color: #f5f5f7 !important;
    }
    section[data-testid="stSidebar"] [data-baseweb="select"] svg { fill: #f5f5f7 !important; }
    section[data-testid="stSidebar"] [data-baseweb="tag"] {
        background: rgba(10,132,255,0.30) !important;
        color: #f5f5f7 !important;
        border: 1px solid rgba(10,132,255,0.50) !important;
    }
    section[data-testid="stSidebar"] [data-baseweb="tag"] * { color: #f5f5f7 !important; }

    /* Labels más legibles */
    section[data-testid="stSidebar"] [data-testid="stWidgetLabel"] p,
    section[data-testid="stSidebar"] label {
        font-size: 0.82rem !important; font-weight: 600;
        color: #f5f5f7 !important;
        opacity: 1;
    }

    /* Expanders */
    section[data-testid="stSidebar"] [data-testid="stExpander"] {
        background: rgba(255,255,255,0.06) !important;
        border: 1px solid rgba(255,255,255,0.12) !important;
        border-radius: 12px;
    }
    section[data-testid="stSidebar"] [data-testid="stExpander"] summary,
    section[data-testid="stSidebar"] [data-testid="stExpander"] summary *,
    section[data-testid="stSidebar"] [data-testid="stExpander"] details *,
    section[data-testid="stSidebar"] [data-testid="stExpander"] svg {
        color: #f5f5f7 !important;
        fill: #f5f5f7 !important;
    }

    /* Botones del sidebar */
    section[data-testid="stSidebar"] .stButton > button {
        background: rgba(255,255,255,0.08) !important;
        border: 1px solid rgba(255,255,255,0.18) !important;
        color: #f5f5f7 !important;
    }
    section[data-testid="stSidebar"] .stButton > button:hover {
        background: rgba(255,255,255,0.14) !important;
    }
    section[data-testid="stSidebar"] .stButton > button[kind="primary"] {
        background: linear-gradient(180deg, #1a8eff, #0a84ff) !important;
        border: 1px solid #0a84ff !important;
    }

    /* Radio en sidebar (selector de tema) */
    section[data-testid="stSidebar"] [data-testid="stRadio"] label {
        background: rgba(255,255,255,0.06);
        border: 1px solid rgba(255,255,255,0.12);
        border-radius: 8px;
        padding: 4px 8px;
        margin-right: 4px;
    }

    /* Popover de selects (cuando se abre el dropdown) */
    [data-baseweb="popover"] [role="listbox"],
    [data-baseweb="popover"] [role="option"],
    [data-baseweb="menu"] li {
        color: var(--text-main) !important;
        background: var(--bg-card) !important;
    }
    [data-baseweb="popover"] [role="option"]:hover {
        background: var(--bg-input) !important;
    }

    /* ═══ TABS (segmented control con liquid glass iOS 26) ═ */
    .stTabs [data-baseweb="tab-list"] {
        gap: 3px;
        background: var(--bg-tab-list);
        backdrop-filter: saturate(200%) blur(24px);
        -webkit-backdrop-filter: saturate(200%) blur(24px);
        padding: 5px;
        border-radius: 14px;
        border: 1px solid var(--border-tab);
        box-shadow:
            inset 0 1px 0 rgba(255,255,255,0.18),
            var(--shadow-card);
        overflow-x: auto;
        scrollbar-width: thin;
        margin-bottom: 6px;
        flex-wrap: nowrap;
    }
    .stTabs [data-baseweb="tab"] {
        font-weight: 600;
        font-size: 0.86rem;
        border-radius: 10px;
        padding: 8px 14px;
        color: var(--text-main) !important;
        white-space: nowrap;
        min-height: 38px;
        display: flex;
        align-items: center;
        gap: 6px;
        transition: all 0.18s ease;
    }
    .stTabs [data-baseweb="tab"]:hover {
        background: var(--bg-input);
    }
    .stTabs [aria-selected="true"] {
        background: var(--bg-tab-active) !important;
        color: var(--text-tabactive) !important;
        box-shadow:
            inset 0 1px 0 rgba(255,255,255,0.30),
            0 2px 8px rgba(0,0,0,0.10);
    }
    /* En pantallas medianas, tabs más compactos */
    @media (max-width: 1280px) {
        .stTabs [data-baseweb="tab"] {
            font-size: 0.8rem;
            padding: 8px 10px;
        }
    }

    /* ═══ INPUTS — TODOS CON LA MISMA ALTURA (44px iOS) ═══ */
    /* Etiqueta consistente */
    [data-testid="stWidgetLabel"],
    [data-testid="stWidgetLabel"] p {
        font-size: 0.78rem; font-weight: 600;
        color: var(--text-sub) !important;
        margin-bottom: 4px !important;
        min-height: 18px;
        line-height: 1.25;
    }
    /* Spacing uniforme entre widgets */
    .stNumberInput, .stTextInput, .stSelectbox, .stTextArea,
    .stDateInput, .stMultiselect {
        margin-bottom: 12px !important;
    }

    /* Altura uniforme — todos los inputs */
    .stTextInput input,
    .stNumberInput input,
    .stDateInput input,
    .stSelectbox > div > div,
    [data-baseweb="select"] > div,
    [data-baseweb="input"] {
        border-radius: 12px !important;
        border: 1px solid var(--border-input) !important;
        background: var(--bg-input) !important;
        color: var(--text-main) !important;
        font-size: 0.93rem !important;
        min-height: 44px !important;
        height: 44px !important;
        box-shadow: inset 0 1px 0 rgba(255,255,255,0.08);
        transition: all 0.15s ease;
    }
    /* Textarea: altura libre pero mismo borde */
    .stTextArea textarea {
        border-radius: 12px !important;
        border: 1px solid var(--border-input) !important;
        background: var(--bg-input) !important;
        color: var(--text-main) !important;
        font-size: 0.93rem !important;
        box-shadow: inset 0 1px 0 rgba(255,255,255,0.08);
    }

    /* Focus uniforme con anillo azul */
    .stTextInput input:focus,
    .stNumberInput input:focus,
    .stDateInput input:focus,
    .stTextArea textarea:focus {
        border-color: var(--accent-blue) !important;
        box-shadow:
            0 0 0 3px rgba(10,132,255,0.18),
            inset 0 1px 0 rgba(255,255,255,0.10) !important;
        outline: none !important;
    }
    [data-baseweb="select"] > div:focus-within {
        border-color: var(--accent-blue) !important;
        box-shadow: 0 0 0 3px rgba(10,132,255,0.18) !important;
    }

    .stTextInput input::placeholder,
    .stNumberInput input::placeholder,
    .stTextArea textarea::placeholder {
        color: var(--text-muted) !important;
        opacity: 0.55;
    }

    /* ── NUMBER INPUT: ocultar X y normalizar botones +/- ── */
    /* La X de "clear" sólo aparece en number_input */
    .stNumberInput [data-baseweb="input"] button[aria-label*="clear" i],
    .stNumberInput [data-baseweb="input"] button[aria-label*="borrar" i],
    .stNumberInput [data-baseweb="input"] button[title*="Clear" i],
    .stNumberInput [data-baseweb="input"] [data-testid*="clear" i] {
        display: none !important;
    }
    /* Botones +/- del number_input: SOLO los que están dentro del input wrapper.
       NUNCA tocar el icono ? del label (que también es un <button>). */
    .stNumberInput [data-baseweb="input"] button {
        min-height: 44px !important;
        height: 44px !important;
        width: 36px !important;
        border-radius: 0 !important;
        border: 1px solid var(--border-input) !important;
        border-left: none !important;
        background: var(--bg-input) !important;
        color: var(--text-main) !important;
        font-weight: 700;
        font-size: 1rem;
        box-shadow: none !important;
    }
    .stNumberInput [data-baseweb="input"] button:hover {
        background: var(--bg-card) !important;
        color: var(--accent-blue) !important;
    }
    /* Quitar background gris de los step buttons */
    .stNumberInput div[data-baseweb="input"] > div:last-child {
        background: transparent !important;
        border-radius: 0 12px 12px 0 !important;
        overflow: hidden;
    }
    /* Container del number_input alineado */
    .stNumberInput [data-baseweb="input"] {
        border-radius: 12px !important;
        overflow: hidden;
    }
    /* Multiselect tags compactas */
    [data-baseweb="select"] [data-baseweb="tag"] {
        color: var(--text-main) !important;
        background: rgba(10,132,255,0.14) !important;
        border: 1px solid rgba(10,132,255,0.35) !important;
        border-radius: 8px !important;
    }
    [data-baseweb="popover"] li,
    [data-baseweb="popover"] [role="option"] {
        color: var(--text-main) !important;
    }

    /* Radio en main: legible en ambos temas */
    .stRadio > div > label { color: var(--text-main) !important; }
    .stRadio > div > label p { color: var(--text-main) !important; }

    /* Checkbox compacto */
    .stCheckbox {
        margin-bottom: 8px;
    }
    .stCheckbox label { font-size: 0.92rem !important; }

    /* ═══ BUTTONS — iOS 26 ═══════════════════════════════════ */
    .stButton > button, .stDownloadButton > button, .stFormSubmitButton > button {
        border-radius: 12px;
        font-weight: 600;
        font-size: 0.92rem;
        padding: 10px 18px;
        min-height: 44px;
        border: 1px solid var(--border-input);
        transition: all 0.18s ease;
        backdrop-filter: blur(20px) saturate(180%);
        -webkit-backdrop-filter: blur(20px) saturate(180%);
        color: var(--text-main);
        background: var(--bg-glass);
        box-shadow:
            inset 0 1px 0 rgba(255,255,255,0.15),
            0 1px 2px rgba(0,0,0,0.04);
    }
    .stButton > button:hover, .stDownloadButton > button:hover {
        transform: translateY(-1px);
        background: var(--bg-card);
    }
    .stButton > button[kind="primary"], .stFormSubmitButton > button[kind="primary"] {
        background: linear-gradient(180deg, var(--accent-blue-h), var(--accent-blue));
        color: #ffffff !important;
        border: 1px solid var(--accent-blue);
        box-shadow:
            inset 0 1px 0 rgba(255,255,255,0.22),
            0 2px 8px rgba(10,132,255,0.35);
    }
    .stButton > button[kind="primary"]:hover {
        transform: translateY(-1px);
        box-shadow:
            inset 0 1px 0 rgba(255,255,255,0.25),
            0 6px 16px rgba(10,132,255,0.45);
    }
    .stButton > button[kind="secondary"] {
        background: var(--bg-card);
        color: var(--text-main) !important;
    }

    /* ═══ EXPANDERS (glass, ambos temas) ═══════════════════ */
    div[data-testid="stExpander"] {
        background: var(--bg-glass);
        backdrop-filter: blur(18px);
        border-radius: 14px;
        border: 1px solid var(--border-light);
        box-shadow: var(--shadow-card);
        margin: 8px 0;
        color: var(--text-main);
    }
    div[data-testid="stExpander"] summary { font-weight: 600; }
    div[data-testid="stExpander"] summary p,
    div[data-testid="stExpander"] summary div { color: var(--text-main) !important; font-weight: 600; }

    /* ═══ DATAFRAME — borde adaptable ══════════════════════ */
    .stDataFrame, [data-testid="stDataFrame"] {
        border-radius: 12px;
        overflow: hidden;
        border: 1px solid var(--border-light);
    }

    /* ═══ FORMULA BOX ═════════════════════════════════════ */
    .formula-box {
        background: var(--bg-formula);
        backdrop-filter: blur(14px);
        border: 1px solid var(--border-light);
        border-radius: 12px;
        padding: 12px 16px;
        margin: 6px 0 12px 0;
        font-family: "SF Mono", "JetBrains Mono", "Menlo", Consolas, monospace;
        font-size: 0.86rem;
        color: var(--text-main);
    }

    /* ═══ PROJECT HEADER ══════════════════════════════════ */
    .project-header-card {
        background: var(--bg-glass);
        backdrop-filter: saturate(180%) blur(20px);
        border: 1px solid var(--border-light);
        border-radius: 16px;
        padding: 12px 18px;
        margin-bottom: 14px;
        box-shadow: var(--shadow-card);
        overflow: hidden;
    }
    .project-header-row {
        display: flex; flex-wrap: wrap;
        gap: 6px 24px;
        font-size: 0.85rem;
    }
    .project-header-row span b { color: var(--text-main); font-weight: 700; }
    .project-header-row span   { color: var(--text-sub); word-break: break-word; }

    /* ═══ EMPTY STATE ════════════════════════════════════ */
    .empty-results {
        text-align: center; padding: 60px 24px;
        color: var(--text-muted);
    }
    .empty-results .empty-icon {
        font-size: 3.5rem; opacity: 0.45; margin-bottom: 8px;
    }
    .empty-results .empty-title {
        font-size: 1.1rem; font-weight: 600;
        color: var(--text-main);
        margin-bottom: 4px;
    }
    .empty-results .empty-sub {
        font-size: 0.88rem; color: var(--text-muted);
    }

    /* ═══ RECOMMENDATION BLOCK ════════════════════════════ */
    .recom {
        background: var(--bg-pill-recom);
        border: 1px solid rgba(0,113,227,0.25);
        border-radius: 14px;
        padding: 12px 16px;
        margin: 10px 0;
    }
    .recom-title {
        font-size: 0.74rem; font-weight: 700;
        text-transform: uppercase; letter-spacing: 0.08em;
        color: var(--accent-blue);
        margin-bottom: 4px;
    }
    .recom-text { font-size: 0.92rem; color: var(--text-main); line-height: 1.4; }

    /* ═══ ICONOS DE AYUDA (?) Y TOOLTIPS ═══════════════════ */
    /* Reset NUCLEAR del ícono ? y todos sus wrappers:
       Streamlit lo renderiza como <button> dentro del label, y como
       otros selectores de buttons podían afectarlo, lo aislamos por
       completo. */
    [data-testid="stTooltipIcon"],
    [data-testid="stTooltipIcon"] *,
    [data-testid="stTooltipHoverTarget"],
    [data-testid="stTooltipHoverTarget"] *,
    [data-testid="stWidgetLabelHelp"],
    [data-testid="stWidgetLabelHelp"] *,
    [data-testid="stWidgetLabelHelpInline"],
    [data-testid="stWidgetLabelHelpInline"] *,
    [data-testid="stWidgetLabel"] button,
    [data-testid="stWidgetLabel"] button * {
        background: transparent !important;
        background-color: transparent !important;
        background-image: none !important;
        border: 0 !important;
        border-radius: 0 !important;
        box-shadow: none !important;
        padding: 0 !important;
        margin: 0 !important;
        outline: none !important;
        width: auto !important;
        height: auto !important;
        min-width: 0 !important;
        min-height: 0 !important;
    }
    /* Estilo final del círculo ? */
    [data-testid="stTooltipIcon"] {
        color: var(--accent-blue) !important;
        opacity: 0.7;
        margin-left: 6px !important;
        cursor: help;
        display: inline-flex !important;
        align-items: center;
        vertical-align: middle;
    }
    [data-testid="stTooltipIcon"]:hover { opacity: 1; }
    [data-testid="stTooltipIcon"] svg {
        fill: var(--accent-blue) !important;
        color: var(--accent-blue) !important;
        width: 14px !important;
        height: 14px !important;
    }
    /* Sidebar (siempre dark): icono ? naranja */
    section[data-testid="stSidebar"] [data-testid="stTooltipIcon"] {
        color: var(--accent-orange) !important;
        opacity: 0.85;
    }
    section[data-testid="stSidebar"] [data-testid="stTooltipIcon"] svg {
        fill: var(--accent-orange) !important;
        color: var(--accent-orange) !important;
    }

    /* Tooltip flotante: caja oscura legible */
    [data-baseweb="tooltip"],
    [role="tooltip"] {
        background: rgba(20,22,30,0.96) !important;
        color: #f5f5f7 !important;
        border-radius: 10px !important;
        padding: 8px 12px !important;
        font-size: 0.82rem !important;
        line-height: 1.4 !important;
        max-width: 320px !important;
        box-shadow: 0 8px 24px rgba(0,0,0,0.30) !important;
        border: 1px solid rgba(255,255,255,0.08) !important;
        backdrop-filter: blur(20px);
        z-index: 9999 !important;
    }
    [data-baseweb="tooltip"] *,
    [role="tooltip"] * {
        color: #f5f5f7 !important;
        background: transparent !important;
    }

    /* ═══ RESPONSIVE — TABLET (≤ 1024px) ═══════════════════ */
    @media (max-width: 1024px) {
        .block-container {
            padding-left: 1rem !important;
            padding-right: 1rem !important;
        }
        .main-title { font-size: 1.55rem; }
        .rc-value { font-size: 1.8rem; }
        .stTabs [data-baseweb="tab"] {
            padding: 7px 12px; font-size: 0.85rem;
        }
        .project-header-row { font-size: 0.8rem; gap: 4px 16px; }
    }

    /* ═══ RESPONSIVE — MÓVIL (≤ 768px) ════════════════════ */
    @media (max-width: 768px) {
        /* Encabezado más compacto */
        .main-title { font-size: 1.35rem; }
        .sub-title  { font-size: 0.82rem; margin-bottom: 8px; }

        /* Las columnas de Streamlit se apilan en móvil */
        [data-testid="column"] {
            width: 100% !important;
            flex: 1 0 100% !important;
            margin-bottom: 6px;
        }
        [data-testid="stHorizontalBlock"] {
            flex-wrap: wrap !important;
            gap: 6px !important;
        }

        /* Cards más compactas */
        .result-card { padding: 14px 16px 12px; border-radius: 14px; }
        .rc-value { font-size: 1.6rem; }
        .rc-icon { width: 30px; height: 30px; font-size: 15px; }
        .metric-card { padding: 10px 12px; border-radius: 12px; }
        .metric-value { font-size: 1.15rem; }

        /* Tabs horizontalmente scrolleables */
        .stTabs [data-baseweb="tab-list"] {
            flex-wrap: nowrap !important;
            overflow-x: auto !important;
            padding: 4px;
        }
        .stTabs [data-baseweb="tab"] {
            padding: 6px 10px; font-size: 0.8rem;
        }

        /* Inputs más altos para tocar mejor */
        .stTextInput input, .stNumberInput input, .stSelectbox > div > div {
            min-height: 38px !important;
            font-size: 0.95rem !important;
        }

        /* Block container con menos padding */
        .block-container {
            padding-top: 0.8rem !important;
            padding-bottom: 1.5rem !important;
        }

        /* Sidebar full width cuando se abre */
        section[data-testid="stSidebar"] {
            width: 88vw !important; min-width: 88vw !important;
        }

        .section-header { font-size: 0.98rem; margin: 12px 0 8px; }
        .subsection-header { font-size: 0.7rem; margin: 10px 0 4px; }

        .project-header-row { font-size: 0.78rem; gap: 4px 12px; }
    }

    /* ═══ RESPONSIVE — PHONE PEQUEÑO (≤ 480px) ════════════ */
    @media (max-width: 480px) {
        .main-title { font-size: 1.2rem; }
        .rc-value   { font-size: 1.4rem; }
        .stTabs [data-baseweb="tab"] {
            padding: 6px 8px; font-size: 0.75rem;
        }
        .block-container {
            padding-left: 0.6rem !important;
            padding-right: 0.6rem !important;
        }
    }
</style>
""", unsafe_allow_html=True)

# ═════════════════════════════════════════════════════════
# SPLASH SCREEN + LOGIN GATE  (eSuit)
# ═════════════════════════════════════════════════════════
def _mostrar_splash_login():
    """Pantalla de bienvenida con logo y formulario de login.
       Bloquea el resto de la app hasta que el usuario se autentique."""
    # CSS específico de splash (usa variables del tema global)
    st.markdown("""
    <style>
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
            letter-spacing: -0.04em;
            color: var(--text-main);
            margin: 0;
            animation: fadeUp 0.6s ease-out 0.2s both;
        }
        .splash-subtitle {
            font-size: 1.05rem;
            color: var(--text-muted);
            margin: 4px 0 28px 0; font-weight: 400;
            animation: fadeUp 0.6s ease-out 0.35s both;
        }
        .splash-foot {
            color: var(--text-muted); font-size: 0.8rem;
            margin-top: 1.5rem;
            animation: fadeUp 0.6s ease-out 0.65s both;
        }
        @media (max-width: 480px) {
            .splash-title    { font-size: 2.2rem; }
            .splash-subtitle { font-size: 0.92rem; }
            .splash-container { padding: 0 12px; margin-top: 1rem; }
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
        # Preservar el tema entre sesiones
        tema_preservado = st.session_state.get("ui_theme")
        for k in list(st.session_state.keys()):
            del st.session_state[k]
        if tema_preservado:
            st.session_state["ui_theme"] = tema_preservado
        st.rerun()

    # Aviso si admin sigue con contraseña por defecto
    if _is_admin and _usr.get("default_pwd"):
        st.warning(
            "⚠️ Estás usando la **contraseña por defecto** (`admin123`). "
            "Ve al panel de administración para cambiarla."
        )

    # ── Selector de TEMA (Claro / Oscuro / Sistema) ──────────
    st.markdown("---")
    _tema_opts = {
        "☀️  Claro": "light",
        "🌙  Oscuro": "dark",
        "⚙️  Sistema": "system",
    }
    _tema_inv = {v: k for k, v in _tema_opts.items()}
    _tema_actual_lbl = _tema_inv.get(_THEME, "⚙️  Sistema")
    _tema_sel = st.radio(
        "Apariencia",
        list(_tema_opts.keys()),
        index=list(_tema_opts.keys()).index(_tema_actual_lbl),
        horizontal=True,
        key="tema_radio",
        label_visibility="visible",
    )
    _tema_nuevo = _tema_opts[_tema_sel]
    if _tema_nuevo != _THEME:
        st.session_state["ui_theme"] = _tema_nuevo
        st.rerun()

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
# Encabezado: logo pequeño + título + subtítulo a la izquierda; resumen proyecto a la derecha
def _esc(s):
    return (s or "").replace("<", "&lt;").replace(">", "&gt;")

col_h1, col_h2 = st.columns([3, 4])
with col_h1:
    st.markdown(
        f"""<div style="display:flex;align-items:center;gap:14px;margin-top:2px;">
            <div style="flex:0 0 auto;">{auth.LOGO_SVG_SMALL}</div>
            <div>
                <div class="main-title">eSuit</div>
                <div class="sub-title">Suite de cálculo eléctrico — NOM-001-SEDE-2012</div>
            </div>
        </div>""",
        unsafe_allow_html=True,
    )
with col_h2:
    st.markdown(
        f"""<div class="project-header-card" style="margin-top:6px;">
            <div class="project-header-row">
                <span><b>Proyecto:</b> {_esc(proyecto.strip() or "—")}</span>
                <span><b>Cliente:</b> {_esc(cliente.strip() or "—")}</span>
            </div>
            <div class="project-header-row" style="margin-top:4px;">
                <span><b>Sistema:</b> {_esc(tipo_sistema)}</span>
                <span><b>Norma:</b> {_esc(norma)}</span>
            </div>
        </div>""",
        unsafe_allow_html=True
    )

# ─────────────────────────────────────────────
# TABS PRINCIPALES
# ─────────────────────────────────────────────
# Pestañas profesionales — herramientas de cálculo (compactas, no se cortan)
_tab_labels = [
    "🔌 Conductor",
    "🔧 Tubería",
    "📊 Tablero",
    "📄 Reporte",
]
if _is_admin:
    _tab_labels.append("👥 Admin")
    tab1, tab2, tab3, tab4, tab_users = st.tabs(_tab_labels)
else:
    tab1, tab2, tab3, tab4 = st.tabs(_tab_labels)
    tab_users = None

# ══════════════════════════════════════════════
# TAB 1 — CAÍDA DE TENSIÓN Y SELECCIÓN DE CONDUCTOR
# Layout: 2 columnas (Inputs | Resultados destacados)
# ══════════════════════════════════════════════
with tab1:
    # ════════════════════════════════════════════════════════════════
    # Mapas y fórmulas (referencias)
    # ════════════════════════════════════════════════════════════════
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

    # ════════════════════════════════════════════════════════════════
    # LAYOUT 2 COLUMNAS — Izquierda: inputs · Derecha: resultados
    # ════════════════════════════════════════════════════════════════
    col_inputs, col_results = st.columns([1, 1], gap="large")

    # ══════════════════════════════════════════════
    # PANEL IZQUIERDO — PARÁMETROS COMPACTOS
    # ══════════════════════════════════════════════
    with col_inputs:
        st.markdown('<div class="subsection-header">Sistema y método</div>',
                    unsafe_allow_html=True)
        c_a, c_b = st.columns(2)
        with c_a:
            config_label = st.selectbox(
                "Configuración del circuito",
                list(config_map.keys()),
                label_visibility="visible",
            )
        with c_b:
            metodo_label = st.selectbox(
                "Método de cálculo",
                list(metodo_map.keys()),
            )
        metodo = metodo_map[metodo_label]
        config = config_map[config_label]

        # Inputs sin valores por defecto — el usuario los completa
        c_v1, c_v2, c_v3 = st.columns(3)
        with c_v1:
            voltaje_nominal = st.number_input(
                "Voltaje (V)", min_value=0, max_value=15000,
                value=None, step=10,
                placeholder="Ej. 220",
                help="Tensión nominal de operación. Mono 2H usa V_F-N; Mono 3H y Trif. usan V_línea.",
            )
        with c_v2:
            potencia = st.number_input(
                "Potencia (W)", min_value=0, max_value=10_000_000,
                value=None, step=100,
                placeholder="Ej. 5000",
                help="Potencia activa de la carga en Watts.",
            )
        with c_v3:
            fp = st.number_input(
                "FP (cosφ)", min_value=0.1, max_value=1.0,
                value=0.90, step=0.01,
                help="Factor de potencia (típico 0.85-0.95 inductivo).",
            )
        sen_phi = round(math.sqrt(max(0.0, 1 - (fp or 0.9)**2)), 4)

        st.markdown('<div class="subsection-header">Conductor</div>',
                    unsafe_allow_html=True)
        c1, c2, c3 = st.columns(3)
        with c1:
            longitud = st.number_input(
                "Longitud (m)", min_value=0, max_value=5000,
                value=None, step=1,
                placeholder="Ej. 40",
                help="Longitud de un solo conductor (ida).",
            )
        with c2:
            material = st.selectbox(
                "Material",
                ["Cobre (Cu)", "Aluminio (Al)"],
            )
        with c3:
            cf = st.number_input(
                "CF (paralelo)", min_value=1, max_value=10, value=1, step=1,
                help="Conductores por fase en paralelo.",
            )

        c4, c5, c6 = st.columns(3)
        with c4:
            cdt_max = st.number_input(
                "C.d.T. máx (%)", min_value=0.5, max_value=10.0, value=3.0, step=0.5,
                help="NOM-001: alimentadores 2%, ramales 3%, total 5%.",
            )
        with c5:
            conductores_activos = st.selectbox(
                "Agrupamiento",
                [1, 2, 3, 4, 5, 6], index=2,
                help="Número de conductores portadores en el tubo.",
            )
        with c6:
            aisl_label = st.selectbox(
                "Aislamiento",
                list(AISLAMIENTOS_T.keys()),
                index=1,
                help="THW=75°C, THHW-2/XHHW-2=90°C.",
            )

        st.markdown('<div class="subsection-header">Canalización</div>',
                    unsafe_allow_html=True)
        c7, c8 = st.columns(2)
        with c7:
            canal_label = st.selectbox(
                "Tipo (Tabla 9 NOM)",
                list(CANALIZACIONES.keys()),
                index=2,
                help="Afecta R y X según material magnético/no-magnético.",
            )
        with c8:
            term_label = st.selectbox(
                "Temp. terminales (Art. 110-14)",
                ["Auto (según In)", "60 °C", "75 °C", "90 °C"],
                help="Auto: 60°C si In ≤ 100A, 75°C si In > 100A.",
            )

        canalizacion = CANALIZACIONES[canal_label]
        temp_aislamiento = AISLAMIENTOS_T[aisl_label]
        temp_term_override = None if term_label.startswith("Auto") else int(term_label.split()[0])
        fc_agrup_val = {1: 1.0, 2: 0.88, 3: 0.82, 4: 0.75, 5: 0.67, 6: 0.58}[conductores_activos]

        # Carga continua + motor
        c9, c10 = st.columns(2)
        with c9:
            carga_continua = st.checkbox(
                "Carga continua ×1.25 (NEC 210.20)",
                value=True,
                help="Cargas que operan ≥ 3 horas continuas.",
            )
        with c10:
            es_motor = st.checkbox(
                "Circuito de motor (NEC 430)",
                value=False,
                help="Fuerza terminal 75°C.",
            )
        factor_demanda = 1.25 if carga_continua else 1.0

        # Ayuda contextual carga continua (expander)
        with st.expander("💡 ¿Cuándo aplicar el factor 1.25?", expanded=False):
            st.markdown("""
**Carga continua** (NEC 210.19 / 210.20):
Carga que opera **3 h o más** ininterrumpidamente.
`I_diseño = I_real × 1.25`

**Aplica para:** alumbrado comercial, HVAC, procesos industriales, tableros principales.

**NO aplica para:** contactos de hogar, electrodomésticos intermitentes, cargas < 3h.

**Motores:** llevan tratamiento aparte (NEC 430.24).
            """)

        st.markdown('<div class="subsection-header">Conductor de tierra (cable desnudo)</div>',
                    unsafe_allow_html=True)
        c_t1, c_t2 = st.columns(2)
        with c_t1:
            tierra_auto = st.checkbox(
                "Auto · Tabla 250-122",
                value=True,
                help="Selecciona el calibre según la capacidad del OCPD.",
            )
            tierra_material = st.selectbox(
                "Material",
                ["Cu (cobre desnudo)", "Al (aluminio)"],
                index=0,
                key="tierra_mat_sel",
            )
        with c_t2:
            if not tierra_auto:
                tierra_manual = st.selectbox(
                    "Calibre manual",
                    ORDEN_CALIBRES,
                    index=ORDEN_CALIBRES.index("8"),
                )
            else:
                tierra_manual = None
                st.caption("📐 Se calcula automáticamente con la protección final.")

        # Tubería integrada (toggle)
        with st.expander("🔧 Calcular tubería (conduit) para este circuito",
                         expanded=False):
            incluir_conduit_calc = st.checkbox(
                "Incluir cálculo de tubería",
                value=False,
                key="chk_incluir_conduit_inline",
            )
            if incluir_conduit_calc:
                cc1, cc2 = st.columns(2)
                num_default = {"mono2h": 3, "mono3h": 4, "trifasico": 5}.get(config, 3)
                with cc1:
                    tipo_tubo_inline = st.selectbox(
                        "Tipo de tubería",
                        list(TABLA_CONDUIT.keys()),
                        key="tubo_inline",
                    )
                    num_conds_inline = st.number_input(
                        "# Conductores (fases+neutro+tierra)",
                        min_value=1, max_value=40,
                        value=num_default, step=1,
                        key="ncond_inline",
                    )
                with cc2:
                    tipo_aisl_inline = st.selectbox(
                        "Aislamiento de los conductores",
                        ["THW (75°C)", "THHW (75/90°C)", "THWN-2 (90°C)", "XHHW-2 (90°C)"],
                        key="aisl_inline",
                    )
                    st.caption("ℹ️ Usa el OD del calibre principal del circuito.")
            else:
                tipo_tubo_inline = None
                num_conds_inline = None
                tipo_aisl_inline = None

        # ── Fórmula utilizada (oculta por defecto) ──
        with st.expander("📐 Ver fórmula utilizada", expanded=False):
            st.markdown(
                f'<div class="formula-box"><b>Fórmula aplicada:</b><br><br>'
                f'{formulas.get((metodo, config), "—")}</div>',
                unsafe_allow_html=True,
            )

    # ══════════════════════════════════════════════
    # CÁLCULO (siempre que los inputs sean válidos)
    # ══════════════════════════════════════════════
    def _inputs_validos():
        return all([
            potencia is not None and potencia > 0,
            voltaje_nominal is not None and voltaje_nominal > 0,
            longitud is not None and longitud > 0,
            fp is not None and 0 < fp <= 1,
            cdt_max is not None and cdt_max > 0,
            cf is not None and cf >= 1,
            proyecto and proyecto.strip(),
        ])

    inputs_ok = _inputs_validos()
    mat = "Cu" if "Cobre" in material else "Al"
    mat_tierra = "Cu" if "Cu" in tierra_material else "Al"

    if inputs_ok:
        try:
            corriente, conductor_min, cdt_real, conductor_verificado, ampacity_corr = (
                calcular_caida_tension(
                    potencia, fp, voltaje_nominal, config, longitud, mat,
                    cdt_max, fc_temp, fc_agrup_val, metodo,
                    factor_demanda=factor_demanda, cf=cf,
                    canalizacion=canalizacion, temp_term=temp_term_override,
                    temp_aislamiento=temp_aislamiento, es_motor=es_motor,
                )
            )
            corriente_diseño = corriente * factor_demanda
            corriente_diseño_cond = corriente_diseño / cf
            ocpd_A, ocpd_status = calcular_proteccion(corriente_diseño_cond, ampacity_corr)
            temp_term_efectiva = temp_term_override or temp_terminales_auto(corriente, es_motor=es_motor)
            proteccion_fmt = formato_proteccion(ocpd_A, config)

            df_tabla = seleccionar_conductor(
                corriente, longitud, voltaje_nominal, config, mat,
                cdt_max, fc_temp, fc_agrup_val, fp, metodo,
                factor_demanda=factor_demanda, cf=cf,
                canalizacion=canalizacion, temp_term=temp_term_override,
                temp_aislamiento=temp_aislamiento, es_motor=es_motor,
            )

            tierra_calibre = (
                calibre_tierra(ocpd_A, mat_tierra) if tierra_auto
                else (tierra_manual or "—")
            )

            # Conduit inline
            conduit_inline_result = None
            if incluir_conduit_calc and tipo_tubo_inline:
                todos_tubos = calcular_conduit_fill(
                    conductor_verificado, num_conds_inline, tipo_tubo_inline
                )
                recom = next((x for x in todos_tubos if x["cumple"]),
                             todos_tubos[-1] if todos_tubos else None)
                conduit_inline_result = {
                    "params": {
                        "calibre": conductor_verificado,
                        "num_conds": num_conds_inline,
                        "tipo_tubo": tipo_tubo_inline,
                        "aislamiento": tipo_aisl_inline,
                    },
                    "resultados": todos_tubos,
                    "resultados_todos": todos_tubos,
                    "recomendada": recom,
                }

            calc_ok = True
        except Exception as ex:
            calc_ok = False
            error_msg = str(ex)
            corriente = corriente_diseño = corriente_diseño_cond = 0
            cdt_real = ampacity_corr = 0
            conductor_verificado = conductor_min = "—"
            ocpd_A = "—"; ocpd_status = "—"; proteccion_fmt = "—"
            df_tabla = None
            tierra_calibre = "—"; conduit_inline_result = None
            temp_term_efectiva = 75
    else:
        calc_ok = False
        corriente = 0
        conductor_verificado = "—"
        cdt_real = ampacity_corr = 0
        ocpd_A = "—"; proteccion_fmt = "—"
        ocpd_status = "—"
        df_tabla = None
        tierra_calibre = "—"; conduit_inline_result = None
        temp_term_efectiva = 75
        conductor_min = "—"
        corriente_diseño = 0; corriente_diseño_cond = 0

    # Snapshot de los inputs actuales para invalidar override manual si cambian
    snap_actual = (
        config, mat, voltaje_nominal, potencia, fp, longitud, cf, cdt_max,
        conductores_activos, canalizacion, temp_aislamiento, term_label,
        carga_continua, es_motor, mat_tierra, tierra_auto,
    )

    # ── Aplicar override manual si existe y los inputs no cambiaron ──
    seleccion_manual_activa = False
    if inputs_ok and calc_ok:
        ov = st.session_state.get("manual_override")
        if ov:
            if tuple(ov.get("snap", ())) == snap_actual:
                # Override válido: recalcular CdT y ampacidad para el calibre manual
                try:
                    cdt_ov = calcular_cdt_calibre(
                        ov["calibre"], corriente, longitud / 1000, mat,
                        fp, voltaje_nominal, config, metodo, cf=cf,
                        canalizacion=canalizacion,
                        temp_term=(temp_term_override or temp_term_efectiva),
                    )
                    amp_ov = ampacidad_base(ov["calibre"], mat, temp_aislamiento) * fc_temp * fc_agrup_val
                    ocpd_A_ov, ocpd_status_ov = calcular_proteccion(
                        corriente_diseño_cond, amp_ov
                    )
                    proteccion_fmt_ov = formato_proteccion(ocpd_A_ov, config)
                    tierra_calibre_ov = (
                        calibre_tierra(ocpd_A_ov, mat_tierra) if tierra_auto
                        else tierra_calibre
                    )
                    # Sobrescribir las variables del cálculo automático
                    conductor_verificado = ov["calibre"]
                    cdt_real = cdt_ov
                    ampacity_corr = amp_ov
                    ocpd_A = ocpd_A_ov
                    ocpd_status = ocpd_status_ov
                    proteccion_fmt = proteccion_fmt_ov
                    tierra_calibre = tierra_calibre_ov
                    seleccion_manual_activa = True
                    # Recalcular conduit con nuevo calibre si aplica
                    if incluir_conduit_calc and tipo_tubo_inline:
                        todos_tubos = calcular_conduit_fill(
                            conductor_verificado, num_conds_inline, tipo_tubo_inline
                        )
                        recom = next((x for x in todos_tubos if x["cumple"]),
                                     todos_tubos[-1] if todos_tubos else None)
                        conduit_inline_result = {
                            "params": {
                                "calibre": conductor_verificado,
                                "num_conds": num_conds_inline,
                                "tipo_tubo": tipo_tubo_inline,
                                "aislamiento": tipo_aisl_inline,
                            },
                            "resultados": todos_tubos,
                            "resultados_todos": todos_tubos,
                            "recomendada": recom,
                        }
                except ValueError:
                    # El calibre manual no es válido con los nuevos inputs → invalidar
                    del st.session_state["manual_override"]
            else:
                # Los inputs cambiaron: el override ya no es válido — descartarlo
                del st.session_state["manual_override"]

    # Guardar siempre el último cálculo (con override aplicado si lo había)
    if inputs_ok and calc_ok:
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
            "seleccion_manual": seleccion_manual_activa,
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
            "_snap": snap_actual,
        }
        st.session_state["tabla_conductor"] = df_tabla

    # ══════════════════════════════════════════════
    # PANEL DERECHO — RESULTADOS DESTACADOS
    # ══════════════════════════════════════════════
    with col_results:
        if not inputs_ok:
            st.markdown(
                '<div class="panel-results"><div class="empty-results">'
                '<div class="empty-icon">⚡</div>'
                '<div class="empty-title">Esperando datos válidos</div>'
                '<div class="empty-sub">Completa los parámetros del circuito '
                'para ver los resultados aquí.</div></div></div>',
                unsafe_allow_html=True,
            )
        elif not calc_ok:
            st.markdown(
                f'<div class="panel-results"><div class="empty-results">'
                f'<div class="empty-icon">⚠️</div>'
                f'<div class="empty-title">Error en el cálculo</div>'
                f'<div class="empty-sub">{error_msg}</div></div></div>',
                unsafe_allow_html=True,
            )
        else:
            r = st.session_state["resultado_conductor"]
            cumple_global = (cdt_real <= cdt_max)
            cls = "ok" if cumple_global else "error"
            estado_txt = "Cumple NOM-001" if cumple_global else "No cumple"
            estado_pill = "ok" if cumple_global else "error"
            cf_label = f"{cf}× " if cf > 1 else ""

            # ── Tarjeta GRANDE: conductor recomendado ──
            st.markdown(f"""
            <div class="result-card {cls}">
                <div class="rc-row">
                    <div class="rc-icon">🔌</div>
                    <div style="flex:1;">
                        <div class="rc-label">Conductor recomendado</div>
                        <div class="rc-value {cls}">{cf_label}{conductor_verificado}<span style="font-size:1.1rem;color:var(--text-muted);font-weight:600;letter-spacing:0;">&nbsp;AWG&nbsp;{mat}</span></div>
                        <div class="rc-unit">{r.get('metodo', '').split('(')[0].strip()} · {r.get('configuracion', '').split('(')[0].strip()}</div>
                    </div>
                </div>
                <span class="rc-pill {estado_pill}">★ {estado_txt}</span>
            </div>
            """, unsafe_allow_html=True)

            # ── Grid 2x2 de métricas clave ──
            g1, g2 = st.columns(2)
            color_cdt_cls = "ok" if cdt_real <= cdt_max else "error"
            color_amp_cls = "ok" if ampacity_corr >= corriente_diseño_cond else "error"
            color_ocpd_cls = {
                "OK": "ok", "REVISAR": "warn", "INSUFICIENTE": "error"
            }.get(ocpd_status, "ok")

            with g1:
                st.markdown(f"""
                <div class="result-card {color_cdt_cls}">
                    <div class="rc-label">Caída de tensión</div>
                    <div class="rc-value {color_cdt_cls}">{cdt_real:.2f}<span style="font-size:1rem;color:var(--text-muted);">&nbsp;%</span></div>
                    <div class="rc-unit">Máx. permitida: {cdt_max}%</div>
                </div>
                """, unsafe_allow_html=True)
                st.markdown(f"""
                <div class="result-card {color_amp_cls}">
                    <div class="rc-label">Ampacidad corregida</div>
                    <div class="rc-value {color_amp_cls}">{ampacity_corr:.1f}<span style="font-size:1rem;color:var(--text-muted);">&nbsp;A</span></div>
                    <div class="rc-unit">Requerida: {corriente_diseño_cond:.1f} A/cond</div>
                </div>
                """, unsafe_allow_html=True)
            with g2:
                st.markdown(f"""
                <div class="result-card {color_ocpd_cls}">
                    <div class="rc-label">Protección OCPD</div>
                    <div class="rc-value {color_ocpd_cls}">{proteccion_fmt}<span style="font-size:1rem;color:var(--text-muted);">&nbsp;A</span></div>
                    <div class="rc-unit">{ocpd_A} A · {ocpd_status}</div>
                </div>
                """, unsafe_allow_html=True)
                tierra_origen = "Tabla 250-122" if tierra_auto else "Manual"
                st.markdown(f"""
                <div class="result-card">
                    <div class="rc-label">Conductor de tierra</div>
                    <div class="rc-value">{tierra_calibre}<span style="font-size:1rem;color:var(--text-muted);">&nbsp;AWG&nbsp;{mat_tierra}</span></div>
                    <div class="rc-unit">{tierra_origen} · cable desnudo</div>
                </div>
                """, unsafe_allow_html=True)

            # ── Corriente (info compacta) ──
            st.markdown(f"""
            <div class="metric-card" style="margin-top:8px;">
                <div class="metric-label">Corriente del circuito</div>
                <div class="metric-value">{corriente:.1f} A
                    <span style="font-size:0.85rem;font-weight:500;color:var(--text-muted);">
                    · diseño {corriente_diseño:.1f} A{(f' · {corriente_diseño_cond:.1f} A/cond' if cf > 1 else '')}
                    </span>
                </div>
            </div>
            """, unsafe_allow_html=True)

            # ── Tubería (si aplica) ──
            if conduit_inline_result:
                rec_t = conduit_inline_result["recomendada"]
                cumple_t = rec_t.get("cumple", False)
                t_cls = "ok" if cumple_t else "warn"
                st.markdown(f"""
                <div class="result-card {t_cls}">
                    <div class="rc-label">Tubería recomendada</div>
                    <div class="rc-value {t_cls}">{rec_t['tubo']}</div>
                    <div class="rc-unit">Relleno {rec_t['fill_pct']:.1f}% · Máx. {rec_t.get('fill_max',40)}%</div>
                </div>
                """, unsafe_allow_html=True)

            # ── Recomendaciones automáticas ──
            recoms = []
            if not cumple_global:
                recoms.append(
                    "🔧 La caída de tensión excede el límite. "
                    "Considera <b>subir un calibre</b> o reducir la longitud."
                )
            if ampacity_corr < corriente_diseño_cond:
                recoms.append(
                    "⚡ La ampacidad corregida es insuficiente. "
                    "Aumenta el calibre o reduce el agrupamiento."
                )
            if ocpd_status == "REVISAR":
                recoms.append(
                    "🛡️ El OCPD estándar supera la ampacidad del conductor. "
                    "Sube un calibre para proteger correctamente."
                )
            elif ocpd_status == "INSUFICIENTE":
                recoms.append(
                    "🛡️ La corriente excede todos los OCPD estándar. "
                    "Revisa el diseño o divide el circuito."
                )
            if cumple_global and ampacity_corr >= corriente_diseño_cond * 1.5:
                recoms.append(
                    "💡 Tienes <b>holgura amplia</b> de ampacidad — podrías "
                    "considerar un calibre menor si quieres optimizar costo."
                )
            if cf == 1 and corriente_diseño_cond > 200:
                recoms.append(
                    "🧵 Para corrientes elevadas, evalúa usar "
                    "<b>conductores en paralelo (CF > 1)</b> para reducir calibre."
                )
            if recoms:
                st.markdown('<div class="recom">', unsafe_allow_html=True)
                st.markdown('<div class="recom-title">💡 Recomendaciones</div>',
                            unsafe_allow_html=True)
                for rec in recoms:
                    st.markdown(f'<div class="recom-text">• {rec}</div>',
                                unsafe_allow_html=True)
                st.markdown('</div>', unsafe_allow_html=True)

            # ── Resumen contextual compacto ──
            st.markdown(
                f"<div style='font-size:0.78rem;color:var(--text-muted);margin-top:6px;'>"
                f"📐 Tabla 9 NOM · canalización <b>{canal_label.split('(')[0].strip()}</b> · "
                f"terminales <b>{temp_term_efectiva}°C</b> · "
                f"aislamiento <b>{temp_aislamiento}°C</b>"
                f"</div>",
                unsafe_allow_html=True,
            )

            # ── Expander: tabla completa de candidatos ──
            with st.expander("📊 Ver tabla completa de conductores candidatos",
                             expanded=False):
                if df_tabla is not None:
                    st.dataframe(df_tabla, use_container_width=True, hide_index=True)
                else:
                    st.info("Sin tabla disponible.")

            # ── Expander: selección manual ──
            taisl_m = temp_aislamiento
            calibres_validos = [
                c for c in ORDEN_CALIBRES
                if ampacidad_base(c, mat, taisl_m) > 0
            ] or ORDEN_CALIBRES
            with st.expander("🔧 Selección manual de conductor", expanded=False):
                default_idx = (
                    calibres_validos.index(r["conductor"])
                    if r["conductor"] in calibres_validos else 0
                )
                calibre_manual = st.selectbox(
                    f"Calibre ({mat})",
                    calibres_validos,
                    index=default_idx,
                    key=f"calibre_manual_{mat}_{taisl_m}",
                )
                cf_r = r.get("cf", 1)
                canal_m = r.get("canalizacion", "AC")
                tterm_m = r.get("temp_term", 75)
                L_km_m = r["longitud"] / 1000
                calc_man_ok = True
                try:
                    cdt_m = calcular_cdt_calibre(
                        calibre_manual, r["corriente"], L_km_m, mat,
                        r["fp"], r["voltaje"], r["config_code"], r["metodo_code"],
                        cf=cf_r, canalizacion=canal_m, temp_term=tterm_m,
                    )
                    amp_m = ampacidad_base(calibre_manual, mat, taisl_m) * fc_temp * fc_agrup_val
                except ValueError as ex:
                    st.error(f"⚠️ {ex}")
                    calc_man_ok = False
                    cdt_m, amp_m = 0.0, 0.0
                I_diseño_cond_m = r["corriente"] * r.get("factor_demanda", 1.25) / cf_r

                if calc_man_ok:
                    cumple_m = cdt_m <= r["cdt_max"] and amp_m >= I_diseño_cond_m
                    cls_m = "ok" if cumple_m else "error"
                    st.markdown(f"""
                    <div class="result-card {cls_m}" style="margin:6px 0;">
                        <div style="display:flex;gap:18px;flex-wrap:wrap;">
                            <div>
                                <div class="rc-label">Calibre</div>
                                <div style="font-size:1.4rem;font-weight:700;">{calibre_manual} AWG</div>
                            </div>
                            <div>
                                <div class="rc-label">C.d.T.</div>
                                <div style="font-size:1.4rem;font-weight:700;" class="rc-value {('ok' if cdt_m <= r['cdt_max'] else 'error')}">{cdt_m:.2f}%</div>
                            </div>
                            <div>
                                <div class="rc-label">Ampacidad</div>
                                <div style="font-size:1.4rem;font-weight:700;" class="rc-value {('ok' if amp_m >= I_diseño_cond_m else 'error')}">{amp_m:.1f} A</div>
                            </div>
                        </div>
                    </div>
                    """, unsafe_allow_html=True)
                    manual_pendiente = (calibre_manual != r.get("conductor"))
                    override_activo = bool(st.session_state.get("manual_override"))
                    if manual_pendiente:
                        st.markdown(f"""<div class="resultado-warn">
                            <b>⚠ Selección sin aplicar.</b> El reporte usará
                            <b>{r.get('conductor','—')} AWG</b>. Presiona el
                            botón para usar <b>{calibre_manual} AWG</b>.
                        </div>""", unsafe_allow_html=True)

                    bcol1, bcol2 = st.columns([2, 1])
                    with bcol1:
                        btn_label = (
                            f"✔️ Aplicar {calibre_manual} AWG"
                            if manual_pendiente
                            else f"✓ Aplicado · {calibre_manual} AWG"
                        )
                        if st.button(btn_label, key="btn_aplicar_manual",
                                      type="primary" if manual_pendiente else "secondary",
                                      disabled=not manual_pendiente,
                                      use_container_width=True):
                            _aplicar_seleccion_manual(
                                calibre_manual, snap_actual,
                                cdt_m, amp_m, I_diseño_cond_m,
                                r.get("config_code", "mono2h"),
                                r.get("tierra_material", "Cu"),
                                r.get("tierra_auto", True),
                            )
                            st.rerun()
                    with bcol2:
                        if override_activo:
                            if st.button("↺ Volver a auto",
                                          key="btn_volver_auto",
                                          use_container_width=True):
                                st.session_state.pop("manual_override", None)
                                st.rerun()

    # ══════════════════════════════════════════════
    # SECCIÓN INFERIOR — GUARDAR AL PROYECTO (full width)
    # ══════════════════════════════════════════════
    if "resultado_conductor" in st.session_state and inputs_ok and calc_ok:
        st.markdown("---")
        r = st.session_state["resultado_conductor"]
        rec_t = r.get("tierra_calibre", "—")
        ci = r.get("conduit_inline")
        conduit_resumen = (
            f"{ci['recomendada']['tubo']} · {ci['recomendada']['fill_pct']:.1f}% relleno"
            if ci else "—"
        )

        st.markdown(
            f"""<div class="glass" style="margin-bottom:12px;">
                <div style="font-size:0.78rem;font-weight:700;color:var(--text-muted);
                    text-transform:uppercase;letter-spacing:0.08em;margin-bottom:6px;">
                    📌 Listo para guardar en el proyecto
                </div>
                <div style="display:flex;gap:18px;flex-wrap:wrap;font-size:0.9rem;">
                    <span><b>{r.get('potencia',0):,} W</b> · {r.get('voltaje','—')} V</span>
                    <span>🧵 <b>{r.get('cf',1)}× {r.get('conductor','—')} AWG {r.get('material','—')}</b></span>
                    <span>🛡️ <b>{r.get('proteccion_fmt','—')} A</b></span>
                    <span>⏚ <b>{rec_t} AWG {r.get('tierra_material','Cu')}</b></span>
                    <span>🔧 <b>{conduit_resumen}</b></span>
                </div>
            </div>""",
            unsafe_allow_html=True,
        )

        n_circ = len(st.session_state["circuitos"]) + 1
        col_add1, col_add2 = st.columns([3, 1])
        with col_add1:
            nombre_circ = st.text_input(
                "Nombre del circuito",
                value=f"Circuito {n_circ:02d}",
                key="nombre_circ_input",
                placeholder="Ej. Iluminación Planta Baja",
                label_visibility="collapsed",
            )
        with col_add2:
            agregar_clicked = st.button(
                "➕ Guardar circuito",
                key="btn_add_circ", type="primary",
                use_container_width=True,
            )

        if agregar_clicked:
            if not nombre_circ.strip():
                st.error("⚠️ Asigna un nombre al circuito antes de guardar.")
            else:
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
                    f"✅ **{nombre_circ}** guardado — {total} circuito(s) en el proyecto."
                )

        # Lista compacta de circuitos guardados
        if st.session_state["circuitos"]:
            with st.expander(
                f"📂 Circuitos en el proyecto ({len(st.session_state['circuitos'])})",
                expanded=False,
            ):
                for c in st.session_state["circuitos"]:
                    cdt_v = c.get("cdt", 0) or 0
                    cdt_mx = c.get("cdt_max", 3.0)
                    icon = "✓" if cdt_v <= cdt_mx else "✗"
                    color_var = "var(--ok-text)" if cdt_v <= cdt_mx else "var(--err-text)"
                    st.markdown(
                        f"<div style='padding:4px 0;font-size:0.9rem;color:var(--text-main);'>"
                        f"<span style='color:{color_var};font-weight:700;'>{icon}</span> "
                        f"<b>#{c.get('id','?')} {c.get('nombre','')}</b> · "
                        f"{c.get('potencia','—')} W · {c.get('cf',1)}×{c.get('conductor','—')} {c.get('material','')} · "
                        f"OCPD {c.get('proteccion_fmt', str(c.get('proteccion_A','—'))+' A')}"
                        f"</div>",
                        unsafe_allow_html=True,
                    )

# ══════════════════════════════════════════════
# TAB 2 — CONDUIT FILL
# ══════════════════════════════════════════════
with tab2:
    st.markdown(
        '<div style="font-size:0.78rem;font-weight:700;color:var(--text-muted);'
        'text-transform:uppercase;letter-spacing:0.08em;margin:4px 0 12px 0;">'
        '🔧 Conduit Fill · NOM-001-SEDE-2012 Cap. 3 (máx. 40% para 3+ conductores)'
        '</div>',
        unsafe_allow_html=True,
    )

    col_in_t, col_res_t = st.columns([1, 1], gap="large")

    # ── Inputs ──────────────────────────────────
    with col_in_t:
        st.markdown('<div class="subsection-header">Conductores forrados (fases + neutro)</div>',
                    unsafe_allow_html=True)
        tipo_tubo = st.selectbox(
            "Tipo de tubería",
            list(TABLA_CONDUIT.keys()),
            key="t2_tipo_tubo",
        )

        ct_a, ct_b = st.columns(2)
        with ct_a:
            calibre_conduit = st.selectbox(
                "Calibre del conductor forrado",
                list(TABLA_CONDUCTORES.keys()),
                key="t2_calibre",
            )
        with ct_b:
            num_conductores = st.number_input(
                "# conductores forrados",
                min_value=0, max_value=40, value=3,
                key="t2_nconds",
                help="Cantidad de conductores con aislamiento (fases + neutro).",
            )

        tipo_aislamiento = st.selectbox(
            "Tipo de aislamiento",
            ["THW (75°C)", "THHW (75/90°C)", "THWN-2 (90°C)", "XHHW-2 (90°C)"],
            key="t2_aisl",
        )

        st.markdown('<div class="subsection-header">Conductor de tierra (cable desnudo)</div>',
                    unsafe_allow_html=True)
        ct_t1, ct_t2 = st.columns(2)
        with ct_t1:
            incluir_tierra_t2 = st.checkbox(
                "Incluir tierra en la tubería",
                value=True,
                key="t2_incluir_tierra",
                help="Suma el área del cable de tierra desnudo al cálculo de relleno.",
            )
        with ct_t2:
            n_tierras_t2 = st.number_input(
                "# cables de tierra",
                min_value=0, max_value=10,
                value=1 if incluir_tierra_t2 else 0,
                key="t2_n_tierras",
                disabled=not incluir_tierra_t2,
            )
        if incluir_tierra_t2:
            ct_t3, ct_t4 = st.columns(2)
            with ct_t3:
                calibre_tierra_t2 = st.selectbox(
                    "Calibre del cable de tierra",
                    ORDEN_CALIBRES,
                    index=ORDEN_CALIBRES.index("8"),
                    key="t2_cal_tierra",
                )
            with ct_t4:
                material_tierra_t2 = st.selectbox(
                    "Material del cable de tierra",
                    ["Cu (cobre desnudo)", "Al (aluminio)"],
                    key="t2_mat_tierra",
                )
        else:
            calibre_tierra_t2 = None
            material_tierra_t2 = None

        st.caption(
            "ℹ️ Relleno máximo NOM-001: 53% (1 cond.), 31% (2 conds.), "
            "40% (3 o más). El cable de tierra desnudo se suma al área total."
        )

    # ── Cálculo: aceptar mix de conductores con áreas distintas ─
    od_forrado = TABLA_CONDUCTORES[calibre_conduit]["od_mm"]
    area_forrados = math.pi * (od_forrado / 2) ** 2 * num_conductores
    area_tierra_total = 0
    n_total_conds = num_conductores
    if incluir_tierra_t2 and calibre_tierra_t2 and n_tierras_t2 > 0:
        od_tierra = TABLA_CONDUCTORES[calibre_tierra_t2]["od_mm"]
        area_tierra_total = math.pi * (od_tierra / 2) ** 2 * n_tierras_t2
        n_total_conds += n_tierras_t2

    # Calcular usando un calibre "equivalente" — pero ajustamos el área manualmente:
    # Llamamos calcular_conduit_fill solo para obtener las áreas de las tuberías,
    # luego reemplazamos area_conds por la suma real.
    resultado_conduit = calcular_conduit_fill(calibre_conduit, num_conductores, tipo_tubo)
    area_conds_total = area_forrados + area_tierra_total
    fill_max = 53 if n_total_conds == 1 else (31 if n_total_conds == 2 else 40)
    primer_ok = None
    for r in resultado_conduit:
        r["area_conds"] = round(area_conds_total, 2)
        r["fill_pct"]   = round(area_conds_total / r["area_tubo"] * 100, 1)
        r["fill_max"]   = fill_max
        r["cumple"]     = r["fill_pct"] <= fill_max
        if r["cumple"] and primer_ok is None:
            primer_ok = r["diametro"]
        r["recomendada"] = (r["diametro"] == primer_ok)

    recomendada = next((r for r in resultado_conduit if r.get("recomendada")), None)
    st.session_state["resultado_conduit"] = resultado_conduit
    st.session_state["conduit_params"] = {
        "calibre": calibre_conduit,
        "num_conds": n_total_conds,
        "tipo_tubo": tipo_tubo,
        "aislamiento": tipo_aislamiento,
        "calibre_tierra": calibre_tierra_t2 if incluir_tierra_t2 else None,
        "n_tierras": n_tierras_t2 if incluir_tierra_t2 else 0,
    }

    # ── Resultados ──────────────────────────────
    with col_res_t:
        if recomendada:
            rec_cls = "ok" if recomendada.get("cumple") else "warn"
            tierra_resumen = (
                f" + {n_tierras_t2}× {calibre_tierra_t2} AWG tierra"
                if incluir_tierra_t2 and calibre_tierra_t2 else ""
            )
            st.markdown(f"""
            <div class="result-card {rec_cls}">
                <div class="rc-row">
                    <div class="rc-icon">🔧</div>
                    <div style="flex:1;">
                        <div class="rc-label">Tubería recomendada</div>
                        <div class="rc-value {rec_cls}">{recomendada['tubo']}</div>
                        <div class="rc-unit">
                            {num_conductores} cond. {calibre_conduit} AWG{tierra_resumen} · {tipo_aislamiento}
                        </div>
                    </div>
                </div>
                <span class="rc-pill {rec_cls}">
                    {'★ CUMPLE' if recomendada.get('cumple') else '⚠ REVISAR'}
                </span>
            </div>
            """, unsafe_allow_html=True)

            g1, g2 = st.columns(2)
            with g1:
                fill_cls = "ok" if recomendada["fill_pct"] <= recomendada["fill_max"] else "error"
                st.markdown(f"""
                <div class="result-card {fill_cls}">
                    <div class="rc-label">Relleno real</div>
                    <div class="rc-value {fill_cls}">{recomendada['fill_pct']:.1f}<span style="font-size:1rem;color:var(--text-muted);">&nbsp;%</span></div>
                    <div class="rc-unit">Máx. permitido: {recomendada['fill_max']}%</div>
                </div>
                """, unsafe_allow_html=True)
            with g2:
                st.markdown(f"""
                <div class="result-card">
                    <div class="rc-label">Área ocupada</div>
                    <div class="rc-value">{recomendada['area_conds']:.1f}<span style="font-size:1rem;color:var(--text-muted);">&nbsp;mm²</span></div>
                    <div class="rc-unit">de {recomendada['area_tubo']:.1f} mm² disponibles</div>
                </div>
                """, unsafe_allow_html=True)
        else:
            st.markdown(f"""
            <div class="result-card error">
                <div class="rc-label">Sin tubería que cumpla</div>
                <div class="rc-value error">⚠</div>
                <div class="rc-unit">Considera otra familia o reducir N conductores.</div>
            </div>
            """, unsafe_allow_html=True)

        # Recomendaciones automáticas
        recoms_t = []
        if not recomendada or not recomendada.get("cumple"):
            recoms_t.append(
                "Aumenta el diámetro de la tubería o cambia a una familia más "
                "grande (RGS / IMC tienen mayor área interna que EMT)."
            )
        # Tubería un tamaño mayor que la recomendada — holgura
        if recomendada and recomendada.get("fill_pct", 0) > 35:
            recoms_t.append(
                "Estás cerca del límite del 40%. Considera el siguiente tamaño "
                "para tener margen de mantenimiento y expansión futura."
            )
        if recoms_t:
            st.markdown('<div class="recom">', unsafe_allow_html=True)
            st.markdown('<div class="recom-title">💡 Recomendaciones</div>',
                        unsafe_allow_html=True)
            for rc in recoms_t:
                st.markdown(f'<div class="recom-text">• {rc}</div>',
                            unsafe_allow_html=True)
            st.markdown('</div>', unsafe_allow_html=True)

    # ── Tabla completa de medidas (debajo, full width) ──
    with st.expander("📊 Tabla completa de medidas (todas las tuberías)",
                     expanded=False):
        df_conduit = pd.DataFrame([
            {
                "Tubería": r["tubo"],
                "Ø nominal": r.get("diametro", "—"),
                "Área int. (mm²)": r["area_tubo"],
                "Área conds. (mm²)": r["area_conds"],
                "Relleno (%)": r["fill_pct"],
                "Máx. (%)": r["fill_max"],
                "Cumple": "Sí" if r["cumple"] else "No",
                "Selección": "Rec." if r.get("recomendada") else "",
            }
            for r in resultado_conduit
        ])
        st.dataframe(df_conduit, use_container_width=True, hide_index=True)

# ══════════════════════════════════════════════
# TAB 3 — PANEL DE CIRCUITOS
# ══════════════════════════════════════════════
with tab3:
    st.markdown(
        '<div style="font-size:0.78rem;font-weight:700;color:var(--text-muted);'
        'text-transform:uppercase;letter-spacing:0.08em;margin:4px 0 12px 0;">'
        '📊 Cuadro de cargas — balanceo de fases · demanda total · NOM-001'
        '</div>',
        unsafe_allow_html=True,
    )

    circuitos = st.session_state["circuitos"]

    if not circuitos:
        st.markdown(
            '<div class="glass-strong"><div class="empty-results">'
            '<div class="empty-icon">📋</div>'
            '<div class="empty-title">No hay circuitos en el proyecto</div>'
            '<div class="empty-sub">Ve a la pestaña <b>🔌 Conductor & C.d.T.</b>, '
            'configura un circuito y haz clic en <b>➕ Guardar circuito</b>. '
            'Los circuitos guardados aparecerán aquí.</div>'
            '</div></div>',
            unsafe_allow_html=True,
        )
    else:
        # ── Asignar fases y calcular balanceo ─────────────────
        circs_fases = asignar_fases(circuitos)
        bal = resumen_balanceo(circs_fases)

        def _cumple(c):
            cdt = c.get("cdt")
            cdt_max = c.get("cdt_max", 3.0)
            return (cdt is not None) and (cdt <= cdt_max)

        ok_count = sum(1 for c in circuitos if _cumple(c))
        todos_cumplen = (ok_count == len(circuitos))
        total_kw = bal["P_total"] / 1000
        fp_prom = sum((c.get("fp", 0.9) or 0.9) for c in circuitos) / max(len(circuitos), 1)
        S_kva = total_kw / fp_prom if fp_prom > 0 else 0

        # ── KPIs principales (4 cards) ────────────────────────
        sm1, sm2, sm3, sm4 = st.columns(4)
        with sm1:
            st.markdown(f"""
            <div class="result-card">
                <div class="rc-label">Circuitos</div>
                <div class="rc-value">{len(circuitos)}</div>
                <div class="rc-unit">{ok_count} cumplen · {len(circuitos)-ok_count} revisar</div>
            </div>""", unsafe_allow_html=True)
        with sm2:
            st.markdown(f"""
            <div class="result-card">
                <div class="rc-label">Carga instalada</div>
                <div class="rc-value">{total_kw:.2f}<span style="font-size:1rem;color:var(--text-muted);">&nbsp;kW</span></div>
                <div class="rc-unit">{S_kva:.2f} kVA · FP prom {fp_prom:.2f}</div>
            </div>""", unsafe_allow_html=True)
        with sm3:
            # Balanceo: verde <10%, amarillo <20%, rojo >=20%
            d = bal["desbalance_pct"]
            cb = "ok" if d < 10 else ("warn" if d < 20 else "error")
            st.markdown(f"""
            <div class="result-card {cb}">
                <div class="rc-label">Desbalance entre fases</div>
                <div class="rc-value {cb}">{d:.1f}<span style="font-size:1rem;color:var(--text-muted);">&nbsp;%</span></div>
                <div class="rc-unit">{bal['calidad']} (NEMA)</div>
            </div>""", unsafe_allow_html=True)
        with sm4:
            est_cls = "ok" if todos_cumplen else "warn"
            est_val = "OK" if todos_cumplen else "Revisar"
            st.markdown(f"""
            <div class="result-card {est_cls}">
                <div class="rc-label">Cumplimiento NOM-001</div>
                <div class="rc-value {est_cls}">{est_val}</div>
                <div class="rc-unit">{ok_count}/{len(circuitos)} circuitos</div>
            </div>""", unsafe_allow_html=True)

        st.markdown("")

        # ── Distribución por fases (visual) ───────────────────
        st.markdown('<div class="subsection-header">Distribución por fase</div>',
                    unsafe_allow_html=True)
        cf1, cf2, cf3 = st.columns(3)
        for col, (fase, val_W, color_bg) in zip(
            (cf1, cf2, cf3),
            [("R", bal["P_R"], "#FFE4E1"),
             ("S", bal["P_S"], "#E4F1FE"),
             ("T", bal["P_T"], "#FFF4E4")]
        ):
            pct = (val_W / bal["P_total"] * 100) if bal["P_total"] > 0 else 0
            with col:
                st.markdown(f"""
                <div class="result-card" style="background:{color_bg};">
                    <div class="rc-row">
                        <div class="rc-icon" style="background:white;">{fase}</div>
                        <div style="flex:1;">
                            <div class="rc-label">Fase {fase}</div>
                            <div class="rc-value">{val_W/1000:.2f}<span style="font-size:0.95rem;color:var(--text-muted);">&nbsp;kW</span></div>
                            <div class="rc-unit">{pct:.1f}% del total · {int(val_W):,} W</div>
                        </div>
                    </div>
                </div>
                """, unsafe_allow_html=True)

        # Recomendaciones de balanceo
        if bal["desbalance_pct"] >= 10:
            recom_bal = []
            if bal["desbalance_pct"] >= 20:
                recom_bal.append(
                    "El desbalance supera el 20%. Es <b>obligatorio reasignar circuitos</b>. "
                    "Considera mover los circuitos monofásicos más grandes a la fase con menor carga."
                )
            elif bal["desbalance_pct"] >= 10:
                recom_bal.append(
                    "Desbalance en zona regular (10-20%). Reasigna circuitos para acercarte a un balanceo óptimo (< 10%)."
                )
            fases_sorted = sorted(
                [("R", bal["P_R"]), ("S", bal["P_S"]), ("T", bal["P_T"])],
                key=lambda x: x[1]
            )
            recom_bal.append(
                f"Fase con menor carga: <b>{fases_sorted[0][0]}</b> ({fases_sorted[0][1]/1000:.2f} kW). "
                f"Fase con mayor carga: <b>{fases_sorted[-1][0]}</b> ({fases_sorted[-1][1]/1000:.2f} kW). "
                f"Diferencia: {(fases_sorted[-1][1]-fases_sorted[0][1])/1000:.2f} kW."
            )
            st.markdown('<div class="recom">', unsafe_allow_html=True)
            st.markdown('<div class="recom-title">💡 Recomendaciones de balanceo</div>',
                        unsafe_allow_html=True)
            for rc in recom_bal:
                st.markdown(f'<div class="recom-text">• {rc}</div>',
                            unsafe_allow_html=True)
            st.markdown('</div>', unsafe_allow_html=True)

        st.markdown("")

        # ── Cuadro de cargas (tabla principal) ────────────────
        st.markdown('<div class="subsection-header">Cuadro de cargas</div>',
                    unsafe_allow_html=True)
        filas_cc = []
        for i, c in enumerate(circs_fases):
            cf = c.get("cf", 1)
            cond = c.get("conductor", "—")
            cond_str = f"{cf}×{cond}" if cf > 1 else cond
            rec = (c.get("conduit_info") or {}).get("recomendada") or {}
            cdt = c.get("cdt")
            cdt_max = c.get("cdt_max", 3.0)
            corriente = c.get("corriente", 0) or 0
            filas_cc.append({
                "No.": c.get("id", i + 1),
                "Descripción": c.get("nombre", f"Circuito {i+1}"),
                "P (W)": int(c.get("potencia", 0) or 0),
                "V (V)": c.get("voltaje", "—"),
                "FP": c.get("fp", 0.9),
                "Sistema": (c.get("configuracion", "—") or "—").split("(")[0].strip(),
                "Fases": "+".join(c.get("fases", [])),
                "P_R (W)": int(c.get("P_R", 0)),
                "P_S (W)": int(c.get("P_S", 0)),
                "P_T (W)": int(c.get("P_T", 0)),
                "I (A)": f"{corriente:.1f}",
                "Conductor": f"{cond_str} AWG {c.get('material', '—')}",
                "Tierra": f"{c.get('tierra_calibre','—')} {c.get('tierra_material','Cu')}",
                "Protección": str(c.get('proteccion_fmt', c.get('proteccion_A','—'))),
                "C.d.T.": f"{cdt:.2f}%" if cdt is not None else "—",
                "Tubería": rec.get("tubo", "—"),
                "Estado": "Cumple" if _cumple(c) else "Revisar",
            })
        st.dataframe(pd.DataFrame(filas_cc), use_container_width=True, hide_index=True)

        # ── Demanda total ────────────────────────────────────
        st.markdown('<div class="subsection-header">Demanda total del tablero</div>',
                    unsafe_allow_html=True)
        cd1, cd2, cd3 = st.columns(3)
        with cd1:
            st.markdown(f"""
            <div class="metric-card">
                <div class="metric-label">Carga total instalada</div>
                <div class="metric-value">{total_kw:.2f} kW</div>
                <div class="metric-unit">{int(bal['P_total']):,} W</div>
            </div>""", unsafe_allow_html=True)
        with cd2:
            st.markdown(f"""
            <div class="metric-card">
                <div class="metric-label">Potencia aparente</div>
                <div class="metric-value">{S_kva:.2f} kVA</div>
                <div class="metric-unit">FP promedio {fp_prom:.2f}</div>
            </div>""", unsafe_allow_html=True)
        with cd3:
            # Corriente de la fase más cargada (criterio dimensionante)
            I_max_fase = max(bal["P_R"], bal["P_S"], bal["P_T"])
            V_fase = 127  # estimación V_F-N estándar MX
            I_demanda = I_max_fase / V_fase if V_fase > 0 else 0
            st.markdown(f"""
            <div class="metric-card">
                <div class="metric-label">I por fase (máx.)</div>
                <div class="metric-value">{I_demanda:.1f} A</div>
                <div class="metric-unit">P_max / V_F-N estimada</div>
            </div>""", unsafe_allow_html=True)

        st.markdown("")

        # ── Acciones: exportar Excel con fórmulas / JSON / Limpiar ──
        ca1, ca2, ca3, ca4 = st.columns(4)
        datos_proy_export = {
            "proyecto": proyecto, "cliente": cliente, "ubicacion": ubicacion,
            "responsable": responsable, "cedula": cedula, "fecha": str(fecha_proj),
            "norma": norma, "sistema": tipo_sistema, "temp_ambiente": temp_ambiente,
        }

        with ca1:
            try:
                excel_bytes = generar_excel_cuadro_cargas(circuitos, datos_proy_export)
                st.download_button(
                    "📊 Exportar Excel (con fórmulas)",
                    excel_bytes,
                    f"CuadroCargas_{proyecto.replace(' ','_')}.xlsx",
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
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
                use_container_width=True,
            )

        with ca3:
            if st.button("🗑️ Limpiar todos", key="btn_limpiar_todos",
                          use_container_width=True):
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
                if st.button("🗑️ Eliminar", key="btn_borrar_uno",
                              use_container_width=True):
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
    st.markdown(
        '<div style="font-size:0.78rem;font-weight:700;color:var(--text-muted);'
        'text-transform:uppercase;letter-spacing:0.08em;margin:4px 0 12px 0;">'
        '📄 Generador de reporte — Memoria técnica profesional'
        '</div>',
        unsafe_allow_html=True,
    )

    sin_resultado = "resultado_conductor" not in st.session_state
    sin_circuitos = not st.session_state["circuitos"]

    if sin_resultado and sin_circuitos:
        st.markdown(
            '<div class="glass-strong"><div class="empty-results">'
            '<div class="empty-icon">📄</div>'
            '<div class="empty-title">No hay datos para reportar</div>'
            '<div class="empty-sub">Calcula al menos un circuito en la pestaña '
            '<b>🔌 Conductor & C.d.T.</b> antes de generar el reporte.</div>'
            '</div></div>',
            unsafe_allow_html=True,
        )
    else:
        datos_proyecto = {
            "proyecto": proyecto, "cliente": cliente, "ubicacion": ubicacion,
            "responsable": responsable, "cedula": cedula, "fecha": str(fecha_proj),
            "norma": norma, "sistema": tipo_sistema, "temp_ambiente": temp_ambiente,
            "descripcion": descripcion_proy,
            "normas_extra": normas_extra,
        }

        col_cfg, col_action = st.columns([3, 2], gap="large")

        with col_cfg:
            st.markdown('<div class="subsection-header">Configuración del reporte</div>',
                        unsafe_allow_html=True)

            # ── Modo de reporte ─────────────────────────────
            modos_disponibles = []
            if not sin_circuitos:
                modos_disponibles.append(
                    f"Proyecto completo ({len(st.session_state['circuitos'])} circuito(s))"
                )
            if not sin_resultado:
                modos_disponibles.append("Circuito actual (memoria detallada)")

            modo = st.radio("Modo", modos_disponibles, horizontal=False, key="modo_rep")
            es_proyecto = "Proyecto completo" in modo

            # ── Formato ─────────────────────────────────────
            formato_lbl = st.radio(
                "Formato",
                ["📄 PDF — profesional, no editable",
                 "📝 Word .docx — editable"],
                horizontal=False, key="fmt_rep",
            )
            es_word = "Word" in formato_lbl

            if not es_proyecto:
                incluir_tabla = st.checkbox(
                    "Incluir tabla completa de conductores", value=True
                )
                incluir_conduit = st.checkbox(
                    "Incluir cálculo de tubería",
                    value="resultado_conduit" in st.session_state,
                )
            else:
                incluir_tabla = True
                incluir_conduit = True

        with col_action:
            st.markdown('<div class="subsection-header">Contenido del reporte</div>',
                        unsafe_allow_html=True)
            if es_proyecto:
                contenido = """
                <div class="glass" style="font-size:0.88rem;line-height:1.6;">
                <b>Reporte de proyecto incluye:</b><br>
                • Portada · Datos del proyecto<br>
                • Índice clickeable<br>
                • Introducción · Objetivo · Descripción<br>
                • Normatividad aplicable<br>
                • Criterios de diseño<br>
                • <b>Tabla resumen</b> de todos los circuitos<br>
                • Memoria por circuito con fórmulas<br>
                • Conclusiones · Firma
                </div>
                """
            else:
                contenido = """
                <div class="glass" style="font-size:0.88rem;line-height:1.6;">
                <b>Reporte de circuito incluye:</b><br>
                • Portada · Datos del proyecto<br>
                • Criterios de diseño y normativa<br>
                • Memoria de cálculo paso a paso<br>
                • Tabla de conductores candidatos<br>
                • Cálculo de tubería (si aplica)<br>
                • Conclusiones · Firma
                </div>
                """
            st.markdown(contenido, unsafe_allow_html=True)

            btn_label = (
                f"{'📝 Generar Word' if es_word else '📄 Generar PDF'}"
                f" — {'Proyecto' if es_proyecto else 'Circuito'}"
            )
            generar_clicked = st.button(
                btn_label, type="primary", key="btn_doc_unified",
                use_container_width=True,
            )

        # ── Generación ──────────────────────────────────────
        if generar_clicked:
            with st.spinner(f"Generando reporte ({'Word' if es_word else 'PDF'})..."):
                if es_proyecto:
                    rc_default = st.session_state.get(
                        "resultado_conductor",
                        st.session_state["circuitos"][0],
                    )
                    if es_word:
                        doc_bytes = generar_reporte_docx(
                            datos_proyecto, rc_default,
                            circuitos=st.session_state["circuitos"],
                        )
                    else:
                        doc_bytes = generar_reporte_pdf(
                            datos_proyecto, rc_default,
                            circuitos=st.session_state["circuitos"],
                        )
                    suffix = "_Proyecto"
                else:
                    tabla_conductor = st.session_state.get("tabla_conductor") if incluir_tabla else None
                    resultado_conduit = st.session_state.get("resultado_conduit") if incluir_conduit else None
                    conduit_params = st.session_state.get("conduit_params") if incluir_conduit else None
                    if es_word:
                        doc_bytes = generar_reporte_docx(
                            datos_proyecto,
                            st.session_state["resultado_conductor"],
                            tabla_conductor, resultado_conduit, conduit_params,
                        )
                    else:
                        doc_bytes = generar_reporte_pdf(
                            datos_proyecto,
                            st.session_state["resultado_conductor"],
                            tabla_conductor, resultado_conduit, conduit_params,
                        )
                    suffix = ""

                if es_word:
                    ext, mime = "docx", (
                        "application/vnd.openxmlformats-officedocument."
                        "wordprocessingml.document"
                    )
                else:
                    ext, mime = "pdf", "application/pdf"

            nombre = f"eSuit_{proyecto.replace(' ','_')}{suffix}.{ext}"
            st.markdown(
                f'<div class="resultado-ok">'
                f'✓ Reporte generado correctamente — {len(doc_bytes):,} bytes'
                f'</div>',
                unsafe_allow_html=True,
            )
            st.download_button(
                f"⬇️ Descargar {nombre}",
                doc_bytes, nombre, mime, type="primary",
                use_container_width=True,
            )

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
