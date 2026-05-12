"""
reporte.py — Generación de reporte PDF profesional
NOM-001-SEDE-2012 / CFE / NMX

Estructura del PDF:
  Portada → Índice (clickeable) → Introducción → Objetivo →
  Descripción → Normatividad → Criterios → Memoria por circuito
  (con desarrollo paso a paso de fórmulas) → Tablero → Conclusiones → Firma

Tipografía:
  Cuerpo:    Garamond 12, interlineado 1.5, justificado, negro
  Títulos:   Courier New 16, azul
  Subtítulos: Courier New 14, azul
  Fórmulas:  Courier New (o Cambria Math), color azul oscuro
"""
import io
import os
import math
from datetime import date

from reportlab.lib.pagesizes import letter
from reportlab.lib import colors
from reportlab.lib.units import cm
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_JUSTIFY
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle,
    HRFlowable, PageBreak, KeepTogether,
)
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

from calculos import (
    TABLA_9_NOM, TABLA_CONDUCTORES, r_x_efectivas, ampacidad_base,
    temp_terminales_auto, formato_proteccion, calibre_tierra,
)

# ═════════════════════════════════════════════════════════
# PALETA DE COLORES
# ═════════════════════════════════════════════════════════
C_AZUL_TIT    = colors.HexColor("#1f3864")  # azul títulos / subtítulos
C_AZUL_FORM   = colors.HexColor("#0f2a5a")  # azul fórmulas
C_AZUL_OSC    = colors.HexColor("#0b1e3f")  # portada
C_AZUL_BAR    = colors.HexColor("#2d4a7a")  # barra portada
C_NEGRO_CUERPO = colors.HexColor("#1a1a1a")  # negro suave para body
C_GRIS_LINEA  = colors.HexColor("#dcdcdc")
C_GRIS_FONDO  = colors.HexColor("#f5f5f7")
C_VERDE       = colors.HexColor("#1d6e3a")
C_ROJO        = colors.HexColor("#a31621")
C_SUBTXT      = colors.HexColor("#5b5b5b")
C_NARANJA     = colors.HexColor("#c2570b")
C_AZUL_LINK   = colors.HexColor("#0066cc")


# ═════════════════════════════════════════════════════════
# REGISTRO DE TIPOGRAFÍAS
# ═════════════════════════════════════════════════════════
def _try_register(name, paths):
    for p in paths:
        if os.path.exists(p):
            try:
                pdfmetrics.registerFont(TTFont(name, p))
                return True
            except Exception:
                pass
    return False


def register_fonts():
    """Registra Garamond y Courier New del sistema si están disponibles.
       Fallback: Times-Roman / Courier.
    """
    fonts = {
        "BODY":         "Times-Roman",
        "BODY_BOLD":    "Times-Bold",
        "BODY_ITALIC":  "Times-Italic",
        "TITULO":       "Courier-Bold",
        "TITULO_PLAIN": "Courier",
        "FORMULA":      "Courier",
        "FORMULA_BOLD": "Courier-Bold",
    }
    win = "C:/Windows/Fonts"
    if _try_register("Garamond",        [f"{win}/GARA.TTF",   f"{win}/garamond.ttf"]):
        fonts["BODY"] = "Garamond"
    if _try_register("Garamond-Bold",   [f"{win}/GARABD.TTF"]):
        fonts["BODY_BOLD"] = "Garamond-Bold"
    if _try_register("Garamond-Italic", [f"{win}/GARAIT.TTF"]):
        fonts["BODY_ITALIC"] = "Garamond-Italic"
    if _try_register("CourierNew",      [f"{win}/cour.ttf"]):
        fonts["FORMULA"] = "CourierNew"
        fonts["TITULO_PLAIN"] = "CourierNew"
    if _try_register("CourierNew-Bold", [f"{win}/courbd.ttf"]):
        fonts["TITULO"] = "CourierNew-Bold"
        fonts["FORMULA_BOLD"] = "CourierNew-Bold"
    return fonts


FONTS = register_fonts()


# ═════════════════════════════════════════════════════════
# ESTILOS
# ═════════════════════════════════════════════════════════
def build_styles():
    L = 18  # interlineado 1.5 sobre 12pt
    return {
        "portada_titulo": ParagraphStyle(
            "portada_titulo",
            fontSize=24, fontName=FONTS["TITULO"],
            textColor=C_AZUL_OSC, alignment=TA_CENTER,
            spaceAfter=8, leading=30,
        ),
        "portada_subtitulo": ParagraphStyle(
            "portada_subtitulo",
            fontSize=14, fontName=FONTS["BODY_ITALIC"],
            textColor=C_AZUL_BAR, alignment=TA_CENTER,
            spaceAfter=6, leading=18,
        ),
        "portada_info": ParagraphStyle(
            "portada_info",
            fontSize=11, fontName=FONTS["BODY"],
            textColor=C_SUBTXT, alignment=TA_CENTER, spaceAfter=3,
        ),

        # Títulos: Courier New 16, azul
        "titulo": ParagraphStyle(
            "titulo",
            fontSize=16, fontName=FONTS["TITULO"],
            textColor=C_AZUL_TIT, spaceBefore=18, spaceAfter=10,
            leading=20,
        ),
        # Subtítulos: Courier New 14, azul
        "subtitulo": ParagraphStyle(
            "subtitulo",
            fontSize=14, fontName=FONTS["TITULO"],
            textColor=C_AZUL_TIT, spaceBefore=12, spaceAfter=6,
            leading=18,
        ),
        # Sub-sub para 7.1.1, etc.
        "subsub": ParagraphStyle(
            "subsub",
            fontSize=12, fontName=FONTS["TITULO_PLAIN"],
            textColor=C_AZUL_TIT, spaceBefore=8, spaceAfter=4,
            leading=15,
        ),

        # Cuerpo: Garamond 12, interlineado 1.5, justificado negro
        "cuerpo": ParagraphStyle(
            "cuerpo",
            fontSize=12, fontName=FONTS["BODY"],
            textColor=C_NEGRO_CUERPO, leading=L,
            alignment=TA_JUSTIFY, spaceAfter=6,
        ),
        "cuerpo_left": ParagraphStyle(
            "cuerpo_left",
            fontSize=12, fontName=FONTS["BODY"],
            textColor=C_NEGRO_CUERPO, leading=L,
            alignment=TA_LEFT, spaceAfter=6,
        ),

        # Fórmulas: Courier New 11, azul oscuro
        "formula": ParagraphStyle(
            "formula",
            fontSize=11, fontName=FONTS["FORMULA"],
            textColor=C_AZUL_FORM, leading=15,
            alignment=TA_LEFT, leftIndent=22, spaceBefore=4, spaceAfter=4,
            backColor=colors.HexColor("#f0f4fa"),
            borderColor=colors.HexColor("#d6e0f0"),
            borderWidth=0.5, borderPadding=6, borderRadius=4,
        ),

        # Resultados
        "resultado_ok": ParagraphStyle(
            "resultado_ok",
            fontSize=11, fontName=FONTS["BODY_BOLD"],
            textColor=C_VERDE, leading=15, spaceBefore=4, spaceAfter=6,
        ),
        "resultado_err": ParagraphStyle(
            "resultado_err",
            fontSize=11, fontName=FONTS["BODY_BOLD"],
            textColor=C_ROJO, leading=15, spaceBefore=4, spaceAfter=6,
        ),

        "pie": ParagraphStyle(
            "pie",
            fontSize=9, fontName=FONTS["BODY"],
            textColor=C_SUBTXT, alignment=TA_CENTER, leading=12,
        ),
        "firma": ParagraphStyle(
            "firma",
            fontSize=11, fontName=FONTS["BODY"],
            textColor=C_NEGRO_CUERPO, alignment=TA_CENTER, spaceAfter=2,
        ),

        # Estilos para índice (más compactos, link)
        "toc_l1": ParagraphStyle(
            "toc_l1",
            fontSize=12, fontName=FONTS["TITULO_PLAIN"],
            textColor=C_AZUL_TIT, leading=18, spaceAfter=2,
        ),
        "toc_l2": ParagraphStyle(
            "toc_l2",
            fontSize=11, fontName=FONTS["BODY"],
            textColor=C_NEGRO_CUERPO, leading=15, leftIndent=18, spaceAfter=1,
        ),

        # Para tablas
        "tab_key": ParagraphStyle(
            "tab_key", fontSize=10, fontName=FONTS["BODY_BOLD"],
            textColor=C_AZUL_TIT, leading=14,
        ),
        "tab_val": ParagraphStyle(
            "tab_val", fontSize=10, fontName=FONTS["BODY"],
            textColor=C_NEGRO_CUERPO, leading=14,
        ),
        "tab_h": ParagraphStyle(
            "tab_h", fontSize=10, fontName=FONTS["TITULO_PLAIN"],
            textColor=colors.white, leading=13, alignment=TA_CENTER,
        ),
        "tab_c": ParagraphStyle(
            "tab_c", fontSize=9, fontName=FONTS["BODY"],
            textColor=C_NEGRO_CUERPO, leading=12, alignment=TA_CENTER,
        ),
    }


# ═════════════════════════════════════════════════════════
# ANCHORS / OUTLINE (índice clickeable + bookmarks PDF)
# ═════════════════════════════════════════════════════════
class OutlinedDocTemplate(SimpleDocTemplate):
    """DocTemplate que captura entradas <anchor name=...> y arma el outline (bookmarks)."""
    def __init__(self, filename, **kw):
        SimpleDocTemplate.__init__(self, filename, **kw)
        self._toc_entries = []  # [(level, label, anchor), ...]

    def afterFlowable(self, flowable):
        # Si el flowable es un Paragraph con un bookmark registrado
        if hasattr(flowable, "_bookmark"):
            level, label, anchor = flowable._bookmark
            self.canv.bookmarkPage(anchor)
            self.canv.addOutlineEntry(label, anchor, level=level, closed=(level > 0))


def _titulo_con_anchor(texto, style, anchor, level=0):
    """Paragraph con anclaje interno + entrada en outline."""
    p = Paragraph(f'<a name="{anchor}"/>{texto}', style)
    p._bookmark = (level, _strip_html(texto), anchor)
    return p


def _strip_html(s):
    out, in_tag = [], False
    for ch in s:
        if ch == "<": in_tag = True
        elif ch == ">": in_tag = False
        elif not in_tag: out.append(ch)
    return "".join(out)


# ═════════════════════════════════════════════════════════
# HELPERS DE TABLAS
# ═════════════════════════════════════════════════════════
def tabla_datos(filas, S, col_widths=None):
    """Tabla 2 columnas: Parámetro | Valor."""
    if col_widths is None:
        col_widths = [6.5 * cm, 10.5 * cm]
    data = [[Paragraph("Parámetro", S["tab_h"]), Paragraph("Valor", S["tab_h"])]]
    for k, v in filas:
        data.append([Paragraph(str(k), S["tab_key"]), Paragraph(str(v), S["tab_val"])])
    t = Table(data, colWidths=col_widths)
    t.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), C_AZUL_TIT),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [C_GRIS_FONDO, colors.white]),
        ("GRID", (0, 0), (-1, -1), 0.4, C_GRIS_LINEA),
        ("LEFTPADDING",  (0, 0), (-1, -1), 8),
        ("RIGHTPADDING", (0, 0), (-1, -1), 8),
        ("TOPPADDING",   (0, 0), (-1, -1), 5),
        ("BOTTOMPADDING",(0, 0), (-1, -1), 5),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
    ]))
    return t


# Sustituciones para caracteres no representables en Garamond/Courier:
# Garamond (Windows) NO incluye ✓ ✗ ★, así que usamos texto en español.
# Los emojis también se sustituyen por glifos disponibles o se eliminan.
_PDF_REPLACEMENTS = {
    # Marcas de tabla (Garamond no tiene estos glifos → usar texto)
    "✓":  "Sí",
    "✗":  "No",
    "★":  "Rec.",
    "✅": "Sí",
    "❌": "No",
    "⭐": "Rec.",
    "✔️": "Sí",
    # Otros emojis decorativos que no aportan al PDF
    "⚠️": "!",
    "⚡":  "*",
    "❤":  "♥",
    "📋": "",   "📐": "",   "📊": "",   "📄": "",   "🔧": "",
    "🔍": "",   "💡": "",   "➕": "+",   "🗑️": "",
    "📂": "",   "📌": "",   "🛡️": "",   "🧵": "",   "⬇️": "",
    "💾": "",
}


import re as _re

# Subíndices / superíndices Unicode → equivalente ASCII
_SUB_MAP = str.maketrans({
    "₀":"0","₁":"1","₂":"2","₃":"3","₄":"4",
    "₅":"5","₆":"6","₇":"7","₈":"8","₉":"9",
})
_SUP_MAP = str.maketrans({
    "⁰":"0","¹":"1","²":"2","³":"3","⁴":"4",
    "⁵":"5","⁶":"6","⁷":"7","⁸":"8","⁹":"9",
})
# Letras pequeñas subscript adicionales que pudieran aparecer
_LETRA_SUB_MAP = str.maketrans({"ₐ":"a","ₑ":"e","ₒ":"o","ₓ":"x","ₕ":"h","ₖ":"k","ₗ":"l","ₘ":"m","ₙ":"n","ₚ":"p","ₛ":"s","ₜ":"t"})

# Runs consecutivos de subíndices o superíndices
_SUB_RE = _re.compile(r"[₀-₉ₐₑₒₓₕₖₗₘₙₚₛₜ]+")
_SUP_RE = _re.compile(r"[⁰¹²³⁴-⁹]+")


def sanitize_pdf_text(s):
    """Reemplaza emojis y glifos no soportados por Garamond/Courier.
       - ✓ → 'Sí', ✗ → 'No', ★ → 'Rec.'
       - Subíndices Unicode  (₀-₉) → <sub>N</sub>
       - Superíndices Unicode (⁰⁴-⁹) → <super>N</super>
       (¹, ², ³ sí están en Garamond — se preservan.)
    """
    out = str(s)
    # Subíndices
    out = _SUB_RE.sub(
        lambda m: f"<sub>{m.group(0).translate(_SUB_MAP).translate(_LETRA_SUB_MAP)}</sub>",
        out
    )
    # Superíndices (solo los que no están en la fuente: ⁰, ⁴-⁹)
    out = _SUP_RE.sub(
        lambda m: f"<super>{m.group(0).translate(_SUP_MAP)}</super>",
        out
    )
    # Reemplazos simples (emojis → texto)
    for k, v in _PDF_REPLACEMENTS.items():
        out = out.replace(k, v)
    return out


def _es_marca_recomendada(texto):
    """True si la celda contiene el marcador de fila recomendada."""
    t = str(texto)
    return ("★" in t) or ("Rec." in t) or ("⭐" in t)


def tabla_dataframe(df, S, col_widths=None):
    """Convierte DataFrame en Table, sanitizando emojis."""
    cols = list(df.columns)
    n = len(cols)
    if col_widths is None:
        col_widths = [17.5 * cm / n] * n
    data = [[Paragraph(sanitize_pdf_text(c), S["tab_h"]) for c in cols]]
    for _, row in df.iterrows():
        data.append([Paragraph(sanitize_pdf_text(v), S["tab_c"]) for v in row])

    t = Table(data, colWidths=col_widths, repeatRows=1)
    # Resalta la fila marcada como "★" (recomendada) en la última columna
    style_cmds = [
        ("BACKGROUND", (0, 0), (-1, 0), C_AZUL_TIT),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [C_GRIS_FONDO, colors.white]),
        ("GRID", (0, 0), (-1, -1), 0.4, C_GRIS_LINEA),
        ("LEFTPADDING",  (0, 0), (-1, -1), 4),
        ("RIGHTPADDING", (0, 0), (-1, -1), 4),
        ("TOPPADDING",   (0, 0), (-1, -1), 4),
        ("BOTTOMPADDING",(0, 0), (-1, -1), 4),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
    ]
    # Pinta la fila recomendada con fondo verde-claro
    for i, (_, row) in enumerate(df.iterrows(), start=1):
        last = str(row.iloc[-1]) if len(row) else ""
        if _es_marca_recomendada(last):
            style_cmds.append(("BACKGROUND", (0, i), (-1, i), colors.HexColor("#e7f5ea")))
            style_cmds.append(("FONTNAME",   (0, i), (-1, i), FONTS["BODY_BOLD"]))
    t.setStyle(TableStyle(style_cmds))
    return t


# ═════════════════════════════════════════════════════════
# SECCIONES PROFESIONALES (introducción, objetivo, etc.)
# ═════════════════════════════════════════════════════════
def seccion_introduccion(S, dp):
    story = [_titulo_con_anchor("1. INTRODUCCIÓN", S["titulo"], "sec_intro", level=0)]
    txt = (
        f"La presente memoria técnica documenta el diseño y cálculo de la instalación "
        f"eléctrica del proyecto <b>{dp.get('proyecto','—')}</b>, conforme a la normativa "
        f"mexicana vigente <b>{dp.get('norma','NOM-001-SEDE-2012')}</b>. El documento describe "
        f"los criterios de diseño, parámetros de entrada, fórmulas utilizadas y desarrollo "
        f"detallado de los cálculos de caída de tensión, selección de conductores, "
        f"dimensionamiento de protecciones contra sobrecorriente y conductores de puesta a tierra. "
        f"El propósito es servir como soporte técnico para la construcción, supervisión y "
        f"verificación de la instalación, garantizando el cumplimiento de los requisitos "
        f"mínimos de seguridad eléctrica establecidos por la norma."
    )
    story.append(Paragraph(txt, S["cuerpo"]))
    return story


def seccion_objetivo(S, dp):
    story = [_titulo_con_anchor("2. OBJETIVO", S["titulo"], "sec_objetivo", level=0)]
    txt = (
        "Determinar, mediante cálculo justificado, los siguientes elementos para cada uno "
        "de los circuitos del proyecto:"
    )
    story.append(Paragraph(txt, S["cuerpo"]))
    items = [
        "<b>a)</b> Corriente de carga y corriente de diseño (incluyendo factor 1.25 si aplica).",
        "<b>b)</b> Calibre del conductor por criterio de ampacidad corregida.",
        "<b>c)</b> Verificación de la caída de tensión (% C.d.T.) dentro del límite normativo.",
        "<b>d)</b> Capacidad nominal de la protección contra sobrecorriente (OCPD) en formato comercial NxA.",
        "<b>e)</b> Calibre del conductor de puesta a tierra de equipos según NOM-001 Tabla 250-122.",
        "<b>f)</b> Tubería (conduit) suficiente para alojar los conductores cumpliendo el % de relleno permitido.",
    ]
    for it in items:
        story.append(Paragraph(it, S["cuerpo"]))
    return story


def seccion_descripcion(S, dp):
    story = [_titulo_con_anchor("3. DESCRIPCIÓN DEL PROYECTO", S["titulo"], "sec_desc", level=0)]
    desc = dp.get("descripcion") or (
        f"Instalación eléctrica para {dp.get('proyecto','—')} localizada en "
        f"{dp.get('ubicacion','—')}, propiedad de {dp.get('cliente','—')}. "
        f"El sistema de alimentación principal es {dp.get('sistema','—')}, "
        f"con temperatura ambiente de diseño de {dp.get('temp_ambiente',35)}°C. "
        f"La instalación contempla circuitos derivados protegidos contra sobrecorriente, "
        f"con conductor de puesta a tierra de equipos y canalizaciones acordes al uso."
    )
    story.append(Paragraph(desc, S["cuerpo"]))

    story.append(Spacer(1, 0.2 * cm))
    story.append(Paragraph("<b>Datos generales</b>", S["cuerpo"]))
    filas = [
        ("Proyecto", dp.get("proyecto", "—")),
        ("Cliente / Propietario", dp.get("cliente", "—")),
        ("Ubicación", dp.get("ubicacion", "—")),
        ("Sistema eléctrico principal", dp.get("sistema", "—")),
        ("Temperatura ambiente", f"{dp.get('temp_ambiente', 35)} °C"),
        ("Responsable técnico", dp.get("responsable", "—")),
        ("Cédula profesional", dp.get("cedula", "—")),
        ("Fecha", str(dp.get("fecha", date.today()))),
    ]
    story.append(tabla_datos(filas, S))
    return story


NORMAS_DESCRIPCION = {
    "NOM-001-SEDE-2012": (
        "Norma Oficial Mexicana de Instalaciones Eléctricas. Documento de "
        "observancia obligatoria que establece las disposiciones para instalaciones "
        "eléctricas destinadas a la utilización de energía eléctrica en territorio nacional."
    ),
    "NOM-001-SEDE-2005": "Versión anterior de la NOM-001-SEDE.",
    "NEC 2023": "National Electrical Code (NFPA 70). Norma estadounidense referente.",
    "CFE DCDIAMT (Instalaciones aéreas MT)": (
        "Especificación CFE para la construcción de instalaciones eléctricas aéreas "
        "en media tensión, aplicable a redes de distribución y acometidas."
    ),
    "CFE DCDIASMT (Instalaciones subterráneas MT)": (
        "Especificación CFE para la construcción de instalaciones eléctricas "
        "subterráneas en media tensión: ductos, registros, cables y empalmes."
    ),
    "NOM-007-ENER-2014 (Eficiencia energética)": (
        "Eficiencia energética para sistemas de alumbrado en edificios no residenciales."
    ),
    "NMX-J-098-ANCE (Tensiones eléctricas estándar)": (
        "Valores normalizados de tensión eléctrica para sistemas de utilización."
    ),
    "NMX-J-235-ANCE (Conductores eléctricos)": (
        "Características y métodos de prueba para conductores eléctricos aislados."
    ),
    "IEEE Std 141 (Red Book)": (
        "Recommended Practice for Electric Power Distribution for Industrial Plants."
    ),
    "IEC 60364 (Instalaciones eléctricas en edificios)": (
        "Estándar internacional para instalaciones eléctricas en edificios."
    ),
}


def seccion_normatividad(S, dp):
    story = [_titulo_con_anchor("4. NORMATIVIDAD APLICABLE", S["titulo"], "sec_norm", level=0)]
    story.append(Paragraph(
        "El presente cálculo se apega a las siguientes normas técnicas y especificaciones:",
        S["cuerpo"]
    ))
    principales = [dp.get("norma", "NOM-001-SEDE-2012")] + list(dp.get("normas_extra", []))
    # Quitar duplicados conservando orden
    seen = set()
    seleccionadas = [n for n in principales if not (n in seen or seen.add(n))]

    for i, n in enumerate(seleccionadas, 1):
        story.append(Paragraph(
            f"<b>{i}. {n}.</b> "
            + NORMAS_DESCRIPCION.get(n, "Norma técnica de referencia."),
            S["cuerpo"]
        ))
    return story


def seccion_criterios(S, dp, ejemplo_circ=None):
    story = [_titulo_con_anchor("5. CRITERIOS DE DISEÑO", S["titulo"], "sec_crit", level=0)]

    txt1 = (
        "La selección del conductor de cada circuito se realiza verificando "
        "simultáneamente dos criterios:"
    )
    story.append(Paragraph(txt1, S["cuerpo"]))
    story.append(Paragraph(
        "<b>(1) Capacidad de conducción (ampacidad).</b> La ampacidad base se toma de la "
        "<b>Tabla 310-15(B)(16)</b> de la NOM-001 según el material del conductor y la "
        "temperatura del aislamiento. Se afecta por el factor de corrección por temperatura "
        "ambiente y por el factor de agrupamiento. Debe satisfacer "
        "<i>Ampacidad ≥ I_diseño / CF</i>.",
        S["cuerpo"]
    ))
    story.append(Paragraph(
        "<b>(2) Caída de tensión.</b> Se calcula con los valores de resistencia (R) y reactancia "
        "(X) de la <b>Tabla 9</b> de la NOM-001, ajustados a la temperatura del terminal del "
        "equipo según el Artículo 110-14. Para corrientes ≤ 100 A se asume terminal de 60 °C; "
        "para corrientes &gt; 100 A se asume 75 °C; para motores (excepto Clase A) se usa 75 °C.",
        S["cuerpo"]
    ))
    story.append(Paragraph(
        "<b>(3) Protección contra sobrecorriente.</b> Se selecciona el siguiente tamaño "
        "estándar (NEC 240.6 / NOM-001 Art. 240) que sea ≥ I_diseño y que no exceda la "
        "ampacidad corregida del conductor. El formato comercial se expresa como N x A "
        "donde N es el número de polos del sistema (1 monofásico, 2 bifásico, 3 trifásico).",
        S["cuerpo"]
    ))
    story.append(Paragraph(
        "<b>(4) Conductor de puesta a tierra de equipos.</b> Se dimensiona según la "
        "<b>Tabla 250-122</b> de la NOM-001, en función de la capacidad nominal de la "
        "protección contra sobrecorriente del circuito. Se utiliza cable de cobre o aluminio "
        "desnudo.",
        S["cuerpo"]
    ))

    # Tabla resumen de constantes
    story.append(Spacer(1, 0.2 * cm))
    story.append(Paragraph("<b>Constantes y parámetros base</b>", S["cuerpo"]))
    filas = [
        ("Base de ampacidad (T<sub>0</sub>)", "30 °C"),
        ("Temperatura ambiente de proyecto", f"{dp.get('temp_ambiente',35)} °C"),
        ("Coeficiente α térmico Cu (R)", "α = 1/234.5 K<super>-1</super>"),
        ("Ajuste de R a temp. terminal",
         "R<sub>T</sub> = R<sub>75</sub> × (T + 234.5) / (75 + 234.5)"),
        ("Factor de carga continua", "1.25 (cargas ≥ 3 h continuas, NEC 210.20)"),
    ]
    story.append(tabla_datos(filas, S))
    return story


# ═════════════════════════════════════════════════════════
# DESARROLLO DE CÁLCULOS POR CIRCUITO
# ═════════════════════════════════════════════════════════
def desarrollo_calculo_circuito(circ, idx, S):
    """Desarrollo paso a paso de las fórmulas para un circuito.
       idx = número del circuito en la memoria (7.1, 7.2, ...).
    """
    story = []
    anchor = f"sec_cto_{idx}"
    nombre = circ.get("nombre", f"Circuito {idx}")

    # Subtítulo del circuito (nivel 1 en outline)
    story.append(_titulo_con_anchor(
        f"7.{idx} {nombre}", S["subtitulo"], anchor, level=1
    ))

    # 7.1.1 Datos de entrada
    story.append(Paragraph(f"7.{idx}.1 Datos de entrada", S["subsub"]))
    mat = circ.get("material", "Cu")
    canal_lbl = circ.get("canalizacion_label", circ.get("canalizacion", "Acero"))
    filas_ent = [
        ("Potencia de la carga", f"P = {circ.get('potencia',0):,} W"),
        ("Tensión de operación", f"V = {circ.get('voltaje','—')} V"),
        ("Factor de potencia", f"cos φ = {circ.get('fp', 0.9)}"),
        ("Configuración", str(circ.get("configuracion", "—"))),
        ("Longitud del circuito", f"L = {circ.get('longitud',0)} m"),
        ("Material conductor", f"{mat} ({'Cobre' if mat=='Cu' else 'Aluminio'})"),
        ("Conductores por fase", f"CF = {circ.get('cf',1)}"),
        ("Canalización (Tabla 9)", canal_lbl),
        ("Temp. ambiente", f"{circ.get('temp_ambiente_proy', '—')} °C"),
        ("Aislamiento del conductor", f"{circ.get('temp_aislamiento',75)} °C"),
        ("C.d.T. máxima permitida", f"{circ.get('cdt_max',3.0)} %"),
    ]
    story.append(tabla_datos(filas_ent, S))

    # 7.1.2 Cálculo de corriente
    story.append(Paragraph(f"7.{idx}.2 Cálculo de corriente nominal", S["subsub"]))
    config_code = circ.get("config_code", "mono2h")
    fp = circ.get("fp", 0.9)
    V = circ.get("voltaje", 120)
    P = circ.get("potencia", 0)
    I = circ.get("corriente", 0)

    if config_code == "trifasico":
        formula_I = "I = P / ( √3 · V · cosφ )"
        sustit = f"I = {P:,} / ( 1.7321 × {V} × {fp} ) = <b>{I:.3f} A</b>"
    else:
        formula_I = "I = P / ( V · cosφ )"
        sustit = f"I = {P:,} / ( {V} × {fp} ) = <b>{I:.3f} A</b>"

    story.append(Paragraph("Fórmula aplicada:", S["cuerpo_left"]))
    story.append(Paragraph(formula_I, S["formula"]))
    story.append(Paragraph("Sustitución:", S["cuerpo_left"]))
    story.append(Paragraph(sustit, S["formula"]))

    # Corriente de diseño
    fd = circ.get("factor_demanda", 1.25)
    I_dis = I * fd
    cf = circ.get("cf", 1)
    I_dis_cond = I_dis / cf
    story.append(Paragraph(
        f"Corriente de diseño (factor de carga continua = {fd}):", S["cuerpo_left"]
    ))
    story.append(Paragraph(
        f"I_diseño = I × FD = {I:.3f} × {fd} = <b>{I_dis:.3f} A</b>",
        S["formula"]
    ))
    if cf > 1:
        story.append(Paragraph(
            f"I por conductor = I_diseño / CF = {I_dis:.3f} / {cf} = <b>{I_dis_cond:.3f} A</b>",
            S["formula"]
        ))

    # 7.1.3 Cálculo de caída de tensión
    story.append(Paragraph(f"7.{idx}.3 Cálculo de caída de tensión", S["subsub"]))
    cond  = circ.get("conductor", "—")
    cdt   = circ.get("cdt", 0)
    cdt_max = circ.get("cdt_max", 3.0)
    canal = circ.get("canalizacion", "AC")
    t_term = circ.get("temp_term", 75)

    # Obtener R, X usados
    try:
        R_eff, X_eff = r_x_efectivas(cond, mat, canal, t_term)
        R75 = R_eff * (75 + 234.5) / (t_term + 234.5)  # R original
    except Exception:
        R_eff, X_eff, R75 = 0, 0, 0
    sen_phi = math.sqrt(max(0.0, 1 - fp**2))

    story.append(Paragraph(
        "De la <b>Tabla 9 NOM-001</b>, para conductor "
        f"<b>{cond} AWG {mat}</b> en canalización <b>{canal_lbl}</b>:",
        S["cuerpo_left"]
    ))
    story.append(Paragraph(
        f"R<sub>75</sub> = {R75:.4f} Ω/km · X = {X_eff:.4f} Ω/km",
        S["formula"]
    ))
    story.append(Paragraph(
        f"Ajuste a temperatura de terminal {t_term}°C (Art. 110-14):",
        S["cuerpo_left"]
    ))
    story.append(Paragraph(
        f"R<sub>T</sub> = R<sub>75</sub> · (T + 234.5)/(75 + 234.5) = "
        f"{R75:.4f} × ({t_term}+234.5)/(75+234.5) = <b>{R_eff:.4f} Ω/km</b>",
        S["formula"]
    ))

    # Fórmula y sustitución según configuración
    story.append(Paragraph(
        "Caída de tensión por el método de impedancia (Tabla 9):",
        S["cuerpo_left"]
    ))
    if config_code == "trifasico":
        formula_cdt = "ΔV = √3 · I · (L/1000) · ( (R/CF)·cosφ + (X/CF)·senφ )"
        denom = V
        formula_pct = "%ΔV = ΔV · 100 / V_línea"
    else:
        formula_cdt = "ΔV = 2 · I · (L/1000) · ( (R/CF)·cosφ + (X/CF)·senφ )"
        denom = V
        formula_pct = ("%ΔV = ΔV · 100 / V_fase-neutro   "
                       "(monofásico 2H)") if config_code == "mono2h" else \
                      "%ΔV = ΔV · 100 / V_línea     (mono 3H / bifásico)"

    L_km = circ.get("longitud", 0) / 1000
    Z = (R_eff / cf) * fp + (X_eff / cf) * sen_phi
    if config_code == "trifasico":
        dV_v = math.sqrt(3) * I * L_km * Z
    else:
        dV_v = 2 * I * L_km * Z
    pct = (dV_v / denom * 100) if denom else 0

    story.append(Paragraph(formula_cdt, S["formula"]))
    story.append(Paragraph("Sustitución de valores:", S["cuerpo_left"]))
    coef = "√3" if config_code == "trifasico" else "2"
    story.append(Paragraph(
        f"ΔV = {coef} × {I:.3f} × ({circ.get('longitud',0)}/1000) × "
        f"( ({R_eff:.4f}/{cf}) × {fp} + ({X_eff:.4f}/{cf}) × {sen_phi:.4f} )",
        S["formula"]
    ))
    story.append(Paragraph(
        f"ΔV = <b>{dV_v:.4f} V</b>", S["formula"]
    ))
    story.append(Paragraph(formula_pct, S["formula"]))
    story.append(Paragraph(
        f"%ΔV = ({dV_v:.4f} × 100) / {denom} = <b>{pct:.3f} %</b>", S["formula"]
    ))

    cumple_cdt = cdt <= cdt_max
    story.append(Paragraph(
        ("Cumple el criterio de caída de tensión: "
         f"%ΔV = {cdt:.2f}% ≤ {cdt_max}% (máx.)")
        if cumple_cdt else
        (f"NO cumple: %ΔV = {cdt:.2f}% &gt; {cdt_max}% (máx.). "
         "Se requiere subir calibre o reducir longitud."),
        S["resultado_ok"] if cumple_cdt else S["resultado_err"]
    ))

    # Tabla completa de conductores candidatos (NOM-001 Tabla 310-15(B)(16) + Tabla 9)
    tabla_cand = circ.get("tabla_conductor_df")
    # Si vino de JSON importado, llega como lista de dicts: convertir a DataFrame.
    if isinstance(tabla_cand, list):
        import pandas as _pd
        tabla_cand = _pd.DataFrame(tabla_cand) if tabla_cand else None
    if tabla_cand is not None and len(tabla_cand) > 0:
        story.append(Spacer(1, 0.2 * cm))
        story.append(Paragraph(
            "<b>Tabla de verificación de conductores candidatos</b> "
            "— se evalúa cada calibre disponible del material. "
            "<i>Rec.</i> = recomendado (fila resaltada); "
            "<i>Sí/No</i> = cumple o no cumple cada criterio.",
            S["cuerpo_left"]
        ))
        # Ancho de columnas balanceado
        col_w = [1.5*cm, 1.7*cm, 1.4*cm, 1.2*cm, 1.6*cm, 1.6*cm, 1.6*cm, 1.2*cm, 1.5*cm, 1.7*cm]
        story.append(tabla_dataframe(tabla_cand, S, col_widths=col_w))

    # 7.1.4 Selección de protección
    story.append(Paragraph(f"7.{idx}.4 Protección contra sobrecorriente", S["subsub"]))
    ocpd_A = circ.get("proteccion_A", "—")
    ocpd_fmt = circ.get("proteccion_fmt", str(ocpd_A))
    polos = circ.get("polos", 1)
    story.append(Paragraph(
        f"Según NOM-001 Art. 240 / NEC 240.6, se selecciona el siguiente tamaño "
        f"estándar de protección que sea ≥ I_diseño/CF y no exceda la ampacidad "
        f"corregida del conductor.",
        S["cuerpo"]
    ))
    story.append(Paragraph(
        f"I_diseño/CF = {I_dis_cond:.3f} A   →   Protección estándar siguiente: "
        f"<b>{ocpd_A} A</b>",
        S["formula"]
    ))
    story.append(Paragraph(
        f"En formato comercial ({polos} polo(s) según sistema): "
        f"<b>{ocpd_fmt} A</b>",
        S["formula"]
    ))

    # 7.1.5 Conductor de tierra
    story.append(Paragraph(f"7.{idx}.5 Conductor de puesta a tierra (cable desnudo)", S["subsub"]))
    tierra_cal = circ.get("tierra_calibre", "—")
    tierra_mat = circ.get("tierra_material", "Cu")
    origen = "automático (Tabla 250-122 NOM-001)" if circ.get("tierra_auto", True) else "selección manual"
    story.append(Paragraph(
        f"Para una protección de <b>{ocpd_A} A</b>, la <b>Tabla 250-122</b> de la "
        f"NOM-001 especifica un conductor de puesta a tierra de equipos calibre "
        f"<b>{tierra_cal} AWG</b> de <b>{tierra_mat} desnudo</b>. "
        f"({origen})",
        S["cuerpo"]
    ))

    # 7.1.6 Tubería (si aplica) — tabla COMPLETA con todas las medidas
    ci = circ.get("conduit_info")
    if ci and ci.get("recomendada"):
        story.append(Paragraph(f"7.{idx}.6 Cálculo de tubería (conduit)", S["subsub"]))
        params = ci.get("params") or {}
        rec = ci.get("recomendada")
        story.append(Paragraph(
            f"Para <b>{params.get('num_conds','—')} conductor(es)</b> calibre "
            f"<b>{params.get('calibre','—')} AWG</b> con aislamiento "
            f"<b>{params.get('aislamiento','—')}</b>, en tubería tipo <b>"
            f"{params.get('tipo_tubo','—')}</b>, el cálculo de relleno se realiza para "
            "todas las medidas comerciales y se elige la primera que satisface el "
            "porcentaje máximo permitido (NOM-001 Cap. 3, Tablas Capítulo 4).",
            S["cuerpo"]
        ))

        # Tabla COMPLETA con todas las medidas
        resultados_todos = ci.get("resultados_todos") or [rec]
        encabezado = ["Tubería", "Ø int. (mm)", "Área int. (mm²)",
                      "Área conds. (mm²)", "Relleno (%)", "Máx. %", "Cumple", "Sel."]
        data = [[Paragraph(sanitize_pdf_text(h), S["tab_h"]) for h in encabezado]]
        for r_t in resultados_todos:
            sel_mark = "Rec." if r_t.get("recomendada", False) else ""
            cumple_mark = "Sí" if r_t.get("cumple", False) else "No"
            diam = r_t.get("diametro", r_t.get("tubo", "—").split(" ")[0])
            row = [
                Paragraph(sanitize_pdf_text(r_t.get("tubo", "—")), S["tab_c"]),
                Paragraph(sanitize_pdf_text(diam), S["tab_c"]),
                Paragraph(f"{r_t.get('area_tubo', 0):.1f}", S["tab_c"]),
                Paragraph(f"{r_t.get('area_conds', 0):.2f}", S["tab_c"]),
                Paragraph(f"{r_t.get('fill_pct', 0):.1f}", S["tab_c"]),
                Paragraph(f"{r_t.get('fill_max', 40)}", S["tab_c"]),
                Paragraph(cumple_mark, S["tab_c"]),
                Paragraph(sel_mark, S["tab_c"]),
            ]
            data.append(row)
        col_w_t = [2.3*cm, 1.8*cm, 2.3*cm, 2.3*cm, 1.9*cm, 1.4*cm, 1.4*cm, 1.4*cm]
        t = Table(data, colWidths=col_w_t, repeatRows=1)
        style_t = [
            ("BACKGROUND", (0, 0), (-1, 0), C_AZUL_TIT),
            ("ROWBACKGROUNDS", (0, 1), (-1, -1), [C_GRIS_FONDO, colors.white]),
            ("GRID", (0, 0), (-1, -1), 0.4, C_GRIS_LINEA),
            ("LEFTPADDING",  (0, 0), (-1, -1), 4),
            ("RIGHTPADDING", (0, 0), (-1, -1), 4),
            ("TOPPADDING",   (0, 0), (-1, -1), 4),
            ("BOTTOMPADDING",(0, 0), (-1, -1), 4),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ]
        for i_r, r_t in enumerate(resultados_todos, start=1):
            if r_t.get("recomendada", False):
                style_t.append(("BACKGROUND", (0, i_r), (-1, i_r), colors.HexColor("#e7f5ea")))
                style_t.append(("FONTNAME",   (0, i_r), (-1, i_r), FONTS["BODY_BOLD"]))
        t.setStyle(TableStyle(style_t))
        story.append(t)

        cumple_t = rec.get("cumple", False)
        story.append(Spacer(1, 0.15 * cm))
        story.append(Paragraph(
            f"Tubería seleccionada: <b>{rec.get('tubo','—')}</b> · relleno "
            f"<b>{rec.get('fill_pct',0):.1f}%</b> ≤ {rec.get('fill_max',40)}%."
            if cumple_t else
            "Ningún tamaño cumple. Considerar tubería de mayor familia o más conductos.",
            S["resultado_ok"] if cumple_t else S["resultado_err"]
        ))

    # Resumen final del circuito
    story.append(Paragraph(f"7.{idx}.7 Resumen", S["subsub"]))
    resumen = [
        ("Conductor seleccionado", f"{circ.get('cf',1)} × {cond} AWG {mat}"),
        ("Aislamiento", f"{circ.get('temp_aislamiento',75)} °C"),
        ("Ampacidad corregida", f"{circ.get('ampacity_corr',0):.1f} A"),
        ("Caída de tensión real", f"{cdt:.2f} % (máx. {cdt_max} %)"),
        ("Protección OCPD", f"{ocpd_fmt} A"),
        ("Conductor de tierra", f"{tierra_cal} AWG {tierra_mat} desnudo"),
    ]
    if ci and ci.get("recomendada"):
        resumen.append(("Tubería", f"{ci['recomendada'].get('tubo','—')} · "
                                    f"{ci['recomendada'].get('fill_pct',0):.1f}% relleno"))
    story.append(tabla_datos(resumen, S))

    return story


# ═════════════════════════════════════════════════════════
# ÍNDICE (clickeable)
# ═════════════════════════════════════════════════════════
def construir_indice(S, secciones):
    """secciones = [(label, anchor, level), ...]"""
    story = [_titulo_con_anchor("ÍNDICE", S["titulo"], "sec_indice", level=0)]
    story.append(Paragraph(
        "<i>Haga clic en cualquier entrada del índice para saltar directamente "
        "a la sección correspondiente.</i>",
        S["cuerpo"]
    ))
    story.append(Spacer(1, 0.3 * cm))
    for label, anchor, level in secciones:
        style = S["toc_l2"] if level > 0 else S["toc_l1"]
        link = (f'<link href="#{anchor}" color="#0066cc">'
                f'<u>{label}</u></link>')
        story.append(Paragraph(link, style))
    return story


# ═════════════════════════════════════════════════════════
# TABLA RESUMEN DE CIRCUITOS (panel-board)
# ═════════════════════════════════════════════════════════
def tabla_resumen_circuitos(circuitos, S):
    encabezado = ["No.", "Circuito", "P (W)", "V", "I (A)", "I.dis", "Cond.",
                  "Amp", "CdT%", "Máx", "OCPD", "Tierra", "Tubería", "Estado"]
    col_w = [0.7*cm, 2.4*cm, 1.3*cm, 1.0*cm, 1.1*cm, 1.1*cm,
             1.8*cm, 1.1*cm, 1.0*cm, 0.9*cm, 1.3*cm, 1.4*cm, 1.7*cm, 1.2*cm]

    data = [[Paragraph(h, S["tab_h"]) for h in encabezado]]
    for i, c in enumerate(circuitos, 1):
        cdt = c.get("cdt", 0)
        cdt_max = c.get("cdt_max", 3.0)
        estado = "OK" if cdt <= cdt_max else "REVISAR"
        ci = c.get("conduit_info")
        tubo = ci["recomendada"]["tubo"] if (ci and ci.get("recomendada")) else "—"
        cond = f"{c.get('cf',1)}×{c.get('conductor','—')} {c.get('material','')}"
        tierra = f"{c.get('tierra_calibre','—')} {c.get('tierra_material','')}".strip()
        row = [
            Paragraph(str(c.get("id", i)), S["tab_c"]),
            Paragraph(str(c.get("nombre", f"Cto {i}")), S["tab_c"]),
            Paragraph(f"{c.get('potencia',0):,}", S["tab_c"]),
            Paragraph(str(c.get("voltaje", "—")), S["tab_c"]),
            Paragraph(f"{c.get('corriente',0):.1f}", S["tab_c"]),
            Paragraph(f"{c.get('corriente_diseño',0):.1f}", S["tab_c"]),
            Paragraph(cond, S["tab_c"]),
            Paragraph(f"{c.get('ampacity_corr',0):.0f}", S["tab_c"]),
            Paragraph(f"{cdt:.2f}", S["tab_c"]),
            Paragraph(f"{cdt_max}", S["tab_c"]),
            Paragraph(str(c.get("proteccion_fmt", c.get("proteccion_A", "—"))), S["tab_c"]),
            Paragraph(tierra, S["tab_c"]),
            Paragraph(tubo, S["tab_c"]),
            Paragraph(estado, S["tab_c"]),
        ]
        data.append(row)
    t = Table(data, colWidths=col_w, repeatRows=1)
    t.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), C_AZUL_TIT),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [C_GRIS_FONDO, colors.white]),
        ("GRID", (0, 0), (-1, -1), 0.4, C_GRIS_LINEA),
        ("LEFTPADDING",  (0, 0), (-1, -1), 3),
        ("RIGHTPADDING", (0, 0), (-1, -1), 3),
        ("TOPPADDING",   (0, 0), (-1, -1), 4),
        ("BOTTOMPADDING",(0, 0), (-1, -1), 4),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
    ]))
    return t


# ═════════════════════════════════════════════════════════
# FUNCIÓN PRINCIPAL
# ═════════════════════════════════════════════════════════
def generar_reporte_pdf(datos_proyecto, resultado_conductor,
                         tabla_conductor=None,
                         resultado_conduit=None,
                         conduit_params=None,
                         circuitos=None) -> bytes:
    """Genera el PDF.
       - Si `circuitos` es None: reporte detallado de un solo circuito.
       - Si `circuitos` está dado: reporte multi-circuito con tablero resumen.
    """
    buffer = io.BytesIO()
    doc = OutlinedDocTemplate(
        buffer,
        pagesize=letter,
        rightMargin=2.2 * cm, leftMargin=2.2 * cm,
        topMargin=2.0 * cm, bottomMargin=2.0 * cm,
        title=datos_proyecto.get("proyecto", "eSuit"),
        author=datos_proyecto.get("responsable", ""),
    )
    S = build_styles()
    story = []
    dp = datos_proyecto
    rc = resultado_conductor

    # Si es modo single-circuit y no llega `circuitos`, fabricamos uno.
    if circuitos is None:
        rc_circ = dict(rc) if rc else {}
        rc_circ.setdefault("id", 1)
        rc_circ.setdefault("nombre", "Circuito principal")
        if resultado_conduit and conduit_params:
            recom = next((x for x in resultado_conduit if x["cumple"]), resultado_conduit[-1])
            rc_circ["conduit_info"] = {
                "params": conduit_params,
                "recomendada": recom,
                "resultados_todos": resultado_conduit,
            }
        if tabla_conductor is not None:
            rc_circ["tabla_conductor_df"] = tabla_conductor
        circuitos_internos = [rc_circ]
    else:
        circuitos_internos = circuitos

    # Inyectar temp_ambiente_proy a cada circuito
    for c in circuitos_internos:
        c.setdefault("temp_ambiente_proy", dp.get("temp_ambiente", 35))

    # ═══════════════════════════════════════════
    # PORTADA
    # ═══════════════════════════════════════════
    story.append(Spacer(1, 2.5 * cm))
    story.append(HRFlowable(width="100%", thickness=4, color=C_AZUL_OSC, spaceAfter=20))
    story.append(Paragraph("MEMORIA TÉCNICA DE CÁLCULO", S["portada_titulo"]))
    story.append(Paragraph("Instalación Eléctrica", S["portada_subtitulo"]))
    story.append(Paragraph(
        "Caída de tensión · Conductores · Protecciones · Puesta a tierra",
        S["portada_info"]
    ))
    story.append(Spacer(1, 1.5 * cm))
    story.append(HRFlowable(width="40%", thickness=1, color=C_GRIS_LINEA, spaceAfter=20))

    filas_portada = [
        ("Proyecto", dp.get("proyecto", "—")),
        ("Cliente / Propietario", dp.get("cliente", "—")),
        ("Ubicación", dp.get("ubicacion", "—")),
        ("Sistema eléctrico", dp.get("sistema", "—")),
        ("Norma principal", dp.get("norma", "NOM-001-SEDE-2012")),
        ("Temperatura ambiente", f"{dp.get('temp_ambiente', 35)} °C"),
        ("Responsable técnico", dp.get("responsable", "—")),
        ("Cédula profesional", dp.get("cedula", "—")),
        ("Fecha de elaboración", str(dp.get("fecha", date.today()))),
    ]
    story.append(tabla_datos(filas_portada, S))

    story.append(Spacer(1, 1 * cm))
    story.append(HRFlowable(width="100%", thickness=4, color=C_AZUL_OSC, spaceAfter=6))
    story.append(Paragraph(
        "Documento técnico generado con eSuit · v2.0",
        S["pie"]
    ))
    story.append(PageBreak())

    # ═══════════════════════════════════════════
    # ÍNDICE
    # ═══════════════════════════════════════════
    secciones = [
        ("1. Introducción", "sec_intro", 0),
        ("2. Objetivo", "sec_objetivo", 0),
        ("3. Descripción del proyecto", "sec_desc", 0),
        ("4. Normatividad aplicable", "sec_norm", 0),
        ("5. Criterios de diseño", "sec_crit", 0),
        ("6. Tabla resumen de circuitos", "sec_tablero", 0),
        ("7. Memoria de cálculo por circuito", "sec_memoria", 0),
    ]
    for i, c in enumerate(circuitos_internos, 1):
        secciones.append((
            f"      7.{i} {c.get('nombre', f'Circuito {i}')}",
            f"sec_cto_{i}", 1,
        ))
    secciones += [
        ("8. Conclusiones", "sec_conclusiones", 0),
        ("9. Firma del responsable técnico", "sec_firma", 0),
    ]

    story += construir_indice(S, secciones)
    story.append(PageBreak())

    # ═══════════════════════════════════════════
    # 1. INTRODUCCIÓN
    # ═══════════════════════════════════════════
    story += seccion_introduccion(S, dp)
    story.append(Spacer(1, 0.3 * cm))

    # ═══════════════════════════════════════════
    # 2. OBJETIVO
    # ═══════════════════════════════════════════
    story += seccion_objetivo(S, dp)
    story.append(Spacer(1, 0.3 * cm))

    # ═══════════════════════════════════════════
    # 3. DESCRIPCIÓN
    # ═══════════════════════════════════════════
    story += seccion_descripcion(S, dp)
    story.append(PageBreak())

    # ═══════════════════════════════════════════
    # 4. NORMATIVIDAD
    # ═══════════════════════════════════════════
    story += seccion_normatividad(S, dp)
    story.append(Spacer(1, 0.3 * cm))

    # ═══════════════════════════════════════════
    # 5. CRITERIOS DE DISEÑO
    # ═══════════════════════════════════════════
    story += seccion_criterios(S, dp, circuitos_internos[0] if circuitos_internos else None)
    story.append(PageBreak())

    # ═══════════════════════════════════════════
    # 6. TABLA RESUMEN
    # ═══════════════════════════════════════════
    story.append(_titulo_con_anchor(
        "6. TABLA RESUMEN DE CIRCUITOS", S["titulo"], "sec_tablero", level=0
    ))
    story.append(Paragraph(
        f"Resumen de los <b>{len(circuitos_internos)} circuito(s)</b> del proyecto. "
        "Cada columna se desarrolla en detalle en la Sección 7.",
        S["cuerpo"]
    ))
    story.append(Spacer(1, 0.3 * cm))
    story.append(tabla_resumen_circuitos(circuitos_internos, S))
    story.append(PageBreak())

    # ═══════════════════════════════════════════
    # 7. MEMORIA POR CIRCUITO
    # ═══════════════════════════════════════════
    story.append(_titulo_con_anchor(
        "7. MEMORIA DE CÁLCULO POR CIRCUITO", S["titulo"], "sec_memoria", level=0
    ))
    story.append(Paragraph(
        "A continuación se presenta el desarrollo paso a paso de cada circuito: "
        "datos de entrada, cálculo de corriente, caída de tensión por método de "
        "impedancia (Tabla 9 NOM con ajuste por temperatura de terminal), "
        "selección de protección, conductor de tierra y tubería.",
        S["cuerpo"]
    ))
    story.append(Spacer(1, 0.3 * cm))

    for i, c in enumerate(circuitos_internos, 1):
        story += desarrollo_calculo_circuito(c, i, S)
        if i < len(circuitos_internos):
            story.append(HRFlowable(width="100%", thickness=0.5, color=C_GRIS_LINEA, spaceAfter=8))

    story.append(PageBreak())

    # ═══════════════════════════════════════════
    # 8. CONCLUSIONES
    # ═══════════════════════════════════════════
    story.append(_titulo_con_anchor(
        "8. CONCLUSIONES", S["titulo"], "sec_conclusiones", level=0
    ))
    total_kw = sum(c.get("potencia", 0) for c in circuitos_internos) / 1000
    ok_count = sum(1 for c in circuitos_internos
                   if c.get("cdt", 0) <= c.get("cdt_max", 3.0))
    story.append(Paragraph(
        f"<b>1.</b> El proyecto comprende <b>{len(circuitos_internos)} circuito(s)</b> con "
        f"una carga total instalada de <b>{total_kw:.2f} kW</b>.",
        S["cuerpo"]
    ))
    story.append(Paragraph(
        f"<b>2.</b> <b>{ok_count} de {len(circuitos_internos)}</b> circuito(s) cumplen "
        f"simultáneamente con los criterios de ampacidad y caída de tensión establecidos "
        f"en la {dp.get('norma','NOM-001-SEDE-2012')}.",
        S["cuerpo"]
    ))
    story.append(Paragraph(
        "<b>3.</b> Los cálculos de resistencia incluyen el ajuste por temperatura del "
        "terminal del equipo conforme al Artículo 110-14 de la NOM-001-SEDE, y los "
        "valores de R y X corresponden a la Tabla 9 de la misma norma.",
        S["cuerpo"]
    ))
    story.append(Paragraph(
        "<b>4.</b> Las protecciones contra sobrecorriente fueron seleccionadas como "
        "los tamaños estándar comerciales (NEC 240.6) inmediatamente superiores a la "
        "corriente de diseño, dentro de la ampacidad corregida del conductor.",
        S["cuerpo"]
    ))
    story.append(Paragraph(
        "<b>5.</b> Los conductores de puesta a tierra de equipos fueron dimensionados "
        "según la Tabla 250-122 de la NOM-001 en función de la capacidad nominal del "
        "OCPD, utilizando cable desnudo del material indicado.",
        S["cuerpo"]
    ))
    if ok_count < len(circuitos_internos):
        story.append(Paragraph(
            "<b>Observación:</b> Los circuitos marcados como REVISAR requieren ajuste "
            "(subir calibre, reducir longitud o redistribuir carga) antes de su construcción.",
            S["resultado_err"]
        ))

    story.append(Spacer(1, 1 * cm))

    # ═══════════════════════════════════════════
    # 9. FIRMA
    # ═══════════════════════════════════════════
    story.append(_titulo_con_anchor(
        "9. FIRMA DEL RESPONSABLE TÉCNICO", S["titulo"], "sec_firma", level=0
    ))
    story.append(Spacer(1, 1.5 * cm))
    story.append(HRFlowable(width="60%", thickness=1, color=C_GRIS_LINEA, spaceAfter=8))

    nombre = dp.get("responsable", "_____________________________")
    cedula = dp.get("cedula", "")
    fecha_doc = str(dp.get("fecha", date.today()))

    firma_data = [
        [Paragraph(f"<b>{nombre}</b>", S["firma"]),
         Paragraph(f"<b>{fecha_doc}</b>", S["firma"])],
        [Paragraph("Responsable Técnico", S["pie"]),
         Paragraph("Fecha", S["pie"])],
        [Paragraph(f"Cédula Profesional: {cedula}" if cedula else "", S["pie"]),
         Paragraph("", S["pie"])],
    ]
    t_firma = Table(firma_data, colWidths=[9 * cm, 8.5 * cm])
    t_firma.setStyle(TableStyle([
        ("ALIGN", (0, 0), (-1, -1), "CENTER"),
        ("TOPPADDING", (0, 0), (-1, 0), 8),
    ]))
    story.append(t_firma)

    story.append(Spacer(1, 0.5 * cm))
    story.append(HRFlowable(width="100%", thickness=3, color=C_AZUL_OSC, spaceAfter=6))
    story.append(Paragraph(
        f"eSuit · {dp.get('proyecto','')} · "
        f"{dp.get('norma','NOM-001-SEDE-2012')}",
        S["pie"]
    ))

    # ═══════════════════════════════════════════
    # BUILD
    # ═══════════════════════════════════════════
    doc.build(story)
    buffer.seek(0)
    return buffer.read()
