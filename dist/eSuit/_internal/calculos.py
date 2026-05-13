"""
calculos.py — Módulo de cálculos eléctricos
NOM-001-SEDE-2012 / NEC

Métodos de caída de tensión:
  - Por Impedancia (Tabla 9 NOM): incluye R, X por canalización y ajuste a temp. terminales (Art. 110-14)
  - Por Sección Transversal: cobre, despreciando reactancia (recomendado <4 AWG)
  - Aéreo: resistividad ρ, conductores NEUTRANEL/ACSR

Configuraciones:
  - Monofásico 2H   (Fase + Neutro, 2 hilos)
  - Monofásico 3H   (Fase + Fase + Neutro, 3 hilos)  ≡ Bifásico
  - Trifásico

Basado en hoja de cálculo del Ing. Juan Carlos Vega Mendoza (ESIME-IPN).
"""
import math
import pandas as pd

# ─────────────────────────────────────────────
# TABLA DE CONDUCTORES (Cu y Al) — AWG/kcmil
# Ampacidades THW 75°C en tubo (NOM-001 / NEC Table 310.15)
# Resistencia AC a 75°C en conduit metálico (Ω/km) — NEC Table 9
# Reactancia inductiva a 60Hz en conduit metálico (Ω/km) — NEC Table 9
# ─────────────────────────────────────────────
TABLA_CONDUCTORES = {
    #         area    ampCu ampAl  rCu    rAl    xCu    xAl    od_mm
    "14":  {"area_mm2": 2.08,  "amp_Cu": 20,  "amp_Al": 15,  "r_Cu": 9.61,  "r_Al": 15.6,  "x_Cu": 0.213, "x_Al": 0.213, "od_mm": 3.68},
    "12":  {"area_mm2": 3.31,  "amp_Cu": 25,  "amp_Al": 20,  "r_Cu": 6.04,  "r_Al": 9.88,  "x_Cu": 0.197, "x_Al": 0.197, "od_mm": 4.11},
    "10":  {"area_mm2": 5.26,  "amp_Cu": 35,  "amp_Al": 30,  "r_Cu": 3.79,  "r_Al": 6.17,  "x_Cu": 0.180, "x_Al": 0.180, "od_mm": 4.67},
    "8":   {"area_mm2": 8.37,  "amp_Cu": 50,  "amp_Al": 40,  "r_Cu": 2.51,  "r_Al": 3.88,  "x_Cu": 0.164, "x_Al": 0.164, "od_mm": 5.94},
    "6":   {"area_mm2": 13.3,  "amp_Cu": 65,  "amp_Al": 50,  "r_Cu": 1.61,  "r_Al": 2.49,  "x_Cu": 0.148, "x_Al": 0.148, "od_mm": 7.42},
    "4":   {"area_mm2": 21.2,  "amp_Cu": 85,  "amp_Al": 65,  "r_Cu": 1.02,  "r_Al": 1.57,  "x_Cu": 0.131, "x_Al": 0.131, "od_mm": 8.76},
    "3":   {"area_mm2": 26.7,  "amp_Cu": 100, "amp_Al": 75,  "r_Cu": 0.814, "r_Al": 1.25,  "x_Cu": 0.131, "x_Al": 0.131, "od_mm": 9.65},
    "2":   {"area_mm2": 33.6,  "amp_Cu": 115, "amp_Al": 90,  "r_Cu": 0.643, "r_Al": 0.991, "x_Cu": 0.128, "x_Al": 0.128, "od_mm": 10.8},
    "1":   {"area_mm2": 42.4,  "amp_Cu": 130, "amp_Al": 100, "r_Cu": 0.515, "r_Al": 0.787, "x_Cu": 0.125, "x_Al": 0.125, "od_mm": 11.9},
    "1/0": {"area_mm2": 53.5,  "amp_Cu": 150, "amp_Al": 120, "r_Cu": 0.407, "r_Al": 0.627, "x_Cu": 0.122, "x_Al": 0.122, "od_mm": 13.3},
    "2/0": {"area_mm2": 67.4,  "amp_Cu": 175, "amp_Al": 135, "r_Cu": 0.323, "r_Al": 0.497, "x_Cu": 0.118, "x_Al": 0.118, "od_mm": 14.9},
    "3/0": {"area_mm2": 85.0,  "amp_Cu": 200, "amp_Al": 155, "r_Cu": 0.256, "r_Al": 0.394, "x_Cu": 0.115, "x_Al": 0.115, "od_mm": 16.7},
    "4/0": {"area_mm2": 107.2, "amp_Cu": 230, "amp_Al": 180, "r_Cu": 0.203, "r_Al": 0.313, "x_Cu": 0.111, "x_Al": 0.111, "od_mm": 18.8},
    "250": {"area_mm2": 126.7, "amp_Cu": 255, "amp_Al": 205, "r_Cu": 0.171, "r_Al": 0.264, "x_Cu": 0.108, "x_Al": 0.108, "od_mm": 20.4},
    "300": {"area_mm2": 152.0, "amp_Cu": 285, "amp_Al": 230, "r_Cu": 0.144, "r_Al": 0.220, "x_Cu": 0.105, "x_Al": 0.105, "od_mm": 22.5},
    "350": {"area_mm2": 177.3, "amp_Cu": 310, "amp_Al": 250, "r_Cu": 0.122, "r_Al": 0.189, "x_Cu": 0.103, "x_Al": 0.103, "od_mm": 24.2},
    "400": {"area_mm2": 202.7, "amp_Cu": 335, "amp_Al": 270, "r_Cu": 0.109, "r_Al": 0.167, "x_Cu": 0.100, "x_Al": 0.100, "od_mm": 25.9},
    "500": {"area_mm2": 253.4, "amp_Cu": 380, "amp_Al": 310, "r_Cu": 0.0869,"r_Al": 0.135, "x_Cu": 0.098, "x_Al": 0.098, "od_mm": 28.6},
}

# Resistividad Cu y Al (Ω·mm²/m) a 75°C — para Método Aéreo
RHO = {"Cu": 0.0217, "Al": 0.0354}

# Orden calibres de menor a mayor
ORDEN_CALIBRES = ["14", "12", "10", "8", "6", "4", "3", "2", "1",
                  "1/0", "2/0", "3/0", "4/0", "250", "300", "350", "400", "500"]

# ─────────────────────────────────────────────
# TABLA CONDUIT (área interior neta, mm²)
# EMT / IMC / RGS por diámetro nominal
# ─────────────────────────────────────────────
TABLA_CONDUIT = {
    "EMT (Conduit metálico ligero)": {
        '1/2"':  78.1,
        '3/4"':  137.1,
        '1"':    228.0,
        '1-1/4"': 387.1,
        '1-1/2"': 526.0,
        '2"':    822.6,
        '2-1/2"': 1314.5,
        '3"':    1963.5,
        '3-1/2"': 2580.0,
        '4"':    3268.0,
    },
    "IMC (Conduit metálico intermedio)": {
        '1/2"':  84.5,
        '3/4"':  148.5,
        '1"':    243.6,
        '1-1/4"': 404.0,
        '1-1/2"': 548.7,
        '2"':    855.5,
        '2-1/2"': 1363.7,
        '3"':    2038.0,
        '3-1/2"': 2666.0,
        '4"':    3366.0,
    },
    "RGS (Rígido galvanizado)": {
        '1/2"':  81.2,
        '3/4"':  144.1,
        '1"':    236.7,
        '1-1/4"': 396.1,
        '1-1/2"': 533.5,
        '2"':    836.5,
        '2-1/2"': 1330.4,
        '3"':    1990.3,
        '3-1/2"': 2581.0,
        '4"':    3287.0,
    },
}


# ═════════════════════════════════════════════════════════
# TABLA 9 NOM-001-SEDE-2012 (R y X efectivas, Ω/km, 75°C)
# Depende de canalización:
#   - "PVC": tubería no metálica (PVC, PE)
#   - "AL" : tubería metálica no magnética (aluminio)
#   - "AC" : tubería metálica magnética (acero, EMT, IMC, RGS)
# X: PVC y AL comparten valor (canalizaciones no magnéticas)
# ═════════════════════════════════════════════════════════
TABLA_9_NOM = {
    "14":  {"area_mm2": 2.08,
            "X":    {"PVC": 0.190, "AL": 0.190, "AC": 0.240},
            "R_Cu": {"PVC": 10.2,  "AL": 10.2,  "AC": 10.2},
            "R_Al": None},
    "12":  {"area_mm2": 3.31,
            "X":    {"PVC": 0.177, "AL": 0.177, "AC": 0.223},
            "R_Cu": {"PVC": 6.6,   "AL": 6.6,   "AC": 6.6},
            "R_Al": None},
    "10":  {"area_mm2": 5.26,
            "X":    {"PVC": 0.164, "AL": 0.164, "AC": 0.207},
            "R_Cu": {"PVC": 3.9,   "AL": 3.9,   "AC": 3.9},
            "R_Al": None},
    "8":   {"area_mm2": 8.37,
            "X":    {"PVC": 0.171, "AL": 0.171, "AC": 0.213},
            "R_Cu": {"PVC": 2.56,  "AL": 2.56,  "AC": 2.56},
            "R_Al": None},
    "6":   {"area_mm2": 13.3,
            "X":    {"PVC": 0.167, "AL": 0.167, "AC": 0.210},
            "R_Cu": {"PVC": 1.61,  "AL": 1.61,  "AC": 1.61},
            "R_Al": {"PVC": 2.66,  "AL": 2.66,  "AC": 2.66}},
    "4":   {"area_mm2": 21.2,
            "X":    {"PVC": 0.157, "AL": 0.157, "AC": 0.197},
            "R_Cu": {"PVC": 1.02,  "AL": 1.02,  "AC": 1.02},
            "R_Al": {"PVC": 1.67,  "AL": 1.67,  "AC": 1.67}},
    "2":   {"area_mm2": 33.6,
            "X":    {"PVC": 0.148, "AL": 0.148, "AC": 0.187},
            "R_Cu": {"PVC": 0.62,  "AL": 0.66,  "AC": 0.66},
            "R_Al": {"PVC": 1.05,  "AL": 1.05,  "AC": 1.05}},
    "1/0": {"area_mm2": 53.49,
            "X":    {"PVC": 0.144, "AL": 0.144, "AC": 0.180},
            "R_Cu": {"PVC": 0.39,  "AL": 0.43,  "AC": 0.39},
            "R_Al": {"PVC": 0.66,  "AL": 0.69,  "AC": 0.66}},
    "2/0": {"area_mm2": 67.43,
            "X":    {"PVC": 0.141, "AL": 0.141, "AC": 0.177},
            "R_Cu": {"PVC": 0.33,  "AL": 0.33,  "AC": 0.33},
            "R_Al": {"PVC": 0.52,  "AL": 0.52,  "AC": 0.52}},
    "3/0": {"area_mm2": 85.01,
            "X":    {"PVC": 0.138, "AL": 0.138, "AC": 0.171},
            "R_Cu": {"PVC": 0.253, "AL": 0.269, "AC": 0.259},
            "R_Al": {"PVC": 0.43,  "AL": 0.43,  "AC": 0.43}},
    "4/0": {"area_mm2": 107.2,
            "X":    {"PVC": 0.135, "AL": 0.135, "AC": 0.167},
            "R_Cu": {"PVC": 0.203, "AL": 0.220, "AC": 0.207},
            "R_Al": {"PVC": 0.33,  "AL": 0.36,  "AC": 0.33}},
    "250": {"area_mm2": 127.0,
            "X":    {"PVC": 0.135, "AL": 0.135, "AC": 0.171},
            "R_Cu": {"PVC": 0.171, "AL": 0.187, "AC": 0.177},
            "R_Al": {"PVC": 0.279, "AL": 0.295, "AC": 0.282}},
    "300": {"area_mm2": 152.0,
            "X":    {"PVC": 0.135, "AL": 0.135, "AC": 0.167},
            "R_Cu": {"PVC": 0.144, "AL": 0.161, "AC": 0.148},
            "R_Al": {"PVC": 0.233, "AL": 0.249, "AC": 0.236}},
    "350": {"area_mm2": 177.0,
            "X":    {"PVC": 0.131, "AL": 0.131, "AC": 0.164},
            "R_Cu": {"PVC": 0.125, "AL": 0.141, "AC": 0.128},
            "R_Al": {"PVC": 0.200, "AL": 0.217, "AC": 0.207}},
    "400": {"area_mm2": 203.0,
            "X":    {"PVC": 0.131, "AL": 0.131, "AC": 0.161},
            "R_Cu": {"PVC": 0.108, "AL": 0.125, "AC": 0.115},
            "R_Al": {"PVC": 0.177, "AL": 0.194, "AC": 0.180}},
    "500": {"area_mm2": 253.0,
            "X":    {"PVC": 0.128, "AL": 0.128, "AC": 0.157},
            "R_Cu": {"PVC": 0.089, "AL": 0.105, "AC": 0.095},
            "R_Al": {"PVC": 0.141, "AL": 0.157, "AC": 0.148}},
}

# ═════════════════════════════════════════════════════════
# TABLA 310-15(B)(16) NOM-001 — Ampacidad en tubo (3 conds)
# (60°C, 75°C, 90°C) por material conductor
# Para THW use 75°C; para THHW-2/XHHW-2 use 90°C
# Recordatorio Art. 110-14: la ampacidad se limita por la
# temperatura del terminal (60° si I ≤ 100A, 75° si > 100A).
# ═════════════════════════════════════════════════════════
TABLA_AMPACIDAD = {
    "14":  {"Cu": (15, 20, 25),    "Al": (None, None, None)},
    "12":  {"Cu": (20, 25, 30),    "Al": (None, None, None)},
    "10":  {"Cu": (30, 35, 40),    "Al": (None, None, None)},
    "8":   {"Cu": (40, 50, 55),    "Al": (None, None, None)},
    "6":   {"Cu": (55, 65, 75),    "Al": (40, 50, 55)},
    "4":   {"Cu": (70, 85, 95),    "Al": (55, 65, 75)},
    "2":   {"Cu": (95, 115, 130),  "Al": (75, 90, 100)},
    "1/0": {"Cu": (125, 150, 170), "Al": (100, 120, 135)},
    "2/0": {"Cu": (145, 175, 195), "Al": (115, 135, 150)},
    "3/0": {"Cu": (165, 200, 225), "Al": (130, 155, 175)},
    "4/0": {"Cu": (195, 230, 260), "Al": (150, 180, 205)},
    "250": {"Cu": (215, 255, 290), "Al": (170, 205, 230)},
    "300": {"Cu": (240, 285, 320), "Al": (195, 230, 260)},
    "350": {"Cu": (260, 310, 350), "Al": (210, 250, 280)},
    "400": {"Cu": (280, 335, 380), "Al": (225, 270, 305)},
    "500": {"Cu": (320, 380, 430), "Al": (260, 310, 350)},
}

# Mapas legibles canalización ↔ código
CANALIZACIONES = {
    "PVC (no metálica)": "PVC",
    "Aluminio (no magnética)": "AL",
    "Acero (EMT/IMC/RGS, magnética)": "AC",
}
AISLAMIENTOS_T = {"60°C": 60, "75°C": 75, "90°C": 90}


def temp_terminales_auto(I_nom: float, es_motor: bool = False) -> int:
    """Art. 110-14(C) NOM-001 / NEC:
       - I ≤ 100 A → terminal 60°C (a menos que estén marcados 75°C)
       - I > 100 A → terminal 75°C
       - Motores (excepto Clase A) → 75°C
    """
    if es_motor:
        return 75
    return 60 if I_nom <= 100 else 75


def ajustar_r_terminales(R_75: float, T_term: int) -> float:
    """Ajuste de resistencia por temperatura del terminal (Art. 110-14):
       R_T = R_75 × (T + 234.5) / (75 + 234.5)
    """
    return R_75 * (T_term + 234.5) / (75 + 234.5)


def r_x_efectivas(calibre: str, material: str,
                   canalizacion: str = "AC", temp_term: int = 75) -> tuple:
    """Retorna (R, X) en Ω/km según Tabla 9 NOM, ajustada a temp del terminal.

       material:     'Cu' | 'Al'
       canalizacion: 'PVC' | 'AL' | 'AC'
       temp_term:    60 | 75 | 90
    """
    if calibre in TABLA_9_NOM:
        d = TABLA_9_NOM[calibre]
        R_map = d.get(f"R_{material}")
        if R_map is None:
            raise ValueError(
                f"Calibre {calibre} no disponible para {material} en Tabla 9 NOM"
            )
        R75 = R_map[canalizacion]
        X = d["X"][canalizacion]
    else:
        # Fallback a tabla legacy (calibres 1, 3)
        d = TABLA_CONDUCTORES.get(calibre)
        if d is None:
            raise ValueError(f"Calibre desconocido: {calibre}")
        R75 = d[f"r_{material}"]
        X = d[f"x_{material}"]
    return ajustar_r_terminales(R75, temp_term), X


def ampacidad_base(calibre: str, material: str, temp_aislamiento: int = 75) -> float:
    """Ampacidad base NOM 310-15(B)(16) por aislamiento (60/75/90 °C).

       Si el calibre no está en la tabla nueva, usa la tabla legacy (THW 75°C).
    """
    if calibre in TABLA_AMPACIDAD:
        idx = {60: 0, 75: 1, 90: 2}[temp_aislamiento]
        val = TABLA_AMPACIDAD[calibre][material][idx]
        return float(val) if val is not None else 0.0
    return float(TABLA_CONDUCTORES.get(calibre, {}).get(f"amp_{material}", 0))


# ─────────────────────────────────────────────
# FACTORES DE TEMPERATURA (THW 75°C base 30°C)
# ─────────────────────────────────────────────
def factor_correccion_temp(temp_c: float, temp_conductor: int = 75) -> float:
    """Factor de corrección por temperatura ambiente (NOM-001-SEDE-2012
       Tabla 310-15(B)(2)(a)).

       Para conductor de 75°C (THW, RHW): FC = sqrt((75 − Tamb) / (75 − 30))
       Es una función monótona decreciente. Para T ambiente ≥ T_conductor,
       el conductor no puede operar (retorna 0.10 como mínimo de seguridad).

       Validada contra la NOM oficial:
         35°C → 0.94 · 40°C → 0.88 · 45°C → 0.82
         50°C → 0.75 · 55°C → 0.67 · 60°C → 0.58
    """
    if temp_c <= 30:
        return 1.00
    if temp_c >= temp_conductor:
        return 0.10  # mínimo de seguridad
    fc = ((temp_conductor - temp_c) / (temp_conductor - 30)) ** 0.5
    return round(max(0.10, fc), 2)


# ─────────────────────────────────────────────
# CÁLCULO DE CORRIENTE
# ─────────────────────────────────────────────
def calcular_corriente(potencia_w: float, fp: float, voltaje: float, configuracion: str) -> float:
    """
    configuracion: 'mono2h' | 'mono3h' | 'trifasico'
    mono2h  = Monofásico 2H  (Fase + Neutro)      I = P / (V × FP)
    mono3h  = Monofásico 3H  (Fase + Fase + Neutro) I = P / (V × FP)  [misma fórmula, distinto circuito]
    trifasico                                      I = P / (√3 × V × FP)
    """
    if configuracion == "trifasico":
        return potencia_w / (math.sqrt(3) * voltaje * fp)
    else:
        return potencia_w / (voltaje * fp)


# ═══════════════════════════════════════════════════════
# MÉTODOS DE CAÍDA DE TENSIÓN
# Fuente: ElectroCalc NOM-001-SEDE / referencia técnica
# ═══════════════════════════════════════════════════════

def cdt_por_impedancia(I: float, L_km: float, R: float, X: float,
                        fp: float, voltaje: float, configuracion: str) -> float:
    """
    Método por Impedancia — incluye componente resistiva y reactiva.

    Mono 2H:   %VD = (200 × I × L × (R·cosφ + X·senφ)) / V
    Mono 3H:   %VD = (200 × I × L × (R·cosφ + X·senφ)) / V
    Trifásico: %VD = (√3 × I × L × (R·cosφ + X·senφ) × 100) / V

    L en km, R y X en Ω/km, V en voltios, I en amperes.
    """
    sen_phi = math.sqrt(1 - fp ** 2)
    Z_efectiva = R * fp + X * sen_phi   # Ω/km

    if configuracion == "trifasico":
        return (math.sqrt(3) * I * L_km * Z_efectiva * 100) / voltaje
    else:
        # Factor 200 = 2 (ida+vuelta) × 100 (para %)
        return (200 * I * L_km * Z_efectiva) / voltaje


def cdt_por_seccion(I: float, L_km: float, S_mm2: float,
                     voltaje: float, configuracion: str) -> float:
    """
    Método por Sección Transversal — usa área del conductor directamente.

    Mono 2H:   %VD = (4 × L[m] × I) / (120 × S)
    Mono 3H:   %VD = (2 × L[m] × I) / (120 × S)
    Trifásico: %VD = (2 × √3 × L[m] × I) / (V × S)

    L en km (×1000 → metros), S en mm², V en voltios.
    Constante 120 ≈ conductividad Cu a 75°C × tensión de referencia.
    """
    L_m = L_km * 1000

    if configuracion == "mono2h":
        return (4 * L_m * I) / (120 * S_mm2)
    elif configuracion == "mono3h":
        return (2 * L_m * I) / (120 * S_mm2)
    else:  # trifasico
        return (2 * math.sqrt(3) * L_m * I) / (voltaje * S_mm2)


def cdt_por_aereo(I: float, L_km: float, S_mm2: float,
                   material: str, configuracion: str) -> float:
    """
    Método Aéreo — usa resistividad ρ del material.
    Para conductores NEUTRANEL, ACSR (reactancia despreciable).

    Monofásico: VD[V] = (2 × L[m] × I × ρ) / S
    Trifásico:  VD[V] = (√3 × L[m] × I × ρ) / S

    Retorna caída en voltios (no en %).
    """
    rho = RHO[material]
    L_m = L_km * 1000

    if configuracion == "trifasico":
        return (math.sqrt(3) * L_m * I * rho) / S_mm2
    else:
        return (2 * L_m * I * rho) / S_mm2


# ─────────────────────────────────────────────
# DESPACHADOR — CDT según método elegido
# ─────────────────────────────────────────────
def calcular_cdt_calibre(calibre: str, I: float, L_km: float,
                          material: str, fp: float, voltaje: float,
                          configuracion: str, metodo: str, cf: int = 1,
                          canalizacion: str = "AC", temp_term: int = 75) -> float:
    """Calcula %CDT para un calibre dado.

    cf:           conductores por fase (paralelos). I por conductor = I / cf.
    canalizacion: 'PVC' | 'AL' | 'AC'  (afecta R y X)
    temp_term:    60 | 75 | 90  (ajusta R por Art. 110-14)
    """
    S = TABLA_9_NOM.get(calibre, TABLA_CONDUCTORES[calibre]).get(
        "area_mm2", TABLA_CONDUCTORES.get(calibre, {}).get("area_mm2", 0)
    )
    I_cond = I / cf

    if metodo == "impedancia":
        R, X = r_x_efectivas(calibre, material, canalizacion, temp_term)
        return cdt_por_impedancia(I_cond, L_km, R, X, fp, voltaje, configuracion)
    elif metodo == "seccion":
        return cdt_por_seccion(I_cond, L_km, S, voltaje, configuracion)
    elif metodo == "aereo":
        vd_v = cdt_por_aereo(I_cond, L_km, S, material, configuracion)
        return (vd_v / voltaje) * 100
    else:
        raise ValueError(f"Método desconocido: {metodo}")


# ─────────────────────────────────────────────
# FUNCIÓN PRINCIPAL: CAÍDA DE TENSIÓN + SELECCIÓN
# ─────────────────────────────────────────────
def calcular_caida_tension(potencia, fp, voltaje, configuracion, longitud,
                            material, cdt_max, fc_temp, fc_agrup, metodo="impedancia",
                            factor_demanda=1.25, cf=1,
                            canalizacion="AC", temp_term=None,
                            temp_aislamiento=75, es_motor=False):
    """
    Retorna: corriente, conductor_min, cdt_real, conductor_final, ampacity_corr

    configuracion:    'mono2h' | 'mono3h' | 'trifasico'
    metodo:           'impedancia' | 'seccion' | 'aereo'
    longitud:         metros → se convierte a km internamente
    factor_demanda:   1.25 para cargas continuas (NEC 210.20 / NOM-001)
    cf:               conductores por fase en paralelo (default 1)
    canalizacion:     'PVC' | 'AL' | 'AC' (Tabla 9 NOM)
    temp_term:        60 | 75 | 90; None = auto (Art. 110-14 por In)
    temp_aislamiento: 60 | 75 | 90 (ampacidad base — THW=75°C)
    es_motor:         True fuerza terminal 75°C (motores no Clase A)
    """
    L_km = longitud / 1000

    # 1. Corriente real de la carga
    I = calcular_corriente(potencia, fp, voltaje, configuracion)
    I_diseño_cond = (I * factor_demanda) / cf

    # 2. Temperatura de terminales: si auto, depende de In
    if temp_term is None:
        temp_term = temp_terminales_auto(I, es_motor=es_motor)

    # 3. Calibres válidos para el material seleccionado (Al sólo desde 6 AWG)
    calibres_disp = [
        c for c in ORDEN_CALIBRES
        if ampacidad_base(c, material, temp_aislamiento) > 0
    ]
    if not calibres_disp:
        calibres_disp = ORDEN_CALIBRES

    # 4. Selección inicial por ampacidad
    conductor_min = None
    for cal in calibres_disp:
        amp_corr = ampacidad_base(cal, material, temp_aislamiento) * fc_temp * fc_agrup
        if amp_corr >= I_diseño_cond:
            conductor_min = cal
            break
    if conductor_min is None:
        conductor_min = calibres_disp[-1]

    # 5. Subir calibre hasta cumplir c.d.t.
    cdt_real = None
    conductor_final = calibres_disp[-1]
    ampacity_corr = 0

    for cal in calibres_disp[calibres_disp.index(conductor_min):]:
        amp_corr = ampacidad_base(cal, material, temp_aislamiento) * fc_temp * fc_agrup
        cdt = calcular_cdt_calibre(cal, I, L_km, material, fp, voltaje,
                                    configuracion, metodo, cf=cf,
                                    canalizacion=canalizacion, temp_term=temp_term)
        if cdt <= cdt_max and amp_corr >= I_diseño_cond:
            conductor_final = cal
            cdt_real = cdt
            ampacity_corr = amp_corr
            break

    if cdt_real is None:
        cal = calibres_disp[-1]
        cdt_real = calcular_cdt_calibre(cal, I, L_km, material, fp, voltaje,
                                         configuracion, metodo, cf=cf,
                                         canalizacion=canalizacion, temp_term=temp_term)
        ampacity_corr = ampacidad_base(cal, material, temp_aislamiento) * fc_temp * fc_agrup
        conductor_final = cal

    return I, conductor_min, cdt_real, conductor_final, ampacity_corr


# ─────────────────────────────────────────────
# TABLA DE VERIFICACIÓN DE CONDUCTORES
# ─────────────────────────────────────────────
def seleccionar_conductor(I, longitud, voltaje, configuracion, material,
                           cdt_max, fc_temp, fc_agrup, fp=0.9, metodo="impedancia",
                           factor_demanda=1.25, cf=1,
                           canalizacion="AC", temp_term=None,
                           temp_aislamiento=75, es_motor=False):
    """Genera tabla comparativa de calibres candidatos.

       Incluye canalización y ajuste a temperatura de terminales (Art. 110-14).
    """
    L_km = longitud / 1000
    I_diseño_cond = (I * factor_demanda) / cf
    if temp_term is None:
        temp_term = temp_terminales_auto(I, es_motor=es_motor)

    calibres_disp = [
        c for c in ORDEN_CALIBRES
        if ampacidad_base(c, material, temp_aislamiento) > 0
    ]
    if not calibres_disp:
        calibres_disp = ORDEN_CALIBRES

    filas = []
    primer_ok = None  # primer calibre que cumple ambos criterios

    for cal in calibres_disp:
        amp_base = ampacidad_base(cal, material, temp_aislamiento)
        amp_corr = amp_base * fc_temp * fc_agrup
        cdt = calcular_cdt_calibre(cal, I, L_km, material, fp, voltaje,
                                    configuracion, metodo, cf=cf,
                                    canalizacion=canalizacion, temp_term=temp_term)
        area = TABLA_9_NOM.get(cal, TABLA_CONDUCTORES.get(cal, {})).get("area_mm2", 0)

        cumple_amp = amp_corr >= I_diseño_cond
        cumple_cdt = cdt <= cdt_max
        cumple_ambos = cumple_amp and cumple_cdt

        if cumple_ambos and primer_ok is None:
            primer_ok = cal

        # Caracteres Unicode estándar (no emoji) que sí están en Garamond/Times/Courier
        marca = "★" if (cumple_ambos and cal == primer_ok) else ("✓" if cumple_ambos else "")

        filas.append({
            "Calibre": cal,
            "Área (mm²)": area,
            "Amp. base": amp_base,
            "FC·FD": f"{fc_temp * fc_agrup:.3f}",
            "Amp. corr.": f"{amp_corr:.1f}",
            "I/CF req.": f"{I_diseño_cond:.1f}",
            "C.d.T. (%)": f"{cdt:.2f}",
            "Amp OK": "✓" if cumple_amp else "✗",
            f"≤ {cdt_max}%": "✓" if cumple_cdt else "✗",
            "Selección": marca,
        })

    return pd.DataFrame(filas)


# ─────────────────────────────────────────────
# CÁLCULO CONDUIT FILL
# ─────────────────────────────────────────────
def calcular_conduit_fill(calibre: str, num_conds: int, tipo_tubo: str) -> list:
    """
    Calcula el % de relleno para todos los diámetros del tipo de tubería.
    Retorna lista completa (todas las medidas), marcando cuál es la recomendada
    (la primera que cumple el % máximo).
    """
    od = TABLA_CONDUCTORES[calibre]["od_mm"]
    area_1_cond = math.pi * (od / 2) ** 2  # mm²
    area_total_conds = area_1_cond * num_conds

    # % máximo de relleno NOM-001 / NEC 310
    if num_conds == 1:
        fill_max = 53.0
    elif num_conds == 2:
        fill_max = 31.0
    else:
        fill_max = 40.0

    tabla_tubo = TABLA_CONDUIT.get(tipo_tubo, TABLA_CONDUIT["EMT (Conduit metálico ligero)"])
    resultados = []
    primer_ok = None

    for diam, area_tubo in tabla_tubo.items():
        fill_pct = (area_total_conds / area_tubo) * 100
        cumple = fill_pct <= fill_max
        if cumple and primer_ok is None:
            primer_ok = diam
        resultados.append({
            "tubo": f"{diam} {tipo_tubo.split()[0]}",
            "diametro": diam,
            "area_tubo": round(area_tubo, 1),
            "area_conds": round(area_total_conds, 2),
            "fill_pct": round(fill_pct, 1),
            "fill_max": fill_max,
            "cumple": cumple,
            "recomendada": (diam == primer_ok),
        })

    return resultados


# ─────────────────────────────────────────────
# PROTECCIONES — OCPD ESTÁNDAR NOM-001 / NEC 240
# ─────────────────────────────────────────────
OCPD_ESTANDARES = [
    15, 20, 25, 30, 35, 40, 45, 50, 60, 70, 80, 90, 100,
    110, 125, 150, 175, 200, 225, 250, 300, 350, 400, 450,
    500, 600, 700, 800, 1000, 1200, 1600, 2000,
]


def calcular_proteccion(I_diseño_cond: float, amp_conductor_corr: float):
    """
    Selecciona el OCPD estándar según NOM-001 Art. 240 / NEC 240.4.

    Regla: siguiente calibre estándar >= I_diseño_cond,
           sin exceder la ampacidad corregida del conductor.

    Returns (ocpd_A: int, status: str)
      'OK'      — OCPD dentro de la ampacidad del conductor
      'REVISAR' — OCPD supera la ampacidad; considerar conductor mayor
    """
    for size in OCPD_ESTANDARES:
        if size >= I_diseño_cond:
            status = "OK" if size <= amp_conductor_corr else "REVISAR"
            return size, status
    return OCPD_ESTANDARES[-1], "INSUFICIENTE"


POLOS_POR_CONFIG = {
    "mono2h": 1,       # 1 polo (Fase + Neutro)
    "mono3h": 2,       # 2 polos (Fase-Fase, comparten neutro)
    "trifasico": 3,    # 3 polos
}


def formato_proteccion(ocpd_A: int, configuracion: str) -> str:
    """Formato comercial 'NxA' según número de polos del sistema.

       mono2h    → 1x{ocpd}   (1 polo)
       mono3h    → 2x{ocpd}   (2 polos)
       trifasico → 3x{ocpd}   (3 polos)
    """
    p = POLOS_POR_CONFIG.get(configuracion, 1)
    return f"{p}x{ocpd_A}"


# ─────────────────────────────────────────────
# CALIBRE DE PUESTA A TIERRA (NOM-001 Tabla 250-122)
# Conductor de puesta a tierra de equipos en función de la
# capacidad de la protección contra sobrecorriente.
# ─────────────────────────────────────────────
# Cada fila: (Amp_proteccion_max, calibre_Cu, calibre_Al)
TABLA_TIERRA = [
    (15,    "14",  "12"),
    (20,    "12",  "10"),
    (60,    "10",  "8"),
    (100,   "8",   "6"),
    (200,   "6",   "4"),
    (300,   "4",   "2"),
    (400,   "3",   "1"),
    (500,   "2",   "1/0"),
    (600,   "1",   "2/0"),
    (800,   "1/0", "3/0"),
    (1000,  "2/0", "4/0"),
    (1200,  "3/0", "250"),
    (1600,  "4/0", "350"),
    (2000,  "250", "400"),
    (2500,  "350", "600"),
    (3000,  "400", "600"),
    (4000,  "500", "750"),
    (5000,  "700", "1200"),
    (6000,  "800", "1200"),
]


def calibre_tierra(amp_proteccion: float, material: str = "Cu") -> str:
    """Calibre del conductor de puesta a tierra (NOM-001 Tabla 250-122).

       amp_proteccion: capacidad nominal del OCPD del circuito (A)
       material:       'Cu' o 'Al'
    """
    col = 1 if material == "Cu" else 2
    for amp_max, cu, al in TABLA_TIERRA:
        if amp_proteccion <= amp_max:
            return (cu, al)[col - 1]
    return (TABLA_TIERRA[-1][1], TABLA_TIERRA[-1][2])[col - 1]
