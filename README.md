# ⚡ eSuit · Cálculo eléctrico profesional

Suite Streamlit para cálculos eléctricos profesionales conforme a la NOM-001-SEDE-2012.

## Características

- **Caída de tensión** por impedancia (Tabla 9 NOM con ajuste por terminal Art. 110-14), sección transversal y aéreo
- **Selección automática de conductor** por ampacidad (Tabla 310-15(B)(16)) + verificación de C.d.T.
- **Selección manual de calibre** con override persistente
- **Protección OCPD** en formato comercial (1×20, 2×40, 3×100 A) según NEC 240.6
- **Conductor de tierra** desnudo (NOM Tabla 250-122) automático o manual
- **Tubería (conduit fill)** EMT / IMC / RGS / PVC con cable de tierra independiente
- **Cuadro de cargas** profesional con balanceo de fases R/S/T (NEMA)
- **Reportes profesionales** en PDF (con índice clickeable) y Word (con TOC nativo)
- **Excel con fórmulas vivas** (no valores) — editable a mano
- **Autenticación** local con roles admin / usuario
- **Temas** Claro / Oscuro / Sistema con liquid glass iOS 26
- **Responsive** para escritorio, tablet y móvil

## Instalación

### Opción 1 — Ejecutable Windows (recomendado)
1. Doble click en `eSuit.bat` (instala dependencias y arranca)
2. O usa el ejecutable empaquetado: `dist/eSuit/eSuit.exe`

### Opción 2 — Desarrollo
```bash
pip install -r requirements.txt
streamlit run app.py
```

Se abre en `http://localhost:8501`.

**Login por defecto**: `admin` / `admin123` (cámbiala en el panel de Administración).

## Construir ejecutable

```bash
python build_exe.py            # carpeta dist/eSuit/ (recomendado)
python build_exe.py --onefile  # un solo .exe (arranque más lento)
```

## Archivos del proyecto

| Archivo | Descripción |
|---|---|
| `app.py` | UI Streamlit (cálculo, tablero, reportes, admin) |
| `calculos.py` | Tablas NOM-001 (Tabla 9, 310-15(B)(16), 250-122) + cálculos |
| `cuadro_cargas.py` | Balanceo R/S/T + Excel con fórmulas |
| `reporte.py` | Reporte PDF profesional con índice y bookmarks |
| `reporte_docx.py` | Reporte Word editable con TOC nativo |
| `auth.py` | Login + gestión de usuarios (SHA-256 + salt) |
| `launcher.py` | Entry point para el ejecutable |
| `eSuit.bat` | Lanzador Windows (modo desarrollo) |
| `build_exe.py` | Script PyInstaller para empaquetar |
| `requirements.txt` | Dependencias de Python |

## Flujo de trabajo

1. **Login** con tus credenciales
2. **Sidebar** → Datos del proyecto (nombre, cliente, norma, sistema, temperatura)
3. **Tab 🔌 Conductor** → Parámetros del circuito · resultados en vivo · guardar al proyecto
4. **Tab 🔧 Tubería** → Cálculo independiente de conduit fill (con tierra opcional)
5. **Tab 📊 Tablero** → Cuadro de cargas con balanceo de fases · exportar Excel
6. **Tab 📄 Reporte** → Generar PDF o Word del proyecto completo o un circuito
7. **Tab 👥 Admin** (solo admin) → Gestión de usuarios

## Normas implementadas

- **NOM-001-SEDE-2012** (principal, México)
- Tabla 9 (R, X conductores)
- Tabla 310-15(B)(16) (ampacidades)
- Tabla 250-122 (conductor de tierra)
- Art. 110-14 (temperatura de terminales)
- Art. 210.20, 240.6, 430.24 (protección OCPD)

Adicionalmente puedes referenciar en el reporte:
- NEC 2023
- CFE DCDIAMT (Instalaciones aéreas MT)
- CFE DCDIASMT (Instalaciones subterráneas MT)
- NOM-007-ENER-2014
- NMX-J-098-ANCE, NMX-J-235-ANCE
- IEEE Std 141 (Red Book)
- IEC 60364

---

Desarrollado para uso profesional en instalaciones eléctricas México.
