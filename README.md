# ⚡ eSuit

Aplicación Streamlit para cálculos eléctricos profesionales:
- Caída de tensión (monofásico, bifásico, trifásico)
- Selección de conductor por ampacidad + verificación c.d.t.
- Cálculo de conduit fill (EMT / IMC / RGS)
- Exportación de reporte PDF profesional

**Norma:** NOM-001-SEDE-2012

---

## Instalación y uso

### 1. Instalar dependencias
```bash
pip install -r requirements.txt
```

### 2. Ejecutar la app
```bash
streamlit run app.py
```

Se abrirá automáticamente en tu navegador en `http://localhost:8501`

---

## Archivos

| Archivo | Descripción |
|---|---|
| `app.py` | Interfaz principal (Streamlit) |
| `calculos.py` | Módulo de cálculos eléctricos + tablas NOM-001 |
| `reporte.py` | Generación del reporte PDF (reportlab) |
| `requirements.txt` | Dependencias de Python |

---

## Flujo de trabajo

1. **Sidebar** → Ingresa datos del proyecto (nombre, cliente, norma, sistema, temperatura)
2. **Tab 1** → Ingresa datos del circuito → presiona **Calcular**
3. **Tab 2** → Verifica conduit fill (opcional)
4. **Tab 3** → Revisa el resumen general
5. **Tab 4** → Descarga el reporte PDF

---

Desarrollado para uso profesional en instalaciones eléctricas México.
