"""
main.py — Ensamblador del Modelo Financiero
Fondo Mutualista de Suelo Urbano
"""
import sys, os
sys.path.insert(0, os.path.dirname(__file__))

import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ── Importar módulos ─────────────────────────────────────────────────────────
from generar_modelo import crear_parametros
from flujo          import crear_flujo
from sensibilidad   import crear_sensibilidad
from dashboard      import crear_dashboard
from escenarios     import crear_escenarios

OUTPUT = os.path.join(os.path.dirname(__file__),
                      "Fondo_Mutualista_Suelo_Urbano.xlsx")

def main():
    print("Creando workbook...")
    wb = openpyxl.Workbook()
    # Eliminar hoja vacía por defecto
    wb.remove(wb.active)

    print("  [1/5] Hoja Parámetros...")
    ws_params, row_map = crear_parametros(wb)

    print("  [2/5] Hoja Flujo del Fondo (180 meses)...")
    ws_flujo = crear_flujo(wb)

    print("  [3/5] Hoja Análisis de Sensibilidad...")
    ws_sens = crear_sensibilidad(wb)

    print("  [4/5] Hoja Dashboard...")
    ws_dash = crear_dashboard(wb)

    print("  [5/5] Hoja Escenarios...")
    ws_esc = crear_escenarios(wb)

    # ── Propiedades del libro ────────────────────────────────────────────────
    wb.properties.title    = "Fondo Mutualista de Suelo Urbano"
    wb.properties.subject  = "Modelo Financiero - Community Land Trust"
    wb.properties.creator  = "Modelo CLT - es-CL"
    wb.properties.keywords = "CLT, suelo urbano, vivienda cooperativa, modelo financiero"
    wb.properties.description = (
        "Modelo financiero para Fondo Mutualista de Suelo Urbano. "
        "Horizonte 180 meses. 3 capas de capital: ahorro familiar, "
        "capital paciente e inversores de impacto."
    )

    # ── Orden de hojas ───────────────────────────────────────────────────────
    sheet_order = [
        "Parámetros",
        "Flujo del Fondo",
        "Análisis de Sensibilidad",
        "Dashboard",
        "Escenarios",
    ]
    for i, name in enumerate(sheet_order):
        if name in wb.sheetnames:
            wb.move_sheet(name, offset=i - wb.sheetnames.index(name))

    # Activar hoja Dashboard al abrir
    for i, ws in enumerate(wb.worksheets):
        ws.sheet_view.tabSelected = (ws.title == "Dashboard")

    print(f"Guardando {OUTPUT}...")
    wb.save(OUTPUT)
    print(f"✓ Archivo generado: {OUTPUT}")
    size_kb = os.path.getsize(OUTPUT) / 1024
    print(f"  Tamaño: {size_kb:.1f} KB")

if __name__ == "__main__":
    main()
