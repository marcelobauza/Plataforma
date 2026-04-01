"""
Modelo Financiero - Fondo Mutualista de Suelo Urbano
Genera archivo .xlsx con 5 hojas completamente funcionales
"""
import openpyxl
from openpyxl.styles import (
    PatternFill, Font, Alignment, Border, Side, numbers
)
from openpyxl.utils import get_column_letter, quote_sheetname
from openpyxl.chart import AreaChart, LineChart, BarChart, Reference
from openpyxl.chart.series import SeriesLabel
from openpyxl.formatting.rule import ColorScaleRule
from openpyxl.comments import Comment
from openpyxl.worksheet.table import Table, TableStyleInfo
import math

# ── Estilos globales ────────────────────────────────────────────────────────
YELLOW   = PatternFill("solid", fgColor="FFFF99")
BLUE_HDR = PatternFill("solid", fgColor="1F3864")
BLUE_SUB = PatternFill("solid", fgColor="2E75B6")
GRAY_ALT = PatternFill("solid", fgColor="EBF3FF")
GREEN_BG = PatternFill("solid", fgColor="E2EFDA")
ORANGE   = PatternFill("solid", fgColor="FCE4D6")
WHITE    = PatternFill("solid", fgColor="FFFFFF")

FONT_HDR   = Font(name="Calibri", bold=True, color="FFFFFF", size=10)
FONT_TITLE = Font(name="Calibri", bold=True, color="1F3864", size=11)
FONT_BOLD  = Font(name="Calibri", bold=True, size=10)
FONT_NORM  = Font(name="Calibri", size=10)
FONT_SMALL = Font(name="Calibri", size=9, color="595959")

ALIGN_C  = Alignment(horizontal="center", vertical="center", wrap_text=True)
ALIGN_L  = Alignment(horizontal="left",   vertical="center")
ALIGN_R  = Alignment(horizontal="right",  vertical="center")

BORDER_ALL = Border(
    left=Side(style="thin"),  right=Side(style="thin"),
    top=Side(style="thin"),   bottom=Side(style="thin")
)
BORDER_BTM = Border(bottom=Side(style="medium"))
BORDER_TOP = Border(top=Side(style="medium"))

FMT_CLP  = '#,##0'          # CLP sin decimales
FMT_PCT  = '0.0%'
FMT_UF   = '#,##0.00'
FMT_INT  = '0'
FMT_MES  = '0'

P = "Parámetros"   # nombre hoja parámetros (referencia corta)

def col(c): return get_column_letter(c)

def add_comment(ws, cell_ref, text):
    c = Comment(text, "Modelo CLT")
    c.width  = 300
    c.height = 100
    ws[cell_ref].comment = c

def style_header_row(ws, row, col_start, col_end, fill=BLUE_HDR, font=FONT_HDR):
    for c in range(col_start, col_end+1):
        cell = ws.cell(row=row, column=c)
        cell.fill      = fill
        cell.font      = font
        cell.alignment = ALIGN_C
        cell.border    = BORDER_ALL

def style_input(ws, row, col_n, value=None, fmt=FMT_CLP, comment=None):
    cell = ws.cell(row=row, column=col_n)
    if value is not None:
        cell.value = value
    cell.fill           = YELLOW
    cell.font           = FONT_NORM
    cell.alignment      = ALIGN_R
    cell.number_format  = fmt
    cell.border         = BORDER_ALL
    if comment:
        add_comment(ws, cell.coordinate, comment)

def style_calc(ws, row, col_n, value=None, fmt=FMT_CLP, bold=False):
    cell = ws.cell(row=row, column=col_n)
    if value is not None:
        cell.value = value
    cell.font          = FONT_BOLD if bold else FONT_NORM
    cell.alignment     = ALIGN_R
    cell.number_format = fmt
    cell.border        = BORDER_ALL

def label(ws, row, col_n, text, bold=False, indent=0):
    cell = ws.cell(row=row, column=col_n)
    cell.value     = ("  " * indent) + text
    cell.font      = FONT_BOLD if bold else FONT_NORM
    cell.alignment = ALIGN_L
    cell.border    = BORDER_ALL

# ══════════════════════════════════════════════════════════════════════════════
# HOJA 1: Parámetros
# ══════════════════════════════════════════════════════════════════════════════
def crear_parametros(wb):
    ws = wb.create_sheet(P)
    ws.sheet_view.showGridLines = False
    ws.column_dimensions["A"].width = 42
    ws.column_dimensions["B"].width = 22
    ws.column_dimensions["C"].width = 18
    ws.column_dimensions["D"].width = 35

    # Título
    ws.merge_cells("A1:D1")
    t = ws["A1"]
    t.value     = "FONDO MUTUALISTA DE SUELO URBANO — Parámetros del Modelo"
    t.font      = Font(name="Calibri", bold=True, color="1F3864", size=14)
    t.alignment = ALIGN_C
    t.fill      = PatternFill("solid", fgColor="D6E4F7")
    ws.row_dimensions[1].height = 30

    ws.merge_cells("A2:D2")
    ws["A2"].value     = "Las celdas amarillas son editables. Todos los cálculos se actualizan automáticamente."
    ws["A2"].font      = FONT_SMALL
    ws["A2"].alignment = ALIGN_C
    ws["A2"].fill      = PatternFill("solid", fgColor="FFF2CC")
    ws.row_dimensions[2].height = 18

    # Encabezados de columna
    headers = ["Parámetro", "Valor", "Unidad", "Nota"]
    for i, h in enumerate(headers, 1):
        ws.cell(row=3, column=i).value = h
    style_header_row(ws, 3, 1, 4)
    ws.row_dimensions[3].height = 22

    row = 4

    def section(title, r):
        ws.merge_cells(f"A{r}:D{r}")
        c = ws[f"A{r}"]
        c.value     = f"▌ {title}"
        c.font      = Font(name="Calibri", bold=True, color="FFFFFF", size=10)
        c.fill      = BLUE_SUB
        c.alignment = ALIGN_L
        c.border    = BORDER_ALL
        ws.row_dimensions[r].height = 20
        return r + 1

    def param(ws, r, lbl, val, unit, fmt=FMT_CLP, note="", named=None, comment=None):
        label(ws, r, 1, lbl)
        style_input(ws, r, 2, val, fmt, comment)
        ws.cell(row=r, column=3).value          = unit
        ws.cell(row=r, column=3).font           = FONT_NORM
        ws.cell(row=r, column=3).alignment      = ALIGN_C
        ws.cell(row=r, column=3).border         = BORDER_ALL
        ws.cell(row=r, column=4).value          = note
        ws.cell(row=r, column=4).font           = FONT_SMALL
        ws.cell(row=r, column=4).alignment      = ALIGN_L
        ws.cell(row=r, column=4).border         = BORDER_ALL
        ws.row_dimensions[r].height = 18
        if named:
            from openpyxl.workbook.defined_name import DefinedName
            dn = DefinedName(name=named, attr_text=f"'{P}'!$B${r}")
            wb.defined_names.add(dn)
        return r + 1

    # ── Sección 1: Demanda ──────────────────────────────────────────────────
    row = section("1. DEMANDA", row)
    row = param(ws, row, "Número de familias participantes", 500, "familias",
                FMT_INT, "Base del fondo", "n_familias",
                "Total de familias que ingresan al fondo al inicio del período")
    row = param(ws, row, "Cuota mensual por familia", 30000, "CLP/mes",
                FMT_CLP, "Aporte mensual", "cuota_mensual",
                "Monto fijo que cada familia aporta mensualmente al fondo")
    row = param(ws, row, "Tasa de deserción anual estimada", 0.05, "anual",
                FMT_PCT, "5% base ajustada", "tasa_desercion",
                "Porcentaje de familias que abandonan el fondo cada año. Ajustado de 10% a 5% para lograr viabilidad en años 3-5")
    row = param(ws, row, "Familias por grupo de sorteo", 25, "familias",
                FMT_INT, "Tamaño lote", "familias_sorteo",
                "Número de familias que acceden al capital en cada sorteo periódico")

    # ── Sección 2: Suelo ────────────────────────────────────────────────────
    row = section("2. SUELO", row)
    row = param(ws, row, "Precio promedio terreno objetivo", 120000000, "CLP",
                FMT_CLP, "Por terreno completo", "precio_terreno",
                "Precio del terreno que aloja a un grupo de familias (viviendas_x_terreno familias)")
    row = param(ws, row, "Superficie promedio terreno", 1000, "m²",
                FMT_INT, "Referencial", "sup_terreno")
    row = param(ws, row, "Viviendas por terreno", 25, "unidades",
                FMT_INT, "Densidad", "viv_x_terreno",
                "Número de viviendas (familias) que caben en cada terreno adquirido")
    row = param(ws, row, "Apreciación anual estimada del suelo", 0.05, "anual",
                FMT_PCT, "Inflación suelo", "apreciacion_suelo",
                "Tasa anual de aumento del precio del suelo. Afecta el costo de terrenos futuros")

    # ── Sección 3: Capital Paciente Capa 2 ─────────────────────────────────
    row = section("3. CAPITAL PACIENTE — CAPA 2", row)
    row = param(ws, row, "Monto comprometido total", 500000000, "CLP",
                FMT_CLP, "Capital paciente total", "cap2_monto",
                "Capital total comprometido por inversores de impacto social (Capa 2)")
    row = param(ws, row, "Ratio desembolso Capa 2 / Capa 1", 3, "x",
                FMT_UF, "Por cada 1 CLP de Capa 1, desembolsa X de Capa 2", "cap2_ratio",
                "Apalancamiento: por cada peso acumulado en ahorro familiar, el capital paciente aporta X pesos adicionales")
    row = param(ws, row, "Retorno anual exigido Capa 2", 0.05, "anual",
                FMT_PCT, "Costo capital", "cap2_retorno",
                "Tasa de retorno anual que exigen los inversores de Capa 2. Determina el servicio de deuda mensual")
    row = param(ws, row, "Plazo del fondo", 15, "años",
                FMT_INT, "Horizonte total", "plazo_anos",
                "Horizonte total del modelo = 180 meses")
    row = param(ws, row, "Período de gracia Capa 2", 12, "meses",
                FMT_INT, "Sin interés inicial", "gracia_meses",
                "Meses iniciales sin cobro de interés sobre el capital paciente desembolsado")

    # ── Sección 4: Subsidio Estatal Capa 3 ─────────────────────────────────
    row = section("4. SUBSIDIO ESTATAL — CAPA 3", row)
    row = param(ws, row, "Subsidio por familia", 800, "UF/familia",
                FMT_UF, "Componente suelo", "subsidio_uf",
                "Subsidio habitacional estatal en UF por familia beneficiaria. Ejemplo: DS-49 tiene ~500 UF, se usa 800 UF para componente suelo")
    row = param(ws, row, "Valor UF", 38000, "CLP/UF",
                FMT_CLP, "Actualizar periódicamente", "valor_uf",
                "Valor de la Unidad de Fomento en CLP. Actualizar según valor vigente del Banco Central de Chile")
    row = param(ws, row, "% familias que acceden a subsidio", 0.60, "%",
                FMT_PCT, "Cobertura estimada", "pct_subsidio",
                "Porcentaje de familias beneficiarias que califican y obtienen el subsidio estatal")
    row = param(ws, row, "Destino del subsidio", 1.0, "%",
                FMT_PCT, "100% a suelo CLT", "subsidio_destino",
                "Fracción del subsidio destinada exclusivamente a la compra de suelo dentro del fideicomiso")

    # ── Sección 5: Canon Ground Lease ───────────────────────────────────────
    row = section("5. CANON DE USO DE SUELO (GROUND LEASE)", row)
    row = param(ws, row, "Canon mensual por familia en CLT", 25000, "CLP/mes",
                FMT_CLP, "Arrendamiento suelo", "canon_mensual",
                "Monto mensual que paga cada familia residente en CLT por el uso del suelo (no son dueños del suelo)")
    row = param(ws, row, "Crecimiento anual del canon (IPC)", 0.03, "anual",
                FMT_PCT, "Indexación IPC", "canon_crecimiento",
                "Tasa de reajuste anual del canon, ligada al IPC estimado")

    # ── Sección 6: Construcción (referencial) ───────────────────────────────
    row = section("6. CONSTRUCCIÓN (referencial — no parte del fondo de suelo)", row)
    row = param(ws, row, "Costo construcción por vivienda", 1200, "UF/vivienda",
                FMT_UF, "Referencial mercado", "costo_construccion",
                "Costo referencial de construcción por unidad. Se financia vía crédito bancario, NO con fondos del modelo")
    row = param(ws, row, "Tasa crédito bancario construcción", 0.04, "anual",
                FMT_PCT, "Tasa referencial", "tasa_credito")
    row = param(ws, row, "Plazo crédito bancario", 20, "años",
                FMT_INT, "Referencial", "plazo_credito")

    # ── Sección 7: Celdas auxiliares para sensibilidad ──────────────────────
    row = section("7. CELDAS AUXILIARES (usadas por tablas de sensibilidad)", row)
    ws.cell(row=row, column=1).value = "  Input variable cuota (Tabla 1)"
    ws.cell(row=row, column=1).font  = FONT_NORM
    ws.cell(row=row, column=1).border = BORDER_ALL
    ws.cell(row=row, column=2).value = f"=B5"   # cuota_mensual
    ws.cell(row=row, column=2).fill  = PatternFill("solid", fgColor="D9EAD3")
    ws.cell(row=row, column=2).number_format = FMT_CLP
    ws.cell(row=row, column=2).border = BORDER_ALL
    ws.cell(row=row, column=2).font  = FONT_NORM
    ws.cell(row=row, column=3).value = "CLP"
    ws.cell(row=row, column=3).border = BORDER_ALL
    ws.cell(row=row, column=4).value = "Variable de entrada para Tabla Sensibilidad 1"
    ws.cell(row=row, column=4).border = BORDER_ALL
    aux_cuota_row = row
    row += 1

    ws.cell(row=row, column=1).value = "  Input variable familias (Tabla 1 y 3)"
    ws.cell(row=row, column=1).font  = FONT_NORM
    ws.cell(row=row, column=1).border = BORDER_ALL
    ws.cell(row=row, column=2).value = f"=B4"   # n_familias
    ws.cell(row=row, column=2).fill  = PatternFill("solid", fgColor="D9EAD3")
    ws.cell(row=row, column=2).number_format = FMT_INT
    ws.cell(row=row, column=2).border = BORDER_ALL
    ws.cell(row=row, column=2).font  = FONT_NORM
    ws.cell(row=row, column=3).value = "familias"
    ws.cell(row=row, column=3).border = BORDER_ALL
    ws.cell(row=row, column=4).value = "Variable de entrada para Tabla Sensibilidad 1 y 3"
    ws.cell(row=row, column=4).border = BORDER_ALL
    aux_fam_row = row
    row += 1

    ws.cell(row=row, column=1).value = "  Input variable canon (Tabla 2)"
    ws.cell(row=row, column=1).font  = FONT_NORM
    ws.cell(row=row, column=1).border = BORDER_ALL
    ws.cell(row=row, column=2).value = f"=B{aux_cuota_row - 16}"   # canon_mensual row ~ 20
    ws.cell(row=row, column=2).fill  = PatternFill("solid", fgColor="D9EAD3")
    ws.cell(row=row, column=2).number_format = FMT_CLP
    ws.cell(row=row, column=2).border = BORDER_ALL
    ws.cell(row=row, column=2).font  = FONT_NORM
    ws.cell(row=row, column=3).value = "CLP"
    ws.cell(row=row, column=3).border = BORDER_ALL
    ws.cell(row=row, column=4).value = "Variable de entrada para Tabla Sensibilidad 2"
    ws.cell(row=row, column=4).border = BORDER_ALL
    aux_canon_row = row
    row += 1

    ws.cell(row=row, column=1).value = "  Input variable precio terreno (Tabla 3)"
    ws.cell(row=row, column=1).font  = FONT_NORM
    ws.cell(row=row, column=1).border = BORDER_ALL
    ws.cell(row=row, column=2).value = f"=B8"   # precio_terreno
    ws.cell(row=row, column=2).fill  = PatternFill("solid", fgColor="D9EAD3")
    ws.cell(row=row, column=2).number_format = FMT_CLP
    ws.cell(row=row, column=2).border = BORDER_ALL
    ws.cell(row=row, column=2).font  = FONT_NORM
    ws.cell(row=row, column=3).value = "CLP"
    ws.cell(row=row, column=3).border = BORDER_ALL
    ws.cell(row=row, column=4).value = "Variable de entrada para Tabla Sensibilidad 3"
    ws.cell(row=row, column=4).border = BORDER_ALL
    aux_precio_row = row
    row += 1

    # Guardar las filas de parámetros para referencia externa
    # Fila real de cada parámetro clave (contando secciones):
    # sec1@4 -> fam=5, cuota=6, desercion=7, sorteo=8
    # sec2@9 -> precio=10, sup=11, viv=12, aprec=13
    # sec3@14 -> cap2m=15, cap2r=16, cap2ret=17, plazo=18, gracia=19
    # sec4@20 -> sub_uf=21, val_uf=22, pct_sub=23, dest=24
    # sec5@25 -> canon=26, crec=27
    # sec6@28 -> costo=29, tasa=30, plazo_c=31
    # sec7@32 -> aux_cuota=33, aux_fam=34, aux_canon=35, aux_precio=36
    return ws, {
        "n_familias":      5,
        "cuota_mensual":   6,
        "tasa_desercion":  7,
        "familias_sorteo": 8,
        "precio_terreno":  10,
        "sup_terreno":     11,
        "viv_x_terreno":   12,
        "apreciacion":     13,
        "cap2_monto":      15,
        "cap2_ratio":      16,
        "cap2_retorno":    17,
        "plazo_anos":      18,
        "gracia_meses":    19,
        "subsidio_uf":     21,
        "valor_uf":        22,
        "pct_subsidio":    23,
        "subsidio_destino":24,
        "canon_mensual":   26,
        "canon_crec":      27,
        "costo_constr":    29,
        "tasa_credito":    30,
        "plazo_credito":   31,
        "aux_cuota":       33,
        "aux_fam":         34,
        "aux_canon":       35,
        "aux_precio":      36,
    }

print("Módulo cargado OK")
