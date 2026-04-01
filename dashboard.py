"""Hoja 4: Dashboard con KPIs y gráficos"""
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.chart import AreaChart, LineChart, BarChart, Reference
from openpyxl.chart.series import SeriesLabel

BLUE_HDR  = PatternFill("solid", fgColor="1F3864")
BLUE_SUB  = PatternFill("solid", fgColor="2E75B6")
BLUE_CARD = PatternFill("solid", fgColor="D6E4F7")
GREEN_BG  = PatternFill("solid", fgColor="E2EFDA")
GREEN_HDR = PatternFill("solid", fgColor="375623")
ORANGE    = PatternFill("solid", fgColor="FCE4D6")
ORANGE_HDR= PatternFill("solid", fgColor="C55A11")
PURPLE_BG = PatternFill("solid", fgColor="E9D7F5")
PURPLE_HDR= PatternFill("solid", fgColor="7030A0")
YELLOW    = PatternFill("solid", fgColor="FFFF99")
GRAY      = PatternFill("solid", fgColor="F2F2F2")

FONT_HDR   = Font(name="Calibri", bold=True, color="FFFFFF", size=10)
FONT_TITLE = Font(name="Calibri", bold=True, color="1F3864", size=11)
FONT_BOLD  = Font(name="Calibri", bold=True, size=10)
FONT_NORM  = Font(name="Calibri", size=10)
FONT_SMALL = Font(name="Calibri", size=9, color="595959", italic=True)
FONT_KPI   = Font(name="Calibri", bold=True, size=14, color="1F3864")
FONT_KPI_S = Font(name="Calibri", size=9, color="595959")

ALIGN_C = Alignment(horizontal="center", vertical="center", wrap_text=True)
ALIGN_R = Alignment(horizontal="right",  vertical="center")
ALIGN_L = Alignment(horizontal="left",   vertical="center")

BORDER_ALL = Border(
    left=Side(style="thin"), right=Side(style="thin"),
    top=Side(style="thin"),  bottom=Side(style="thin")
)
BORDER_MED = Border(
    left=Side(style="medium"), right=Side(style="medium"),
    top=Side(style="medium"),  bottom=Side(style="medium")
)

P = "Parámetros"
F = "Flujo del Fondo"

R = {
    "n_familias":      5,
    "cuota_mensual":   6,
    "tasa_desercion":  7,
    "precio_terreno":  10,
    "viv_x_terreno":   12,
    "apreciacion":     13,
    "cap2_monto":      15,
    "cap2_ratio":      16,
    "cap2_retorno":    17,
    "gracia_meses":    19,
    "subsidio_uf":     21,
    "valor_uf":        22,
    "pct_subsidio":    23,
    "subsidio_destino":24,
    "canon_mensual":   26,
    "canon_crec":      27,
}
def p(key): return f"'{P}'!$B${R[key]}"

# Mes en la hoja Flujo: fila DATA_START + mes - 1 = 4 + mes - 1 = mes + 3
def flujo_row(mes): return mes + 3

F_DATA = 4   # primera fila de datos en Flujo

def kpi_card(ws, row, col_start, title, formula, fmt, subtitle="", fill=BLUE_CARD, title_fill=BLUE_HDR):
    """KPI card: ocupa 2 columnas, 3 filas"""
    c1, c2 = col_start, col_start + 1
    # Header
    ws.merge_cells(start_row=row, start_column=c1, end_row=row, end_column=c2)
    h = ws.cell(row=row, column=c1)
    h.value     = title
    h.font      = Font(name="Calibri", bold=True, color="FFFFFF", size=9)
    h.fill      = title_fill
    h.alignment = ALIGN_C
    h.border    = BORDER_ALL
    ws.row_dimensions[row].height = 18

    # Valor
    ws.merge_cells(start_row=row+1, start_column=c1, end_row=row+1, end_column=c2)
    v = ws.cell(row=row+1, column=c1)
    v.value         = formula
    v.font          = FONT_KPI
    v.fill          = fill
    v.alignment     = ALIGN_C
    v.number_format = fmt
    v.border        = BORDER_ALL
    ws.row_dimensions[row+1].height = 28

    # Subtítulo
    ws.merge_cells(start_row=row+2, start_column=c1, end_row=row+2, end_column=c2)
    s = ws.cell(row=row+2, column=c1)
    s.value     = subtitle
    s.font      = FONT_KPI_S
    s.fill      = fill
    s.alignment = ALIGN_C
    s.border    = BORDER_ALL
    ws.row_dimensions[row+2].height = 16

def crear_dashboard(wb):
    ws = wb.create_sheet("Dashboard")
    ws.sheet_view.showGridLines = False

    # Anchos de columna
    for ci in range(1, 20):
        ws.column_dimensions[get_column_letter(ci)].width = 14
    ws.column_dimensions["A"].width = 5

    # Título principal
    ws.merge_cells("A1:R1")
    t = ws["A1"]
    t.value     = "FONDO MUTUALISTA DE SUELO URBANO — Dashboard de KPIs"
    t.font      = Font(name="Calibri", bold=True, color="FFFFFF", size=15)
    t.alignment = ALIGN_C
    t.fill      = BLUE_HDR
    ws.row_dimensions[1].height = 32

    # Subtítulo fecha
    ws.merge_cells("A2:R2")
    ws["A2"].value     = "Resultados calculados automáticamente desde la hoja 'Parámetros'. Horizonte: 180 meses (15 años)."
    ws["A2"].font      = FONT_SMALL
    ws["A2"].alignment = ALIGN_C
    ws["A2"].fill      = PatternFill("solid", fgColor="D6E4F7")
    ws.row_dimensions[2].height = 18

    # ── SECCIÓN 1: Terrenos y Familias ───────────────────────────────────────
    row_s1 = 4

    ws.merge_cells(f"B{row_s1}:I{row_s1}")
    h = ws[f"B{row_s1}"]
    h.value     = "IMPACTO — Terrenos y Familias"
    h.font      = FONT_HDR
    h.fill      = BLUE_HDR
    h.alignment = ALIGN_C
    h.border    = BORDER_ALL
    ws.row_dimensions[row_s1].height = 20

    # Terrenos a 3, 5, 10, 15 años (meses 36, 60, 120, 180)
    # En Flujo del Fondo, col J = terrenos, fila = mes + 3
    for idx, (anos, mes) in enumerate([(3,36),(5,60),(10,120),(15,180)]):
        c_start = 2 + idx * 2
        frow = flujo_row(mes)
        kpi_card(ws, row_s1+1,
                 c_start,
                 f"Terrenos a {anos} años",
                 f"='{F}'!J{frow}",
                 '0" terrenos"',
                 f"(mes {mes})",
                 BLUE_CARD, BLUE_SUB)

    row_s2 = row_s1 + 4 + 1

    ws.merge_cells(f"B{row_s2}:I{row_s2}")
    h2 = ws[f"B{row_s2}"]
    h2.value     = "FAMILIAS UBICADAS EN CLT"
    h2.font      = FONT_HDR
    h2.fill      = GREEN_HDR
    h2.alignment = ALIGN_C
    h2.border    = BORDER_ALL
    ws.row_dimensions[row_s2].height = 20

    for idx, (anos, mes) in enumerate([(3,36),(5,60),(10,120),(15,180)]):
        c_start = 2 + idx * 2
        frow = flujo_row(mes)
        kpi_card(ws, row_s2+1,
                 c_start,
                 f"Familias ubicadas a {anos} años",
                 f"='{F}'!K{frow}",
                 '#,##0" familias"',
                 f"(mes {mes})",
                 GREEN_BG,
                 PatternFill("solid", fgColor="548235"))

    # ── SECCIÓN 2: Capital y Apalancamiento ──────────────────────────────────
    row_s3 = row_s2 + 5 + 1

    ws.merge_cells(f"B{row_s3}:I{row_s3}")
    h3 = ws[f"B{row_s3}"]
    h3.value     = "CAPITAL MOVILIZADO Y APALANCAMIENTO"
    h3.font      = FONT_HDR
    h3.fill      = ORANGE_HDR
    h3.alignment = ALIGN_C
    h3.border    = BORDER_ALL
    ws.row_dimensions[row_s3].height = 20

    frow180 = flujo_row(180)
    frow36  = flujo_row(36)

    # Capital total movilizado (C1+C2+C3 al mes 180)
    kpi_card(ws, row_s3+1, 2,
             "Capital Total Movilizado",
             f"='{F}'!I{frow180}",
             '"$"#,##0',
             "Suma 3 capas (mes 180)",
             ORANGE, ORANGE_HDR)

    # Apalancamiento (cap total / cap 1)
    kpi_card(ws, row_s3+1, 4,
             "Apalancamiento Total",
             f"=IF('{F}'!E{frow180}=0,0,'{F}'!I{frow180}/'{F}'!E{frow180})",
             '0.0"x"',
             "Cap. total / Capa 1",
             ORANGE, ORANGE_HDR)

    # Meses hasta primer sorteo (primer mes donde terrenos >= 1)
    kpi_card(ws, row_s3+1, 6,
             "Meses hasta Primer Sorteo",
             (f"=LET(t,COINCIDIR(1,'{F}'!J{F_DATA}:'{F}'!J{frow180},1),"
              f"IF(ESERROR(t),\">180\",t))"),
             '0" meses"',
             "1er terreno adquirido",
             ORANGE, ORANGE_HDR)

    # Costo por familia ubicada
    kpi_card(ws, row_s3+1, 8,
             "Costo por Familia Ubicada",
             f"=IF('{F}'!K{frow180}=0,0,'{F}'!I{frow180}/'{F}'!K{frow180})",
             '"$"#,##0',
             "Cap. total / fam. ubicadas",
             ORANGE, ORANGE_HDR)

    # ── SECCIÓN 3: Servicio de Deuda y Cobertura ─────────────────────────────
    row_s4 = row_s3 + 5 + 1

    ws.merge_cells(f"B{row_s4}:I{row_s4}")
    h4 = ws[f"B{row_s4}"]
    h4.value     = "COBERTURA DEUDA CAPA 2 (Canon / Servicio Deuda)"
    h4.font      = FONT_HDR
    h4.fill      = PURPLE_HDR
    h4.alignment = ALIGN_C
    h4.border    = BORDER_ALL
    ws.row_dimensions[row_s4].height = 20

    for idx, (anos, mes) in enumerate([(5,60),(10,120),(15,180)]):
        c_start = 2 + idx * 2
        frow = flujo_row(mes)
        # Cobertura = Canon acumulado en periodo / Servicio deuda acumulado en periodo
        # Aproximamos: canon_mes / serv_mes (cobertura mensual en ese mes)
        kpi_card(ws, row_s4+1, c_start,
                 f"Cobertura a {anos} años",
                 f"=IF('{F}'!O{frow}=0,\">10x\",'{F}'!M{frow}/'{F}'!O{frow})",
                 '0.00"x"',
                 f"(mes {mes})",
                 PURPLE_BG, PURPLE_HDR)

    # TIR Capa 2 (aproximada: retorno exigido si flujo neto acumulado > 0)
    kpi_card(ws, row_s4+1, 8,
             "Retorno exigido Capa 2",
             f"={p('cap2_retorno')}",
             '0.0%',
             "(configurado en Parámetros)",
             PURPLE_BG, PURPLE_HDR)

    # ── SECCIÓN 4: Comparación DS-49 ─────────────────────────────────────────
    row_s5 = row_s4 + 5 + 1

    ws.merge_cells(f"B{row_s5}:I{row_s5}")
    h5 = ws[f"B{row_s5}"]
    h5.value     = "COMPARACIÓN VS. SUBSIDIO DS-49"
    h5.font      = FONT_HDR
    h5.fill      = PatternFill("solid", fgColor="833C00")
    h5.alignment = ALIGN_C
    h5.border    = BORDER_ALL
    ws.row_dimensions[row_s5].height = 20

    # DS-49 promedio = 1.400 UF → en CLP
    kpi_card(ws, row_s5+1, 2,
             "Subsidio DS-49 referencial",
             f"=1400*{p('valor_uf')}",
             '"$"#,##0',
             "1.400 UF × valor UF",
             PatternFill("solid", fgColor="F4CCAC"),
             PatternFill("solid", fgColor="C55A11"))

    # Costo por familia en este modelo
    kpi_card(ws, row_s5+1, 4,
             "Costo por familia (CLT)",
             f"=IF('{F}'!K{frow180}=0,0,'{F}'!I{frow180}/'{F}'!K{frow180})",
             '"$"#,##0',
             "Cap. total / fam. ubicadas",
             PatternFill("solid", fgColor="F4CCAC"),
             PatternFill("solid", fgColor="C55A11"))

    # Ahorro vs DS-49
    kpi_card(ws, row_s5+1, 6,
             "Diferencia vs DS-49",
             (f"=IF('{F}'!K{frow180}=0,0,"
              f"('{F}'!I{frow180}/'{F}'!K{frow180})-(1400*{p('valor_uf')}))"),
             '"$"#,##0;[Rojo]-"$"#,##0',
             "(+ costoso / - costoso que DS-49)",
             PatternFill("solid", fgColor="F4CCAC"),
             PatternFill("solid", fgColor="C55A11"))

    # Flujo neto acumulado a 15 años
    kpi_card(ws, row_s5+1, 8,
             "Flujo Neto Acumulado 15 años",
             f"='{F}'!Q{frow180}",
             '"$"#,##0',
             "(positivo = autosustentable)",
             PatternFill("solid", fgColor="F4CCAC"),
             PatternFill("solid", fgColor="C55A11"))

    # ── GRÁFICOS ─────────────────────────────────────────────────────────────
    # Preparar mini-tabla de datos para gráficos (en filas ocultas al final)
    CHART_DATA_ROW = row_s5 + 8

    ws.merge_cells(f"B{CHART_DATA_ROW}:D{CHART_DATA_ROW}")
    ws[f"B{CHART_DATA_ROW}"].value = "Datos para gráficos (referencias a Flujo del Fondo)"
    ws[f"B{CHART_DATA_ROW}"].font  = FONT_SMALL

    headers_g = ["Año", "Acum.C1(M$)", "Acum.C2(M$)", "Acum.C3(M$)", "Fam.Ubic.", "Fam.Espera", "Terrenos"]
    for ci, h in enumerate(headers_g, 2):
        ws.cell(row=CHART_DATA_ROW+1, column=ci).value         = h
        ws.cell(row=CHART_DATA_ROW+1, column=ci).font          = FONT_BOLD
        ws.cell(row=CHART_DATA_ROW+1, column=ci).alignment     = ALIGN_C
        ws.cell(row=CHART_DATA_ROW+1, column=ci).border        = BORDER_ALL
        ws.cell(row=CHART_DATA_ROW+1, column=ci).fill          = BLUE_HDR
        ws.cell(row=CHART_DATA_ROW+1, column=ci).font          = FONT_HDR

    # Datos anuales (mes 12, 24, ... 180)
    for yi in range(1, 16):
        mes = yi * 12
        frow = flujo_row(mes)
        gr  = CHART_DATA_ROW + 1 + yi

        ws.cell(row=gr, column=2).value          = yi
        ws.cell(row=gr, column=2).number_format  = '0'
        ws.cell(row=gr, column=2).alignment      = ALIGN_C

        ws.cell(row=gr, column=3).value          = f"=REDONDEAR('{F}'!E{frow}/1000000,1)"
        ws.cell(row=gr, column=3).number_format  = '#,##0.0'

        ws.cell(row=gr, column=4).value          = f"=REDONDEAR('{F}'!G{frow}/1000000,1)"
        ws.cell(row=gr, column=4).number_format  = '#,##0.0'

        ws.cell(row=gr, column=5).value          = f"=REDONDEAR('{F}'!H{frow}/1000000,1)"
        ws.cell(row=gr, column=5).number_format  = '#,##0.0'

        ws.cell(row=gr, column=6).value          = f"='{F}'!K{frow}"
        ws.cell(row=gr, column=6).number_format  = '0'

        ws.cell(row=gr, column=7).value          = f"='{F}'!L{frow}"
        ws.cell(row=gr, column=7).number_format  = '0'

        ws.cell(row=gr, column=8).value          = f"='{F}'!J{frow}"
        ws.cell(row=gr, column=8).number_format  = '0'

        for ci in range(2, 9):
            ws.cell(row=gr, column=ci).border    = BORDER_ALL
            ws.cell(row=gr, column=ci).font      = FONT_NORM
            ws.cell(row=gr, column=ci).alignment = ALIGN_C

    # ── Gráfico 1: Área apilada — evolución 3 capas ──────────────────────────
    chart1 = AreaChart()
    chart1.title    = "Evolución del Capital por Capa (millones CLP)"
    chart1.style    = 10
    chart1.grouping = "stacked"
    chart1.x_axis.title = "Año"
    chart1.y_axis.title = "Millones CLP"

    cats = Reference(ws, min_col=2, max_col=2,
                     min_row=CHART_DATA_ROW+2, max_row=CHART_DATA_ROW+16)

    from openpyxl.chart import Series
    for ci, name in [(3,"Capa 1 (Ahorro familiar)"),
                     (4,"Capa 2 (Capital paciente)"),
                     (5,"Capa 3 (Subsidio estatal)")]:
        data_ref = Reference(ws, min_col=ci, max_col=ci,
                             min_row=CHART_DATA_ROW+1, max_row=CHART_DATA_ROW+16)
        s = Series(data_ref, title=name)
        chart1.append(s)

    chart1.set_categories(cats)
    chart1.width  = 18
    chart1.height = 12
    ws.add_chart(chart1, f"B{row_s5 + 12}")

    # ── Gráfico 2: Línea — familias ubicadas vs en espera ────────────────────
    chart2 = LineChart()
    chart2.title    = "Familias Ubicadas vs. En Espera"
    chart2.style    = 10
    chart2.x_axis.title = "Año"
    chart2.y_axis.title = "Familias"

    from openpyxl.chart import Series
    s_ubic = Series(
        Reference(ws, min_col=6, max_col=6,
                  min_row=CHART_DATA_ROW+1, max_row=CHART_DATA_ROW+16),
        title="Familias Ubicadas"
    )
    s_esp = Series(
        Reference(ws, min_col=7, max_col=7,
                  min_row=CHART_DATA_ROW+1, max_row=CHART_DATA_ROW+16),
        title="Familias en Espera"
    )
    chart2.append(s_ubic)
    chart2.append(s_esp)
    chart2.set_categories(cats)
    chart2.width  = 18
    chart2.height = 12
    ws.add_chart(chart2, f"K{row_s5 + 12}")

    # ── Gráfico 3: Barras — terrenos adquiridos por año ───────────────────────
    chart3 = BarChart()
    chart3.title    = "Terrenos Adquiridos Acumulados por Año"
    chart3.style    = 10
    chart3.type     = "col"
    chart3.x_axis.title = "Año"
    chart3.y_axis.title = "Terrenos"

    s_ter = Series(
        Reference(ws, min_col=8, max_col=8,
                  min_row=CHART_DATA_ROW+1, max_row=CHART_DATA_ROW+16),
        title="Terrenos Adquiridos"
    )
    chart3.append(s_ter)
    chart3.set_categories(cats)
    chart3.width  = 18
    chart3.height = 12
    ws.add_chart(chart3, f"B{row_s5 + 28}")

    return ws

print("dashboard.py cargado OK")
