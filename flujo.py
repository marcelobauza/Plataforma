"""Hoja 2: Flujo del Fondo — 180 meses con fórmulas Excel"""
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.views import SheetView, Pane

YELLOW   = PatternFill("solid", fgColor="FFFF99")
BLUE_HDR = PatternFill("solid", fgColor="1F3864")
BLUE_SUB = PatternFill("solid", fgColor="2E75B6")
GRAY_ALT = PatternFill("solid", fgColor="EBF3FF")
GREEN_BG = PatternFill("solid", fgColor="E2EFDA")
ORANGE   = PatternFill("solid", fgColor="FCE4D6")

FONT_HDR   = Font(name="Calibri", bold=True, color="FFFFFF", size=9)
FONT_NORM  = Font(name="Calibri", size=9)
FONT_BOLD  = Font(name="Calibri", bold=True, size=9)

ALIGN_C = Alignment(horizontal="center", vertical="center", wrap_text=True)
ALIGN_R = Alignment(horizontal="right",  vertical="center")
ALIGN_L = Alignment(horizontal="left",   vertical="center")

BORDER_ALL = Border(
    left=Side(style="thin"), right=Side(style="thin"),
    top=Side(style="thin"),  bottom=Side(style="thin")
)

FMT_CLP = '#,##0'
FMT_INT = '0'
FMT_PCT = '0.0%'

def col(c): return get_column_letter(c)

P = "Parámetros"

# Filas de parámetros (de crear_parametros)
R = {
    "n_familias":      5,
    "cuota_mensual":   6,
    "tasa_desercion":  7,
    "familias_sorteo": 8,
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

COLS = [
    # (titulo corto,          titulo largo,                           ancho, fmt)
    ("Mes",          "Mes",                                            5,  FMT_INT),
    ("Año",          "Año",                                            5,  FMT_INT),
    ("Fam. Activas", "Familias Activas",                              10,  FMT_INT),
    ("Ap. Capa 1",   "Aporte Mensual Capa 1 (CLP)",                  14,  FMT_CLP),
    ("Acum. Capa 1", "Acumulado Capa 1 (CLP)",                       15,  FMT_CLP),
    ("Desemb. C2",   "Desembolso Capa 2 Mensual (CLP)",              15,  FMT_CLP),
    ("Acum. C2",     "Capital Capa 2 Acumulado (CLP)",               15,  FMT_CLP),
    ("Acum. C3",     "Subsidio Capa 3 Acumulado (CLP)",              15,  FMT_CLP),
    ("Cap. Total",   "Capital Total Disponible (CLP)",                15,  FMT_CLP),
    ("Terrenos",     "Terrenos Adquiridos Acumulados",                10,  FMT_INT),
    ("Fam. Ubic.",   "Familias Ubicadas Acumuladas",                  10,  FMT_INT),
    ("En Espera",    "Familias en Espera",                            10,  FMT_INT),
    ("Canon Mes.",   "Canon Mensual Recibido (CLP)",                  14,  FMT_CLP),
    ("Acum. Canon",  "Canon Acumulado (CLP)",                         15,  FMT_CLP),
    ("Serv. C2",     "Servicio Deuda Capa 2 Mensual (CLP)",          15,  FMT_CLP),
    ("Flujo Neto",   "Flujo Neto Mensual (CLP)",                     14,  FMT_CLP),
    ("Acum. Neto",   "Flujo Neto Acumulado (CLP)",                    15,  FMT_CLP),
]

def crear_flujo(wb):
    ws = wb.create_sheet("Flujo del Fondo")
    ws.sheet_view.showGridLines = False

    # Título
    ws.merge_cells(f"A1:{col(len(COLS))}1")
    t = ws["A1"]
    t.value     = "FLUJO DEL FONDO — Horizonte 180 Meses (15 Años)"
    t.font      = Font(name="Calibri", bold=True, color="1F3864", size=13)
    t.alignment = ALIGN_C
    t.fill      = PatternFill("solid", fgColor="D6E4F7")
    ws.row_dimensions[1].height = 26

    # Fila 2: subtítulo corto y fila 3: título largo
    for ci, (short, long, width, fmt) in enumerate(COLS, 1):
        c2 = ws.cell(row=2, column=ci)
        c2.value     = short
        c2.fill      = BLUE_HDR
        c2.font      = FONT_HDR
        c2.alignment = ALIGN_C
        c2.border    = BORDER_ALL

        c3 = ws.cell(row=3, column=ci)
        c3.value     = long
        c3.fill      = BLUE_SUB
        c3.font      = FONT_HDR
        c3.alignment = ALIGN_C
        c3.border    = BORDER_ALL

        ws.column_dimensions[col(ci)].width = width

    ws.row_dimensions[2].height = 22
    ws.row_dimensions[3].height = 36

    # ── Fórmulas para cada mes ────────────────────────────────────────────
    DATA_START = 4   # primera fila de datos

    for m in range(1, 181):
        r = DATA_START + m - 1
        is_alt = (m % 2 == 0)
        WHITE_BG = PatternFill("solid", fgColor="FFFFFF")
        bg = GRAY_ALT if is_alt else WHITE_BG

        def cf(ci, formula, fmt=FMT_CLP, bold=False):
            cell = ws.cell(row=r, column=ci)
            cell.value          = formula
            cell.number_format  = fmt
            cell.font           = FONT_BOLD if bold else FONT_NORM
            cell.alignment      = ALIGN_R if ci > 2 else ALIGN_C
            cell.border         = BORDER_ALL
            cell.fill           = GRAY_ALT if is_alt else PatternFill("solid", fgColor="FFFFFF")

        # Col 1: Mes
        cf(1, m, FMT_INT)
        ws.cell(row=r, column=1).alignment = ALIGN_C

        # Col 2: Año
        cf(2, f"=INT(({m}-1)/12)+1", FMT_INT)
        ws.cell(row=r, column=2).alignment = ALIGN_C

        # Col 3: Familias activas (deserción mensual = (1 - tasa_anual)^(1/12))
        if m == 1:
            cf(3, f"={p('n_familias')}", FMT_INT)
        else:
            cf(3, f"=REDONDEAR(C{r-1}*(1-{p('tasa_desercion')})^(1/12),0)", FMT_INT)

        # Col 4: Aporte mensual Capa 1
        cf(4, f"=C{r}*{p('cuota_mensual')}", FMT_CLP)

        # Col 5: Acumulado Capa 1
        if m == 1:
            cf(5, f"=D{r}", FMT_CLP)
        else:
            cf(5, f"=E{r-1}+D{r}", FMT_CLP)

        # Col 6: Desembolso Capa 2 mensual
        # Ratio: por cada 1 CLP acumulado Capa 1, desembolsa ratio CLP de Capa 2
        # Desembolso incremental = aporte_c1 * ratio, limitado al monto comprometido total
        if m == 1:
            cf(6, f"=MIN({p('cap2_monto')}, D{r}*{p('cap2_ratio')})", FMT_CLP)
        else:
            cf(6, f"=MAX(0, MIN({p('cap2_monto')}-G{r-1}, D{r}*{p('cap2_ratio')}))", FMT_CLP)

        # Col 7: Capital Capa 2 acumulado desembolsado
        if m == 1:
            cf(7, f"=F{r}", FMT_CLP)
        else:
            cf(7, f"=G{r-1}+F{r}", FMT_CLP)

        # Col 8: Subsidio Capa 3 acumulado
        # El subsidio se activa cada vez que se puede adquirir un nuevo terreno
        # Para simplificar: subsidio por terreno = viv_x_terreno * pct_subsidio * subsidio_uf * valor_uf
        # Se añade proporcionalmente al capital disponible cuando se sortea (cada vez que terrenos aumenta)
        subsidio_por_terreno = (
            f"{p('viv_x_terreno')}*{p('pct_subsidio')}*{p('subsidio_uf')}"
            f"*{p('valor_uf')}*{p('subsidio_destino')}"
        )
        if m == 1:
            # No hay sorteo en mes 1
            cf(8, f"=0", FMT_CLP)
        else:
            # Cuando terrenos adquiridos aumenta (J), se agrega subsidio por el nuevo terreno
            cf(8, f"=H{r-1}+IF(J{r}>J{r-1},(J{r}-J{r-1})*({subsidio_por_terreno}),0)", FMT_CLP)

        # Col 9: Capital total disponible = Acum C1 + Acum C2 + Acum C3
        cf(9, f"=E{r}+G{r}+H{r}", FMT_CLP, bold=True)

        # Col 10: Terrenos adquiridos acumulados
        # precio terreno se aprecia: precio * (1+apreciacion)^((m-1)/12)
        precio_ajustado = (
            f"{p('precio_terreno')}*(1+{p('apreciacion')})^(({m}-1)/12)"
        )
        cf(10, f"=ENTERO(I{r}/({precio_ajustado}))", FMT_INT)
        ws.cell(row=r, column=10).alignment = ALIGN_C

        # Col 11: Familias ubicadas acumuladas
        cf(11, f"=J{r}*{p('viv_x_terreno')}", FMT_INT)
        ws.cell(row=r, column=11).alignment = ALIGN_C

        # Col 12: Familias en espera
        cf(12, f"=MAX(0,C{r}-K{r})", FMT_INT)
        ws.cell(row=r, column=12).alignment = ALIGN_C

        # Col 13: Canon mensual recibido
        # Canon ajustado por IPC mensual = canon_base*(1+crec)^((m-1)/12)
        canon_ajustado = (
            f"{p('canon_mensual')}*(1+{p('canon_crec')})^(({m}-1)/12)"
        )
        cf(13, f"=K{r}*({canon_ajustado})", FMT_CLP)

        # Col 14: Canon acumulado
        if m == 1:
            cf(14, f"=M{r}", FMT_CLP)
        else:
            cf(14, f"=N{r-1}+M{r}", FMT_CLP)

        # Col 15: Servicio deuda Capa 2 (interés mensual sobre capital desembolsado)
        # Período de gracia: sin interés los primeros gracia_meses
        cf(15, (
            f"=IF({m}<={p('gracia_meses')},0,"
            f"G{r}*{p('cap2_retorno')}/12)"
        ), FMT_CLP)

        # Col 16: Flujo neto mensual = aportes C1 + canon - servicio deuda C2
        cf(16, f"=D{r}+M{r}-O{r}", FMT_CLP, bold=True)
        # Color condicional via fill
        # (no se puede hacer condicional real en Python fácilmente, se deja neutro)

        # Col 17: Flujo neto acumulado
        if m == 1:
            cf(17, f"=P{r}", FMT_CLP)
        else:
            cf(17, f"=Q{r-1}+P{r}", FMT_CLP, bold=True)

    # Congelar paneles (col A + filas 1-3)
    ws.freeze_panes = "B4"

    # Filtros automáticos
    ws.auto_filter.ref = f"A3:{col(len(COLS))}3"

    return ws

print("flujo.py cargado OK")
