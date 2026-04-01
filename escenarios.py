"""Hoja 5: Escenarios comparativos"""
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

BLUE_HDR  = PatternFill("solid", fgColor="1F3864")
BLUE_SUB  = PatternFill("solid", fgColor="2E75B6")
BLUE_BG   = PatternFill("solid", fgColor="D6E4F7")
GREEN_HDR = PatternFill("solid", fgColor="375623")
GREEN_BG  = PatternFill("solid", fgColor="E2EFDA")
ORANGE_HDR= PatternFill("solid", fgColor="C55A11")
ORANGE_BG = PatternFill("solid", fgColor="FCE4D6")
PURPLE_HDR= PatternFill("solid", fgColor="7030A0")
PURPLE_BG = PatternFill("solid", fgColor="E9D7F5")
GRAY      = PatternFill("solid", fgColor="F2F2F2")
YELLOW    = PatternFill("solid", fgColor="FFFF99")

FONT_HDR  = Font(name="Calibri", bold=True, color="FFFFFF", size=10)
FONT_BOLD = Font(name="Calibri", bold=True, size=10)
FONT_NORM = Font(name="Calibri", size=10)
FONT_SMALL= Font(name="Calibri", size=9, color="595959", italic=True)
FONT_KPI  = Font(name="Calibri", bold=True, size=13, color="1F3864")

ALIGN_C = Alignment(horizontal="center", vertical="center", wrap_text=True)
ALIGN_R = Alignment(horizontal="right",  vertical="center")
ALIGN_L = Alignment(horizontal="left",   vertical="center")

BORDER_ALL = Border(
    left=Side(style="thin"), right=Side(style="thin"),
    top=Side(style="thin"),  bottom=Side(style="thin")
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

def flujo_row(mes): return mes + 3

ESCENARIOS = [
    {
        "nombre":     "Escenario 1\nMÍNIMO VIABLE",
        "fill_hdr":   PatternFill("solid", fgColor="1F5C9A"),
        "fill_bg":    PatternFill("solid", fgColor="D6E4F7"),
        "familias":   200,
        "cuota":      30000,
        "cap2_monto": 200000000,
        "pct_sub":    0.0,
        "desc":       "200 familias, sin subsidio estatal, capital paciente 200M",
    },
    {
        "nombre":     "Escenario 2\nBASE",
        "fill_hdr":   GREEN_HDR,
        "fill_bg":    GREEN_BG,
        "familias":   500,
        "cuota":      30000,
        "cap2_monto": 500000000,
        "pct_sub":    0.60,
        "desc":       "500 familias, subsidio 60%, capital paciente 500M",
    },
    {
        "nombre":     "Escenario 3\nESCALA",
        "fill_hdr":   PURPLE_HDR,
        "fill_bg":    PURPLE_BG,
        "familias":   1000,
        "cuota":      35000,
        "cap2_monto": 1500000000,
        "pct_sub":    0.80,
        "desc":       "1.000 familias, subsidio 80%, capital paciente 1.500M",
    },
]

def crear_escenarios(wb):
    ws = wb.create_sheet("Escenarios")
    ws.sheet_view.showGridLines = False

    ws.column_dimensions["A"].width = 32
    ws.column_dimensions["B"].width = 18
    ws.column_dimensions["C"].width = 18
    ws.column_dimensions["D"].width = 18
    ws.column_dimensions["E"].width = 3   # separador

    # Título
    ws.merge_cells("A1:D1")
    t = ws["A1"]
    t.value     = "ESCENARIOS COMPARATIVOS — Fondo Mutualista de Suelo Urbano"
    t.font      = Font(name="Calibri", bold=True, color="FFFFFF", size=14)
    t.alignment = ALIGN_C
    t.fill      = BLUE_HDR
    ws.row_dimensions[1].height = 30

    ws.merge_cells("A2:D2")
    ws["A2"].value     = ("Los escenarios calculan KPIs directamente con sus parámetros propios. "
                          "NO sobreescriben la hoja Parámetros.")
    ws["A2"].font      = FONT_SMALL
    ws["A2"].alignment = ALIGN_C
    ws["A2"].fill      = PatternFill("solid", fgColor="FFF2CC")
    ws.row_dimensions[2].height = 18

    # Encabezado de columnas
    row = 4
    ws.cell(row=row, column=1).value     = "KPI / Indicador"
    ws.cell(row=row, column=1).font      = FONT_BOLD
    ws.cell(row=row, column=1).fill      = BLUE_HDR
    ws.cell(row=row, column=1).alignment = ALIGN_C
    ws.cell(row=row, column=1).border    = BORDER_ALL

    for ci, esc in enumerate(ESCENARIOS, 2):
        ws.cell(row=row, column=ci).value     = esc["nombre"]
        ws.cell(row=row, column=ci).font      = FONT_HDR
        ws.cell(row=row, column=ci).fill      = esc["fill_hdr"]
        ws.cell(row=row, column=ci).alignment = ALIGN_C
        ws.cell(row=row, column=ci).border    = BORDER_ALL
        ws.row_dimensions[row].height = 36

    # Parámetros de cada escenario
    def esc_params(esc):
        """Devuelve dict de parámetros del escenario."""
        return {
            "n_fam":     esc["familias"],
            "cuota":     esc["cuota"],
            "des":       0.05,             # tasa deserción fija
            "cap2_r":    3,                # ratio capa 2 fijo
            "cap2_m":    esc["cap2_monto"],
            "cap2_ret":  0.05,             # retorno exigido fijo
            "pct_sub":   esc["pct_sub"],
            "sub_uf":    800,
            "val_uf":    38000,
            "sub_dest":  1.0,
            "precio":    120000000,
            "viv_x_t":   25,
            "aprec":     0.05,
            "canon_m":   25000,
            "canon_c":   0.03,
            "gracia":    12,
        }

    def kpi_formula_esc(esc_p, mes):
        """
        Genera fórmulas LET para cada KPI sin depender de la hoja Flujo.
        Usa LET de Excel para calcular todo dentro de la celda.
        """
        n   = esc_p["n_fam"]
        cu  = esc_p["cuota"]
        des = esc_p["des"]
        r   = esc_p["cap2_r"]
        cm  = esc_p["cap2_m"]
        cr  = esc_p["cap2_ret"]
        ps  = esc_p["pct_sub"]
        su  = esc_p["sub_uf"]
        vu  = esc_p["val_uf"]
        sd  = esc_p["sub_dest"]
        pr  = esc_p["precio"]
        vt  = esc_p["viv_x_t"]
        ap  = esc_p["aprec"]
        can = esc_p["canon_m"]
        cc  = esc_p["canon_c"]
        gr  = esc_p["gracia"]

        base = (
            f"LET("
            f"des_mes,(1-{des})^(1/12),"
            f"acum_c1,{cu}*{n}*(1-des_mes^{mes})/(1-des_mes),"
            f"acum_c2,MIN({cm},acum_c1*{r}),"
            f"sub_por_t,{vt}*{ps}*{su}*{vu}*{sd},"
            f"precio_adj,{pr}*(1+{ap})^({mes}/12),"
            f"cap_pre,acum_c1+acum_c2,"
            f"terr_pre,MAX(0,ENTERO(cap_pre/precio_adj)),"
            f"cap_total,acum_c1+acum_c2+terr_pre*sub_por_t,"
            f"terrenos,MAX(0,ENTERO(cap_total/precio_adj)),"
            f"fam_ubic,terrenos*{vt},"
            f"fam_act,REDONDEAR({n}*des_mes^{mes},0),"
            f"fam_esp,MAX(0,fam_act-fam_ubic),"
            f"canon_adj,{can}*(1+{cc})^({mes}/12),"
            f"canon_mes,fam_ubic*canon_adj,"
            f"serv_deuda,IF({mes}<={gr},0,acum_c2*{cr}/12),"
            f"flujo_neto,{cu}*fam_act+canon_mes-serv_deuda"
        )
        # función auxiliar no usada directamente; mk() reemplaza su lógica
        return base

    def mk(esc_p, mes, output_var, fmt):
        """Crea fórmula LET completa para una variable de salida."""
        n   = esc_p["n_fam"]
        cu  = esc_p["cuota"]
        des = esc_p["des"]
        r2  = esc_p["cap2_r"]
        cm  = esc_p["cap2_m"]
        cr  = esc_p["cap2_ret"]
        ps  = esc_p["pct_sub"]
        su  = esc_p["sub_uf"]
        vu  = esc_p["val_uf"]
        sd  = esc_p["sub_dest"]
        pr  = esc_p["precio"]
        vt  = esc_p["viv_x_t"]
        ap  = esc_p["aprec"]
        can = esc_p["canon_m"]
        cc  = esc_p["canon_c"]
        gr  = esc_p["gracia"]

        return (
            f"=LET("
            f"des_mes,(1-{des})^(1/12),"
            f"acum_c1,{cu}*{n}*(1-des_mes^{mes})/(1-des_mes),"
            f"acum_c2,MIN({cm},acum_c1*{r2}),"
            f"sub_por_t,{vt}*{ps}*{su}*{vu}*{sd},"
            f"precio_adj,{pr}*(1+{ap})^({mes}/12),"
            f"cap_pre,acum_c1+acum_c2,"
            f"terr_pre,MAX(0,ENTERO(cap_pre/precio_adj)),"
            f"cap_total,acum_c1+acum_c2+terr_pre*sub_por_t,"
            f"terrenos,MAX(0,ENTERO(cap_total/precio_adj)),"
            f"fam_ubic,terrenos*{vt},"
            f"fam_act,REDONDEAR({n}*des_mes^{mes},0),"
            f"fam_esp,MAX(0,fam_act-fam_ubic),"
            f"canon_adj,{can}*(1+{cc})^({mes}/12),"
            f"canon_mes,fam_ubic*canon_adj,"
            f"serv_deuda,IF({mes}<={gr},0,acum_c2*{cr}/12),"
            f"flujo_neto,{cu}*fam_act+canon_mes-serv_deuda,"
            f"{output_var}"
            f")"
        )

    def mk_meses_sorteo(esc_p):
        """Meses hasta primer terreno."""
        n   = esc_p["n_fam"]
        cu  = esc_p["cuota"]
        des = esc_p["des"]
        r2  = esc_p["cap2_r"]
        cm  = esc_p["cap2_m"]
        pr  = esc_p["precio"]
        ap  = esc_p["aprec"]
        return (
            f"=LET("
            f"des_mes,(1-{des})^(1/12),"
            f"meses,SECUENCIA(180),"
            f"acum_c1,{cu}*{n}*(1-des_mes^meses)/(1-des_mes),"
            f"acum_c2,MIN({cm},acum_c1*{r2}),"
            f"cap_t,acum_c1+acum_c2,"
            f"precio_adj,{pr}*(1+{ap})^((meses-1)/12),"
            f"idx,COINCIDIR(VERDADERO,cap_t>=precio_adj,0),"
            f"IF(ESERROR(idx),\">180\",idx)"
            f")"
        )

    # ── Sección: Parámetros del escenario ──────────────────────────────────
    row += 1
    sec_params = [
        ("── PARÁMETROS ──",        None,  None,    None),
        ("Familias participantes",  "n_fam",  "#,##0",  "familias"),
        ("Cuota mensual (CLP)",     "cuota",  "#,##0",  "CLP/mes"),
        ("Capital Capa 2 (CLP)",    "cap2_m", "#,##0",  "CLP"),
        ("% familias con subsidio", "pct_sub","0.0%",   "%"),
        ("Descripción",             None,     None,     None),
    ]

    for lbl, key, fmt, unit in sec_params:
        is_section = (key is None and fmt is None)
        c_lbl = ws.cell(row=row, column=1)
        c_lbl.value     = lbl
        c_lbl.font      = FONT_BOLD if is_section else FONT_NORM
        c_lbl.fill      = BLUE_SUB if is_section else GRAY
        c_lbl.border    = BORDER_ALL
        if is_section:
            c_lbl.font  = FONT_HDR
        c_lbl.alignment = ALIGN_L

        for ci, esc in enumerate(ESCENARIOS, 2):
            ep = esc_params(esc)
            cell = ws.cell(row=row, column=ci)
            if is_section:
                cell.value     = ""
                cell.fill      = esc["fill_hdr"]
                cell.border    = BORDER_ALL
            elif key == "pct_sub" and lbl == "Descripción":
                cell.value     = esc["desc"]
                cell.font      = FONT_SMALL
                cell.fill      = esc["fill_bg"]
                cell.border    = BORDER_ALL
                cell.alignment = ALIGN_C
                cell.number_format = "@"
            else:
                if key is not None:
                    cell.value = ep.get(key, "")
                elif lbl == "Descripción":
                    cell.value = esc["desc"]
                cell.font          = FONT_NORM
                cell.fill          = esc["fill_bg"]
                cell.border        = BORDER_ALL
                cell.alignment     = ALIGN_R
                if fmt:
                    cell.number_format = fmt

        ws.row_dimensions[row].height = 18
        row += 1

    # Descripción en fila separada
    ws.cell(row=row, column=1).value     = "Descripción"
    ws.cell(row=row, column=1).fill      = GRAY
    ws.cell(row=row, column=1).border    = BORDER_ALL
    ws.cell(row=row, column=1).font      = FONT_NORM
    ws.cell(row=row, column=1).alignment = ALIGN_L
    for ci, esc in enumerate(ESCENARIOS, 2):
        c = ws.cell(row=row, column=ci)
        c.value         = esc["desc"]
        c.font          = FONT_SMALL
        c.fill          = esc["fill_bg"]
        c.border        = BORDER_ALL
        c.alignment     = ALIGN_C
        c.number_format = "@"
    ws.row_dimensions[row].height = 30
    row += 1

    # ── Sección: KPIs por horizonte ─────────────────────────────────────────
    kpis = [
        # (label, output_var, mes, fmt, es_seccion)
        ("── TERRENOS ADQUIRIDOS ──", None, None, None, True),
        ("Terrenos a 3 años (mes 36)",  "terrenos",  36,  "0",      False),
        ("Terrenos a 5 años (mes 60)",  "terrenos",  60,  "0",      False),
        ("Terrenos a 10 años (mes 120)","terrenos", 120,  "0",      False),
        ("Terrenos a 15 años (mes 180)","terrenos", 180,  "0",      False),
        ("── FAMILIAS UBICADAS ──",   None, None, None, True),
        ("Familias ubicadas a 3 años",  "fam_ubic",  36,  "#,##0",  False),
        ("Familias ubicadas a 5 años",  "fam_ubic",  60,  "#,##0",  False),
        ("Familias ubicadas a 10 años", "fam_ubic", 120,  "#,##0",  False),
        ("Familias ubicadas a 15 años", "fam_ubic", 180,  "#,##0",  False),
        ("── CAPITAL ──",             None, None, None, True),
        ("Capital total mes 60 (CLP)",  "cap_total", 60,  "#,##0",  False),
        ("Capital total mes 180 (CLP)", "cap_total", 180, "#,##0",  False),
        ("── PRIMER SORTEO ──",       None, None, None, True),
        ("Meses hasta 1er sorteo",      None,       None, "0",      False),
        ("── FLUJO NETO ──",          None, None, None, True),
        ("Flujo neto mensual mes 60",   "flujo_neto", 60, "#,##0",  False),
        ("Flujo neto mensual mes 120",  "flujo_neto",120, "#,##0",  False),
        ("Flujo neto mensual mes 180",  "flujo_neto",180, "#,##0",  False),
        ("── COBERTURA DEUDA ──",     None, None, None, True),
        ("Cobertura mes 60 (canon/svc)","cobertura", 60,  "0.00",   False),
        ("Cobertura mes 120",           "cobertura",120,  "0.00",   False),
        ("Cobertura mes 180",           "cobertura",180,  "0.00",   False),
    ]

    def mk_cobertura(esc_p, mes):
        n   = esc_p["n_fam"]
        cu  = esc_p["cuota"]
        des = esc_p["des"]
        r2  = esc_p["cap2_r"]
        cm  = esc_p["cap2_m"]
        cr  = esc_p["cap2_ret"]
        ps  = esc_p["pct_sub"]
        su  = esc_p["sub_uf"]
        vu  = esc_p["val_uf"]
        sd  = esc_p["sub_dest"]
        pr  = esc_p["precio"]
        vt  = esc_p["viv_x_t"]
        ap  = esc_p["aprec"]
        can = esc_p["canon_m"]
        cc  = esc_p["canon_c"]
        gr  = esc_p["gracia"]
        return (
            f"=LET("
            f"des_mes,(1-{des})^(1/12),"
            f"acum_c1,{cu}*{n}*(1-des_mes^{mes})/(1-des_mes),"
            f"acum_c2,MIN({cm},acum_c1*{r2}),"
            f"sub_por_t,{vt}*{ps}*{su}*{vu}*{sd},"
            f"precio_adj,{pr}*(1+{ap})^({mes}/12),"
            f"cap_pre,acum_c1+acum_c2,"
            f"terr_pre,MAX(0,ENTERO(cap_pre/precio_adj)),"
            f"cap_total,acum_c1+acum_c2+terr_pre*sub_por_t,"
            f"terrenos,MAX(0,ENTERO(cap_total/precio_adj)),"
            f"fam_ubic,terrenos*{vt},"
            f"canon_adj,{can}*(1+{cc})^({mes}/12),"
            f"canon_mes,fam_ubic*canon_adj,"
            f"serv_deuda,IF({mes}<={gr},0,acum_c2*{cr}/12),"
            f"IF(serv_deuda=0,999,REDONDEAR(canon_mes/serv_deuda,2))"
            f")"
        )

    for lbl, output_var, mes, fmt, is_section in kpis:
        c_lbl = ws.cell(row=row, column=1)
        c_lbl.value     = lbl
        c_lbl.border    = BORDER_ALL
        c_lbl.alignment = ALIGN_L

        if is_section:
            c_lbl.font = FONT_HDR
            c_lbl.fill = BLUE_HDR
            ws.row_dimensions[row].height = 20
            for ci in range(2, 5):
                ws.cell(row=row, column=ci).fill   = BLUE_HDR
                ws.cell(row=row, column=ci).border = BORDER_ALL
        else:
            c_lbl.font = FONT_NORM
            c_lbl.fill = GRAY
            ws.row_dimensions[row].height = 18

            for ci, esc in enumerate(ESCENARIOS, 2):
                ep = esc_params(esc)
                cell = ws.cell(row=row, column=ci)
                cell.fill      = esc["fill_bg"]
                cell.border    = BORDER_ALL
                cell.font      = FONT_BOLD
                cell.alignment = ALIGN_C

                if output_var == "cobertura":
                    cell.value         = mk_cobertura(ep, mes)
                    cell.number_format = fmt + '"x"'
                elif output_var is None and lbl.startswith("Meses"):
                    cell.value         = mk_meses_sorteo(ep)
                    cell.number_format = fmt
                else:
                    cell.value         = mk(ep, mes, output_var, fmt)
                    cell.number_format = fmt

        row += 1

    # Leyenda
    row += 1
    ws.merge_cells(f"A{row}:D{row}")
    ws[f"A{row}"].value     = ("* Los KPIs se calculan con fórmulas LET directamente en cada celda. "
                               "Parámetros fijos: deserción 5%, precio terreno 120M, apreciación 5%, canon 25.000 CLP, crecimiento canon 3%, ratio C2 3x, retorno C2 5%.")
    ws[f"A{row}"].font      = FONT_SMALL
    ws[f"A{row}"].alignment = ALIGN_L
    ws[f"A{row}"].fill      = PatternFill("solid", fgColor="FFF2CC")
    ws.row_dimensions[row].height = 36
    for ci in range(1, 5):
        ws.cell(row=row, column=ci).border = BORDER_ALL

    ws.freeze_panes = "B5"
    return ws

print("escenarios.py cargado OK")
