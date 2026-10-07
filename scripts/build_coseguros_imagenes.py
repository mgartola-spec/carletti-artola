#!/usr/bin/env python3
"""
OSAP — Coseguros de practicas de imagenes / con aparatologia
Construye la base limpia (sin nombres, afiliado por ben_id) y el Excel de analisis.

Uso:  python3 scripts/build_coseguros_imagenes.py <Base_Coseg_Agosto.xlsx> <salida.xlsx>
"""
import sys, re, unicodedata
import pandas as pd

# ---------------------------------------------------------------- clasificador
# Cada modalidad con su patron. El ORDEN DE ESTA LISTA es el orden de prioridad:
# gana la primera que matchea, de lo mas especifico a lo mas general.
# Es deliberado: "RETINOGRAFIA CON TRES PLACAS" matchea PLACA (Radiografia) y RETINOGRAF
# (Oftalmo) — tiene que ganar Oftalmo. Y "TOMOGRAFIA NERVIO OPTICO (H.R.T.)" es oftalmologica,
# no una TAC. El lookbehind de TAC evita el falso positivo de "diluTACion"/"dilaTACion".
MODALIDADES = [
    ('Oftalmo c/aparato',     r'\bOCT\b|COHERENCIA OPTICA|RETINOGRAF|CAMPO VISUAL|TOPOGRAFIA CORNEAL|'
                              r'PAQUIMETRIA|REFRACTOMET|OFTALMOSCOPIA|FONDO DE OJO|ANGIOFLUORESCEIN|'
                              r'TONOMETR|GONIOSCOP|ECOBIOMETR|ECOMETRIA OCULAR|TOMOGRAFIA NERVIO OPTICO|'
                              r'H\.R\.T|NERVIO OPTICO|ESQUIASCOPIA|BIOMICROSCOP'),
    ('Mamografia',            r'MAMOGRAF|SENOGRAF'),
    ('Densitometria',         r'DENSITOMET'),
    ('Medicina Nuclear',      r'CENTELLO|GAMMAGRAF|SPECT|\bPET\b|RADIOISOTOP|CAMARA GAMMA'),
    ('Angio / Hemodinamia',   r'ANGIOGRAF|CATETERISMO|CINECORONARIO|ARTERIOGRAF|FLEBOGRAF'),
    ('Resonancia',            r'RESON'),
    ('Tomografia Computada',  r'(?<![A-Z])T\.?A\.?C(?![A-Z])|\bTOMOGRAF'),
    ('Endoscopia / Video',    r'ENDOSCOP|COLONOSCOP|FIBROBRONCO|RECTOSIGMOIDEO|CISTOSCOP|VIDEO|'
                              r'LARINGOSCOP|COLPOSCOP|ARTROSCOP|HISTEROSCOP'),
    ('Eco / Doppler',         r'ECOGRAF|ECODOPPLER|ECO-DOPPLER|ECO DOPPLER|DOPPLER|ECOCARDIO|\bECO\b'),
    ('Radiografia',           r'RADIOGRAF|TELERRADIOGRAF|PLACA|SERIADA|COLANGIOGRAF|UROGRAF|'
                              r'COLON POR ENEMA'),
    ('Cardio c/aparato',      r'ELECTROCARDIOGRAMA|\bECG\b|ERGOMETR|HOLTER|PRESURIZACION|\bMAPA\b|'
                              r'PRUEBA DE ESFUERZO'),
    ('Neuro / Func. c/aparato', r'ELECTROENCEFALOGRAMA|\bEEG\b|ELECTROMIOGRAF|POTENCIALES EVOCADOS|'
                                r'AUDIOMETR|IMPEDANCIOMETR|LOGOAUDIOMETR|ESPIROMETR|POLISOMNOGRAF'),
]

# Orden en que se muestran en los informes (de mas a menos "imagen pesada").
ORDEN_INFORME = [
    'Resonancia', 'Tomografia Computada', 'Eco / Doppler', 'Radiografia', 'Mamografia',
    'Densitometria', 'Medicina Nuclear', 'Endoscopia / Video', 'Angio / Hemodinamia',
    'Oftalmo c/aparato', 'Cardio c/aparato', 'Neuro / Func. c/aparato',
]

ORDEN_MODALIDAD = ORDEN_INFORME

# "Todo lo que es imagen": estudios que producen una imagen diagnostica.
# Los trazados funcionales usan aparatologia pero no generan imagen: quedan en la base,
# marcados aparte, para poder incluirlos o excluirlos con el filtro.
PRODUCEN_IMAGEN = {
    'Resonancia', 'Tomografia Computada', 'Eco / Doppler', 'Radiografia', 'Mamografia',
    'Densitometria', 'Medicina Nuclear', 'Endoscopia / Video', 'Angio / Hemodinamia',
    'Oftalmo c/aparato',
}
TRAZADOS = {'Cardio c/aparato', 'Neuro / Func. c/aparato'}
IMAGEN, TRAZADO = 'Imagen', 'Trazado funcional'

# Columnas que se descartan y por que (documentado para auditoria)
DESCARTADAS = {
    'privacidad':  ['BEN_Nombre', 'agecta_nombre', 'doc_id', 'Doc_AgeCta'],
    'vacias':      ['nroafi', 'mot_id', 'Ori_CUIT', 'Pre_CUIT', 'Age_CUI',
                    'it_importacion_obs', 'ModiExp_id', 'FiltroCero'],
    'constantes':  ['ord_peri', 'Multiplicador1', 'Multiplicador', 'area', 'conc_id', 'exp_id',
                    'tiva_id', 'it_por', 'it_estado', 'it_grupo', 'coseguro', 'bonificacion',
                    'ea_id', 'to_id', 'aut_estado', 'agrupra_id'],
    'duplicadas':  ['Valor1', 'it_can1', 'cod_auto', 'ord_est', 'Age_Nom'],
    'peor_opcion': ['it_itot', 'tCob_Observ', 'id_tipoCob', 'ben_gr_id'],
    'tecnicas':    ['it_id', 'ord_it_id', 'nom_codd', 'nom_codh', 'gru_id', 'nom_id', 'usuario',
                    'ord_sol', 'mensa_ipam', 'Especialidad', 'diagnostico', 'AZTER', 'ven_id',
                    'os_id', 'ord_ori_dup'],
}

COLS = [
    ('Orden',               'ord_numero'),
    ('Fecha Orden',         'ord_fec'),
    ('Afiliado ID',         'ben_id'),
    ('Grupo Familiar ID',   'agecta_id'),
    ('Edad',                'BEN_EDAD'),
    ('Ciudad Afiliado',     '__ciudad'),
    ('Convenio',            'Convenio'),
    ('Tipo Coseguro',       'TipoCoseguro'),
    ('Plan',                'plan_id'),
    ('Cobertura Especial',  'tCob_Nombre'),
    ('Modalidad',           '__modalidad'),
    ('Codigo',              'it_cod'),
    ('Practica',            'nom_nom'),
    ('Regla (Agrupacion)',  'Agrupacion'),
    ('% Aplicacion',        '% Aplicacion'),
    ('Modo',                'ModoCoseguro'),
    ('Proceso',             'Proceso'),
    ('Coseguro Unitario',   'Valor'),
    ('Cantidad',            'it_can'),
    ('Coseguro Total',      'Valor_Total'),
    ('Valor Prestacion',    'ValorPractica'),
    ('% Real',              '__pct'),          # formula
    ('Impacto Proyectado',  'Impacto'),
    ('Prestador ID',        'ord_ori'),
    ('Prestador Facturador','Ori_Nom'),
    ('Prescriptor',         'Pre_Nom'),
    ('Tipo de Estudio',     '__tipoest'),
]


def clasificar(df):
    """Asigna modalidad: gana la primera de MODALIDADES que matchea (mas especifica primero)."""
    nm = (df['nom_nom'].astype(str).str.upper()
            .str.normalize('NFKD').str.encode('ascii', 'ignore').str.decode('ascii'))
    df['__modalidad'] = '(no imagen)'
    sin_asignar = df['__modalidad'] == '(no imagen)'
    for name, pat in MODALIDADES:
        hit = sin_asignar & nm.str.contains(pat, regex=True, na=False)
        df.loc[hit, '__modalidad'] = name
        sin_asignar = sin_asignar & ~hit
    return df


def preparar(ruta):
    df = pd.read_excel(ruta, sheet_name='Hoja2')
    total_filas, total_coseg = len(df), df['Valor_Total'].sum()
    df = clasificar(df)
    img = df[df['__modalidad'] != '(no imagen)'].copy()

    # Afiliado identificado por ben_id: nunca por nombre ni DNI.
    img['__ciudad'] = 'PENDIENTE PADRON'
    img['TipoCoseguro'] = img['TipoCoseguro'].astype(str).str.strip()
    img['Convenio'] = img['Convenio'].astype(str).str.strip()
    img['Ori_Nom'] = img['Ori_Nom'].astype(str).str.strip()
    img['Pre_Nom'] = img['Pre_Nom'].astype(str).str.strip()
    img['it_cod'] = img['it_cod'].astype(str).str.strip()
    img['nom_nom'] = img['nom_nom'].astype(str).str.replace(r'\s+', ' ', regex=True).str.strip()
    img['ord_fec'] = pd.to_datetime(img['ord_fec'], errors='coerce')
    img['__tipoest'] = img['__modalidad'].map(
        lambda m: IMAGEN if m in PRODUCEN_IMAGEN else TRAZADO)
    img['__pct'] = None
    img['__modcat'] = pd.Categorical(img['__modalidad'], ORDEN_MODALIDAD, ordered=True)
    img = img.sort_values(['__modcat', 'Valor_Total'], ascending=[True, False])
    return df, img, total_filas, total_coseg


# ------------------------------------------------------------------ estilo
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.formatting.rule import CellIsRule

NAVY   = '2E4A6B'
CLARO  = 'DCE6F1'
ALERTA = 'C0504D'
AMBAR  = 'F2C14E'
F      = 'Arial'

H1   = Font(name=F, size=16, bold=True, color='FFFFFF')
H2   = Font(name=F, size=11, bold=True, color='FFFFFF')
TXT  = Font(name=F, size=10)
BOLD = Font(name=F, size=10, bold=True)
KPIV = Font(name=F, size=18, bold=True, color=NAVY)
KPIL = Font(name=F, size=9,  color='555555')
NOTA = Font(name=F, size=9,  italic=True, color='555555')

FNAVY  = PatternFill('solid', fgColor=NAVY)
FCLARO = PatternFill('solid', fgColor=CLARO)
FBLANC = PatternFill('solid', fgColor='FFFFFF')
BOX    = Border(*[Side('thin', color='B7C9DD')] * 4)
CTR    = Alignment(horizontal='center', vertical='center')
WRAP   = Alignment(horizontal='center', vertical='center', wrap_text=True)

MONEY = '"$"#,##0'
PCT   = '0.0"%"'
NUM   = '#,##0'


def titulo(ws, texto, ancho, fila=1):
    ws.merge_cells(start_row=fila, start_column=1, end_row=fila, end_column=ancho)
    c = ws.cell(fila, 1, texto)
    c.font, c.fill, c.alignment = H1, FNAVY, Alignment(horizontal='left', vertical='center', indent=1)
    ws.row_dimensions[fila].height = 30


def cabecera(ws, fila, encabezados, anchos=None):
    for j, h in enumerate(encabezados, start=1):
        c = ws.cell(fila, j, h)
        c.font, c.fill, c.alignment, c.border = H2, FNAVY, WRAP, BOX
    ws.row_dimensions[fila].height = 32
    if anchos:
        for j, a in enumerate(anchos, start=1):
            ws.column_dimensions[get_column_letter(j)].width = a
    ws.freeze_panes = ws.cell(fila + 1, 1)


def kpi(ws, fila, col, etiqueta, valor, fmt=MONEY):
    ws.merge_cells(start_row=fila, start_column=col, end_row=fila, end_column=col + 1)
    c = ws.cell(fila, col, etiqueta)
    c.font, c.alignment, c.fill = KPIL, CTR, FCLARO
    ws.merge_cells(start_row=fila + 1, start_column=col, end_row=fila + 1, end_column=col + 1)
    v = ws.cell(fila + 1, col, valor)
    v.font, v.alignment, v.number_format, v.fill = KPIV, CTR, fmt, FBLANC
    for r in (fila, fila + 1):
        for cc in range(col, col + 2):
            ws.cell(r, cc).border = BOX
    ws.row_dimensions[fila + 1].height = 26


# ------------------------------------------------------------------ hojas
BASE = "'Base Imagenes'"


def filtro_img(ult):
    """Criterio extra para que toda agregacion cuente solo estudios que producen imagen."""
    return f',{BASE}!$AA$2:$AA${ult},"{IMAGEN}"'


def hoja_base(wb, img):
    ws = wb.create_sheet('Base Imagenes')
    encabezados = [h for h, _ in COLS]
    anchos = [11, 12, 11, 13, 7, 16, 20, 9, 7, 24, 20, 10, 52, 17, 11, 7, 20,
              13, 9, 13, 14, 9, 13, 11, 40, 32, 18]
    cabecera(ws, 1, encabezados, anchos)
    for i, (_, src) in enumerate(COLS):
        pass
    r = 2
    for _, row in img.iterrows():
        for j, (_, src) in enumerate(COLS, start=1):
            if src == '__pct':
                c = ws.cell(r, j, f'=IFERROR(IF(U{r}=0,"",T{r}/U{r}*100),"")')
                c.number_format = PCT
            else:
                v = row[src]
                if pd.isna(v):
                    v = None
                c = ws.cell(r, j, v)
                if src in ('Valor', 'Valor_Total', 'ValorPractica', 'Impacto'):
                    c.number_format = MONEY
                elif src == 'ord_fec':
                    c.number_format = 'dd/mm/yyyy'
                elif src in ('it_can', 'ben_id', 'agecta_id', 'ord_numero', 'ord_ori'):
                    c.number_format = NUM
                elif src == 'BEN_EDAD':
                    c.number_format = '0.0'
            c.font = TXT
        r += 1
    ultima = r - 1
    ws.add_table(Table(displayName='TablaBaseImagenes',
                       ref=f'A1:{get_column_letter(len(COLS))}{ultima}',
                       tableStyleInfo=TableStyleInfo(name='TableStyleLight9',
                                                     showRowStripes=True)))
    ws.conditional_formatting.add(f'V2:V{ultima}',
        CellIsRule(operator='greaterThanOrEqual', formula=['80'],
                   fill=PatternFill('solid', fgColor='F2C9C8'),
                   font=Font(name=F, size=10, bold=True, color=ALERTA)))
    ws.conditional_formatting.add(f'V2:V{ultima}',
        CellIsRule(operator='between', formula=['50', '79.999'],
                   fill=PatternFill('solid', fgColor='FCEBC8')))
    return ultima


def hoja_resumen(wb, img, ult, total_filas, total_coseg):
    ws = wb.create_sheet('Resumen', 0)
    titulo(ws, 'OSAP — COSEGUROS DE PRACTICAS DE IMAGEN — Agosto 2026', 10)
    ws['A2'] = ('Afiliado identificado por Afiliado ID (ben_id): sin nombres ni DNI. Ciudad pendiente '
                'de cruce con el padron. Los KPI son de estudios que producen imagen.')
    ws['A2'].font = NOTA
    ws.merge_cells('A2:J2')

    im = img[img['__tipoest'] == IMAGEN]
    tz = img[img['__tipoest'] == TRAZADO]
    CI = f'{BASE}!$AA$2:$AA${ult},"{IMAGEN}"'

    kpi(ws, 4, 1, 'COSEGURO DE IMAGEN',  f'=SUMIFS({BASE}!$T$2:$T${ult},{CI})')
    kpi(ws, 4, 3, 'VALOR PRESTACIONES',  f'=SUMIFS({BASE}!$U$2:$U${ult},{CI})')
    kpi(ws, 4, 5, '% REAL DE COSEGURO',  '=A5/C5*100', PCT)
    kpi(ws, 4, 7, 'ESTUDIOS (LINEAS)',   f'=COUNTIFS({CI})', NUM)
    kpi(ws, 7, 1, 'AFILIADOS ALCANZADOS',   im['ben_id'].nunique(), NUM)
    kpi(ws, 7, 3, 'GRUPOS FAMILIARES',      im['agecta_id'].nunique(), NUM)
    kpi(ws, 7, 5, 'PRACTICAS DISTINTAS',    im['it_cod'].nunique(), NUM)
    kpi(ws, 7, 7, '% DEL COSEGURO TOTAL OSAP',
        round(100 * im['Valor_Total'].sum() / total_coseg, 1), PCT)

    ws['A10'] = (f'Universo del mes: {total_filas:,} lineas y ${total_coseg:,.0f} de coseguro en toda la '
                 f'obra social. Imagen: ${im["Valor_Total"].sum():,.0f}. Trazados funcionales (ECG, Holter, '
                 f'EEG, EMG, espirometria, audiometria): ${tz["Valor_Total"].sum():,.0f} — estan en la Base '
                 f'marcados como "{TRAZADO}" en la columna Tipo de Estudio, para incluirlos o sacarlos '
                 f'con el filtro.').replace(',', '.')
    ws['A10'].font = NOTA
    ws.merge_cells('A10:J11')
    ws['A10'].alignment = Alignment(wrap_text=True, vertical='top')
    ws.row_dimensions[10].height = 15
    ws.row_dimensions[11].height = 15

    fila = 13
    ws.cell(fila, 1, 'COSEGURO POR MODALIDAD').font = Font(name=F, size=12, bold=True, color=NAVY)
    fila += 1
    enc = ['Modalidad', 'Tipo de Estudio', 'Lineas', 'Afiliados', 'Codigos', 'Coseguro',
           'Valor Prestacion', '% Real', '% del Coseguro de Imagen', 'Coseguro prom. x estudio']
    for j, h in enumerate(enc, start=1):
        c = ws.cell(fila, j, h); c.font, c.fill, c.alignment, c.border = H2, FNAVY, WRAP, BOX
    ws.row_dimensions[fila].height = 32
    for j, a in enumerate([26, 16, 9, 10, 9, 16, 18, 10, 16, 18], start=1):
        ws.column_dimensions[get_column_letter(j)].width = a

    def fila_mod(r, m, sub, total_ref):
        ws.cell(r, 1, m).font = BOLD
        ws.cell(r, 2, IMAGEN if m in PRODUCEN_IMAGEN else TRAZADO).font = TXT
        ws.cell(r, 2).alignment = CTR
        ws.cell(r, 3, f'=COUNTIFS({BASE}!$K$2:$K${ult},$A{r})').number_format = NUM  # por modalidad
        ws.cell(r, 4, sub['ben_id'].nunique()).number_format = NUM
        ws.cell(r, 5, sub['it_cod'].nunique()).number_format = NUM
        ws.cell(r, 6, f'=SUMIFS({BASE}!$T$2:$T${ult},{BASE}!$K$2:$K${ult},$A{r})').number_format = MONEY
        ws.cell(r, 7, f'=SUMIFS({BASE}!$U$2:$U${ult},{BASE}!$K$2:$K${ult},$A{r})').number_format = MONEY
        ws.cell(r, 8, f'=IFERROR($F{r}/$G{r}*100,"")').number_format = PCT
        ws.cell(r, 9, f'=IFERROR($F{r}/{total_ref}*100,"")').number_format = PCT
        ws.cell(r, 10, f'=IFERROR($F{r}/$C{r},"")').number_format = MONEY
        for j in range(1, 11):
            ws.cell(r, j).border = BOX
            if j > 2:
                ws.cell(r, j).font = TXT

    def fila_sub(r, etiq, grupo, rango, total_ref, fuerte=True):
        color = 'FFFFFF' if fuerte else NAVY
        relleno = FNAVY if fuerte else FCLARO
        ws.cell(r, 1, etiq)
        for j in range(1, 11):
            c = ws.cell(r, j)
            c.font = Font(name=F, size=10, bold=True, color=color)
            c.fill, c.border = relleno, BOX
        a, b = rango
        ws.cell(r, 3, f'=SUM(C{a}:C{b})').number_format = NUM
        ws.cell(r, 4, grupo['ben_id'].nunique()).number_format = NUM
        ws.cell(r, 5, grupo['it_cod'].nunique()).number_format = NUM
        ws.cell(r, 6, f'=SUM(F{a}:F{b})').number_format = MONEY
        ws.cell(r, 7, f'=SUM(G{a}:G{b})').number_format = MONEY
        ws.cell(r, 8, f'=IFERROR($F{r}/$G{r}*100,"")').number_format = PCT
        if total_ref:
            ws.cell(r, 9, f'=IFERROR($F{r}/{total_ref}*100,"")').number_format = PCT
        ws.cell(r, 10, f'=IFERROR($F{r}/$C{r},"")').number_format = MONEY

    total_img = float(im['Valor_Total'].sum())
    r = fila + 1
    mods_img = [m for m in ORDEN_MODALIDAD if m in PRODUCEN_IMAGEN and m in set(img['__modalidad'])]
    ini_img = r
    for m in mods_img:
        fila_mod(r, m, img[img['__modalidad'] == m], total_img); r += 1
    fila_sub(r, 'TOTAL IMAGEN', im, (ini_img, r - 1), total_img); r += 2

    mods_tz = [m for m in ORDEN_MODALIDAD if m in TRAZADOS and m in set(img['__modalidad'])]
    if mods_tz:
        ws.cell(r, 1, 'TRAZADOS FUNCIONALES (no producen imagen — fuera del alcance)').font = \
            Font(name=F, size=10, bold=True, italic=True, color='777777')
        r += 1
        ini_tz = r
        for m in mods_tz:
            fila_mod(r, m, img[img['__modalidad'] == m], total_img); r += 1
        fila_sub(r, 'TOTAL TRAZADOS', tz, (ini_tz, r - 1), total_img, fuerte=False); r += 1

    r += 1
    ws.cell(r, 1, 'COMO LEER EL % REAL').font = Font(name=F, size=11, bold=True, color=NAVY)
    for k, t in enumerate([
        'El % Real es el coseguro cobrado dividido el valor de la prestacion: lo que de verdad paga el',
        'afiliado de su bolsillo sobre el precio del estudio, no el porcentaje nominal de la regla.',
        'Un % Real muy por encima del % Aplicacion nominal significa que el coseguro se cobra sobre una',
        'tabla de valores propia que quedo desfasada del arancel del convenio.',
        'En rojo: el afiliado paga 80% o mas del valor del estudio. En ambar: entre 50% y 80%.',
    ]):
        ws.cell(r + 1 + k, 1, t).font = NOTA
        ws.merge_cells(start_row=r + 1 + k, start_column=1, end_row=r + 1 + k, end_column=10)
    return r


def hoja_reglas(wb, img, ult):
    ws = wb.create_sheet('Reglas Coseguro')
    titulo(ws, 'REGLAS DE COSEGURO VIGENTES — nominal vs. real (solo practicas con aparatologia)', 10)
    ws['A2'] = ('Cada fila es una regla tal como esta configurada hoy. La columna Desvio muestra cuanto '
                'se aparta el cobro real del porcentaje nominal: ahi se decide que modificar.')
    ws['A2'].font = NOTA
    ws.merge_cells('A2:J2')

    enc = ['Regla (Agrupacion)', '% Aplicacion', 'Modo', 'Lineas', 'Afiliados', 'Coseguro',
           'Valor Prestacion', '% Real', '% Nominal', 'Desvio (pp)']
    cabecera(ws, 4, enc, [22, 13, 8, 10, 11, 16, 18, 10, 11, 12])

    g = (img.groupby(['Agrupacion', '% Aplicacion', 'ModoCoseguro'], dropna=False)
            .agg(afil=('ben_id', 'nunique'), coseg=('Valor_Total', 'sum'))
            .reset_index().sort_values('coseg', ascending=False))
    r = 5
    for _, row in g.iterrows():
        ag, pa, mo = row['Agrupacion'], row['% Aplicacion'], row['ModoCoseguro']
        crit = (f'{BASE}!$N$2:$N${ult},$A{r},{BASE}!$O$2:$O${ult},$B{r},'
                f'{BASE}!$P$2:$P${ult},$C{r}') + filtro_img(ult)
        ws.cell(r, 1, ag).font = BOLD
        ws.cell(r, 2, pa).alignment = CTR
        ws.cell(r, 3, mo).alignment = CTR
        ws.cell(r, 4, f'=COUNTIFS({crit})').number_format = NUM
        ws.cell(r, 5, row['afil']).number_format = NUM
        ws.cell(r, 6, f'=SUMIFS({BASE}!$T$2:$T${ult},{crit})').number_format = MONEY
        ws.cell(r, 7, f'=SUMIFS({BASE}!$U$2:$U${ult},{crit})').number_format = MONEY
        ws.cell(r, 8, f'=IFERROR($F{r}/$G{r}*100,"")').number_format = PCT
        ws.cell(r, 9, f'=IF(ISNUMBER($B{r}),$B{r},"")').number_format = PCT
        ws.cell(r, 10, f'=IFERROR(IF(ISNUMBER($B{r}),$H{r}-$B{r},""),"")').number_format = '+0.0;-0.0'
        for j in range(1, 11):
            ws.cell(r, j).border = BOX
            if j != 1:
                ws.cell(r, j).font = TXT
        r += 1
    ws.conditional_formatting.add(f'J5:J{r - 1}',
        CellIsRule(operator='greaterThanOrEqual', formula=['10'],
                   fill=PatternFill('solid', fgColor='F2C9C8'),
                   font=Font(name=F, size=10, bold=True, color=ALERTA)))

    # Corte por categoria de afiliado: A (activos) vs H (jubilados/adherentes)
    r += 2
    ws.cell(r, 1, 'EL MISMO COSEGURO POR CATEGORIA DE AFILIADO').font = Font(name=F, size=12, bold=True, color=NAVY)
    r += 1
    enc2 = ['Tipo Coseguro', 'Quienes son', 'Lineas', 'Afiliados', 'Coseguro', 'Valor Prestacion',
            '% Real', '% del total', 'Coseguro prom. x afiliado']
    for j, h in enumerate(enc2, start=1):
        c = ws.cell(r, j, h); c.font, c.fill, c.alignment, c.border = H2, FNAVY, WRAP, BOX
    ws.row_dimensions[r].height = 32
    QUIEN = {'A': 'Activos, aportantes, corporativo, PMO', 'H': 'Jubilados y adherentes',
             'E': 'Exentos / casos especiales', 'AC': 'Activos c/cobertura especial',
             'HC': 'Jubilados c/cobertura especial', 'S': 'Otros'}
    gt = (img.groupby('TipoCoseguro').agg(afil=('ben_id', 'nunique'), coseg=('Valor_Total', 'sum'))
             .reset_index().sort_values('coseg', ascending=False))
    r0 = r + 1
    for i, (_, row) in enumerate(gt.iterrows()):
        rr = r0 + i
        t = row['TipoCoseguro']
        ws.cell(rr, 1, t).font = BOLD
        ws.cell(rr, 1).alignment = CTR
        ws.cell(rr, 2, QUIEN.get(t, '—')).font = TXT
        ct = f'{BASE}!$H$2:$H${ult},$A{rr}' + filtro_img(ult)
        ws.cell(rr, 3, f'=COUNTIFS({ct})').number_format = NUM
        ws.cell(rr, 4, row['afil']).number_format = NUM
        ws.cell(rr, 5, f'=SUMIFS({BASE}!$T$2:$T${ult},{ct})').number_format = MONEY
        ws.cell(rr, 6, f'=SUMIFS({BASE}!$U$2:$U${ult},{ct})').number_format = MONEY
        ws.cell(rr, 7, f'=IFERROR($E{rr}/$F{rr}*100,"")').number_format = PCT
        ws.cell(rr, 8, f'=IFERROR($E{rr}/SUM($E${r0}:$E${r0 + len(gt) - 1})*100,"")').number_format = PCT
        ws.cell(rr, 9, f'=IFERROR($E{rr}/$D{rr},"")').number_format = MONEY
        for j in range(1, 10):
            ws.cell(rr, j).border = BOX
            if j > 2:
                ws.cell(rr, j).font = TXT
    rr = r0 + len(gt) + 1
    ws.cell(rr, 1, 'Hoy la tabla de coseguros es la misma para activos y para jubilados: '
                   'el valor unitario de cada practica no distingue categoria.').font = NOTA
    ws.merge_cells(start_row=rr, start_column=1, end_row=rr, end_column=9)


def hoja_practicas(wb, img, ult):
    ws = wb.create_sheet('Practicas')
    titulo(ws, 'PRACTICAS CON APARATOLOGIA — donde esta la distorsion', 10)
    ws['A2'] = ('Ordenado por coseguro. Las filas en rojo son las practicas donde el afiliado paga '
                '80% o mas del valor de la prestacion: ahi el coseguro dejo de ser copago.')
    ws['A2'].font = NOTA
    ws.merge_cells('A2:J2')
    enc = ['Codigo', 'Practica', 'Modalidad', 'Regla', 'Lineas', 'Afiliados',
           'Coseguro Unit.', 'Coseguro', 'Valor Prestacion', '% Real']
    cabecera(ws, 4, enc, [10, 56, 22, 18, 9, 10, 14, 15, 17, 10])
    g = (img.groupby(['it_cod', 'nom_nom', '__modalidad', 'Agrupacion'])
            .agg(afil=('ben_id', 'nunique'), vu=('Valor', 'median'), coseg=('Valor_Total', 'sum'))
            .reset_index().sort_values('coseg', ascending=False))
    r = 5
    for _, row in g.iterrows():
        crit = f'{BASE}!$L$2:$L${ult},$A{r},{BASE}!$M$2:$M${ult},$B{r}' + filtro_img(ult)
        ws.cell(r, 1, row['it_cod']).alignment = CTR
        ws.cell(r, 2, row['nom_nom'])
        ws.cell(r, 3, row['__modalidad'])
        ws.cell(r, 4, row['Agrupacion'])
        ws.cell(r, 5, f'=COUNTIFS({crit})').number_format = NUM
        ws.cell(r, 6, row['afil']).number_format = NUM
        ws.cell(r, 7, row['vu']).number_format = MONEY
        ws.cell(r, 8, f'=SUMIFS({BASE}!$T$2:$T${ult},{crit})').number_format = MONEY
        ws.cell(r, 9, f'=SUMIFS({BASE}!$U$2:$U${ult},{crit})').number_format = MONEY
        ws.cell(r, 10, f'=IFERROR($H{r}/$I{r}*100,"")').number_format = PCT
        for j in range(1, 11):
            ws.cell(r, j).border = BOX
            ws.cell(r, j).font = TXT
        r += 1
    ws.cell(r - 1 + 2, 1, f'{len(g)} combinaciones de codigo y practica.').font = NOTA
    ws.conditional_formatting.add(f'J5:J{r - 1}',
        CellIsRule(operator='greaterThanOrEqual', formula=['80'],
                   font=Font(name=F, size=10, bold=True, color=ALERTA)))
    ws.conditional_formatting.add(f'J5:J{r - 1}',
        CellIsRule(operator='between', formula=['50', '79.999'],
                   fill=PatternFill('solid', fgColor='FCEBC8')))
    ws.auto_filter.ref = f'A4:J{r - 1}'


def hoja_prestadores(wb, img, ult, top=25):
    ws = wb.create_sheet('Prestadores')
    titulo(ws, 'PRESTADORES FACTURADORES — coseguro de imagenes que se cobra en cada uno', 9)
    ws['A2'] = ('Prestador = quien FACTURA a OSAP. Comparar precios solo dentro de la misma '
                'modalidad y, cuando este el padron, dentro de la misma ciudad.')
    ws['A2'].font = NOTA
    ws.merge_cells('A2:I2')
    enc = ['Prestador Facturador', 'ID', 'Lineas', 'Afiliados', 'Coseguro', 'Valor Prestacion',
           '% Real', '% del Coseguro de Imagenes', 'Modalidad principal']
    cabecera(ws, 4, enc, [44, 9, 9, 10, 16, 18, 10, 16, 24])
    g = (img.groupby('Ori_Nom').agg(oid=('ord_ori', 'first'), afil=('ben_id', 'nunique'),
                                    coseg=('Valor_Total', 'sum'))
            .reset_index().sort_values('coseg', ascending=False).head(top))
    principal = img.groupby('Ori_Nom')['__modalidad'].agg(
        lambda s: s.value_counts().idxmax()).to_dict()
    total = img['Valor_Total'].sum()
    r = 5
    for _, row in g.iterrows():
        crit = f'{BASE}!$Y$2:$Y${ult},$A{r}' + filtro_img(ult)
        ws.cell(r, 1, row['Ori_Nom']).font = BOLD
        ws.cell(r, 2, row['oid']).number_format = NUM
        ws.cell(r, 3, f'=COUNTIFS({crit})').number_format = NUM
        ws.cell(r, 4, row['afil']).number_format = NUM
        ws.cell(r, 5, f'=SUMIFS({BASE}!$T$2:$T${ult},{crit})').number_format = MONEY
        ws.cell(r, 6, f'=SUMIFS({BASE}!$U$2:$U${ult},{crit})').number_format = MONEY
        ws.cell(r, 7, f'=IFERROR($E{r}/$F{r}*100,"")').number_format = PCT
        ws.cell(r, 8, f'=IFERROR($E{r}/{total}*100,"")').number_format = PCT
        ws.cell(r, 9, principal.get(row['Ori_Nom'], '—'))
        for j in range(1, 10):
            ws.cell(r, j).border = BOX
            if j > 1:
                ws.cell(r, j).font = TXT
        r += 1
    ws.cell(r + 1, 1, f'Top {top} de {img["Ori_Nom"].nunique()} prestadores facturadores.').font = NOTA


def hoja_topes(wb, img, ult):
    """Simulacion de tope mensual: cuanto resigna OSAP y a cuantos alcanza."""
    ws = wb.create_sheet('Topes')
    titulo(ws, 'SIMULACION DE TOPE MENSUAL DE COSEGURO DE IMAGENES', 8)
    ws['A2'] = ('El tope se calcula sobre el coseguro de practicas con aparatologia del mes. '
                'Los calculos salen de la hoja oculta Maestro Afiliados.')
    ws['A2'].font = NOTA
    ws.merge_cells('A2:H2')

    pa = img.groupby('ben_id')['Valor_Total'].sum().sort_values(ascending=False)
    pg = img.groupby('agecta_id')['Valor_Total'].sum().sort_values(ascending=False)
    mw = wb.create_sheet('Maestro Afiliados')
    mw.sheet_state = 'hidden'
    cabecera(mw, 1, ['Afiliado ID', 'Coseguro Imagenes', 'Grupo Familiar ID', 'Coseguro Grupo'],
             [14, 20, 18, 18])
    for i, (bid, v) in enumerate(pa.items(), start=2):
        mw.cell(i, 1, int(bid)); mw.cell(i, 2, float(v)).number_format = MONEY
    for i, (gid, v) in enumerate(pg.items(), start=2):
        mw.cell(i, 3, int(gid)); mw.cell(i, 4, float(v)).number_format = MONEY
    fa, fg = len(pa) + 1, len(pg) + 1
    RA, RG = f"'Maestro Afiliados'!$B$2:$B${fa}", f"'Maestro Afiliados'!$D$2:$D${fg}"
    total = float(img['Valor_Total'].sum())

    for base_row, (etiq, rng, n) in enumerate([
            ('POR GRUPO FAMILIAR', RG, len(pg)), ('POR AFILIADO', RA, len(pa))]):
        r = 4 + base_row * 13
        ws.cell(r, 1, etiq).font = Font(name=F, size=12, bold=True, color=NAVY)
        r += 1
        enc = ['Tope mensual', 'Alcanzados', '% de los alcanzables', 'Coseguro que resigna OSAP',
               '% del coseguro de imagenes', 'Coseguro que queda', 'Ahorro promedio x alcanzado']
        for j, h in enumerate(enc, start=1):
            c = ws.cell(r, j, h); c.font, c.fill, c.alignment, c.border = H2, FNAVY, WRAP, BOX
        ws.row_dimensions[r].height = 32
        for j, a in enumerate([15, 12, 17, 24, 20, 18, 20], start=1):
            ws.column_dimensions[get_column_letter(j)].width = a
        for k, tope in enumerate([20000, 30000, 40000, 50000, 60000, 80000, 100000, 150000]):
            rr = r + 1 + k
            ws.cell(rr, 1, tope).number_format = MONEY
            ws.cell(rr, 2, f'=COUNTIF({rng},">"&$A{rr})').number_format = NUM
            ws.cell(rr, 3, f'=IFERROR($B{rr}/{n}*100,"")').number_format = PCT
            ws.cell(rr, 4, f'=SUMPRODUCT(({rng}>$A{rr})*({rng}-$A{rr}))').number_format = MONEY
            ws.cell(rr, 5, f'=IFERROR($D{rr}/{total}*100,"")').number_format = PCT
            ws.cell(rr, 6, f'={total}-$D{rr}').number_format = MONEY
            ws.cell(rr, 7, f'=IFERROR($D{rr}/$B{rr},"")').number_format = MONEY
            for j in range(1, 8):
                ws.cell(rr, j).border = BOX
                ws.cell(rr, j).font = BOLD if j == 1 else TXT
    r = 4 + 2 * 13
    for k, t in enumerate([
        'El tope por grupo familiar es la medida mas eficiente: concentra el alivio en las familias',
        'que acumularon varios estudios en el mes, que son las que realmente sufren el coseguro.',
        'Cambiar el tope en la columna A recalcula toda la fila.',
    ]):
        ws.cell(r + k, 1, t).font = NOTA
        ws.merge_cells(start_row=r + k, start_column=1, end_row=r + k, end_column=7)


def hoja_maestro_prestadores(wb, img):
    ws = wb.create_sheet('Maestro Prestadores')
    ws.sheet_state = 'hidden'
    cabecera(ws, 1, ['Prestador ID', 'Prestador Facturador', 'Ciudad (pendiente padron)'],
             [14, 46, 26])
    g = img.groupby('ord_ori')['Ori_Nom'].agg(lambda s: s.value_counts().idxmax())
    for i, (oid, nom) in enumerate(g.items(), start=2):
        ws.cell(i, 1, int(oid)); ws.cell(i, 2, nom); ws.cell(i, 3, 'PENDIENTE PADRON')


def hoja_criterios(wb, total_filas, total_coseg, img):
    ws = wb.create_sheet('Criterios')
    titulo(ws, 'CRITERIOS DEL ANALISIS — que se incluyo, que se saco y por que', 6)
    r = 3
    def bloque(tit, lineas):
        nonlocal r
        ws.cell(r, 1, tit).font = Font(name=F, size=11, bold=True, color=NAVY); r += 1
        for t in lineas:
            ws.cell(r, 1, t).font = TXT
            ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=6); r += 1
        r += 1
    ws.column_dimensions['A'].width = 120
    bloque('UNIVERSO', [
        f'Archivo origen: Base_Coseg_Agosto.xlsx, hoja Hoja2. Periodo unico 202608 (agosto 2026).',
        f'{total_filas:,} lineas en total, ${total_coseg:,.0f} de coseguro de toda la obra social.'.replace(',', '.'),
        f'Se aislaron {len(img):,} lineas de practicas con aparatologia por ${img["Valor_Total"].sum():,.0f}.'.replace(',', '.'),
    ])
    bloque('IDENTIFICACION DEL AFILIADO', [
        'Se usa Afiliado ID (ben_id) y Grupo Familiar ID (agecta_id). No hay nombres ni DNI en la base.',
        'La columna nroafi del archivo original esta 100% vacia: no servia como numero de afiliado.',
    ])
    bloque('CIUDAD', [
        'El archivo original no trae localidad en ninguna de sus 76 columnas.',
        'La columna Ciudad Afiliado queda en PENDIENTE PADRON hasta cruzar ben_id contra el padron',
        '(OSAP_Benef: ben_id, loc_nombre, par_nombre). El analisis cubre todas las ciudades.',
    ])
    bloque('MODALIDADES INCLUIDAS', [' / '.join(ORDEN_MODALIDAD)])
    bloque('COLUMNAS DESCARTADAS DEL ARCHIVO ORIGINAL', [
        f'Por privacidad ({len(DESCARTADAS["privacidad"])}): ' + ', '.join(DESCARTADAS['privacidad']),
        f'Vacias al 100% ({len(DESCARTADAS["vacias"])}): ' + ', '.join(DESCARTADAS['vacias']),
        f'Un unico valor en todo el archivo ({len(DESCARTADAS["constantes"])}): ' + ', '.join(DESCARTADAS['constantes']),
        f'Duplicadas exactas ({len(DESCARTADAS["duplicadas"])}): ' + ', '.join(DESCARTADAS['duplicadas']),
        f'Peor opcion que su reemplazo ({len(DESCARTADAS["peor_opcion"])}): ' + ', '.join(DESCARTADAS['peor_opcion']),
        f'Tecnicas sin aporte analitico: ' + ', '.join(DESCARTADAS['tecnicas'][:12]),
    ])
    bloque('VERIFICACIONES HECHAS SOBRE EL ARCHIVO ORIGINAL', [
        'Coseguro Total = Coseguro Unitario x Cantidad: exacto en las 23.617 lineas.',
        'Valor1 = Valor, it_can1 = it_can, cod_auto = AZTER, ord_est = ord_ori: identicas al 100%.',
        'Age_Nom y Ori_Nom son mapeo 1 a 1 (280 valores cada una): se conserva solo Ori_Nom.',
        'it_itot esta vacia en el 12% de las filas y difiere de ValorPractica en el 14%:',
        'se usa ValorPractica, que esta completa.',
        'Impacto Proyectado = Coseguro Total x 1,05 o x 1,10 segun la fila (proyeccion de ajuste).',
    ])


# ------------------------------------------------------------------- main
def main():
    entrada, salida = sys.argv[1], sys.argv[2]
    df, img, total_filas, total_coseg = preparar(entrada)
    im = img[img['__tipoest'] == IMAGEN].copy()

    wb = Workbook()
    wb.remove(wb.active)
    ult = hoja_base(wb, img)                     # base completa, trazados marcados aparte
    hoja_resumen(wb, img, ult, total_filas, total_coseg)
    hoja_reglas(wb, im, ult)                     # alcance = imagen
    hoja_practicas(wb, im, ult)
    hoja_prestadores(wb, im, ult)
    hoja_topes(wb, im, ult)
    hoja_maestro_prestadores(wb, im)
    hoja_criterios(wb, total_filas, total_coseg, im)
    wb._sheets = sorted(wb._sheets, key=lambda w: [
        'Resumen', 'Reglas Coseguro', 'Practicas', 'Prestadores', 'Topes',
        'Base Imagenes', 'Criterios', 'Maestro Prestadores', 'Maestro Afiliados'
    ].index(w.title))
    for w in wb.worksheets:
        w.sheet_view.showGridLines = False
    wb.save(salida)

    print(f'OK  {salida}')
    print(f'    base completa (imagen + trazados): {len(img):,} lineas')
    print(f'    IMAGEN: {len(im):,} lineas | ${im["Valor_Total"].sum():,.0f} coseguro | '
          f'${im["ValorPractica"].sum():,.0f} valor prestacion | '
          f'{100 * im["Valor_Total"].sum() / im["ValorPractica"].sum():.1f}% real | '
          f'{im["ben_id"].nunique():,} afiliados')
    tz = img[img['__tipoest'] == TRAZADO]
    print(f'    TRAZADOS: {len(tz):,} lineas | ${tz["Valor_Total"].sum():,.0f}')


if __name__ == '__main__':
    main()
