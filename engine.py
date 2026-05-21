"""
engine.py — Motor de análise de estoque
Toda a lógica de cálculo isolada da interface.
"""
import pandas as pd
import numpy as np
import math
from datetime import date, timedelta
from io import BytesIO
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter


# ── CONSTANTES ────────────────────────────────────────────────────────────────
CFOP_VENDA    = [5102, 6102]
CFOP_REMESSA  = [5904, 6904, 5949, 6949, 5910, 6910]
CFOP_RETORNO  = [1904, 2904, 1949, 2949, 1152, 2152, 1556, 2551, 2556]
CFOP_ENTRADA  = [1102, 2102, 1202, 2202, 1551, 5602, 5605, 1933]

Z_MAP         = {'A': 1.65, 'B': 1.28, 'C': 1.00}
FREQ_DIAS     = {'A': 30,   'B': 60,   'C': 90}
FREQ_LABEL    = {'A': 'Mensal', 'B': 'Bimestral', 'C': 'Trimestral'}


# ── LEITURA E LIMPEZA ────────────────────────────────────────────────────────
def carregar_inventario(file):
    """Lê o CSV de inventário e retorna TP=11 limpo e TP=21."""
    inv = pd.read_csv(file, sep=None, engine='python', encoding='latin1')

    def limpar_num(df, cols):
        for c in cols:
            df[c] = df[c].astype(str).str.replace('.', '', regex=False).str.replace(',', '.', regex=False)
            df[c] = pd.to_numeric(df[c], errors='coerce')
        return df

    inv_tp11 = inv[inv['TP'] == 11].copy()
    inv_tp11 = inv_tp11[inv_tp11['Produto'].notna()].copy()
    inv_tp11 = limpar_num(inv_tp11, ['Quantidade', 'V.Unitario', 'Total'])
    inv_tp11['Produto'] = inv_tp11['Produto'].astype(str).str.strip()

    inv_tp21 = inv[inv['TP'] == 21].copy()
    inv_tp21 = inv_tp21[
        inv_tp21['Produto'].notna() &
        ~inv_tp21['Produto'].astype(str).str.match(r'^\d{11,}$')
    ].copy()
    inv_tp21 = limpar_num(inv_tp21, ['Quantidade', 'V.Unitario', 'Total'])
    inv_tp21['Produto'] = inv_tp21['Produto'].astype(str).str.strip()

    return inv_tp11, inv_tp21


def carregar_saidas(file):
    """Lê o CSV de movimentações e classifica cada linha."""
    sai = pd.read_csv(file, sep=None, engine='python', encoding='latin1')
    sai['Data_dt'] = pd.to_datetime(sai['Data'], dayfirst=True, errors='coerce')

    for col in ['Quant.', 'Total', 'Valor NF/OP']:
        if col in sai.columns:
            sai[col + '_num'] = sai[col].astype(str).str.replace('.', '', regex=False).str.replace(',', '.', regex=False)
            sai[col + '_num'] = pd.to_numeric(sai[col + '_num'], errors='coerce')

    sai['Produto'] = sai['Produto'].astype(str).str.strip()
    sai['is_lely'] = sai['Cliente/Fornecedor'].astype(str).str.upper().str.contains('LELY', na=False)
    sai['Mes'] = sai['Data_dt'].dt.to_period('M')

    def classifica(row):
        if row['CFOP'] in CFOP_VENDA and not row['is_lely']:  return 'venda_cliente_final'
        if row['CFOP'] in CFOP_VENDA and row['is_lely']:      return 'faturamento_lely'
        if row['CFOP'] in CFOP_REMESSA and row['is_lely']:    return 'remessa_lely'
        if row['CFOP'] in CFOP_REMESSA and not row['is_lely']:return 'remessa_outros'
        if row['CFOP'] in CFOP_RETORNO:                       return 'retorno'
        if row['CFOP'] in CFOP_ENTRADA:                       return 'entrada'
        return 'outro'

    sai['tipo'] = sai.apply(classifica, axis=1)
    return sai


def detectar_meses_completos(sai):
    """Retorna lista de períodos com meses completos (exclui meses parciais)."""
    meses = sorted(sai['Mes'].dropna().unique())
    if len(meses) < 2:
        return [str(m) for m in meses]
    # Excluir primeiro e último se forem parciais
    primeiro = sai[sai['Mes'] == meses[0]]['Data_dt']
    ultimo   = sai[sai['Mes'] == meses[-1]]['Data_dt']
    completos = []
    for i, m in enumerate(meses):
        if i == 0 and primeiro.dt.day.min() > 1:
            continue
        if i == len(meses) - 1 and ultimo.dt.day.max() < 25:
            continue
        completos.append(str(m))
    return completos if completos else [str(m) for m in meses]


# ── CÁLCULOS ─────────────────────────────────────────────────────────────────
def calcular_cmm(sai, meses_cmm):
    """Calcula CMM, desvio padrão e totais por produto."""
    sai_consumo = sai[sai['tipo'] == 'venda_cliente_final'].copy()
    sai_periodo = sai_consumo[sai_consumo['Mes'].astype(str).isin(meses_cmm)]

    consumo_mensal = sai_periodo.groupby(['Produto', 'Mes'])['Quant._num'].sum().reset_index()

    cmm = sai_periodo.groupby('Produto').agg(
        Total_Consumo=('Quant._num', 'sum'),
        Meses_Ativos=('Mes', 'nunique'),
        Total_Valor=('Total_num', 'sum'),
        Descricao_Saida=('Descricao', 'first'),
        Unidade=('Und.', 'first'),
        Categoria=('Categoria.Produto', 'first'),
        Id_Estoque=('Id.Estoque', 'first')
    ).reset_index()

    n_meses = len(meses_cmm)
    cmm['CMM'] = cmm['Total_Consumo'] / cmm['Meses_Ativos']

    # Desvio padrão com zeros para meses sem consumo
    todos = pd.MultiIndex.from_product([cmm['Produto'], meses_cmm], names=['Produto', 'Mes'])
    consumo_full = consumo_mensal.copy()
    consumo_full['Mes'] = consumo_full['Mes'].astype(str)
    consumo_full = consumo_full.set_index(['Produto', 'Mes']).reindex(todos, fill_value=0).reset_index()
    desvio = consumo_full.groupby('Produto')['Quant._num'].std().reset_index()
    desvio.columns = ['Produto', 'Desvio_Pad']

    cmm = cmm.merge(desvio, on='Produto', how='left')
    cmm['Desvio_Pad'] = cmm['Desvio_Pad'].fillna(0)
    return cmm, n_meses


def calcular_parametros(df, lead_time_dias=30, custo_pedido=50, taxa_carr=0.25):
    """Calcula ES, PP, LEC, Est_Max e Status para cada item."""
    df = df.copy()
    df['Lead_Time_dias'] = lead_time_dias
    df['LT_meses'] = lead_time_dias / 30
    df['Fator_Z'] = df['ABC'].map(Z_MAP)

    df['ES'] = (df['Fator_Z'] * df['Desvio_Pad'] * np.sqrt(df['LT_meses'])).apply(
        lambda x: math.ceil(x) if not np.isnan(x) else 0)
    df['PP'] = (df['CMM'] * df['LT_meses'] + df['ES']).apply(lambda x: math.ceil(x))

    df['D_anual'] = df['CMM'] * 12
    df['H'] = df['V.Unitario'].fillna(0) * taxa_carr
    df['LEC'] = df.apply(
        lambda r: math.ceil(math.sqrt(2 * r['D_anual'] * custo_pedido / r['H']))
        if r['H'] > 0 and r['D_anual'] > 0 else 0, axis=1)

    df['Est_Max'] = df['PP'] + df['LEC']
    df['Estoque_Atual'] = df['Quantidade'].fillna(0)
    df['Status'] = df.apply(
        lambda r: 'ACIMA DO MÁXIMO' if r['Estoque_Atual'] > r['Est_Max']
        else ('ABAIXO DO PP' if r['Estoque_Atual'] < r['PP'] else 'OK'), axis=1)
    df['Excesso_Un'] = (df['Estoque_Atual'] - df['Est_Max']).clip(lower=0)
    df['Excesso_R$'] = df['Excesso_Un'] * df['V.Unitario'].fillna(0)
    return df


def calcular_abc(df):
    """Classifica itens em A/B/C por valor de consumo anual."""
    df = df.copy()
    df['Valor_Consumo_Anual'] = df['CMM'] * 12 * df['V.Unitario'].fillna(0)
    df = df.sort_values('Valor_Consumo_Anual', ascending=False).reset_index(drop=True)
    total = df['Valor_Consumo_Anual'].sum()
    df['Pct_Acum'] = df['Valor_Consumo_Anual'].cumsum() / total if total > 0 else 0
    df['ABC'] = df['Pct_Acum'].apply(lambda x: 'A' if x <= 0.8 else ('B' if x <= 0.95 else 'C'))
    return df


def calcular_remessas(sai):
    """Calcula saldo de remessas enviadas vs retornadas/faturadas."""
    rem = sai[sai['tipo'] == 'remessa_lely']
    ret = sai[sai['tipo'] == 'retorno']
    fat = sai[sai['tipo'] == 'faturamento_lely']

    s_env = rem.groupby('Produto').agg(
        Qtd_Enviada=('Quant._num', 'sum'), Valor_Enviado=('Total_num', 'sum'),
        Ultima_Remessa=('Data_dt', 'max'), N_Remessas=('Doc.Id', 'count'),
        Descricao=('Descricao', 'first')).reset_index()
    s_ret = ret.groupby('Produto').agg(
        Qtd_Retornada=('Quant._num', 'sum'), Valor_Retornado=('Total_num', 'sum')).reset_index()
    s_fat = fat.groupby('Produto').agg(
        Qtd_Faturada=('Quant._num', 'sum'), Valor_Faturado=('Total_num', 'sum')).reset_index()

    remessas = s_env.merge(s_ret, on='Produto', how='left').merge(s_fat, on='Produto', how='left')
    for c in ['Qtd_Retornada', 'Qtd_Faturada', 'Valor_Retornado', 'Valor_Faturado']:
        remessas[c] = remessas[c].fillna(0)

    remessas['Qtd_Pendente']   = remessas['Qtd_Enviada'] - remessas['Qtd_Retornada'] - remessas['Qtd_Faturada']
    remessas['Valor_Pendente'] = remessas['Valor_Enviado'] - remessas['Valor_Retornado'] - remessas['Valor_Faturado']
    em_aberto = remessas[remessas['Qtd_Pendente'] > 0].sort_values('Valor_Pendente', ascending=False)
    return remessas, em_aberto


def identificar_giro_direto(df_abc, inv_clean, sai):
    """Separa itens sem saldo próprio (giro direto via LELY)."""
    inv_prods = set(inv_clean['Produto'].astype(str))
    sem_tp11  = ~df_abc['Produto'].astype(str).isin(inv_prods)
    df_giro   = df_abc[sem_tp11].copy()
    df_param  = df_abc[~sem_tp11].copy()

    # Enriquecer giro com histórico de vendas
    giro_prods = set(df_giro['Produto'].astype(str))
    vendas = sai[sai['Produto'].isin(giro_prods) & sai['CFOP'].isin(CFOP_VENDA) & ~sai['is_lely']].groupby('Produto').agg(
        Qtd_Vendida=('Quant._num', 'sum'), Valor_Vendido=('Total_num', 'sum'),
        N_Vendas=('CFOP', 'count'), Ultima_Venda=('Data', 'last')).reset_index()
    entradas = sai[sai['Produto'].isin(giro_prods) & sai['CFOP'].isin(CFOP_RETORNO)].groupby('Produto').agg(
        Qtd_Recebida=('Quant._num', 'sum')).reset_index()

    df_giro = df_giro.merge(vendas, on='Produto', how='left').merge(entradas, on='Produto', how='left')
    for c in ['Qtd_Vendida', 'Qtd_Recebida', 'N_Vendas']:
        df_giro[c] = df_giro[c].fillna(0)

    return df_param, df_giro


# ── PIPELINE COMPLETO ────────────────────────────────────────────────────────
def rodar_analise(inv_file, sai_file, lead_time=30, custo_pedido=50, taxa_carr=0.25):
    """
    Pipeline principal. Recebe os dois arquivos e parâmetros,
    retorna dict com todos os DataFrames e metadados.
    """
    inv_clean, inv_tp21 = carregar_inventario(inv_file)
    sai = carregar_saidas(sai_file)

    meses_cmm = detectar_meses_completos(sai)
    cmm, n_meses = calcular_cmm(sai, meses_cmm)

    # Merge com inventário
    df = cmm.merge(
        inv_clean[['Produto', 'Id.Estoque', 'Descricao', 'Unidade', 'Quantidade',
                   'V.Unitario', 'Total', 'Tipo Produto', 'Classificação', 'End. Físico']],
        on='Produto', how='left')
    df['Descricao']  = df['Descricao'].fillna(df['Descricao_Saida'])
    df['Quantidade'] = df['Quantidade'].fillna(0)

    df = calcular_abc(df)
    df = calcular_parametros(df, lead_time, custo_pedido, taxa_carr)

    df_param, df_giro = identificar_giro_direto(df, inv_clean, sai)

    remessas, em_aberto = calcular_remessas(sai)

    sem_saida = inv_clean[~inv_clean['Produto'].isin(set(cmm['Produto']))].sort_values('Total', ascending=False)

    return {
        'df_param':    df_param,
        'df_giro':     df_giro,
        'sem_saida':   sem_saida,
        'remessas':    remessas,
        'em_aberto':   em_aberto,
        'inv_clean':   inv_clean,
        'inv_tp21':    inv_tp21,
        'sai':         sai,
        'meses_cmm':   meses_cmm,
        'n_meses':     n_meses,
        'params': {
            'lead_time':    lead_time,
            'custo_pedido': custo_pedido,
            'taxa_carr':    taxa_carr,
        }
    }


# ── GERAÇÃO DO EXCEL ──────────────────────────────────────────────────────────
DARK="1A1714"; GREEN="1A7A4A"; GREEN_L="E8F5EE"
ORANGE="B84C00"; ORANGE_L="FDF0E6"; RED="8A1A1A"; RED_L="F8E6E6"
BLUE="1A4F8A"; BLUE_L="E6EEF8"; GREY="F5F5F5"
AMBER="7A4A00"; AMBER_L="FDF5E6"; PURPLE="4A1A7A"; PURPLE_L="F3EEF8"
TEAL="0F6E56"; TEAL_L="E6F5F1"

def _bd():
    t = Side(style='thin', color='DDDDDD')
    return Border(left=t, right=t, top=t, bottom=t)

def _hdr(ws, row, cols, color, txt="FFFFFF"):
    for c, (label, width) in enumerate(cols, 1):
        cell = ws.cell(row=row, column=c, value=label)
        cell.font = Font(bold=True, color=txt, size=9, name="Arial")
        cell.fill = PatternFill("solid", fgColor=color)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = _bd()
        ws.column_dimensions[get_column_letter(c)].width = width
    ws.row_dimensions[row].height = 28

def _drow(ws, r, vals, fills=None, fmts=None, bolds=None, centers=None):
    for c, val in enumerate(vals, 1):
        cell = ws.cell(row=r, column=c, value=val)
        cell.font = Font(name="Arial", size=9, bold=(bolds[c-1] if bolds else False))
        cell.border = _bd()
        ha = "center" if (centers and c in centers) else "left"
        cell.alignment = Alignment(vertical="center", horizontal=ha)
        if fills and fills[c-1]: cell.fill = PatternFill("solid", fgColor=fills[c-1])
        if fmts  and fmts[c-1]:  cell.number_format = fmts[c-1]
    ws.row_dimensions[r].height = 16

def _title(ws, r, c, txt, size=12, color=DARK):
    cell = ws.cell(r, c, txt)
    cell.font = Font(name="Arial", bold=True, size=size, color=color)
    cell.alignment = Alignment(vertical="center")
    ws.row_dimensions[r].height = 22

def _sub(ws, r, c, txt):
    cell = ws.cell(r, c, txt)
    cell.font = Font(name="Arial", size=8, color="888888")
    ws.row_dimensions[r].height = 15

ST_FILL = {"OK": GREEN_L, "ACIMA DO MÁXIMO": ORANGE_L, "ABAIXO DO PP": RED_L}
ABC_BG  = {"A": GREEN_L,  "B": BLUE_L, "C": GREY}
ABC_FG  = {"A": GREEN,    "B": BLUE,   "C": "555555"}


def gerar_excel(resultado):
    """Gera o workbook Excel completo e retorna BytesIO."""
    r = resultado
    df_param  = r['df_param']
    df_giro   = r['df_giro']
    sem_saida = r['sem_saida']
    remessas  = r['remessas']
    em_aberto = r['em_aberto']
    inv_clean = r['inv_clean']
    inv_tp21  = r['inv_tp21']
    meses_cmm = r['meses_cmm']
    n_meses   = r['n_meses']
    today     = date.today()

    wb = Workbook()

    # ── ABA 1: RESUMO ──────────────────────────────────────────────────────
    ws1 = wb.active; ws1.title = "RESUMO"
    ws1.sheet_view.showGridLines = False
    _title(ws1, 1, 1, "ANÁLISE DE ESTOQUE — PARÂMETROS DE REPOSIÇÃO", 13)
    _sub(ws1, 2, 1, f"CMM: {meses_cmm[0]} a {meses_cmm[-1]} ({n_meses} meses)  |  Remessas LELY excluídas  |  LT: {r['params']['lead_time']}d  |  Custo pedido: R${r['params']['custo_pedido']}  |  Taxa: {int(r['params']['taxa_carr']*100)}%")

    kpis = [
        ("Itens no Inventário",     len(inv_clean),               "",   DARK),
        ("Itens c/ Consumo Real",   len(df_param),                "",   BLUE),
        ("Giro Direto",             len(df_giro),                 "",   TEAL),
        ("Itens SEM Movimento",     len(sem_saida),               "",   "555555"),
        ("Valor Total Inventário",  inv_clean['Total'].sum(),     "R$", DARK),
        ("Excesso Estimado",        df_param['Excesso_R$'].sum(), "R$", RED),
    ]
    row = 4; _title(ws1, row, 1, "INDICADORES GERAIS", 10); row += 1
    for i, (label, val, prefix, color) in enumerate(kpis, 1):
        col = ((i-1) % 3) * 2 + 1; rr = row + ((i-1) // 3) * 3
        for rc in [rr, rr+1]:
            for cc in [col, col+1]:
                ws1.cell(rc, cc).border = _bd()
                ws1.cell(rc, cc).fill = PatternFill("solid", fgColor="F7F7F7")
        lc = ws1.cell(rr, col, label)
        lc.font = Font(name="Arial", size=8, color="888888")
        lc.alignment = Alignment(horizontal="center", vertical="center")
        txt = f"R$ {val:,.0f}" if prefix == "R$" else f"{int(val):,}"
        vc = ws1.cell(rr+1, col, txt)
        vc.font = Font(name="Arial", bold=True, size=14, color=color)
        vc.alignment = Alignment(horizontal="center", vertical="center")
        ws1.row_dimensions[rr+1].height = 26
    row += 7

    _title(ws1, row, 1, "CURVA ABC", 10); row += 1
    _hdr(ws1, row, [("Classe",8),("Qtd Itens",12),("Valor Cons. Anual (R$)",24),("% do Total",13),("Fator Z",10),("Freq. Contagem",22)], DARK); row += 1
    tv = df_param['Valor_Consumo_Anual'].sum()
    for cls, freq in [("A","Mensal — 30 dias"),("B","Bimestral — 60 dias"),("C","Trimestral — 90 dias")]:
        sub = df_param[df_param['ABC'] == cls]
        pct = sub['Valor_Consumo_Anual'].sum() / max(tv, 1)
        clr = {"A": GREEN, "B": BLUE, "C": "777777"}[cls]
        _drow(ws1, row, [cls, len(sub), sub['Valor_Consumo_Anual'].sum(), pct, Z_MAP[cls], freq],
              fmts=[None,"#,##0","R$ #,##0.00","0.0%","0.00",None], centers={1,2,4,5})
        ws1.cell(row,1).fill = PatternFill("solid", fgColor=clr)
        ws1.cell(row,1).font = Font(name="Arial", bold=True, size=9, color="FFFFFF")
        row += 1
    row += 1

    _title(ws1, row, 1, "STATUS ESTOQUE vs PARÂMETROS", 10); row += 1
    _hdr(ws1, row, [("Status",22),("Qtd Itens",12),("% dos Itens",18)], DARK); row += 1
    for st, clr in [("OK",GREEN_L),("ACIMA DO MÁXIMO",ORANGE_L),("ABAIXO DO PP",RED_L)]:
        cnt = df_param['Status'].value_counts().get(st, 0)
        _drow(ws1, row, [st, cnt, cnt/max(len(df_param),1)],
              fills=[clr,clr,clr], fmts=[None,"#,##0","0.0%"], centers={2,3}); row += 1
    for col in range(1, 9): ws1.column_dimensions[get_column_letter(col)].width = 20

    # ── ABA 2: PARÂMETROS ─────────────────────────────────────────────────
    ws2 = wb.create_sheet("PARÂMETROS")
    ws2.sheet_view.showGridLines = False; ws2.freeze_panes = "A2"
    cols2 = [("Cód. Produto",16),("Descrição",44),("Und.",7),("Classificação",22),("ABC",6),
             ("End. Físico",12),("Estoque Atual",13),("CMM",10),("Desvio Pad.",11),("Fator Z",8),
             ("LT (dias)",9),("Est. Segurança",13),("Ponto Pedido",13),("LEC",8),("Est. Máximo",12),
             ("Custo Unit. (R$)",15),("Cons. Anual (R$)",16),("Status",18),("Excesso (un)",11),("Excesso (R$)",13)]
    _hdr(ws2, 1, cols2, DARK)
    for i, (_, rd) in enumerate(df_param.iterrows(), 2):
        vu  = rd.get('V.Unitario', 0) if pd.notna(rd.get('V.Unitario', np.nan)) else 0
        und = str(rd.get('Unidade_y') or rd.get('Unidade_x') or rd.get('Unidade',''))
        vals = [str(rd['Produto']), rd['Descricao'], und,
                rd.get('Classificação',''), rd['ABC'],
                str(rd.get('End. Físico','') or ''),
                rd['Estoque_Atual'], round(rd['CMM'],2), round(rd['Desvio_Pad'],2),
                rd['Fator_Z'], int(rd['Lead_Time_dias']), int(rd['ES']), int(rd['PP']),
                int(rd['LEC']), int(rd['Est_Max']), vu, rd['Valor_Consumo_Anual'],
                rd['Status'], int(rd['Excesso_Un']), rd['Excesso_R$']]
        fmts = [None,None,None,None,None,None,"#,##0.00","#,##0.00","#,##0.00","0.00","#,##0",
                "#,##0","#,##0","#,##0","#,##0","R$ #,##0.00","R$ #,##0.00",None,"#,##0","R$ #,##0.00"]
        sf = ST_FILL.get(rd['Status']); fills = [None]*17 + [sf, None, None]
        _drow(ws2, i, vals, fills=fills, fmts=fmts, centers={3,5,7,8,9,10,11,12,13,14,15,19})
        ws2.cell(i,5).fill = PatternFill("solid", fgColor=ABC_BG.get(rd['ABC'],"FFFFFF"))
        ws2.cell(i,5).font = Font(name="Arial",bold=True,size=9,color=ABC_FG.get(rd['ABC'],"000000"))
    ws2.auto_filter.ref = f"A1:{get_column_letter(len(cols2))}1"

    # ── ABA 3: GIRO DIRETO ────────────────────────────────────────────────
    ws3 = wb.create_sheet("🔁 GIRO DIRETO")
    ws3.sheet_view.showGridLines = False; ws3.freeze_panes = "A3"
    _title(ws3, 1, 1, f"Itens sem saldo próprio — entram via LELY e saem direto ao cliente ({len(df_giro)} itens)", 11, TEAL)
    _sub(ws3, 2, 1, "Esses itens NÃO entram na contagem física. Controle feito pelo fluxo de remessas LELY.")
    cols3 = [("Cód. Produto",16),("Descrição",44),("CMM",10),("Qtd Vendida (período)",18),
             ("N° Vendas",10),("Última Venda",13),("Qtd Recebida LELY",17),("Custo Unit. (R$)",15),("Cons. Anual Est. (R$)",20)]
    _hdr(ws3, 2, cols3, TEAL)
    for i, (_, row_data) in enumerate(df_giro.sort_values('CMM', ascending=False).iterrows(), 3):
        vu = row_data.get('V.Unitario', 0) if pd.notna(row_data.get('V.Unitario', np.nan)) else 0
        vals = [str(row_data['Produto']), row_data['Descricao'], round(row_data['CMM'],2),
                row_data.get('Qtd_Vendida',0), int(row_data.get('N_Vendas',0)),
                str(row_data.get('Ultima_Venda','')), row_data.get('Qtd_Recebida',0),
                vu, row_data['CMM']*12*vu]
        fmts = [None,None,"#,##0.00","#,##0.00","#,##0",None,"#,##0.00","R$ #,##0.00","R$ #,##0.00"]
        _drow(ws3, i, vals, fills=[None,TEAL_L]+[None]*7, fmts=fmts, centers={3,4,5,6,7})

    # ── ABA 4: REMESSAS ───────────────────────────────────────────────────
    ws4 = wb.create_sheet("🔄 REMESSAS EM ABERTO")
    ws4.sheet_view.showGridLines = False; ws4.freeze_panes = "A3"
    _title(ws4, 1, 1, f"Mercadorias enviadas à LELY Filial sem retorno/faturamento — Pendente: R$ {em_aberto['Valor_Pendente'].sum():,.2f}", 11, AMBER)
    cols4 = [("Cód. Produto",16),("Descrição",44),("Qtd Enviada",12),("Qtd Retornada",13),
             ("Qtd Faturada",12),("Qtd Pendente",12),("Valor Enviado",14),("Valor Retornado",14),
             ("Valor Faturado",13),("Valor Pendente",14),("Última Remessa",14)]
    _hdr(ws4, 2, cols4, AMBER)
    for i, (_, row_data) in enumerate(remessas.sort_values('Valor_Pendente', ascending=False).iterrows(), 3):
        pend = row_data['Qtd_Pendente']
        bg = AMBER_L if pend > 0 else GREEN_L
        ultima = row_data['Ultima_Remessa'].strftime('%d/%m/%Y') if pd.notna(row_data['Ultima_Remessa']) else ''
        vals = [str(row_data['Produto']), row_data.get('Descricao',''),
                row_data['Qtd_Enviada'], row_data['Qtd_Retornada'], row_data['Qtd_Faturada'], pend,
                row_data['Valor_Enviado'], row_data['Valor_Retornado'], row_data['Valor_Faturado'], row_data['Valor_Pendente'], ultima]
        fmts = [None,None,"#,##0","#,##0","#,##0","#,##0","R$ #,##0.00","R$ #,##0.00","R$ #,##0.00","R$ #,##0.00",None]
        fills = [None, bg if pend>0 else None]+[None]*4+[None]*3+[AMBER_L if pend>0 else None, None]
        _drow(ws4, i, vals, fills=fills, fmts=fmts, centers={3,4,5,6,11})
        if pend > 0:
            ws4.cell(i,6).font = Font(name="Arial",size=9,bold=True,color=AMBER)
            ws4.cell(i,10).font = Font(name="Arial",size=9,bold=True,color=AMBER)
    ws4.auto_filter.ref = f"A2:{get_column_letter(len(cols4))}2"

    # ── ABA 5: EM PODER TERCEIROS ─────────────────────────────────────────
    ws5 = wb.create_sheet("📦 EM PODER TERCEIROS")
    ws5.sheet_view.showGridLines = False; ws5.freeze_panes = "A3"
    _title(ws5, 1, 1, f"Itens de sua propriedade na LELY Filial (TP=21) — {len(inv_tp21)} itens | Valor: R$ {inv_tp21['Total'].sum():,.2f}", 11, BLUE)
    cols5 = [("Cód. Produto",16),("Descrição",44),("Tipo Produto",16),("Classificação",22),
             ("Und.",7),("Qtd em Terceiro",14),("Custo Unit. (R$)",15),("Valor Total (R$)",15)]
    _hdr(ws5, 2, cols5, BLUE)
    for i, (_, row_data) in enumerate(inv_tp21.sort_values('Total', ascending=False).iterrows(), 3):
        bg = BLUE_L if row_data['Total'] > 1000 else None
        vals = [str(row_data['Produto']), row_data['Descricao'], row_data.get('Tipo Produto',''),
                row_data.get('Classificação',''), str(row_data.get('Unidade','')),
                row_data['Quantidade'], row_data['V.Unitario'], row_data['Total']]
        fmts = [None,None,None,None,None,"#,##0.00","R$ #,##0.00","R$ #,##0.00"]
        _drow(ws5, i, vals, fills=[None,bg]+[None]*6, fmts=fmts, centers={3,5,6})

    # ── ABA 6: EXCESSO ────────────────────────────────────────────────────
    ws6 = wb.create_sheet("⚠️ EXCESSO DE ESTOQUE")
    ws6.sheet_view.showGridLines = False; ws6.freeze_panes = "A3"
    excesso = df_param[df_param['Excesso_R$'] > 0].sort_values('Excesso_R$', ascending=False)
    _title(ws6, 1, 1, f"Estoque ACIMA do máximo — Capital em excesso: R$ {excesso['Excesso_R$'].sum():,.2f}", 11, RED)
    cols6 = [("Cód. Produto",16),("Descrição",44),("ABC",6),("Custo Unit. (R$)",15),
             ("Estoque Atual",13),("Est. Máximo",12),("Excesso (un)",11),
             ("Excesso (R$)",14),("CMM",10),("Cobertura Atual (meses)",22)]
    _hdr(ws6, 2, cols6, RED)
    for i, (_, row_data) in enumerate(excesso.iterrows(), 3):
        vu = row_data.get('V.Unitario',0) if pd.notna(row_data.get('V.Unitario',np.nan)) else 0
        cob = round(row_data['Estoque_Atual']/row_data['CMM'],1) if row_data['CMM']>0 else 0
        vals = [str(row_data['Produto']), row_data['Descricao'], row_data['ABC'], vu,
                row_data['Estoque_Atual'], int(row_data['Est_Max']), int(row_data['Excesso_Un']),
                row_data['Excesso_R$'], round(row_data['CMM'],2), cob]
        fmts = [None,None,None,"R$ #,##0.00","#,##0.00","#,##0","#,##0","R$ #,##0.00","#,##0.00","#,##0.0"]
        bg = ORANGE_L if row_data['Excesso_R$'] > 5000 else "FFFBF5"
        _drow(ws6, i, vals, fills=[None,bg]+[None]*8, fmts=fmts, centers={3,5,6,7,9,10})

    # ── ABA 7: RUPTURA ────────────────────────────────────────────────────
    ws7 = wb.create_sheet("🔴 RISCO DE RUPTURA")
    ws7.sheet_view.showGridLines = False; ws7.freeze_panes = "A3"
    ruptura = df_param[df_param['Status']=='ABAIXO DO PP'].sort_values('Valor_Consumo_Anual', ascending=False)
    _title(ws7, 1, 1, f"{len(ruptura)} itens com estoque ABAIXO do ponto de pedido", 11, RED)
    cols7 = [("Cód. Produto",16),("Descrição",44),("ABC",6),("Custo Unit. (R$)",15),
             ("Estoque Atual",13),("Ponto de Pedido",14),("Déficit (un)",11),
             ("CMM",10),("Cons. Anual (R$)",16)]
    _hdr(ws7, 2, cols7, RED)
    for i, (_, row_data) in enumerate(ruptura.iterrows(), 3):
        vu = row_data.get('V.Unitario',0) if pd.notna(row_data.get('V.Unitario',np.nan)) else 0
        deficit = int(row_data['PP']) - int(row_data['Estoque_Atual'])
        vals = [str(row_data['Produto']), row_data['Descricao'], row_data['ABC'], vu,
                row_data['Estoque_Atual'], int(row_data['PP']), deficit,
                round(row_data['CMM'],2), row_data['Valor_Consumo_Anual']]
        fmts = [None,None,None,"R$ #,##0.00","#,##0.00","#,##0","#,##0","#,##0.00","R$ #,##0.00"]
        bg = RED_L if row_data['ABC']=='A' else "FFF5F5"
        _drow(ws7, i, vals, fills=[None,bg]+[None]*7, fmts=fmts, centers={3,5,6,7,8})

    # ── ABA 8: SEM MOVIMENTO ──────────────────────────────────────────────
    ws8 = wb.create_sheet("📦 SEM MOVIMENTO")
    ws8.sheet_view.showGridLines = False; ws8.freeze_panes = "A3"
    _title(ws8, 1, 1, f"Itens SEM consumo real — {len(sem_saida)} itens | Valor: R$ {sem_saida['Total'].sum():,.2f}", 11, ORANGE)
    cols8 = [("Cód. Produto",16),("Descrição",44),("Tipo Produto",16),("Classificação",22),
             ("Und.",7),("Estoque Atual",13),("Custo Unit. (R$)",15),("Valor Total (R$)",15)]
    _hdr(ws8, 2, cols8, ORANGE)
    for i, (_, row_data) in enumerate(sem_saida.iterrows(), 3):
        bg = ORANGE_L if row_data['Total'] > 1000 else "FFFBF5"
        vals = [str(row_data['Produto']), row_data['Descricao'], row_data.get('Tipo Produto',''),
                row_data.get('Classificação',''), str(row_data.get('Unidade','')),
                row_data['Quantidade'], row_data['V.Unitario'], row_data['Total']]
        fmts = [None,None,None,None,None,"#,##0.00","R$ #,##0.00","R$ #,##0.00"]
        _drow(ws8, i, vals, fills=[None,bg]+[None]*6, fmts=fmts, centers={3,5,6})

    # ── ABA 9: PLANO DE CONTAGEM ──────────────────────────────────────────
    ws9 = wb.create_sheet("📋 PLANO DE CONTAGEM")
    ws9.sheet_view.showGridLines = False; ws9.freeze_panes = "A3"
    _title(ws9, 1, 1, "PLANO DE CONTAGEM CÍCLICA — Apenas itens com saldo físico no almoxarifado", 12)
    _sub(ws9, 2, 1, "Preencha apenas col. J (Qtd. Contada ✏️). Divergência, Acurácia e Status calculam automaticamente. Oculte col. I antes de imprimir.")
    cols9 = [("Prioridade",12),("Cód. Produto",16),("Descrição",44),("Und.",7),
             ("End. Físico",13),("ABC",6),("Freq.",13),("Próx. Contagem",14),
             ("Qtd. Sistema",12),("Qtd. Contada ✏️",14),
             ("Divergência",12),("Acurácia %",11),("Status Contagem",16),("Observação",22)]
    _hdr(ws9, 2, cols9, DARK)
    for col_i, clr in [(9,PURPLE),(10,GREEN),(11,BLUE),(12,BLUE),(13,BLUE)]:
        ws9.cell(2, col_i).fill = PatternFill("solid", fgColor=clr)

    df_p2 = df_param.copy()
    df_p2['prio_sort'] = df_p2['Status'].map({'ABAIXO DO PP':0,'ACIMA DO MÁXIMO':1,'OK':2})
    df_sorted = df_p2.sort_values(['prio_sort','ABC','Valor_Consumo_Anual'], ascending=[True,True,False])

    rn = 3
    for _, row_data in df_sorted.iterrows():
        cls  = row_data['ABC']
        prio = "🔴 URGENTE" if row_data['Status']=='ABAIXO DO PP' else ("⚠️ EXCESSO" if row_data['Status']=='ACIMA DO MÁXIMO' else f"Classe {cls}")
        prox = (today + timedelta(days=FREQ_DIAS[cls])).strftime('%d/%m/%Y')
        und  = str(row_data.get('Unidade_y') or row_data.get('Unidade_x') or '')
        end  = str(row_data.get('End. Físico','') or '')
        bg   = {"ABAIXO DO PP":RED_L,"ACIMA DO MÁXIMO":ORANGE_L,"OK":None}.get(row_data['Status'])
        vals = [prio, str(row_data['Produto']), row_data['Descricao'], und, end, cls,
                FREQ_LABEL[cls], prox, row_data['Estoque_Atual'],"","","","",""]
        fmts = [None,None,None,None,None,None,None,"DD/MM/YYYY","#,##0.00",None,None,"0.0%",None,None]
        fills = [bg]*8 + [PURPLE_L,GREEN_L,BLUE_L,BLUE_L,BLUE_L,None]
        _drow(ws9, rn, vals, fills=fills, fmts=fmts, centers={1,4,6,7,8,9,10,11,12,13})
        for col_i, formula, fmt in [
            (11, f'=IF(J{rn}="","",J{rn}-I{rn})',                                                                  '#,##0.00;[Red]-#,##0.00'),
            (12, f'=IF(J{rn}="","",IF(I{rn}=0,IF(J{rn}=0,1,0),1-ABS(J{rn}-I{rn})/I{rn}))',                       '0.0%'),
            (13, f'=IF(J{rn}="","Não contado",IF(ABS(J{rn}-I{rn})<=0.01,"✅ OK",IF(ABS(J{rn}-I{rn})/MAX(I{rn},0.01)<=0.05,"⚠️ Dif. Pequena","❌ Divergência")))', ''),
        ]:
            c = ws9.cell(rn, col_i); c.value=formula; c.number_format=fmt
            c.font=Font(name="Arial",size=9); c.border=_bd()
            c.alignment=Alignment(horizontal="center",vertical="center")
            c.fill=PatternFill("solid",fgColor=BLUE_L)
        rn += 1

    for _, row_data in sem_saida.iterrows():
        vals = ["Classe D", str(row_data['Produto']), row_data['Descricao'],
                str(row_data.get('Unidade','')), str(row_data.get('End. Físico','') or ''),
                "D","Trimestral",(today+timedelta(days=90)).strftime('%d/%m/%Y'),
                row_data['Quantidade'],"","","","",""]
        fmts = [None,None,None,None,None,None,None,"DD/MM/YYYY","#,##0.00",None,None,"0.0%",None,None]
        fills = ["F5F5F5"]*8 + [PURPLE_L,GREEN_L,BLUE_L,BLUE_L,BLUE_L,None]
        _drow(ws9, rn, vals, fills=fills, fmts=fmts, centers={1,4,6,7,8,9,10,11,12,13})
        for col_i, formula, fmt in [
            (11, f'=IF(J{rn}="","",J{rn}-I{rn})',                                                                   '#,##0.00;[Red]-#,##0.00'),
            (12, f'=IF(J{rn}="","",IF(I{rn}=0,IF(J{rn}=0,1,0),1-ABS(J{rn}-I{rn})/I{rn}))',                        '0.0%'),
            (13, f'=IF(J{rn}="","Não contado",IF(ABS(J{rn}-I{rn})<=0.01,"✅ OK",IF(ABS(J{rn}-I{rn})/MAX(I{rn},0.01)<=0.05,"⚠️ Dif. Pequena","❌ Divergência")))', ''),
        ]:
            c = ws9.cell(rn, col_i); c.value=formula; c.number_format=fmt
            c.font=Font(name="Arial",size=9); c.border=_bd()
            c.alignment=Alignment(horizontal="center",vertical="center")
            c.fill=PatternFill("solid",fgColor=BLUE_L)
        rn += 1

    ws9.auto_filter.ref = f"A2:{get_column_letter(len(cols9))}2"
    LAST_ROW = rn - 1

    # ── ABA 10: PAINEL ACURÁCIA ───────────────────────────────────────────
    ws10 = wb.create_sheet("📊 PAINEL ACURÁCIA")
    ws10.sheet_view.showGridLines = False
    _title(ws10, 1, 1, "PAINEL DE ACURÁCIA DO INVENTÁRIO", 13)
    _sub(ws10, 2, 1, "Atualiza automaticamente conforme o PLANO DE CONTAGEM é preenchido")
    r2 = 4; _title(ws10, r2, 1, "INDICADORES", 10); r2 += 1
    _hdr(ws10, r2, [("Indicador",42),("Resultado",22)], DARK); r2 += 1

    plan = "'📋 PLANO DE CONTAGEM'"
    indicadores = [
        ("Total de itens no plano",           f"=COUNTA({plan}!B3:B{LAST_ROW})",  "#,##0", GREY),
        ("Itens já contados",                 f"=COUNTIF({plan}!J3:J{LAST_ROW},\"<>\"&\"\")", "#,##0", GREY),
        ("% do plano contado",                f"=COUNTIF({plan}!J3:J{LAST_ROW},\"<>\"&\"\")/COUNTA({plan}!B3:B{LAST_ROW})", "0.0%", GREY),
        ("✅ Itens OK",                        f"=COUNTIF({plan}!M3:M{LAST_ROW},\"✅ OK\")", "#,##0", GREEN_L),
        ("⚠️ Diferença Pequena (≤5%)",        f"=COUNTIF({plan}!M3:M{LAST_ROW},\"⚠️ Dif. Pequena\")", "#,##0", AMBER_L),
        ("❌ Divergência (>5%)",               f"=COUNTIF({plan}!M3:M{LAST_ROW},\"❌ Divergência\")", "#,##0", RED_L),
        ("─── Acurácia por Classe ───",       "", "", DARK),
        ("Acurácia Geral",                    f"=AVERAGEIF({plan}!J3:J{LAST_ROW},\"<>\"&\"\",{plan}!L3:L{LAST_ROW})", "0.0%", BLUE_L),
        ("Acurácia Classe A",                 f"=IFERROR(AVERAGEIFS({plan}!L3:L{LAST_ROW},{plan}!J3:J{LAST_ROW},\"<>\"&\"\",{plan}!F3:F{LAST_ROW},\"A\"),\"Sem dados\")", "0.0%", GREEN_L),
        ("Acurácia Classe B",                 f"=IFERROR(AVERAGEIFS({plan}!L3:L{LAST_ROW},{plan}!J3:J{LAST_ROW},\"<>\"&\"\",{plan}!F3:F{LAST_ROW},\"B\"),\"Sem dados\")", "0.0%", BLUE_L),
        ("Acurácia Classe C",                 f"=IFERROR(AVERAGEIFS({plan}!L3:L{LAST_ROW},{plan}!J3:J{LAST_ROW},\"<>\"&\"\",{plan}!F3:F{LAST_ROW},\"C\"),\"Sem dados\")", "0.0%", GREY),
        ("Acurácia Classe D (sem movimento)", f"=IFERROR(AVERAGEIFS({plan}!L3:L{LAST_ROW},{plan}!J3:J{LAST_ROW},\"<>\"&\"\",{plan}!F3:F{LAST_ROW},\"D\"),\"Sem dados\")", "0.0%", "F5F5F5"),
        ("─── Metas ───",                     "", "", DARK),
        ("Meta Classe A",  0.98, "0.0%", GREEN_L),
        ("Meta Classe B/C",0.95, "0.0%", BLUE_L),
        ("Meta Geral",     0.95, "0.0%", GREY),
    ]
    for label, formula, fmt, bg in indicadores:
        is_sep = label.startswith("───")
        cl = ws10.cell(r2, 1, label)
        cl.border = _bd(); cl.alignment = Alignment(vertical="center")
        cl.fill = PatternFill("solid", fgColor=DARK if is_sep else bg)
        cl.font = Font(name="Arial",size=9,bold=is_sep,color="FFFFFF" if is_sep else DARK)
        cv = ws10.cell(r2, 2, formula if formula else "")
        cv.border = _bd(); cv.alignment = Alignment(horizontal="center",vertical="center")
        cv.fill = PatternFill("solid", fgColor=DARK if is_sep else bg)
        cv.font = Font(name="Arial",bold=not is_sep,size=11 if not is_sep else 9,color="FFFFFF" if is_sep else DARK)
        if fmt: cv.number_format = fmt
        ws10.row_dimensions[r2].height = 24; r2 += 1
    ws10.column_dimensions['A'].width = 44; ws10.column_dimensions['B'].width = 22

    # Salvar em memória
    output = BytesIO()
    wb.save(output)
    output.seek(0)
    return output
