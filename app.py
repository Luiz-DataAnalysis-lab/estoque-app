"""
app.py — Interface Streamlit para análise de estoque
Execute com: streamlit run app.py
"""
import streamlit as st
import pandas as pd
import json
import os
from datetime import date, datetime
from pathlib import Path
from engine import rodar_analise, gerar_excel

# ── CONFIGURAÇÃO DA PÁGINA ────────────────────────────────────────────────────
st.set_page_config(
    page_title="Gestão de Estoque",
    page_icon="📦",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ── ESTILOS ───────────────────────────────────────────────────────────────────
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=DM+Serif+Display:ital@0;1&family=JetBrains+Mono:wght@400;500&family=DM+Sans:wght@300;400;500;600&display=swap');

    html, body, [class*="css"] {
        font-family: 'DM Sans', sans-serif;
    }
    .main { background: #f7f4ef; }
    .block-container { padding-top: 2rem; padding-bottom: 2rem; }

    /* Header */
    .app-header {
        background: #1a1714;
        color: #f7f4ef;
        padding: 2rem 2.5rem;
        border-radius: 12px;
        margin-bottom: 2rem;
        position: relative;
        overflow: hidden;
    }
    .app-header::after {
        content: '📦';
        position: absolute;
        right: 2rem; top: 50%;
        transform: translateY(-50%);
        font-size: 4rem;
        opacity: 0.15;
    }
    .app-header h1 {
        font-family: 'DM Serif Display', serif;
        font-size: 2rem;
        margin: 0 0 0.3rem 0;
        font-weight: 400;
    }
    .app-header p { color: #9a9488; margin: 0; font-size: 0.9rem; }
    .app-header .tag {
        display: inline-block;
        background: #2a2d35;
        color: #c8e87a;
        font-family: 'JetBrains Mono', monospace;
        font-size: 0.65rem;
        letter-spacing: 0.12em;
        text-transform: uppercase;
        padding: 0.2rem 0.6rem;
        border-radius: 3px;
        margin-bottom: 0.8rem;
    }

    /* KPI Cards */
    .kpi-grid { display: grid; grid-template-columns: repeat(3, 1fr); gap: 1rem; margin-bottom: 2rem; }
    .kpi-card {
        background: white;
        border: 1px solid #ddd8ce;
        border-radius: 8px;
        padding: 1.2rem;
    }
    .kpi-label { font-size: 0.75rem; color: #7a7268; text-transform: uppercase; letter-spacing: 0.08em; margin-bottom: 0.4rem; }
    .kpi-value { font-family: 'DM Serif Display', serif; font-size: 1.8rem; color: #1a1714; line-height: 1; }
    .kpi-value.green  { color: #1a7a4a; }
    .kpi-value.red    { color: #8a1a1a; }
    .kpi-value.orange { color: #b84c00; }
    .kpi-value.blue   { color: #1a4f8a; }
    .kpi-value.teal   { color: #0f6e56; }

    /* Status badges */
    .badge {
        display: inline-block;
        font-family: 'JetBrains Mono', monospace;
        font-size: 0.65rem;
        letter-spacing: 0.08em;
        text-transform: uppercase;
        padding: 0.2rem 0.6rem;
        border-radius: 3px;
        margin-right: 0.3rem;
    }
    .badge-a  { background: #e8f5ee; color: #1a7a4a; }
    .badge-b  { background: #e6eef8; color: #1a4f8a; }
    .badge-c  { background: #f5f5f5; color: #555555; }
    .badge-ok { background: #e8f5ee; color: #1a7a4a; }
    .badge-exc{ background: #fdf0e6; color: #b84c00; }
    .badge-rup{ background: #f8e6e6; color: #8a1a1a; }

    /* Historico */
    .hist-item {
        background: white;
        border: 1px solid #ddd8ce;
        border-radius: 8px;
        padding: 1rem 1.2rem;
        margin-bottom: 0.6rem;
        display: flex;
        justify-content: space-between;
        align-items: center;
    }
    .hist-date { font-family: 'JetBrains Mono', monospace; font-size: 0.75rem; color: #7a7268; }
    .hist-name { font-weight: 500; font-size: 0.9rem; }
    .hist-stats { font-size: 0.8rem; color: #7a7268; }

    /* Upload zone */
    [data-testid="stFileUploader"] {
        background: white;
        border: 2px dashed #ddd8ce;
        border-radius: 8px;
        padding: 0.5rem;
    }

    /* Buttons */
    .stButton > button {
        background: #1a1714 !important;
        color: #f7f4ef !important;
        border: none !important;
        border-radius: 6px !important;
        font-family: 'DM Sans', sans-serif !important;
        font-weight: 500 !important;
        padding: 0.6rem 1.5rem !important;
        transition: all 0.2s !important;
    }
    .stButton > button:hover { background: #2a2d35 !important; }

    /* Sidebar */
    [data-testid="stSidebar"] {
        background: #1a1714;
    }
    [data-testid="stSidebar"] * { color: #f7f4ef !important; }
    [data-testid="stSidebar"] .stSlider > div > div > div { background: #c8e87a !important; }

    /* Section titles */
    .section-title {
        font-family: 'DM Serif Display', serif;
        font-size: 1.3rem;
        color: #1a1714;
        font-weight: 400;
        margin: 1.5rem 0 1rem;
        padding-bottom: 0.5rem;
        border-bottom: 1px solid #ddd8ce;
    }

    /* Alert boxes */
    .alert-box {
        border-radius: 6px;
        padding: 0.8rem 1rem;
        margin-bottom: 1rem;
        font-size: 0.88rem;
    }
    .alert-red    { background: #f8e6e6; border-left: 3px solid #8a1a1a; color: #8a1a1a; }
    .alert-orange { background: #fdf0e6; border-left: 3px solid #b84c00; color: #b84c00; }
    .alert-green  { background: #e8f5ee; border-left: 3px solid #1a7a4a; color: #1a7a4a; }
    .alert-blue   { background: #e6eef8; border-left: 3px solid #1a4f8a; color: #1a4f8a; }

    /* Meses pill */
    .mes-pill {
        display: inline-block;
        background: #e6eef8;
        color: #1a4f8a;
        font-family: 'JetBrains Mono', monospace;
        font-size: 0.7rem;
        padding: 0.2rem 0.5rem;
        border-radius: 3px;
        margin: 0.1rem;
    }
</style>
""", unsafe_allow_html=True)

# ── HISTÓRICO ─────────────────────────────────────────────────────────────────
HIST_FILE = Path("historico.json")

def carregar_historico():
    if HIST_FILE.exists():
        with open(HIST_FILE) as f:
            return json.load(f)
    return []

def salvar_historico(entrada):
    hist = carregar_historico()
    hist.insert(0, entrada)
    hist = hist[:20]  # Manter últimas 20
    with open(HIST_FILE, 'w') as f:
        json.dump(hist, f, ensure_ascii=False, indent=2)

# ── SIDEBAR ───────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("### ⚙️ Parâmetros")
    st.markdown("---")

    lead_time = st.slider(
        "Lead Time padrão (dias)",
        min_value=7, max_value=90, value=30, step=1,
        help="Dias entre o pedido e a entrega do fornecedor"
    )
    custo_pedido = st.number_input(
        "Custo por pedido (R$)",
        min_value=10, max_value=500, value=50, step=10,
        help="Custo administrativo de emitir um pedido de compra"
    )
    taxa_carr = st.slider(
        "Taxa de carregamento (%/ano)",
        min_value=10, max_value=40, value=25, step=5,
        help="Custo anual de manter o estoque (% do valor)"
    ) / 100

    st.markdown("---")
    st.markdown("### 📊 Níveis de Serviço")
    st.markdown("""
    <div style='font-size:0.8rem; color:#9a9488; line-height:1.8'>
    <b style='color:#f7f4ef'>Classe A</b> → Z=1,65 → 95%<br>
    <b style='color:#f7f4ef'>Classe B</b> → Z=1,28 → 90%<br>
    <b style='color:#f7f4ef'>Classe C</b> → Z=1,00 → 84%
    </div>
    """, unsafe_allow_html=True)

    st.markdown("---")
    st.markdown("### 📁 Histórico")
    hist = carregar_historico()
    if hist:
        for h in hist[:5]:
            st.markdown(f"""
            <div style='background:#2a2d35;border-radius:6px;padding:0.6rem 0.8rem;margin-bottom:0.4rem;'>
                <div style='font-size:0.7rem;color:#9a9488;font-family:monospace'>{h['data']}</div>
                <div style='font-size:0.85rem;font-weight:500'>{h['nome']}</div>
                <div style='font-size:0.75rem;color:#9a9488'>{h['itens_inventario']} itens · R$ {h['valor_inventario']:,.0f}</div>
            </div>
            """, unsafe_allow_html=True)
    else:
        st.markdown("<p style='font-size:0.8rem;color:#9a9488'>Nenhuma análise ainda</p>", unsafe_allow_html=True)

# ── HEADER ────────────────────────────────────────────────────────────────────
st.markdown("""
<div class="app-header">
    <div class="tag">Supply Chain · Análise de Estoque</div>
    <h1>Sistema de Gestão de Estoque</h1>
    <p>Carregue os arquivos CSV para gerar a análise completa com parâmetros de reposição, curva ABC, plano de contagem e painel de acurácia.</p>
</div>
""", unsafe_allow_html=True)

# ── UPLOAD ────────────────────────────────────────────────────────────────────
st.markdown('<div class="section-title">📂 Carregar Arquivos</div>', unsafe_allow_html=True)

col1, col2 = st.columns(2)
with col1:
    st.markdown("**Inventário** — exportação do saldo atual")
    inv_file = st.file_uploader("CSV de Inventário", type=['csv'], key="inv",
                                label_visibility="collapsed")
    if inv_file:
        st.markdown('<div class="alert-box alert-green">✅ Inventário carregado</div>', unsafe_allow_html=True)

with col2:
    st.markdown("**Movimentações** — saídas e entradas do período")
    sai_file = st.file_uploader("CSV de Movimentações", type=['csv'], key="sai",
                                label_visibility="collapsed")
    if sai_file:
        st.markdown('<div class="alert-box alert-green">✅ Movimentações carregadas</div>', unsafe_allow_html=True)

# ── NOME DA ANÁLISE ───────────────────────────────────────────────────────────
nome_analise = st.text_input(
    "Nome desta análise (para o histórico)",
    value=f"Análise {date.today().strftime('%b/%Y')}",
    help="Ex: 'Análise Mai/2026' ou 'Fechamento Q2'"
)

# ── BOTÃO GERAR ───────────────────────────────────────────────────────────────
st.markdown("")
gerar = st.button("🚀 Gerar Análise Completa", use_container_width=True)

if gerar:
    if not inv_file or not sai_file:
        st.markdown('<div class="alert-box alert-red">❌ Carregue os dois arquivos antes de gerar.</div>', unsafe_allow_html=True)
    else:
        with st.spinner("Processando dados..."):
            try:
                resultado = rodar_analise(inv_file, sai_file, lead_time, custo_pedido, taxa_carr)
                st.session_state['resultado'] = resultado
                st.session_state['nome'] = nome_analise

                # Salvar no histórico
                salvar_historico({
                    'data': datetime.now().strftime('%d/%m/%Y %H:%M'),
                    'nome': nome_analise,
                    'itens_inventario': len(resultado['inv_clean']),
                    'itens_consumo': len(resultado['df_param']),
                    'valor_inventario': float(resultado['inv_clean']['Total'].sum()),
                    'excesso': float(resultado['df_param']['Excesso_R$'].sum()),
                    'meses': resultado['meses_cmm'],
                    'lead_time': lead_time,
                    'custo_pedido': custo_pedido,
                    'taxa_carr': taxa_carr,
                })
                excel_bytes = gerar_excel(resultado)
                st.session_state['excel_bytes'] = excel_bytes.read()
                st.success("✅ Análise concluída!")
            except Exception as e:
                st.error(f"Erro ao processar: {e}")
                st.exception(e)

# ── RESULTADOS ────────────────────────────────────────────────────────────────
if 'resultado' in st.session_state:
    res = st.session_state['resultado']
    df_param  = res['df_param']
    df_giro   = res['df_giro']
    sem_saida = res['sem_saida']
    em_aberto = res['em_aberto']
    inv_clean = res['inv_clean']
    meses_cmm = res['meses_cmm']

    st.markdown(f'<div class="section-title">📊 Resultados — {st.session_state["nome"]}</div>', unsafe_allow_html=True)

    # Meses utilizados
    meses_html = " ".join([f'<span class="mes-pill">{m}</span>' for m in meses_cmm])
    st.markdown(f"<p style='font-size:0.85rem;color:#7a7268;margin-bottom:1rem'>CMM calculado com {res['n_meses']} meses completos: {meses_html}</p>", unsafe_allow_html=True)

    # ── KPIs ──────────────────────────────────────────────────────────────
    c1,c2,c3,c4,c5,c6 = st.columns(6)
    with c1:
        st.metric("Itens Inventário", f"{len(inv_clean):,}")
    with c2:
        st.metric("Com Consumo Real", f"{len(df_param):,}")
    with c3:
        st.metric("Giro Direto", f"{len(df_giro):,}")
    with c4:
        st.metric("Sem Movimento", f"{len(sem_saida):,}")
    with c5:
        st.metric("Valor Total", f"R$ {inv_clean['Total'].sum():,.0f}")
    with c6:
        st.metric("Excesso Estimado", f"R$ {df_param['Excesso_R$'].sum():,.0f}",
                  delta=f"-{df_param['Excesso_R$'].sum():,.0f}", delta_color="inverse")

    st.markdown("---")

    # ── ABC ───────────────────────────────────────────────────────────────
    col_abc, col_status = st.columns(2)
    with col_abc:
        st.markdown("**Curva ABC**")
        tv = df_param['Valor_Consumo_Anual'].sum()
        for cls, cor in [('A','🟢'),('B','🔵'),('C','⚫')]:
            sub = df_param[df_param['ABC']==cls]
            pct = sub['Valor_Consumo_Anual'].sum()/max(tv,1)*100
            st.markdown(f"{cor} **Classe {cls}** — {len(sub)} itens · {pct:.1f}% do valor")

    with col_status:
        st.markdown("**Status do Estoque**")
        for st_label, emoji in [('OK','✅'),('ACIMA DO MÁXIMO','⚠️'),('ABAIXO DO PP','🔴')]:
            cnt = df_param['Status'].value_counts().get(st_label, 0)
            pct = cnt/max(len(df_param),1)*100
            st.markdown(f"{emoji} **{st_label}** — {cnt} itens ({pct:.0f}%)")

    st.markdown("---")

    # ── ALERTAS RÁPIDOS ───────────────────────────────────────────────────
    ruptura = df_param[df_param['Status']=='ABAIXO DO PP']
    excesso = df_param[df_param['Excesso_R$']>0]

    if len(ruptura) > 0:
        st.markdown(f'<div class="alert-box alert-red">🔴 <b>{len(ruptura)} itens</b> abaixo do ponto de pedido — risco de ruptura imediato. Verifique a aba "RISCO DE RUPTURA" no Excel.</div>', unsafe_allow_html=True)

    if len(excesso) > 0:
        st.markdown(f'<div class="alert-box alert-orange">⚠️ <b>{len(excesso)} itens</b> acima do estoque máximo — R$ {excesso["Excesso_R$"].sum():,.0f} em excesso imobilizado.</div>', unsafe_allow_html=True)

    if len(em_aberto) > 0:
        st.markdown(f'<div class="alert-box alert-blue">🔄 <b>{len(em_aberto)} produtos</b> com remessas em aberto na LELY Filial — R$ {em_aberto["Valor_Pendente"].sum():,.0f} pendentes.</div>', unsafe_allow_html=True)

    st.markdown("---")

    # ── PRÉVIA DOS DADOS ──────────────────────────────────────────────────
    tab1, tab2, tab3, tab4 = st.tabs(["📋 Top Itens A", "🔴 Risco de Ruptura", "⚠️ Excesso", "🔁 Giro Direto"])

    with tab1:
        top_a = df_param[df_param['ABC']=='A'][['Produto','Descricao','CMM','PP','Est_Max','Estoque_Atual','Status','Valor_Consumo_Anual']].head(15)
        top_a.columns = ['Produto','Descrição','CMM','PP','Est. Máx','Estoque Atual','Status','Cons. Anual (R$)']
        st.dataframe(top_a, use_container_width=True, hide_index=True)

    with tab2:
        if len(ruptura) > 0:
            rup_show = ruptura[['Produto','Descricao','ABC','Estoque_Atual','PP','CMM','V.Unitario']].head(20)
            rup_show.columns = ['Produto','Descrição','ABC','Estoque Atual','PP','CMM','Custo Unit.']
            st.dataframe(rup_show, use_container_width=True, hide_index=True)
        else:
            st.success("Nenhum item abaixo do ponto de pedido! 🎉")

    with tab3:
        if len(excesso) > 0:
            exc_show = excesso[['Produto','Descricao','ABC','Estoque_Atual','Est_Max','Excesso_Un','Excesso_R$']].head(20)
            exc_show.columns = ['Produto','Descrição','ABC','Estoque Atual','Est. Máx','Excesso (un)','Excesso (R$)']
            st.dataframe(exc_show, use_container_width=True, hide_index=True)
        else:
            st.success("Nenhum item acima do estoque máximo! 🎉")

    with tab4:
        if len(df_giro) > 0:
            giro_show = df_giro[['Produto','Descricao','CMM','Qtd_Vendida','N_Vendas']].head(20)
            giro_show.columns = ['Produto','Descrição','CMM','Qtd Vendida','N° Vendas']
            st.dataframe(giro_show, use_container_width=True, hide_index=True)

    st.markdown("---")

    # ── DOWNLOAD ──────────────────────────────────────────────────────────
    st.markdown('<div class="section-title">📥 Baixar Excel Completo</div>', unsafe_allow_html=True)
    st.markdown("O arquivo inclui **10 abas**: Resumo · Parâmetros · Giro Direto · Remessas · Em Poder Terceiros · Excesso · Ruptura · Sem Movimento · Plano de Contagem · Painel Acurácia")

    nome_arquivo = f"Estoque_{date.today().strftime('%Y%m%d')}_{st.session_state['nome'].replace(' ','_').replace('/','_')}.xlsx"

    if 'excel_bytes' in st.session_state:
        st.download_button(
            label="⬇️ Baixar Excel Completo",
            data=st.session_state['excel_bytes'],
            file_name=nome_arquivo,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
    else:
        st.info("Clique em 'Gerar Análise Completa' para gerar o Excel.")
