"""
app.py — Sistema de Gestão de Estoque
Duas abas: Análise (gera Excel) + Contagem (mobile, Google Sheets)
"""
import streamlit as st
import pandas as pd
import numpy as np
from datetime import date, datetime, timedelta
import json, math

from engine import rodar_analise, gerar_excel, calc_acuracia, calcular_acuracia_contagens, FREQ_DIAS
from sheets import (ler_contagens, salvar_contagem, remover_contagem,
                    limpar_todas_contagens, salvar_itens, ler_itens, sheets_configurado)

# ── CONFIG ─────────────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Gestão de Estoque",
    page_icon="📦",
    layout="wide",
    initial_sidebar_state="collapsed",
)

# ── CSS ────────────────────────────────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@400;500;600;700&family=DM+Serif+Display:ital@0;1&display=swap');
html, body, [class*="css"] { font-family: 'DM Sans', sans-serif; }
.block-container { padding-top: 1.5rem; padding-bottom: 2rem; }

/* KPI Cards */
.kpi-row { display: flex; gap: 12px; flex-wrap: wrap; margin-bottom: 1.2rem; }
.kpi { background: white; border: 1px solid #E0DDD8; border-radius: 10px;
       padding: 14px 18px; flex: 1; min-width: 130px; text-align: center; }
.kpi-val { font-size: 1.8rem; font-weight: 800; line-height: 1.1; }
.kpi-lbl { font-size: 0.72rem; color: #7A7268; text-transform: uppercase;
           letter-spacing: .5px; margin-top: 3px; font-weight: 600; }
.kpi.green .kpi-val { color: #1A7A4A; }
.kpi.red   .kpi-val { color: #8A1A1A; }
.kpi.blue  .kpi-val { color: #1A4F8A; }
.kpi.orange .kpi-val { color: #B84C00; }
.kpi.amber  .kpi-val { color: #7A4A00; }
.kpi.teal   .kpi-val { color: #0F6E56; }

/* Alert items */
.alert-item { background: white; border-radius: 10px; border: 1px solid #E0DDD8;
              padding: 12px 14px; margin-bottom: 8px; cursor: pointer;
              border-left: 4px solid #E0DDD8; transition: .15s; }
.alert-item:hover { box-shadow: 0 2px 8px rgba(0,0,0,.08); }
.alert-item.urgente { border-left-color: #8A1A1A; }
.alert-item.excesso { border-left-color: #B84C00; }
.alert-item.normal  { border-left-color: #1A4F8A; }
.alert-code { font-family: monospace; font-size: 11px; color: #7A7268; }
.alert-desc { font-size: 13px; font-weight: 600; margin: 2px 0; }
.alert-meta { font-size: 11px; color: #7A7268; }
.tag { display: inline-block; font-size: 10px; font-weight: 700;
       padding: 2px 8px; border-radius: 20px; }
.tag-A { background: #E8F5EE; color: #1A7A4A; }
.tag-B { background: #E6EEF8; color: #1A4F8A; }
.tag-C { background: #F5F5F5; color: #555; }
.tag-D { background: #F0F0F0; color: #999; }

/* Acc badge */
.acc-ok   { color: #1A7A4A; font-weight: 800; }
.acc-warn { color: #B84C00; font-weight: 800; }
.acc-bad  { color: #8A1A1A; font-weight: 800; }

/* Section headers */
.section-hdr { font-size: 13px; font-weight: 700; color: #1A1714;
               margin: 1rem 0 .5rem; display: flex; justify-content: space-between; }

/* Mobile-friendly buttons */
.stButton > button { border-radius: 8px !important; font-weight: 600 !important; }

/* Tabs */
.stTabs [data-baseweb="tab"] { font-size: 15px; font-weight: 600; }

/* Acc progress */
.progress-bg { background: #F0EDE8; border-radius: 8px; height: 10px; overflow: hidden; margin: 6px 0; }
.progress-fill { height: 100%; border-radius: 8px; transition: .4s; }
</style>
""", unsafe_allow_html=True)


# ── HELPERS ────────────────────────────────────────────────────────────────────
def fmt_acc(v):
    if v is None: return "–"
    return f"{v*100:.1f}%"

def acc_class(v):
    if v is None: return ""
    if v >= 0.95: return "acc-ok"
    if v >= 0.80: return "acc-warn"
    return "acc-bad"

def acc_color(v):
    if v is None: return "#7A7268"
    if v >= 0.95: return "#1A7A4A"
    if v >= 0.80: return "#B84C00"
    return "#8A1A1A"

def prioridade_contagem(item, contagens_df):
    """Calcula prioridade de contagem de um item."""
    codigo = item['codigo']
    cont = None
    if contagens_df is not None and len(contagens_df) > 0:
        row = contagens_df[contagens_df['codigo'] == codigo]
        if len(row) > 0:
            cont = row.iloc[0]

    score = 0
    # 1. Urgência de estoque
    if item['prioridade'] == 'urgente': score += 100
    elif item['prioridade'] == 'excesso': score += 50
    # 2. Classe ABC
    score += {'A': 40, 'B': 20, 'C': 10, 'D': 0}.get(item['abc'], 0)
    # 3. Nunca contado = prioridade alta
    if cont is None: score += 30
    else:
        # Contado há muito tempo
        try:
            dt = datetime.fromisoformat(cont['data'])
            dias = (datetime.now() - dt).days
            freq = FREQ_DIAS.get(item['abc'], 90)
            if dias > freq: score += 25
        except:
            pass
        # Acurácia ruim = recontar
        acc = calc_acuracia(item['qtdSistema'], cont['qtd'])
        if acc is not None and acc < 0.80: score += 35

    return score


# ── SESSION STATE ──────────────────────────────────────────────────────────────
if 'resultado' not in st.session_state:
    st.session_state['resultado'] = None
if 'excel_bytes' not in st.session_state:
    st.session_state['excel_bytes'] = None
if 'usuario' not in st.session_state:
    st.session_state['usuario'] = 'Almoxarife'
if 'item_selecionado' not in st.session_state:
    st.session_state['item_selecionado'] = None
if 'busca' not in st.session_state:
    st.session_state['busca'] = ''


# ── HEADER ─────────────────────────────────────────────────────────────────────
col_h1, col_h2 = st.columns([3, 1])
with col_h1:
    st.markdown("## 📦 Gestão de Estoque — Lely Center Carambeí")
with col_h2:
    sheets_ok = sheets_configurado()
    if sheets_ok:
        st.success("🔗 Google Sheets conectado", icon="✅")
    else:
        st.warning("⚠️ Google Sheets não configurado")

# ── ABAS PRINCIPAIS ────────────────────────────────────────────────────────────
aba_analise, aba_contagem = st.tabs(["📊 Análise & Relatórios", "📋 Contagem de Estoque"])


# ══════════════════════════════════════════════════════════════════════════════
# ABA 1 — ANÁLISE
# ══════════════════════════════════════════════════════════════════════════════
with aba_analise:

    # Sidebar params
    with st.sidebar:
        st.markdown("### ⚙️ Parâmetros")
        lead_time   = st.slider("Lead Time padrão (dias)", 7, 90, 30)
        custo_ped   = st.number_input("Custo por pedido (R$)", 10, 500, 50, 10)
        taxa_carr   = st.slider("Taxa de carregamento (%/ano)", 10, 40, 25, 5) / 100

    col1, col2 = st.columns(2)
    with col1:
        st.markdown("**Inventário** — saldo atual")
        inv_file = st.file_uploader("CSV Inventário", type=['csv'], key="inv",
                                    label_visibility="collapsed")
        if inv_file: st.success("✅ Inventário carregado")
    with col2:
        st.markdown("**Movimentações** — todas as NFs")
        sai_file = st.file_uploader("CSV Movimentações", type=['csv'], key="sai",
                                    label_visibility="collapsed")
        if sai_file: st.success("✅ Movimentações carregadas")

    nome = st.text_input("Nome desta análise", f"Análise {date.today().strftime('%b/%Y')}")

    if st.button("🚀 Gerar Análise Completa", use_container_width=True, type="primary"):
        if not inv_file or not sai_file:
            st.error("Carregue os dois arquivos antes de gerar.")
        else:
            with st.spinner("Processando..."):
                try:
                    resultado = rodar_analise(inv_file, sai_file, custo_ped, taxa_carr)
                    st.session_state['resultado'] = resultado

                    # Salvar itens no Google Sheets para a aba de contagem
                    if sheets_ok:
                        with st.spinner("Sincronizando itens com Google Sheets..."):
                            salvar_itens(resultado['itens_contagem'])
                        st.success(f"✅ {len(resultado['itens_contagem'])} itens sincronizados com a aba Contagem!")

                    # Ler contagens existentes e gerar Excel
                    cont_df = ler_contagens() if sheets_ok else None
                    excel   = gerar_excel(resultado, cont_df if (cont_df is not None and len(cont_df)>0) else None)
                    st.session_state['excel_bytes'] = excel.read()
                    st.success("✅ Análise concluída!")
                except Exception as e:
                    st.error(f"Erro: {e}")
                    st.exception(e)

    # Resultados
    if st.session_state['resultado']:
        res = st.session_state['resultado']
        df  = res['df_param']
        inv = res['inv_clean']
        emb = res['em_aberto']
        gs  = res['sem_saida']

        # Acurácia real das contagens
        cont_df = ler_contagens() if sheets_ok else None
        acc = calcular_acuracia_contagens(
            df[['Produto','ABC','Estoque_Atual']].rename(columns={'Produto':'codigo'}),
            cont_df) if cont_df is not None else None

        st.markdown("---")

        # KPIs
        st.markdown('<div class="kpi-row">', unsafe_allow_html=True)
        kpis_html = ""
        kpis = [
            ("Inventário", f"{len(inv):,}", "blue"),
            ("Consumo Real", f"{len(df):,}", "teal"),
            ("Sem Movimento", f"{len(gs):,}", "orange"),
            ("Valor Total", f"R$ {inv['Total'].sum()/1e6:.2f}M", "blue"),
            ("Excesso", f"R$ {df['Excesso_R$'].sum()/1e3:.0f}K", "red"),
        ]
        if acc and acc['geral'] is not None:
            kpis.append((f"Acurácia ({acc['contados']} ctd.)", fmt_acc(acc['geral']),
                         "green" if acc['geral']>=0.95 else "red"))
        for lbl, val, cls in kpis:
            kpis_html += f'<div class="kpi {cls}"><div class="kpi-val">{val}</div><div class="kpi-lbl">{lbl}</div></div>'
        st.markdown(kpis_html + '</div>', unsafe_allow_html=True)

        # Acurácia por classe
        if acc and acc['geral'] is not None:
            st.markdown("**Acurácia por Classe**")
            c1, c2, c3, c4 = st.columns(4)
            for col, cls, meta in [(c1,'A',0.98),(c2,'B',0.95),(c3,'C',0.95),(c4,'D',0.95)]:
                with col:
                    v = acc.get(cls)
                    delta = f"{(v-meta)*100:+.1f}pp vs meta" if v is not None else None
                    st.metric(f"Classe {cls}", fmt_acc(v), delta,
                              delta_color="normal" if v and v>=meta else "inverse")

        # Preview tabs
        t1, t2, t3 = st.tabs(["🔴 Ruptura", "⚠️ Excesso", "🔄 Remessas"])
        with t1:
            rup = df[df['Status']=='ABAIXO DO PP'][['Produto','Descricao','ABC','Estoque_Atual','PP','CMM','V.Unitario']].head(20)
            rup.columns = ['Produto','Descrição','ABC','Estoque','PP','CMM','Custo']
            st.dataframe(rup, use_container_width=True, hide_index=True)
        with t2:
            exc = df[df['Status']=='ACIMA DO MÁXIMO'][['Produto','Descricao','ABC','Estoque_Atual','Est_Max','Excesso_Un','Excesso_R$']].head(20)
            exc.columns = ['Produto','Descrição','ABC','Estoque','Est.Máx','Excesso (un)','Excesso (R$)']
            st.dataframe(exc, use_container_width=True, hide_index=True)
        with t3:
            st.dataframe(emb[['Produto','Descricao','Qtd_Pendente','Valor_Pendente']].head(20),
                         use_container_width=True, hide_index=True)

        # Download
        st.markdown("---")
        st.markdown("**Excel completo com 6 abas:** Painel · Parâmetros · Ações · Conciliação · Contagem · Sem Movimento")
        if st.session_state['excel_bytes']:
            st.download_button(
                "⬇️ Baixar Excel",
                data=st.session_state['excel_bytes'],
                file_name=f"Estoque_{date.today().strftime('%Y%m%d')}_{nome.replace(' ','_').replace('/','_')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True, type="primary")


# ══════════════════════════════════════════════════════════════════════════════
# ABA 2 — CONTAGEM
# ══════════════════════════════════════════════════════════════════════════════
with aba_contagem:

    # ── Verificar conexão ────────────────────────────────────────────────────
    if not sheets_ok:
        st.error("⚠️ Google Sheets não configurado. Siga o README para configurar os secrets.")
        st.stop()

    # ── Carregar dados ───────────────────────────────────────────────────────
    contagens_df = ler_contagens()
    itens_df     = ler_itens()

    # Usar itens da sessão se disponíveis (análise acabou de rodar)
    if st.session_state['resultado'] and len(itens_df) == 0:
        itens_raw = st.session_state['resultado']['itens_contagem']
    elif len(itens_df) > 0:
        itens_raw = itens_df.to_dict('records')
    else:
        itens_raw = []

    # ── PAINEL SUPERIOR ──────────────────────────────────────────────────────
    total = len(itens_raw)
    contados = len(contagens_df) if len(contagens_df) > 0 else 0

    # Calcular acurácia real
    if contados > 0 and len(itens_raw) > 0:
        itens_acc_df = pd.DataFrame([
            {'codigo': i['codigo'], 'ABC': i['abc'], 'Estoque_Atual': i['qtdSistema']}
            for i in itens_raw])
        acc = calcular_acuracia_contagens(
            itens_acc_df.rename(columns={'Estoque_Atual':'Estoque_Atual'}),
            contagens_df)
    else:
        acc = None

    # KPIs do painel de contagem
    st.markdown('<div class="kpi-row">', unsafe_allow_html=True)
    pct = contados/total*100 if total > 0 else 0
    acc_geral = acc['geral'] if acc and acc['geral'] is not None else None

    html_kpis = f"""
    <div class="kpi {'green' if acc_geral and acc_geral>=0.95 else 'red' if acc_geral else 'blue'}">
      <div class="kpi-val">{fmt_acc(acc_geral)}</div>
      <div class="kpi-lbl">Acurácia Geral</div>
    </div>
    <div class="kpi blue">
      <div class="kpi-val">{contados}/{total}</div>
      <div class="kpi-lbl">Itens Contados</div>
    </div>
    <div class="kpi {'green' if pct>=50 else 'orange'}">
      <div class="kpi-val">{pct:.0f}%</div>
      <div class="kpi-lbl">% Concluído</div>
    </div>
    <div class="kpi red">
      <div class="kpi-val">{sum(1 for i in itens_raw if i.get('prioridade')=='urgente')}</div>
      <div class="kpi-lbl">🔴 Comprar</div>
    </div>
    """
    st.markdown(html_kpis + '</div>', unsafe_allow_html=True)

    # Barra de progresso geral
    acc_pct = max(0, (acc_geral or 0)) * 100
    cor_bar = "#1A7A4A" if acc_pct >= 95 else "#B84C00" if acc_pct >= 80 else "#8A1A1A"
    st.markdown(f"""
    <div class="progress-bg">
      <div class="progress-fill" style="width:{acc_pct:.1f}%;background:{cor_bar}"></div>
    </div>
    <div style="font-size:11px;color:#7A7268;margin-bottom:12px">
      Meta: 95% &nbsp;·&nbsp; {contados} de {total} itens contados
    </div>""", unsafe_allow_html=True)

    # Acurácia por classe
    if acc:
        st.markdown("**Acurácia por Classe**")
        cols_acc = st.columns(4)
        for i, (cls, meta) in enumerate([('A',0.98),('B',0.95),('C',0.95),('D',0.95)]):
            with cols_acc[i]:
                v = acc.get(cls)
                delta_txt = f"{(v-meta)*100:+.1f}pp" if v is not None else None
                st.metric(f"Classe {cls}", fmt_acc(v), delta_txt,
                          delta_color="normal" if v and v>=meta else "inverse")

    st.markdown("---")

    # ── USUÁRIO ──────────────────────────────────────────────────────────────
    col_u1, col_u2 = st.columns([2, 1])
    with col_u1:
        st.session_state['usuario'] = st.selectbox(
            "👤 Quem está contando?",
            ["Almoxarife", "Gestor", "Auxiliar", "Outro"],
            index=["Almoxarife","Gestor","Auxiliar","Outro"].index(
                st.session_state['usuario']) if st.session_state['usuario'] in
                ["Almoxarife","Gestor","Auxiliar","Outro"] else 0)
    with col_u2:
        if st.button("🔄 Atualizar", use_container_width=True):
            ler_contagens.clear()
            ler_itens.clear()
            st.rerun()

    st.markdown("---")

    # ── TABS DA CONTAGEM ──────────────────────────────────────────────────────
    tab_alerta, tab_busca, tab_hist = st.tabs(
        ["🚨 O que Contar Agora", "🔍 Buscar & Registrar", "📅 Histórico"])

    # ── TAB: O QUE CONTAR AGORA ───────────────────────────────────────────────
    with tab_alerta:
        if not itens_raw:
            st.info("Gere a análise na aba **Análise & Relatórios** para ver os itens.")
        else:
            # Calcular prioridade para cada item
            itens_com_score = []
            for item in itens_raw:
                score = prioridade_contagem(item, contagens_df if len(contagens_df)>0 else None)
                itens_com_score.append((score, item))
            itens_com_score.sort(key=lambda x: -x[0])

            # Filtro
            filtro = st.radio("Filtrar por:", ["Todos urgentes", "Nunca contados", "Com divergência", "Classe A", "Classe B"],
                              horizontal=True, label_visibility="collapsed")

            lista_filtrada = []
            for score, item in itens_com_score:
                codigo = item['codigo']
                cont_row = None
                if len(contagens_df) > 0:
                    r = contagens_df[contagens_df['codigo']==codigo]
                    if len(r) > 0: cont_row = r.iloc[0]

                if filtro == "Todos urgentes" and item['prioridade'] not in ('urgente','excesso') and score < 50:
                    continue
                if filtro == "Nunca contados" and cont_row is not None:
                    continue
                if filtro == "Com divergência":
                    if cont_row is None: continue
                    acc_i = calc_acuracia(item['qtdSistema'], cont_row['qtd'])
                    if acc_i is None or acc_i >= 0.95: continue
                if filtro == "Classe A" and item['abc'] != 'A':
                    continue
                if filtro == "Classe B" and item['abc'] != 'B':
                    continue
                lista_filtrada.append((score, item, cont_row))

            st.markdown(f"**{len(lista_filtrada)} itens** · Clique para registrar a contagem")

            for score, item, cont_row in lista_filtrada[:50]:
                acc_i = calc_acuracia(item['qtdSistema'], cont_row['qtd']) if cont_row is not None else None
                acc_txt = fmt_acc(acc_i) if acc_i is not None else "Não contado"
                acc_cls = acc_class(acc_i) if acc_i is not None else ""
                status_txt = ""
                if item['prioridade'] == 'urgente':
                    status_txt = ' <span style="color:#8A1A1A;font-weight:700">🔴 COMPRAR</span>'
                elif item['prioridade'] == 'excesso':
                    status_txt = ' <span style="color:#B84C00;font-weight:700">⚠️ EXCESSO</span>'

                with st.expander(
                    f"[{item['abc']}] {item['codigo']} · {item['descricao'][:60]}... · "
                    f"Sist: **{item['qtdSistema']:.0f}** · {acc_txt}",
                    expanded=False):

                    cc1, cc2 = st.columns([3, 2])
                    with cc1:
                        st.markdown(f"**Código:** `{item['codigo']}`")
                        st.markdown(f"**Descrição:** {item['descricao']}")
                        st.markdown(f"**Endereço:** {item['endereco'] or '–'}")
                        st.markdown(f"**Unidade:** {item['unidade']} · **Classe:** {item['abc']} · **Freq:** {item['frequencia']}")
                        if cont_row is not None:
                            try:
                                dt = datetime.fromisoformat(cont_row['data'])
                                st.markdown(f"**Última contagem:** {dt.strftime('%d/%m/%Y %H:%M')} por {cont_row['usuario']}")
                            except: pass

                    with cc2:
                        st.markdown(f"**Saldo Sistema:** `{item['qtdSistema']:.1f} {item['unidade']}`")
                        if acc_i is not None:
                            div = (cont_row['qtd'] - item['qtdSistema'])
                            st.markdown(f"**Contagem anterior:** `{cont_row['qtd']:.1f}`")
                            st.markdown(f"**Divergência:** `{div:+.1f}`")
                            st.markdown(f"**Acurácia:** `{fmt_acc(acc_i)}`")

                    qty_key = f"qty_{item['codigo']}"
                    obs_key = f"obs_{item['codigo']}"
                    default_qty = float(cont_row['qtd']) if cont_row is not None else float(item['qtdSistema'])

                    nova_qty = st.number_input(
                        f"Quantidade Contada ({item['unidade']})",
                        min_value=0.0, value=default_qty, step=1.0,
                        key=qty_key)
                    obs = st.text_input("Observação (opcional)", key=obs_key,
                                        value=cont_row.get('observacao','') if cont_row is not None else '')

                    # Preview da acurácia em tempo real
                    nova_acc = calc_acuracia(item['qtdSistema'], nova_qty)
                    div_nova = nova_qty - item['qtdSistema']
                    cor_acc  = acc_color(nova_acc)
                    st.markdown(f"""
                    <div style="background:#F7F7F7;border-radius:8px;padding:10px;text-align:center;margin:8px 0">
                      <span style="font-size:1.4rem;font-weight:800;color:{cor_acc}">{fmt_acc(nova_acc)}</span>
                      &nbsp;&nbsp;
                      <span style="color:#7A7268;font-size:13px">Div: {div_nova:+.1f}</span>
                    </div>""", unsafe_allow_html=True)

                    b1, b2 = st.columns(2)
                    with b1:
                        if st.button("✅ Salvar Contagem", key=f"save_{item['codigo']}", use_container_width=True, type="primary"):
                            if salvar_contagem(item['codigo'], nova_qty, st.session_state['usuario'], obs):
                                st.success("Salvo!")
                                st.rerun()
                    with b2:
                        if cont_row is not None:
                            if st.button("🗑 Remover", key=f"del_{item['codigo']}", use_container_width=True):
                                if remover_contagem(item['codigo']):
                                    st.success("Removido!")
                                    st.rerun()

    # ── TAB: BUSCAR & REGISTRAR ───────────────────────────────────────────────
    with tab_busca:
        busca = st.text_input("🔍 Buscar por código ou descrição",
                              placeholder="Ex: 9.1138 ou VÁLVULA...")
        if busca and len(busca) >= 2:
            ql = busca.lower()
            results = [i for i in itens_raw
                       if ql in i['codigo'].lower() or ql in i['descricao'].lower()][:10]

            if not results:
                st.info("Nenhum item encontrado.")
            else:
                for item in results:
                    cont_row = None
                    if len(contagens_df) > 0:
                        r = contagens_df[contagens_df['codigo']==item['codigo']]
                        if len(r) > 0: cont_row = r.iloc[0]

                    acc_i = calc_acuracia(item['qtdSistema'], cont_row['qtd']) if cont_row is not None else None
                    with st.expander(
                        f"`{item['codigo']}` · {item['descricao'][:55]}... · Sist: {item['qtdSistema']:.0f}",
                        expanded=True):

                        c1, c2 = st.columns([3,2])
                        with c1:
                            st.markdown(f"**End.:** {item['endereco'] or '–'} · **Classe:** {item['abc']} · **Freq:** {item['frequencia']}")
                            if item['prioridade'] == 'urgente':
                                st.error("🔴 ESTOQUE ABAIXO DO PONTO DE PEDIDO")
                            elif item['prioridade'] == 'excesso':
                                st.warning("⚠️ ESTOQUE ACIMA DO MÁXIMO")
                        with c2:
                            if acc_i is not None:
                                st.metric("Acurácia anterior", fmt_acc(acc_i))

                        default = float(cont_row['qtd']) if cont_row is not None else float(item['qtdSistema'])
                        nova = st.number_input(f"Qtd. Contada ({item['unidade']})",
                                               min_value=0.0, value=default, step=1.0,
                                               key=f"b_{item['codigo']}")
                        obs = st.text_input("Obs.", key=f"bo_{item['codigo']}",
                                            value=cont_row.get('observacao','') if cont_row is not None else '')

                        nova_acc = calc_acuracia(item['qtdSistema'], nova)
                        cor = acc_color(nova_acc)
                        st.markdown(f'<div style="text-align:center;font-size:1.3rem;font-weight:800;color:{cor}">{fmt_acc(nova_acc)}</div>', unsafe_allow_html=True)

                        if st.button("✅ Salvar", key=f"bs_{item['codigo']}", use_container_width=True, type="primary"):
                            if salvar_contagem(item['codigo'], nova, st.session_state['usuario'], obs):
                                st.success("✅ Salvo!")
                                st.rerun()
        else:
            st.markdown("Digite pelo menos 2 caracteres para buscar.")

    # ── TAB: HISTÓRICO ────────────────────────────────────────────────────────
    with tab_hist:
        if len(contagens_df) == 0:
            st.info("Nenhuma contagem registrada ainda.")
        else:
            # Enriquecer com dados dos itens
            cont_enr = contagens_df.copy()
            if itens_raw:
                itens_idx = {i['codigo']: i for i in itens_raw}
                cont_enr['descricao'] = cont_enr['codigo'].map(
                    lambda c: itens_idx.get(c, {}).get('descricao', '')[:60])
                cont_enr['abc']       = cont_enr['codigo'].map(
                    lambda c: itens_idx.get(c, {}).get('abc', '–'))
                cont_enr['qtdSist']   = cont_enr['codigo'].map(
                    lambda c: itens_idx.get(c, {}).get('qtdSistema', 0))
                cont_enr['acc']       = cont_enr.apply(
                    lambda r: calc_acuracia(r['qtdSist'], r['qtd']), axis=1)

            col_f1, col_f2, col_f3 = st.columns(3)
            with col_f1:
                f_cls = st.selectbox("Classe", ["Todas","A","B","C","D"], label_visibility="collapsed")
            with col_f2:
                f_stat = st.selectbox("Status", ["Todos","✅ OK","❌ Divergência"], label_visibility="collapsed")
            with col_f3:
                f_user = st.selectbox("Usuário",
                    ["Todos"] + sorted(cont_enr['usuario'].unique().tolist()),
                    label_visibility="collapsed")

            df_hist = cont_enr.copy()
            if f_cls != "Todas" and 'abc' in df_hist.columns:
                df_hist = df_hist[df_hist['abc'] == f_cls]
            if f_user != "Todos":
                df_hist = df_hist[df_hist['usuario'] == f_user]
            if f_stat != "Todos" and 'acc' in df_hist.columns:
                if f_stat == "✅ OK":    df_hist = df_hist[df_hist['acc'] >= 0.99]
                if f_stat == "❌ Divergência": df_hist = df_hist[df_hist['acc'] < 0.95]

            st.markdown(f"**{len(df_hist)} contagens** · Acurácia média: {fmt_acc(df_hist['acc'].mean() if 'acc' in df_hist.columns else None)}")

            cols_show = ['codigo','descricao','abc','qtdSist','qtd','acc','usuario','data']
            cols_show = [c for c in cols_show if c in df_hist.columns]
            rename_map = {'codigo':'Código','descricao':'Descrição','abc':'Classe',
                          'qtdSist':'Sist.','qtd':'Contado','acc':'Acurácia',
                          'usuario':'Usuário','data':'Data/Hora'}
            df_show = df_hist[cols_show].rename(columns=rename_map).sort_values('Data/Hora', ascending=False)
            if 'Acurácia' in df_show.columns:
                df_show['Acurácia'] = df_show['Acurácia'].apply(
                    lambda x: fmt_acc(x) if pd.notna(x) else '–')
            st.dataframe(df_show, use_container_width=True, hide_index=True)

            col_dl, col_cl = st.columns(2)
            with col_dl:
                csv = df_show.to_csv(index=False).encode('utf-8-sig')
                st.download_button("⬇️ Exportar CSV", csv,
                                   f"contagens_{date.today()}.csv", "text/csv",
                                   use_container_width=True)
            with col_cl:
                if st.button("🗑 Limpar todas as contagens", use_container_width=True):
                    if st.checkbox("Confirmar exclusão de TODAS as contagens"):
                        if limpar_todas_contagens():
                            st.success("Contagens removidas!")
                            st.rerun()
