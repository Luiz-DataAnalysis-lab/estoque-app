"""
app.py — Sistema de Gestão de Estoque · Lely Center Carambeí
Aba 1: Análise & Relatórios | Aba 2: Contagem de Estoque
Storage: arquivo JSON local no servidor Streamlit (sem Google Cloud)
"""
import streamlit as st
import pandas as pd
import numpy as np
from datetime import date, datetime, timedelta
import math

from engine import (rodar_analise, gerar_excel, calc_acuracia,
                    calcular_acuracia_contagens, FREQ_DIAS, FREQ_LABEL)
from storage import (ler_contagens, salvar_contagem, remover_contagem,
                     limpar_todas_contagens, salvar_itens, ler_itens,
                     importar_contagens_csv)

# ── CONFIG ────────────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Gestão de Estoque · Lely",
    page_icon="📦",
    layout="wide",
    initial_sidebar_state="collapsed",
)

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@400;500;600;700&display=swap');
html,body,[class*="css"]{font-family:'DM Sans',sans-serif}
.block-container{padding-top:1.2rem;padding-bottom:2rem}

.kpi-row{display:flex;gap:10px;flex-wrap:wrap;margin-bottom:1rem}
.kpi{background:#fff;border:1px solid #E0DDD8;border-radius:10px;
     padding:12px 16px;flex:1;min-width:120px;text-align:center}
.kpi-val{font-size:1.7rem;font-weight:800;line-height:1.1}
.kpi-lbl{font-size:.7rem;color:#7A7268;text-transform:uppercase;
         letter-spacing:.5px;margin-top:3px;font-weight:600}
.kpi.green .kpi-val{color:#1A7A4A} .kpi.red .kpi-val{color:#8A1A1A}
.kpi.blue  .kpi-val{color:#1A4F8A} .kpi.orange .kpi-val{color:#B84C00}
.kpi.amber .kpi-val{color:#7A4A00} .kpi.teal .kpi-val{color:#0F6E56}

.item-card{background:#fff;border:1px solid #E0DDD8;border-radius:10px;
           padding:12px 14px;margin-bottom:8px;border-left:4px solid #E0DDD8}
.item-card.urgente{border-left-color:#8A1A1A;background:#FFF8F8}
.item-card.excesso{border-left-color:#B84C00;background:#FFFAF6}
.item-card.normal{border-left-color:#1A4F8A}
.item-code{font-family:monospace;font-size:11px;color:#7A7268}
.item-desc{font-size:13px;font-weight:600;margin:2px 0}
.item-meta{font-size:11px;color:#7A7268}

.acc-ok  {color:#1A7A4A;font-weight:800}
.acc-warn{color:#B84C00;font-weight:800}
.acc-bad {color:#8A1A1A;font-weight:800}

.prog-bg{background:#F0EDE8;border-radius:8px;height:10px;overflow:hidden;margin:4px 0}
.prog-fill{height:100%;border-radius:8px;transition:.4s}

.stTabs [data-baseweb="tab"]{font-size:15px;font-weight:600}
.stButton>button{border-radius:8px!important;font-weight:600!important}
</style>
""", unsafe_allow_html=True)


# ── HELPERS ───────────────────────────────────────────────────────────────────
def fmt_acc(v):
    return f"{v*100:.1f}%" if v is not None else "–"

def acc_color(v):
    if v is None: return "#7A7268"
    if v >= 0.95: return "#1A7A4A"
    if v >= 0.80: return "#B84C00"
    return "#8A1A1A"

def acc_class(v):
    if v is None: return ""
    if v >= 0.95: return "acc-ok"
    if v >= 0.80: return "acc-warn"
    return "acc-bad"

def prioridade_score(item, cont_row):
    """Calcula score de prioridade para a fila de contagem."""
    score = 0
    if item['prioridade'] == 'urgente': score += 100
    elif item['prioridade'] == 'excesso': score += 50
    score += {'A':40,'B':20,'C':10,'D':0}.get(item['abc'], 0)
    if cont_row is None:
        score += 30
    else:
        try:
            dias = (datetime.now() - datetime.fromisoformat(cont_row['data'])).days
            if dias > FREQ_DIAS.get(item['abc'], 90): score += 25
        except: pass
        acc = calc_acuracia(item['qtdSistema'], cont_row['qtd'])
        if acc is not None and acc < 0.80: score += 35
    return score


# ── SESSION STATE ──────────────────────────────────────────────────────────────
for k, v in [('resultado', None), ('excel_bytes', None),
              ('usuario', 'Almoxarife'), ('analise_nome', '')]:
    if k not in st.session_state:
        st.session_state[k] = v


# ── HEADER ─────────────────────────────────────────────────────────────────────
c1, c2 = st.columns([4, 1])
with c1:
    st.markdown("## 📦 Gestão de Estoque — Lely Center Carambeí")
with c2:
    st.success("✅ Pronto", icon="📦")


# ── ABAS ──────────────────────────────────────────────────────────────────────
aba_analise, aba_contagem = st.tabs(["📊 Análise & Relatórios", "📋 Contagem de Estoque"])


# ══════════════════════════════════════════════════════════════════════════
# ABA 1 — ANÁLISE
# ══════════════════════════════════════════════════════════════════════════
with aba_analise:

    with st.sidebar:
        st.markdown("### ⚙️ Parâmetros")
        custo_ped = st.number_input("Custo por pedido (R$)", 10, 500, 50, 10)
        taxa_carr = st.slider("Taxa carregamento (%/ano)", 10, 40, 25, 5) / 100

    col1, col2 = st.columns(2)
    with col1:
        st.markdown("**Inventário** — saldo atual do sistema")
        inv_file = st.file_uploader("CSV Inventário", type=['csv'], key="inv",
                                    label_visibility="collapsed")
        if inv_file: st.success("✅ Inventário carregado")
    with col2:
        st.markdown("**Movimentações** — todas as NFs (entradas e saídas)")
        sai_file = st.file_uploader("CSV Movimentações", type=['csv'], key="sai",
                                    label_visibility="collapsed")
        if sai_file: st.success("✅ Movimentações carregadas")

    st.session_state['analise_nome'] = st.text_input(
        "Nome desta análise", f"Análise {date.today().strftime('%b/%Y')}")

    if st.button("🚀 Gerar Análise Completa", use_container_width=True, type="primary"):
        if not inv_file or not sai_file:
            st.error("Carregue os dois arquivos antes de gerar.")
        else:
            with st.spinner("Processando dados..."):
                try:
                    resultado = rodar_analise(inv_file, sai_file, custo_ped, taxa_carr)
                    st.session_state['resultado'] = resultado

                    # Sincronizar itens para aba de contagem
                    with st.spinner("Disponibilizando itens para contagem..."):
                        salvar_itens(resultado['itens_contagem'])

                    # Gerar Excel com contagens existentes
                    cont_df = ler_contagens()
                    excel   = gerar_excel(
                        resultado,
                        cont_df if len(cont_df) > 0 else None)
                    st.session_state['excel_bytes'] = excel.read()
                    st.success(f"✅ Análise concluída! {len(resultado['itens_contagem'])} itens disponíveis para contagem.")
                except Exception as e:
                    st.error(f"Erro ao processar: {e}")
                    st.exception(e)

    # ── Resultados ────────────────────────────────────────────────────────
    if st.session_state['resultado']:
        res = st.session_state['resultado']
        df  = res['df_param']
        inv = res['inv_clean']
        emb = res['em_aberto']
        gs  = res['sem_saida']

        cont_df = ler_contagens()
        acc = calcular_acuracia_contagens(
            df[['Produto','ABC','Estoque_Atual']].rename(columns={'Produto':'codigo'}),
            cont_df if len(cont_df) > 0 else None)

        st.markdown("---")

        # KPIs
        kpi_html = '<div class="kpi-row">'
        kpis_data = [
            ("Inventário",    f"{len(inv):,}",                               "blue"),
            ("Consumo Real",  f"{len(df):,}",                                "teal"),
            ("Sem Movimento", f"{len(gs):,}",                                "orange"),
            ("Valor Total",   f"R$ {inv['Total'].sum()/1e6:.2f}M",           "blue"),
            ("Excesso",       f"R$ {df['Excesso_R$'].sum()/1e3:.0f}K",       "red"),
            ("Remessas",      f"R$ {emb['Valor_Pendente'].sum()/1e3:.0f}K",  "amber"),
        ]
        if acc and acc['geral'] is not None:
            cor = "green" if acc['geral'] >= 0.95 else "red"
            kpis_data.append((f"Acurácia ({acc['contados']} ctd.)",
                              fmt_acc(acc['geral']), cor))
        for lbl, val, cls in kpis_data:
            kpi_html += f'<div class="kpi {cls}"><div class="kpi-val">{val}</div><div class="kpi-lbl">{lbl}</div></div>'
        kpi_html += '</div>'
        st.markdown(kpi_html, unsafe_allow_html=True)

        # Acurácia por classe
        if acc and acc['geral'] is not None:
            st.markdown("**Acurácia por Classe**")
            cols_acc = st.columns(4)
            for col, cls, meta in zip(cols_acc, ['A','B','C','D'],
                                      [0.98, 0.95, 0.95, 0.95]):
                with col:
                    v = acc.get(cls)
                    delta = f"{(v-meta)*100:+.1f}pp" if v is not None else None
                    st.metric(f"Classe {cls}", fmt_acc(v), delta,
                              delta_color="normal" if v and v >= meta else "inverse")

        # Tabelas preview
        t1, t2, t3 = st.tabs(["🔴 Comprar Agora", "⚠️ Excesso", "🔄 Remessas"])
        with t1:
            r = df[df['Status']=='ABAIXO DO PP'][
                ['Produto','Descricao','ABC','LT_dias','Estoque_Atual','PP','CMM']].head(20)
            r.columns=['Produto','Descrição','ABC','LT(dias)','Estoque','PP','CMM']
            st.dataframe(r, use_container_width=True, hide_index=True)
        with t2:
            e = df[df['Status']=='ACIMA DO MÁXIMO'][
                ['Produto','Descricao','ABC','Estoque_Atual','Est_Max','Excesso_Un','Excesso_R$']].head(20)
            e.columns=['Produto','Descrição','ABC','Estoque','Est.Máx','Excesso(un)','Excesso(R$)']
            st.dataframe(e, use_container_width=True, hide_index=True)
        with t3:
            st.dataframe(
                emb[['Produto','Descricao','Qtd_Pendente','Valor_Pendente']].head(20),
                use_container_width=True, hide_index=True)

        # Download Excel
        st.markdown("---")
        st.markdown("**Excel com 6 abas:** Painel · Parâmetros · Ações · Conciliação · Contagem · Sem Movimento")
        if st.session_state['excel_bytes']:
            nome_arq = f"Estoque_{date.today().strftime('%Y%m%d')}_{st.session_state['analise_nome'].replace(' ','_').replace('/','_')}.xlsx"
            st.download_button(
                "⬇️ Baixar Excel Completo",
                data=st.session_state['excel_bytes'],
                file_name=nome_arq,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True, type="primary")


# ══════════════════════════════════════════════════════════════════════════
# ABA 2 — CONTAGEM
# ══════════════════════════════════════════════════════════════════════════
with aba_contagem:

    # Carregar dados
    contagens_df = ler_contagens()
    itens_df     = ler_itens()

    # Usar itens da sessão se Sheets ainda vazio
    if len(itens_df) == 0 and st.session_state['resultado']:
        itens_raw = st.session_state['resultado']['itens_contagem']
    else:
        itens_raw = itens_df.to_dict('records') if len(itens_df) > 0 else []

    total    = len(itens_raw)
    contados = len(contagens_df)

    # Mapa rápido código → contagem
    cont_map = {str(r['codigo']): r for _, r in contagens_df.iterrows()} \
               if contados > 0 else {}

    # Acurácia geral
    if contados > 0 and total > 0:
        itens_acc = pd.DataFrame([
            {'codigo': i['codigo'], 'ABC': i['abc'], 'Estoque_Atual': i['qtdSistema']}
            for i in itens_raw])
        acc_cont = calcular_acuracia_contagens(itens_acc, contagens_df)
    else:
        acc_cont = None

    acc_geral = acc_cont['geral'] if acc_cont else None
    pct_conc  = contados / total * 100 if total > 0 else 0

    # ── PAINEL ────────────────────────────────────────────────────────────
    kpi_c = '<div class="kpi-row">'
    cor_acc = "green" if acc_geral and acc_geral >= 0.95 else "red" if acc_geral else "blue"
    urgentes = sum(1 for i in itens_raw if i.get('prioridade') == 'urgente')
    excessos = sum(1 for i in itens_raw if i.get('prioridade') == 'excesso')
    ok_cnt   = sum(1 for c in cont_map.values()
                   if calc_acuracia(
                       next((i['qtdSistema'] for i in itens_raw if i['codigo']==c['codigo']), 0),
                       c['qtd']) is not None and
                   calc_acuracia(
                       next((i['qtdSistema'] for i in itens_raw if i['codigo']==c['codigo']), 0),
                       c['qtd']) >= 0.99)

    for v, l, c in [
        (fmt_acc(acc_geral), "Acurácia Geral", cor_acc),
        (f"{contados}/{total}", "Itens Contados", "blue"),
        (f"{pct_conc:.0f}%", "% Concluído", "teal" if pct_conc >= 50 else "orange"),
        (str(urgentes), "🔴 Comprar", "red"),
        (str(excessos), "⚠️ Excesso", "orange"),
        (str(ok_cnt), "✅ OK", "green"),
    ]:
        kpi_c += f'<div class="kpi {c}"><div class="kpi-val">{v}</div><div class="kpi-lbl">{l}</div></div>'
    kpi_c += '</div>'
    st.markdown(kpi_c, unsafe_allow_html=True)

    # Barra de progresso acurácia
    acc_pct = max(0, (acc_geral or 0)) * 100
    cor_bar = "#1A7A4A" if acc_pct >= 95 else "#B84C00" if acc_pct >= 80 else "#8A1A1A"
    st.markdown(f"""
    <div class="prog-bg">
      <div class="prog-fill" style="width:{acc_pct:.1f}%;background:{cor_bar}"></div>
    </div>
    <div style="font-size:11px;color:#7A7268;margin-bottom:8px">
      Acurácia · Meta: 95% · {contados} de {total} itens contados
    </div>""", unsafe_allow_html=True)

    # Acurácia por classe
    if acc_cont and acc_cont['geral'] is not None:
        st.markdown("**Acurácia por Classe**")
        cols_c = st.columns(4)
        for col, (cls, meta) in zip(cols_c, [('A',0.98),('B',0.95),('C',0.95),('D',0.95)]):
            with col:
                v = acc_cont.get(cls)
                d = f"{(v-meta)*100:+.1f}pp" if v is not None else None
                st.metric(f"Classe {cls}", fmt_acc(v), d,
                          delta_color="normal" if v and v >= meta else "inverse")

    st.markdown("---")

    # Usuário + Atualizar
    cu1, cu2 = st.columns([3, 1])
    with cu1:
        usuarios = ["Almoxarife", "Gestor", "Auxiliar", "Outro"]
        idx = usuarios.index(st.session_state['usuario']) \
              if st.session_state['usuario'] in usuarios else 0
        st.session_state['usuario'] = st.selectbox(
            "👤 Quem está contando?", usuarios, index=idx)
    with cu2:
        if st.button("🔄 Atualizar", use_container_width=True):
            st.rerun()

    st.markdown("---")

    # ── SUB-ABAS DA CONTAGEM ──────────────────────────────────────────────
    tab_fila, tab_busca, tab_hist = st.tabs(
        ["🚨 O que Contar Agora", "🔍 Buscar & Registrar", "📅 Histórico"])

    # ── FILA PRIORIZADA ───────────────────────────────────────────────────
    with tab_fila:
        if not itens_raw:
            st.info("Gere a análise na aba **Análise & Relatórios** para ver os itens.")
        else:
            # Filtro rápido
            filtro = st.radio(
                "Mostrar:",
                ["🔴 Urgentes + Excesso", "Nunca Contados", "Com Divergência", "Classe A", "Classe B", "Todos"],
                horizontal=True, label_visibility="collapsed")

            # Calcular scores e aplicar filtro
            fila = []
            for item in itens_raw:
                cont_row = cont_map.get(str(item['codigo']))
                acc_i    = calc_acuracia(item['qtdSistema'], cont_row['qtd']) \
                           if cont_row is not None else None
                score    = prioridade_score(item, cont_row)

                # Aplicar filtro
                if filtro == "🔴 Urgentes + Excesso" and item['prioridade'] not in ('urgente','excesso'):
                    continue
                if filtro == "Nunca Contados" and cont_row is not None:
                    continue
                if filtro == "Com Divergência" and (acc_i is None or acc_i >= 0.95):
                    continue
                if filtro == "Classe A" and item['abc'] != 'A':
                    continue
                if filtro == "Classe B" and item['abc'] != 'B':
                    continue

                fila.append((score, item, cont_row, acc_i))

            fila.sort(key=lambda x: -x[0])

            st.markdown(f"**{len(fila)} itens** — clique para registrar a contagem")

            for score, item, cont_row, acc_i in fila[:60]:
                prio = item['prioridade']
                acc_txt = fmt_acc(acc_i) if acc_i is not None else "Não contado"
                tag_prio = ""
                if prio == 'urgente': tag_prio = " 🔴"
                elif prio == 'excesso': tag_prio = " ⚠️"

                with st.expander(
                    f"[{item['abc']}]{tag_prio} {item['codigo']} · "
                    f"{item['descricao'][:55]}... · "
                    f"Sist: **{item['qtdSistema']:.0f}** · {acc_txt}"):

                    cc1, cc2 = st.columns([3, 2])
                    with cc1:
                        st.markdown(f"**Código:** `{item['codigo']}`")
                        st.markdown(f"**Endereço:** {item['endereco'] or '–'}")
                        st.markdown(f"**Classe:** {item['abc']} · **Freq:** {item['frequencia']}")
                        if prio == 'urgente':
                            st.error("🔴 Estoque ABAIXO do ponto de pedido — comprar!")
                        elif prio == 'excesso':
                            st.warning("⚠️ Estoque ACIMA do máximo — investigar")
                        if cont_row is not None:
                            try:
                                dt  = datetime.fromisoformat(cont_row['data'])
                                div = cont_row['qtd'] - item['qtdSistema']
                                st.markdown(
                                    f"**Última contagem:** {dt.strftime('%d/%m/%Y %H:%M')} "
                                    f"por {cont_row['usuario']} · "
                                    f"Qtd: {cont_row['qtd']:.0f} · Div: {div:+.1f}")
                            except: pass
                    with cc2:
                        st.markdown(f"**Saldo Sistema:** `{item['qtdSistema']:.1f} {item['unidade']}`")
                        if acc_i is not None:
                            st.markdown(f"**Acurácia anterior:** `{fmt_acc(acc_i)}`")

                    default_qty = float(cont_row['qtd']) if cont_row is not None \
                                  else float(item['qtdSistema'])
                    nova_qty = st.number_input(
                        f"Quantidade Contada ({item['unidade']})",
                        min_value=0.0, value=default_qty, step=1.0,
                        key=f"q_{item['codigo']}")
                    obs = st.text_input(
                        "Observação (opcional)", key=f"o_{item['codigo']}",
                        value=cont_row.get('observacao','') if cont_row else '')

                    # Preview acurácia em tempo real
                    nova_acc = calc_acuracia(item['qtdSistema'], nova_qty)
                    div_nova = nova_qty - item['qtdSistema']
                    cor      = acc_color(nova_acc)
                    st.markdown(
                        f'<div style="background:#F7F7F7;border-radius:8px;padding:10px;'
                        f'text-align:center;margin:6px 0">'
                        f'<span style="font-size:1.4rem;font-weight:800;color:{cor}">'
                        f'{fmt_acc(nova_acc)}</span>'
                        f'&nbsp;&nbsp;<span style="color:#7A7268;font-size:13px">'
                        f'Div: {div_nova:+.1f}</span></div>',
                        unsafe_allow_html=True)

                    b1, b2 = st.columns(2)
                    with b1:
                        if st.button("✅ Salvar", key=f"s_{item['codigo']}",
                                     use_container_width=True, type="primary"):
                            if salvar_contagem(item['codigo'], float(nova_qty),
                                               st.session_state['usuario'], obs):
                                st.success("Salvo!"); st.rerun()
                    with b2:
                        if cont_row is not None:
                            if st.button("🗑 Remover", key=f"d_{item['codigo']}",
                                         use_container_width=True):
                                if remover_contagem(item['codigo']):
                                    st.success("Removido!"); st.rerun()

    # ── BUSCA ─────────────────────────────────────────────────────────────
    with tab_busca:
        busca = st.text_input("🔍 Código ou descrição",
                              placeholder="Ex: 9.1138 ou VÁLVULA...")
        if busca and len(busca) >= 2:
            ql = busca.lower()
            results = [i for i in itens_raw
                       if ql in i['codigo'].lower() or ql in i['descricao'].lower()][:8]
            if not results:
                st.info("Nenhum item encontrado.")
            else:
                for item in results:
                    cont_row = cont_map.get(str(item['codigo']))
                    acc_i    = calc_acuracia(item['qtdSistema'], cont_row['qtd']) \
                               if cont_row else None
                    with st.expander(
                        f"`{item['codigo']}` · {item['descricao'][:55]}... · "
                        f"Sist: {item['qtdSistema']:.0f} · {fmt_acc(acc_i)}",
                        expanded=True):

                        c1, c2 = st.columns([3, 2])
                        with c1:
                            st.markdown(f"**End.:** {item['endereco'] or '–'} · **{item['abc']}** · {item['frequencia']}")
                            if item['prioridade'] == 'urgente': st.error("🔴 COMPRAR")
                            elif item['prioridade'] == 'excesso': st.warning("⚠️ EXCESSO")
                        with c2:
                            if acc_i is not None:
                                st.metric("Acurácia anterior", fmt_acc(acc_i))

                        default = float(cont_row['qtd']) if cont_row else float(item['qtdSistema'])
                        nova = st.number_input(
                            f"Qtd. Contada ({item['unidade']})",
                            min_value=0.0, value=default, step=1.0,
                            key=f"b_{item['codigo']}")
                        obs = st.text_input("Obs.", key=f"bo_{item['codigo']}",
                                            value=cont_row.get('observacao','') if cont_row else '')

                        nova_acc = calc_acuracia(item['qtdSistema'], nova)
                        cor = acc_color(nova_acc)
                        st.markdown(
                            f'<div style="text-align:center;font-size:1.3rem;'
                            f'font-weight:800;color:{cor};padding:8px 0">'
                            f'{fmt_acc(nova_acc)} · Div: {nova - item["qtdSistema"]:+.1f}</div>',
                            unsafe_allow_html=True)
                        if st.button("✅ Salvar", key=f"bs_{item['codigo']}",
                                     use_container_width=True, type="primary"):
                            if salvar_contagem(item['codigo'], float(nova),
                                               st.session_state['usuario'], obs):
                                st.success("✅ Salvo!"); st.rerun()
        elif busca:
            st.markdown("*Digite pelo menos 2 caracteres.*")

    # ── HISTÓRICO ─────────────────────────────────────────────────────────
    with tab_hist:
        if contados == 0:
            st.info("Nenhuma contagem registrada ainda.")
        else:
            # Enriquecer contagens com dados dos itens
            itens_idx = {i['codigo']: i for i in itens_raw}
            hist = contagens_df.copy()
            hist['descricao'] = hist['codigo'].map(
                lambda c: itens_idx.get(c, {}).get('descricao', '')[:60])
            hist['abc']       = hist['codigo'].map(
                lambda c: itens_idx.get(c, {}).get('abc', '–'))
            hist['qtdSist']   = hist['codigo'].map(
                lambda c: itens_idx.get(c, {}).get('qtdSistema', 0))
            hist['acc']       = hist.apply(
                lambda r: calc_acuracia(r['qtdSist'], r['qtd']), axis=1)
            hist['div']       = hist['qtd'] - hist['qtdSist']

            # Filtros
            f1, f2, f3 = st.columns(3)
            with f1:
                f_cls = st.selectbox("Classe", ["Todas","A","B","C","D"])
            with f2:
                f_stat = st.selectbox("Status", ["Todos","✅ OK","❌ Divergência"])
            with f3:
                f_user = st.selectbox(
                    "Usuário", ["Todos"] + sorted(hist['usuario'].dropna().unique().tolist()))

            df_h = hist.copy()
            if f_cls  != "Todas": df_h = df_h[df_h['abc'] == f_cls]
            if f_user != "Todos": df_h = df_h[df_h['usuario'] == f_user]
            if f_stat == "✅ OK":          df_h = df_h[df_h['acc'] >= 0.99]
            if f_stat == "❌ Divergência": df_h = df_h[df_h['acc'] < 0.95]

            acc_med = df_h['acc'].mean()
            st.markdown(f"**{len(df_h)} contagens** · Acurácia média: **{fmt_acc(acc_med)}**")

            show = df_h.sort_values('data', ascending=False)[
                ['codigo','descricao','abc','qtdSist','qtd','div','acc','usuario','data','observacao']
            ].rename(columns={
                'codigo':'Código','descricao':'Descrição','abc':'Classe',
                'qtdSist':'Sistema','qtd':'Contado','div':'Div.',
                'acc':'Acurácia','usuario':'Usuário','data':'Data/Hora','observacao':'Obs.'
            })
            show['Acurácia'] = show['Acurácia'].apply(fmt_acc)
            show['Div.']     = show['Div.'].apply(lambda x: f"{x:+.1f}" if pd.notna(x) else '')
            st.dataframe(show, use_container_width=True, hide_index=True)

            col_dl, col_cl = st.columns(2)
            with col_dl:
                csv = show.to_csv(index=False).encode('utf-8-sig')
                st.download_button(
                    "⬇️ Exportar CSV", csv,
                    f"contagens_{date.today()}.csv", "text/csv",
                    use_container_width=True)
            with col_cl:
                if st.button("🗑 Limpar todas", use_container_width=True):
                    if st.checkbox("Confirmar exclusão de TODAS as contagens"):
                        limpar_todas_contagens()
                        st.success("Contagens removidas!")
                        st.rerun()

            # ── Importar histórico ─────────────────────────────────────────
            st.markdown("---")
            st.markdown("**📥 Importar contagens de arquivo CSV**")
            st.caption("Formato esperado: colunas `codigo`, `qtd`, `usuario`, `data`, `observacao`")
            arq_imp = st.file_uploader("Selecione o CSV de contagens",
                                       type=['csv'], key="imp_csv")
            if arq_imp:
                df_imp = pd.read_csv(arq_imp, encoding='utf-8-sig')
                st.dataframe(df_imp.head(5), use_container_width=True, hide_index=True)
                if st.button("⬆️ Importar contagens", use_container_width=True, type="primary"):
                    ok, ig = importar_contagens_csv(df_imp)
                    st.success(f"✅ {ok} contagens importadas · {ig} ignoradas")
                    st.rerun()
