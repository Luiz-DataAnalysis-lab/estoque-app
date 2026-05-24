"""
sheets.py — Integração com Google Sheets
Salva e lê contagens em tempo real para múltiplos usuários.
"""
import streamlit as st
import pandas as pd
from datetime import datetime
import gspread
from google.oauth2.service_account import Credentials

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

# Colunas do Google Sheets
COLS_CONTAGENS = ["codigo", "qtd", "usuario", "data", "observacao"]
COLS_ITENS     = ["codigo", "descricao", "unidade", "endereco", "abc",
                  "frequencia", "proxContagem", "qtdSistema", "custoUnit",
                  "status", "prioridade"]


@st.cache_resource(ttl=0)
def get_client():
    """Cria cliente Google Sheets autenticado via secrets."""
    try:
        creds = Credentials.from_service_account_info(
            st.secrets["gcp_service_account"], scopes=SCOPES)
        return gspread.authorize(creds)
    except Exception as e:
        return None


def get_sheet(sheet_name: str):
    """Retorna a worksheet pelo nome."""
    gc = get_client()
    if gc is None:
        return None
    try:
        spreadsheet = gc.open_by_key(st.secrets["google_sheets"]["spreadsheet_id"])
        try:
            return spreadsheet.worksheet(sheet_name)
        except gspread.WorksheetNotFound:
            ws = spreadsheet.add_worksheet(title=sheet_name, rows=2000, cols=20)
            return ws
    except Exception as e:
        st.error(f"Erro ao conectar ao Google Sheets: {e}")
        return None


# ── CONTAGENS ────────────────────────────────────────────────────────────────

@st.cache_data(ttl=30)
def ler_contagens() -> pd.DataFrame:
    """Lê todas as contagens do Google Sheets. Cache de 30s."""
    ws = get_sheet("contagens")
    if ws is None:
        return pd.DataFrame(columns=COLS_CONTAGENS)
    try:
        data = ws.get_all_records()
        if not data:
            return pd.DataFrame(columns=COLS_CONTAGENS)
        df = pd.DataFrame(data)
        df['qtd'] = pd.to_numeric(df['qtd'], errors='coerce')
        # Manter apenas a última contagem por código
        df = df.sort_values('data').drop_duplicates(subset=['codigo'], keep='last')
        return df.reset_index(drop=True)
    except Exception as e:
        st.error(f"Erro ao ler contagens: {e}")
        return pd.DataFrame(columns=COLS_CONTAGENS)


def salvar_contagem(codigo: str, qtd: float, usuario: str, observacao: str = ""):
    """Salva ou atualiza uma contagem no Google Sheets."""
    ws = get_sheet("contagens")
    if ws is None:
        return False
    try:
        agora = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        # Verificar se já existe linha para este código
        try:
            cell = ws.find(codigo, in_column=1)
            ws.update(f"A{cell.row}:E{cell.row}",
                      [[codigo, qtd, usuario, agora, observacao]])
        except gspread.CellNotFound:
            # Verificar se tem cabeçalho
            all_vals = ws.get_all_values()
            if not all_vals or all_vals[0] != COLS_CONTAGENS:
                ws.insert_row(COLS_CONTAGENS, 1)
            ws.append_row([codigo, qtd, usuario, agora, observacao])

        ler_contagens.clear()
        return True
    except Exception as e:
        st.error(f"Erro ao salvar: {e}")
        return False


def remover_contagem(codigo: str):
    """Remove uma contagem do Google Sheets."""
    ws = get_sheet("contagens")
    if ws is None:
        return False
    try:
        cell = ws.find(codigo, in_column=1)
        ws.delete_rows(cell.row)
        ler_contagens.clear()
        return True
    except gspread.CellNotFound:
        return True
    except Exception as e:
        st.error(f"Erro ao remover: {e}")
        return False


def limpar_todas_contagens():
    """Remove todas as contagens (mantém cabeçalho)."""
    ws = get_sheet("contagens")
    if ws is None:
        return False
    try:
        ws.clear()
        ws.append_row(COLS_CONTAGENS)
        ler_contagens.clear()
        return True
    except Exception as e:
        st.error(f"Erro ao limpar: {e}")
        return False


# ── ITENS (plano de contagem) ─────────────────────────────────────────────────

def salvar_itens(itens: list):
    """Salva a lista de itens no Google Sheets (sobrescreve tudo)."""
    ws = get_sheet("itens")
    if ws is None:
        return False
    try:
        ws.clear()
        rows = [COLS_ITENS]
        for item in itens:
            rows.append([
                item.get('codigo',''),
                item.get('descricao',''),
                item.get('unidade',''),
                item.get('endereco',''),
                item.get('abc',''),
                item.get('frequencia',''),
                item.get('proxContagem',''),
                item.get('qtdSistema',0),
                item.get('custoUnit',0),
                item.get('status',''),
                item.get('prioridade',''),
            ])
        ws.update(rows)
        return True
    except Exception as e:
        st.error(f"Erro ao salvar itens: {e}")
        return False


@st.cache_data(ttl=300)
def ler_itens() -> pd.DataFrame:
    """Lê o plano de itens do Google Sheets. Cache de 5 min."""
    ws = get_sheet("itens")
    if ws is None:
        return pd.DataFrame(columns=COLS_ITENS)
    try:
        data = ws.get_all_records()
        if not data:
            return pd.DataFrame(columns=COLS_ITENS)
        df = pd.DataFrame(data)
        df['qtdSistema'] = pd.to_numeric(df['qtdSistema'], errors='coerce').fillna(0)
        df['custoUnit']  = pd.to_numeric(df['custoUnit'],  errors='coerce').fillna(0)
        return df
    except Exception as e:
        return pd.DataFrame(columns=COLS_ITENS)


def sheets_configurado() -> bool:
    """Verifica se o Google Sheets está configurado nos secrets."""
    try:
        return (
            "gcp_service_account" in st.secrets and
            "google_sheets" in st.secrets and
            "spreadsheet_id" in st.secrets["google_sheets"]
        )
    except Exception:
        return False
