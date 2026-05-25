"""
storage.py — Persistência simples via arquivo JSON no servidor Streamlit.
Sem Google Cloud, sem banco de dados, sem configuração extra.
Funciona para múltiplos usuários em tempo real no Streamlit Cloud.
"""
import streamlit as st
import pandas as pd
import json
import os
from datetime import datetime
from pathlib import Path
import threading

# Arquivo de dados persistente no servidor
DATA_DIR  = Path("/tmp/estoque_data")
CONT_FILE = DATA_DIR / "contagens.json"
ITENS_FILE= DATA_DIR / "itens.json"

_lock = threading.Lock()

def _ensure_dir():
    DATA_DIR.mkdir(parents=True, exist_ok=True)

def _ler_json(path: Path) -> list:
    try:
        if path.exists():
            with open(path, 'r', encoding='utf-8') as f:
                return json.load(f)
    except Exception:
        pass
    return []

def _salvar_json(path: Path, data: list):
    _ensure_dir()
    tmp = path.with_suffix('.tmp')
    with open(tmp, 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, indent=2, default=str)
    tmp.replace(path)  # atomic write


# ── CONTAGENS ─────────────────────────────────────────────────────────────────

def ler_contagens() -> pd.DataFrame:
    """Lê todas as contagens salvas."""
    cols = ["codigo", "qtd", "usuario", "data", "observacao"]
    data = _ler_json(CONT_FILE)
    if not data:
        return pd.DataFrame(columns=cols)
    df = pd.DataFrame(data)
    df['qtd'] = pd.to_numeric(df.get('qtd', 0), errors='coerce')
    # Mantém só a última por código
    df = df.sort_values('data').drop_duplicates(subset=['codigo'], keep='last')
    return df.reset_index(drop=True)


def salvar_contagem(codigo: str, qtd: float, usuario: str, observacao: str = "") -> bool:
    """Salva ou atualiza uma contagem."""
    with _lock:
        data = _ler_json(CONT_FILE)
        # Remove entrada antiga do mesmo código
        data = [r for r in data if r.get('codigo') != codigo]
        data.append({
            "codigo":     codigo,
            "qtd":        qtd,
            "usuario":    usuario,
            "data":       datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "observacao": observacao,
        })
        _salvar_json(CONT_FILE, data)
    return True


def remover_contagem(codigo: str) -> bool:
    """Remove a contagem de um item."""
    with _lock:
        data = _ler_json(CONT_FILE)
        data = [r for r in data if r.get('codigo') != codigo]
        _salvar_json(CONT_FILE, data)
    return True


def limpar_todas_contagens() -> bool:
    """Remove todas as contagens."""
    with _lock:
        _salvar_json(CONT_FILE, [])
    return True


# ── ITENS ─────────────────────────────────────────────────────────────────────

def salvar_itens(itens: list) -> bool:
    """Salva o plano de itens (sobrescreve tudo)."""
    with _lock:
        _salvar_json(ITENS_FILE, itens)
    return True


def ler_itens() -> pd.DataFrame:
    """Lê o plano de itens."""
    cols = ["codigo","descricao","unidade","endereco","abc",
            "frequencia","proxContagem","qtdSistema","custoUnit","status","prioridade"]
    data = _ler_json(ITENS_FILE)
    if not data:
        return pd.DataFrame(columns=cols)
    df = pd.DataFrame(data)
    df['qtdSistema'] = pd.to_numeric(df.get('qtdSistema', 0), errors='coerce').fillna(0)
    df['custoUnit']  = pd.to_numeric(df.get('custoUnit',  0), errors='coerce').fillna(0)
    return df


def storage_configurado() -> bool:
    """Sempre disponível — não precisa de configuração."""
    return True


def status_storage() -> dict:
    """Retorna stats do storage para debug."""
    cont = _ler_json(CONT_FILE)
    itens= _ler_json(ITENS_FILE)
    return {
        "contagens": len(cont),
        "itens":     len(itens),
        "cont_file": str(CONT_FILE),
        "itens_file":str(ITENS_FILE),
    }


# ── IMPORTAÇÃO DE HISTÓRICO ───────────────────────────────────────────────────

def importar_contagens_csv(df: "pd.DataFrame") -> tuple:
    """
    Importa contagens de um DataFrame CSV.
    Colunas esperadas: codigo, qtd, usuario, data, observacao
    Retorna (importados, ignorados).
    """
    import pandas as pd
    cols_req = {'codigo', 'qtd'}
    if not cols_req.issubset(set(df.columns)):
        return 0, len(df)

    with _lock:
        existentes = _ler_json(CONT_FILE)
        # índice por código para update
        idx = {r['codigo']: i for i, r in enumerate(existentes)}

        importados = 0
        ignorados  = 0
        for _, row in df.iterrows():
            codigo = str(row.get('codigo', '')).strip()
            qtd    = pd.to_numeric(row.get('qtd', None), errors='coerce')
            if not codigo or pd.isna(qtd):
                ignorados += 1
                continue
            registro = {
                "codigo":     codigo,
                "qtd":        float(qtd),
                "usuario":    str(row.get('usuario', 'Importado')),
                "data":       str(row.get('data', '')),
                "observacao": str(row.get('observacao', '') or ''),
            }
            if codigo in idx:
                existentes[idx[codigo]] = registro
            else:
                existentes.append(registro)
                idx[codigo] = len(existentes) - 1
            importados += 1

        _salvar_json(CONT_FILE, existentes)
    return importados, ignorados
