# 📦 Sistema de Gestão de Estoque — Lely Center Carambeí

App web com duas abas: **Análise** (gera Excel) + **Contagem** (registra contagens em tempo real).

**Sem Google Cloud. Sem banco de dados externo. Sem configuração de API.**

---

## 🚀 Como publicar no Streamlit Cloud (5 minutos)

### Passo 1 — Subir os arquivos no GitHub

No repositório `estoque-app`, substitua todos os arquivos pelos do ZIP:
- `app.py`
- `engine.py`
- `storage.py`
- `requirements.txt`
- `README.md`

> Apague o `sheets.py` antigo se existir.

### Passo 2 — Deploy

1. Acesse **share.streamlit.io**
2. **New app → From existing repo**
3. Selecione `estoque-app` → Branch `main` → Main file: `app.py`
4. **Deploy** — sem nenhum secret ou configuração extra

---

## 📱 Como usar

### Análise (uma vez por mês)
1. Aba **Análise & Relatórios**
2. Carregue os dois CSVs
3. Clique **Gerar Análise Completa**
4. Baixe o Excel — itens ficam disponíveis automaticamente na aba Contagem

### Contagem (almoxarife no celular)
1. Aba **Contagem de Estoque**
2. Selecione seu nome
3. **O que Contar Agora** mostra fila priorizada automaticamente:
   - 🔴 Abaixo do ponto de pedido
   - ⚠️ Acima do estoque máximo
   - Nunca contados
   - Acurácia ruim na última contagem
   - Fora do prazo ABC
4. Toque no item → digite a quantidade → **Salvar**

---

## ⚠️ Persistência dos dados

Os dados ficam em arquivos JSON no servidor Streamlit Cloud.
O plano gratuito pode hibernar apps sem uso por 7+ dias e limpar os arquivos.
**Exporte o CSV regularmente** pela aba Histórico como backup.

---

## 🗂️ Arquivos

```
estoque-app/
├── app.py         ← Interface (Análise + Contagem)
├── engine.py      ← Motor de cálculo
├── storage.py     ← Persistência local (sem dependências externas)
├── requirements.txt
└── README.md
```
