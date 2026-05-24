# 📦 Sistema de Gestão de Estoque — Lely Center Carambeí

App web com duas abas: **Análise** (gera Excel) + **Contagem** (registra contagens em tempo real via Google Sheets).

---

## 🚀 Como publicar no Streamlit Cloud

### Passo 1 — Subir os arquivos no GitHub

No seu repositório `estoque-app`, substitua os arquivos existentes por:
- `app.py`
- `engine.py`
- `sheets.py`
- `requirements.txt`
- `README.md`

### Passo 2 — Criar o Google Sheets

1. Acesse **sheets.google.com** e crie uma nova planilha
2. Nomeie como **"Estoque Lely Carambeí"**
3. Copie o **ID** da URL: `docs.google.com/spreadsheets/d/`**SEU_ID_AQUI**`/edit`
4. A planilha vai ter duas abas criadas automaticamente pelo app:
   - `contagens` — histórico de todas as contagens
   - `itens` — plano de contagem (preenchido quando você gera a análise)

### Passo 3 — Criar a Service Account no Google Cloud

1. Acesse **console.cloud.google.com**
2. Crie um projeto (ou use um existente)
3. Ative as APIs:
   - **Google Sheets API**
   - **Google Drive API**
4. Vá em **IAM & Admin → Service Accounts → Create Service Account**
5. Nome: `estoque-lely`
6. Clique em **Create and Continue** (sem precisar de roles especiais)
7. Em **Keys → Add Key → Create new key → JSON** — faça o download do arquivo

### Passo 4 — Compartilhar o Google Sheets com a Service Account

1. Abra o arquivo JSON baixado
2. Copie o campo `client_email` (ex: `estoque-lely@seu-projeto.iam.gserviceaccount.com`)
3. No Google Sheets, clique em **Compartilhar** e cole esse e-mail com permissão de **Editor**

### Passo 5 — Configurar os Secrets no Streamlit Cloud

1. No **share.streamlit.io**, vá no seu app → **Settings → Secrets**
2. Cole exatamente isto (substituindo pelos seus dados):

```toml
[google_sheets]
spreadsheet_id = "COLE_SEU_ID_AQUI"

[gcp_service_account]
type = "service_account"
project_id = "SEU_PROJETO"
private_key_id = "..."
private_key = "-----BEGIN RSA PRIVATE KEY-----\n...\n-----END RSA PRIVATE KEY-----\n"
client_email = "estoque-lely@seu-projeto.iam.gserviceaccount.com"
client_id = "..."
auth_uri = "https://accounts.google.com/o/oauth2/auth"
token_uri = "https://oauth2.googleapis.com/token"
auth_provider_x509_cert_url = "https://www.googleapis.com/oauth2/v1/certs"
client_x509_cert_url = "..."
```

> ⚠️ Copie os valores **exatamente** do arquivo JSON baixado no Passo 3.
> O campo `private_key` deve ter os `\n` preservados — copie com cuidado.

### Passo 6 — Fazer o deploy

1. No Streamlit Cloud → **New app → From existing repo**
2. Selecione `estoque-app` → Branch `main` → Main file: `app.py`
3. URL: `estoque-lely-center.streamlit.app` (ou o nome que escolher)
4. **Deploy** — aguarde 2-3 minutos

---

## 📱 Como usar o app de contagem

### Fluxo básico:

1. **Aba Análise** — carregue os CSVs de inventário e movimentações → clique **Gerar Análise Completa**
   - Isso processa tudo e **sincroniza automaticamente** os 817+ itens com o Google Sheets
   - Gera o Excel com as 6 abas

2. **Aba Contagem** — abra no celular do almoxarife
   - **O que Contar Agora**: lista priorizada automaticamente por:
     - 🔴 Itens abaixo do ponto de pedido (urgente)
     - ⚠️ Itens acima do estoque máximo
     - Itens nunca contados
     - Itens com acurácia ruim na última contagem
     - Frequência ABC (A=mensal, B=bimestral, C=trimestral)
   - Clique no item → digite a quantidade → **Salvar**
   - A acurácia aparece em tempo real antes de salvar

3. **Histórico** — veja todas as contagens, exporte CSV, filtre por classe/status/usuário

### Múltiplos usuários:
- Cada um acessa o link no próprio celular
- Seleciona o nome antes de contar
- As contagens são salvas no Google Sheets e ficam visíveis para todos em ~30 segundos

---

## 📊 Estrutura do Google Sheets

| Aba | Colunas | Atualizado quando |
|-----|---------|-------------------|
| `itens` | codigo, descricao, unidade, endereco, abc, frequencia, proxContagem, qtdSistema, custoUnit, status, prioridade | Ao gerar análise |
| `contagens` | codigo, qtd, usuario, data, observacao | A cada contagem registrada |

Você pode abrir o Google Sheets a qualquer momento para ver ou editar as contagens manualmente.

---

## 🔧 Arquivos do projeto

```
estoque-app/
├── app.py          ← Interface principal (Análise + Contagem)
├── engine.py       ← Motor de cálculo (CMM, ABC, parâmetros, Excel)
├── sheets.py       ← Integração Google Sheets (ler/salvar contagens)
├── requirements.txt
└── README.md
```

---

## ❓ Problemas comuns

**"Google Sheets não configurado"**
→ Verifique os Secrets no Streamlit Cloud — todos os campos são obrigatórios

**"Erro ao conectar ao Google Sheets"**
→ Confirme que a Service Account foi compartilhada como Editor na planilha

**A aba Contagem não mostra itens**
→ Gere a análise primeiro na aba Análise — isso sincroniza os itens com o Sheets

**Contagens não atualizam para outros usuários**
→ Clique em 🔄 Atualizar — o cache é de 30 segundos por design para não sobrecarregar a API
