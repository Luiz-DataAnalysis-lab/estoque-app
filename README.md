# 📦 Sistema de Gestão de Estoque

App web local para análise automática de estoque a partir dos CSVs do sistema.

---

## 🚀 Como instalar e rodar (só precisa fazer isso uma vez)

### Pré-requisito: Python instalado
Se não tiver Python, baixe em: https://www.python.org/downloads/
(marque a opção "Add Python to PATH" durante a instalação)

---

### Passo 1 — Abrir o terminal na pasta do app

**Windows:**
- Abra a pasta `estoque_app` no explorador de arquivos
- Clique na barra de endereço, digite `cmd` e pressione Enter

**Mac:**
- Abra o Terminal
- Digite `cd ` (com espaço) e arraste a pasta `estoque_app` para o terminal

---

### Passo 2 — Instalar as dependências (só na primeira vez)

```
pip install -r requirements.txt
```

Aguarde a instalação terminar (1-2 minutos).

---

### Passo 3 — Rodar o app

```
streamlit run app.py
```

O navegador abrirá automaticamente em `http://localhost:8501`

---

## 📋 Como usar

1. **Acesse** `http://localhost:8501` no navegador
2. **Ajuste os parâmetros** na barra lateral esquerda (Lead Time, Custo de Pedido, Taxa de Carregamento)
3. **Carregue os dois CSVs:**
   - `Inventario.csv` — exportação do saldo atual do sistema
   - Movimentações CSV — saídas e entradas do período
4. **Dê um nome** para a análise (ex: "Análise Mai/2026")
5. **Clique em "Gerar Análise Completa"**
6. **Confira os resultados** na tela e **baixe o Excel** completo

---

## 📁 Arquivos do sistema

```
estoque_app/
├── app.py              ← Interface (não editar)
├── engine.py           ← Motor de cálculo (não editar)
├── requirements.txt    ← Dependências Python
├── historico.json      ← Criado automaticamente com o histórico de análises
└── README.md           ← Este arquivo
```

---

## ❓ Problemas comuns

**"streamlit não é reconhecido"**
→ Feche e reabra o terminal após instalar o requirements.txt

**"Erro ao processar"**
→ Verifique se os CSVs são exportados com encoding Latin1 e separador ponto-e-vírgula ou vírgula

**O navegador não abriu automaticamente**
→ Abra manualmente: `http://localhost:8501`

---

## 🔄 Atualizar os parâmetros futuramente

Os parâmetros (Lead Time, Custo de Pedido, Taxa) podem ser ajustados na barra lateral a qualquer momento — não é preciso reinstalar nada.

---

*Desenvolvido para análise de estoque com classificação ABC, parâmetros de reposição (ES, PP, LEC, Est. Máximo) e plano de contagem cíclica com acurácia automática.*
