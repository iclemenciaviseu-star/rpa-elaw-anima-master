# RPA eLaw Ânima — Complemento de Cadastro

Automação da tarefa "Complemento de cadastro — Escritório (Judicial)" no eLaw da Ânima.

## Como usar

1. Duplo-clique em `run.bat`
2. Na primeira vez: instala Python deps + Chromium do Playwright (~3 min)
3. Navegador abre automaticamente em `http://127.0.0.1:8765`
4. Siga o wizard de 5 passos:
   - **1. Base de casos** — upload da planilha
   - **2. Pasta de PDFs** — caminho da pasta com os documentos
   - **3. Credenciais** — usuário/senha + intervalo + opções
   - **4. Revisão** — confere os dados
   - **5. Dashboard** — execução em tempo real + relatório CSV ao fim

## Requisitos

- Windows 10/11
- Python 3.10+ (https://www.python.org/downloads/) com "Add Python to PATH" marcado

## Colunas esperadas na planilha

| Coluna | Exemplo |
|---|---|
| `(Processo) Número` | 1234567-89.2024.8.26.0100 |
| `Tipo de Ação` | Ação Indenizatória |
| `Procedimento` | Procedimento Comum Cível |
| `Fase processual` | Conhecimento |
| `Resumo da Ação` | (texto livre) |
| `Pedidos` | Danos Morais, Obrigação de Fazer |
| `Valor da Causa` | 10000.00 |
| `Prazo para Defesa` | 15/05/2026 |
| `Deseja solicitar subsídios?` | Sim ou Não |

## Estrutura

```
RPA_Elaw_Anima/
├── run.bat              ← inicia tudo
├── server.py            ← backend FastAPI
├── rpa.py               ← script Playwright
├── index.html           ← interface (wizard)
├── requirements.txt     ← deps Python
├── README.md
├── uploads/             ← planilhas enviadas
├── logs/                ← log de cada execução
└── .venv/               ← ambiente virtual (criado automaticamente)
```

## Comportamento

- **Subsídios** lê da coluna `Deseja solicitar subsídios?`. Se vazia ou ausente, usa "Não".
- **Preservação** (ligada por padrão): antes de escrever em qualquer campo, verifica se já está preenchido. Se sim, pula e loga `[SKIP]`. Pode desligar no passo 3.
- **Logs**: cada execução gera um arquivo em `logs/run_YYYYMMDD_HHMMSS_xxx.log`.

## Problemas comuns

**Python não encontrado:** reinstale marcando "Add Python to PATH".

**Falha ao instalar deps:** firewall/proxy. Apague `.venv\.installed` e tente de novo, ou configure proxy com `set HTTPS_PROXY=...` antes de rodar.

**Charmap codec error:** já tratado — o launcher força UTF-8 (`chcp 65001` + `PYTHONUTF8=1`).

**Reinstalar do zero:** apague a pasta `.venv` e rode `run.bat`.

## CLI direto (sem interface)

```bash
.venv\Scripts\python rpa.py config.json
```

com `config.json`:

```json
{
  "usuario": "...",
  "senha": "...",
  "planilha": "C:/...",
  "pasta_pdfs": "C:/...",
  "inicio": 0,
  "fim": null,
  "headless": false,
  "preservar_campos": true,
  "log_file": "logs/run.log"
}
```
