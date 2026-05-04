# Playbook RPA eLaw — Lições do Primeiro Projeto

Documento de referência para criar novos RPAs de tarefas no eLaw da Ânima (ou similares baseados em PrimeFaces/JSF).

---

## 1. Arquitetura final que funcionou

```
RPA_Elaw_Anima/
├── run.bat              ← launcher Windows (Python + venv + Chromium)
├── server.py            ← FastAPI local em 127.0.0.1:8765
├── rpa.py               ← script Playwright (lê config.json, emite JSON Lines)
├── index.html           ← wizard 5 passos (HTML/CSS/JS standalone)
├── requirements.txt     ← versões soltas (>=) por causa do Python 3.14
├── uploads/             ← planilhas enviadas
├── logs/                ← run_*.log + screenshots de erro
└── .venv/               ← criado automaticamente
```

**Comunicação:** server spawna `rpa.py` como subprocess. RPA emite JSON Lines no stdout, server reenvia via SSE pra interface mostrar em tempo real.

---

## 2. Wizard de 5 passos (interface)

1. **Base de casos** — upload de planilha (.xlsx/.xls/.csv)
2. **Pasta de PDFs** — caminho da pasta com os anexos
3. **Credenciais** — usuário/senha eLaw + intervalo de linhas + opções (headless, preservar campos)
4. **Revisão** — KPIs e dados pra confirmar
5. **Dashboard & Relatório** — execução em tempo real com:
   - 4 KPIs: Total / Sucesso / Erros / Tempo (cronômetro mm:ss)
   - Barra de progresso
   - **Tabela caso-a-caso** atualizando ao vivo (#, Processo, Status, Detalhe)
   - Console de log bruto (colapsável)
   - Botões: Parar / Exportar CSV / Executar novamente
   - Banner final colorido (verde/amarelo/vermelho)

---

## 3. Gotchas críticos do eLaw / PrimeFaces

### IDs auto-gerados são INSTÁVEIS
Os IDs `j_id_6v_2_18_2_f_5_2_1_2_1_input` mudam a cada deploy do eLaw. **NÃO use** esses IDs como seletor primário. Use:
- **`title="..."` no widget `.ui-selectonemenu`** (estável, descobri por análise do HTML)
- **`<span style="font-weight: bold">Label:</span>`** seguido de `<select>`/`<input>` na mesma `<tr>` (padrão consistente)

### Selects são HIDDEN (`aria-hidden="true"`)
PrimeFaces esconde o `<select>` real e mostra um widget custom. **Playwright `select_option()` falha** porque verifica visibilidade. Alternativas:
- **Simular clique** no `.ui-selectonemenu-label`, esperar `_panel` abrir, clicar no `<li>`
- **JS puro:** `sel.value = X; sel.dispatchEvent(new Event('change', {bubbles: true}))` — bypassa visibilidade

### Save sem navegação imediata
Após clicar Confirmar:
- A página pode redirecionar (navegação) → `Execution context was destroyed` (TRATAR COMO SUCESSO)
- Pode mostrar growl de sucesso/erro sem mudar URL
- Pode mostrar mensagens de validação inline

**Estratégia:** loop ativo de 90s verificando `URL mudou OU growl-error OU growl-success`.

### Validações vêm com prefixos técnicos
Mensagens são tipo `MSG: Tipo de ação obrigatório | ALERT: ... | ALERT: Erro Tipo de ação obrigatório`. Limpar prefixos `MSG:`, `ALERT:`, `Erro ` e deduplicar antes de mostrar pro usuário.

### Upload de arquivo é em "modo auto"
- Selecionar tipo de documento PRIMEIRO (combo `eFileTipoCombo`)
- Setar arquivo no input (`uploadGedEFile_input`)
- **Aguardar barra de progresso sumir** em `.ui-fileupload-files` (até 2 minutos)
- **Aguardar arquivo aparecer** na tabela `gedEFileDataTable`
- **Aguardar `networkidle`** (até 20s) — server processando
- Sem isso, o save reclama "[Anexos do processo] obrigatório"

### Busca de processo precisa de Enter
O input com placeholder "Pesquise por aqui!" só dispara busca com Enter — sem Enter, não aparece autocomplete.

### Tarefa correta no listing
Lista de tarefas tem múltiplos botões por linha (lupa, editar, **check**, calendário, info, X, ordenar, download). O botão certo pra abrir o formulário de uma tarefa é o **check (✓)**. Encontre por:
1. `title` contendo "confirm"
2. Ícone com classe `pi-check`/`fa-check`/`ui-icon-check`
3. Posição 3 dos botões da row

### Dropdowns AJAX cascateados
Selecionar Tipo de Ação → eLaw faz AJAX e carrega Procedimento. Esperar **1.5s + networkidle** entre selects que disparam cascade.

---

## 4. Gotchas técnicos do ambiente

### OneDrive corrompe arquivos durante edição
Editar arquivos na pasta sincronizada do OneDrive causa corrupção (bytes nulos, arquivos truncados). Soluções:
- **Pausar OneDrive** durante desenvolvimento
- Editar via bash com cat heredoc → /tmp → cp para destino
- **Sempre validar** após escrever: `python3 -c "import ast; ast.parse(open('rpa.py').read())"`

### Python 3.14 + pacotes binários
Pacotes com extensões C/Rust (pandas, pydantic) podem não ter wheels pra 3.14. Use `>=` em vez de `==` no `requirements.txt`.

### Charmap codec error no Windows
Console cmd usa cp1252 por padrão e quebra com emojis/acentos no `print()`. Soluções **acumulativas**:
1. `chcp 65001` no início do `run.bat`
2. `sys.stdout.reconfigure(encoding="utf-8", errors="replace")` no Python
3. `PYTHONIOENCODING=utf-8` e `PYTHONUTF8=1` no env do subprocess

### `wait_until="networkidle"` é exigente demais
Sites com polling/analytics nunca atingem networkidle. Use `wait_until="commit"` (mais leniente — só espera começar a navegação) e depois `wait_for_selector` no campo que você precisa.

### Login com `ERR_ABORTED`
Redirecionamentos podem causar `ERR_ABORTED`. Solução: **retry x3 com `wait_until="commit"`** + esperar pelo seletor do campo de login.

---

## 5. Padrões de código que valeram a pena

### Função `mensagem_erro_amigavel(exc)`
Converte exceptions técnicas em mensagens claras pra usuário não-técnico. Casos cobertos:
- `eLaw rejeitou o save: MSG: ...` → `Validacao eLaw: Tipo de ação obrigatório / ...`
- `PDF nao localizado` → `PDF nao encontrado na pasta - anexar manualmente`
- `Timeout` → `Tempo esgotado durante operacao`
- `ERR_CONNECTION` → `Falha de conexao com o eLaw`
- Genérico → primeira linha truncada em 180 chars

### Função `_find_field_by_span(page, label, kind)`
Encontra elemento editável pelo padrão eLaw: `<span style="font-weight:bold">Label:</span>` + próxima `<td>` com select/input/textarea/datepicker. Mais estável que IDs.

### `fill_select_robusto` / `fill_input_robusto`
Wrapper com 3 estratégias em cascata:
1. Click humano via `title=` do widget
2. Span-based via `_find_field_by_span`
3. Fallback por ID parcial (j_id_*)

### Espera ATIVA pós-save
Loop em JS de até 90s verificando 4 condições simultâneas (URL mudou / growl erro / growl sucesso / timeout) com checagem a cada 500ms. Trata `Execution context destroyed` como sucesso.

### Diagnóstico automático
- Dump do HTML da página em `logs/form_dump_*.html` no primeiro caso (pra debug futuro)
- Screenshot em `logs/erro_save_*.png` quando save falha
- Screenshot em `logs/falhou_busca_*.png` quando processo não encontrado
- Log estruturado em JSON Lines (interface) + texto humano (arquivo)

---

## 6. Configurações de timing que funcionaram

- `slow_mo=600` no Playwright (delay entre ações)
- Wait entre casos: **3-5s**
- Wait após cada select: **800ms** (1500ms se cascade AJAX)
- Wait após upload terminar: **networkidle 20s + 2s margem**
- Wait pré-save: **3s** pra UI assentar
- Espera ativa pós-save: **até 90s**

---

## 7. Como começar um novo RPA aproveitando esta base

1. **Copiar a pasta** `RPA_Elaw_Anima` pra `RPA_Elaw_<NomeDaTarefa>`
2. **Adaptar o que muda:**
   - URL da tarefa específica no `processar()`
   - Texto que identifica a tarefa na lista (atualmente: `"Complemento de cadastro - Escritório (Judicial)"`)
   - Campos do formulário em `preencher_formulario()` — usar `fill_select_robusto` e `fill_input_robusto`
   - Colunas esperadas em `server.py` na lista `required`
3. **Executar 1 caso** com `form_dump` ativo pra ver o HTML real e mapear os `title=` dos widgets
4. **Validar** que os 4 fluxos críticos funcionam: busca, abertura da tarefa, preenchimento, save
5. **Empacotar** mantendo `run.bat`, `requirements.txt`, e estrutura de pastas

---

## 8. Checklist de validação final

- [ ] Login com retry funciona (testar com rede instável)
- [ ] Busca de processo abre o resultado certo
- [ ] Botão correto da tarefa é clicado (não a lupa)
- [ ] **Todos os campos** do formulário preenchem (verificar visualmente)
- [ ] Upload aguarda barra de progresso sumir
- [ ] Save detecta erro e marca como falha (não como sucesso falso)
- [ ] PDF não encontrado dá mensagem clara
- [ ] Validações do eLaw aparecem limpas no Detalhe
- [ ] CSV final tem todas as colunas
- [ ] Dashboard atualiza ao vivo
