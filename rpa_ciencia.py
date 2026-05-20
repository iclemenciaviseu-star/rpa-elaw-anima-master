"""
RPA eLaw Anima - Ciencia de Novo Processo
Fluxo: busca processo -> localiza tarefa -> clica confirmar -> pagina de confirmacao
       -> clica Confirmar novamente -> aguarda gravacao.
Planilha: qualquer planilha; o usuario indica qual coluna contem o numero do processo.
"""
from __future__ import annotations

import asyncio
import json
import os
import re
import sys
import traceback
from datetime import datetime
from pathlib import Path

if sys.platform == "win32":
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
        sys.stderr.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

import pandas as pd
from playwright.async_api import async_playwright, TimeoutError as PWTimeout


URL_BASE = "https://anima.elaw.com.br"
BASE = Path(__file__).resolve().parent
LOGS = BASE / "logs"
LOGS.mkdir(exist_ok=True)

# Textos possiveis da tarefa no eLaw (PT e EN, variantes do sistema)
TASK_TARGETS = [
    "Ciencia de novo processo",
    "Ciência de novo processo",
    "Escritorio Externo: Ciencia de novo processo",
    "Escritório Externo: Ciência de novo processo",
    "Ciencia",
    "novo processo",
]


class Logger:
    def __init__(self, log_file=None):
        self.log_file = log_file
        if log_file:
            Path(log_file).parent.mkdir(parents=True, exist_ok=True)

    def _emit(self, level, msg, **kw):
        payload = {"ts": datetime.now().isoformat(), "level": level, "msg": msg}
        payload.update(kw)
        line = json.dumps(payload, ensure_ascii=False)
        print(line, flush=True)
        if self.log_file:
            try:
                with open(self.log_file, "a", encoding="utf-8") as f:
                    f.write(line + "\n")
            except Exception:
                pass

    def info(self, msg, **kw):    self._emit("info",    msg, **kw)
    def success(self, msg, **kw): self._emit("success", msg, **kw)
    def warn(self, msg, **kw):    self._emit("warn",    msg, **kw)
    def error(self, msg, **kw):   self._emit("error",   msg, **kw)


def carregar_planilha(caminho, coluna_numero=None):
    """Carrega qualquer planilha. Se coluna_numero for informado, usa essa coluna
    como numero do processo; caso contrario tenta auto-detectar pelo nome padrao."""
    c = caminho.lower()
    if c.endswith(".xls"):
        df = pd.read_excel(caminho, engine="xlrd")
    elif c.endswith(".csv"):
        df = pd.read_csv(caminho)
    else:
        df = pd.read_excel(caminho)
    df.columns = df.columns.str.strip()

    TARGET = "(Processo) Número"

    if coluna_numero:
        # Busca a coluna indicada pelo usuario (comparacao exata primeiro, depois case-insensitive)
        match = next((x for x in df.columns if x.strip() == coluna_numero), None)
        if match is None:
            match = next((x for x in df.columns
                          if x.strip().lower() == coluna_numero.strip().lower()), None)
        if match is None:
            raise ValueError(
                f"Coluna '{coluna_numero}' nao encontrada na planilha. "
                f"Colunas disponiveis: {list(df.columns)}"
            )
        if match != TARGET:
            df = df.rename(columns={match: TARGET})
    else:
        # Auto-deteccao pelo nome padrao (com ou sem acento)
        match = next(
            (x for x in df.columns
             if x.strip().lower().replace("ú", "u") == "processo) numero"
             or "(processo) n" in x.lower()),
            None
        )
        if match and match != TARGET:
            df = df.rename(columns={match: TARGET})
        elif TARGET not in df.columns:
            raise ValueError(
                "Nao foi possivel detectar a coluna com o numero do processo. "
                "Selecione a coluna correta na interface."
            )

    df = df.dropna(subset=[TARGET])
    df = df[df[TARGET].astype(str).str.strip() != ""]
    df = df.reset_index(drop=True)
    return df


def mensagem_erro_amigavel(e):
    s = str(e)
    if "ERR_NAME_NOT_RESOLVED" in s or "net::ERR" in s:
        return "Sem conexao com o eLaw. Verifique a internet e tente novamente."
    if "TimeoutError" in type(e).__name__ or "Timeout" in s:
        if "wait_for_url" in s:
            return "eLaw nao redirecionou apos a acao (timeout). O botao foi clicado mas o sistema nao respondeu."
        if "wait_for_selector" in s or "wait_for" in s:
            return "Elemento nao apareceu na pagina dentro do tempo esperado."
        return "Operacao expirou (timeout). O eLaw pode estar lento."
    if "Login falhou" in s:
        return s
    if "MFA" in s:
        return s
    if "Nao foi possivel encontrar" in s or "Nao consegui localizar" in s:
        return s
    if "eLaw rejeitou o save" in s:
        return s
    if "Save aparentemente nao funcionou" in s:
        return s
    return f"Erro inesperado: {s[:200]}"


async def _abrir_processo_da_pagina(page, numero, log):
    """Tenta clicar no processo na pagina atual (resultados ou dashboard)."""
    prefixo = numero.split("-")[0].strip()
    estrategias = [
        '[href*="processoView"]',
        f'tr:has-text("{numero}") a',
        f'a:has-text("{numero}")',
        f'td:has-text("{numero}")',
        f'tr:has-text("{prefixo}") a',
        f'a:has-text("{prefixo}")',
    ]
    for sel in estrategias:
        try:
            loc = page.locator(sel).first
            n = await loc.count()
            if n == 0:
                continue
            await loc.scroll_into_view_if_needed()
            await loc.click(timeout=5000)
            await page.wait_for_url("**/processoView.elaw**", timeout=30000)
            log.info(f"Processo aberto (seletor: {sel[:60]})")
            return True
        except Exception:
            continue
    return False


async def processar(page, row, idx, total, log):
    numero = str(row["(Processo) Número"]).strip()
    log.info(f"--- [{idx+1}/{total}] {numero} ---", step="start_case", caso=numero)
    try:
        # ── 1. Navegar ao dashboard e buscar processo ──────────────────────
        await page.goto(f"{URL_BASE}/contenciosoDashboard.elaw", wait_until="domcontentloaded", timeout=45000)

        campo = page.locator('input[placeholder="Pesquise por aqui!"]')
        await campo.wait_for(state="visible", timeout=15000)
        await campo.click()
        await page.keyboard.press("Control+a")
        await page.keyboard.press("Delete")
        await page.wait_for_timeout(300)
        await campo.press_sequentially(numero, delay=50)
        await page.wait_for_timeout(1000)

        clicou = False

        # Estrategia 1: autocomplete
        try:
            ac = page.locator(".ui-autocomplete-panel li, .ui-autocomplete-items li").first
            await ac.wait_for(state="visible", timeout=10000)
            await ac.click()
            try:
                await page.wait_for_url("**/processoView.elaw**", timeout=30000)
            except Exception:
                pass
            if "processoView.elaw" in page.url:
                clicou = True
                log.info("Processo aberto via autocomplete")
        except Exception:
            pass

        # Estrategia 2: Enter
        if not clicou:
            await campo.focus()
            await campo.press("Enter")
            try:
                await page.wait_for_load_state("load", timeout=30000)
            except Exception:
                pass
            await page.wait_for_timeout(3000)

            if "processoView.elaw" in page.url:
                clicou = True
                log.info("Processo aberto via Enter (direto)")
            else:
                log.info(f"Pagina apos Enter: {page.url} — buscando processo na pagina")
                clicou = await _abrir_processo_da_pagina(page, numero, log)

        # Estrategia 3: busca na pagina atual
        if not clicou:
            log.warn(f"Estrategias 1 e 2 falharam — tentando localizar na pagina (URL: {page.url})")
            clicou = await _abrir_processo_da_pagina(page, numero, log)

        # Estrategia 4: aguarda navegacao tardia
        if not clicou:
            log.info("Aguardando possivel navegacao em andamento (15s)...")
            try:
                await page.wait_for_url("**/processoView.elaw**", timeout=15000)
                clicou = True
                log.info("Processo aberto (navegacao tardia detectada)")
            except Exception:
                pass

        if not clicou:
            try:
                safe = re.sub(r"[^0-9a-zA-Z]", "_", numero)
                shot = LOGS / f"falhou_busca_{safe}.png"
                await page.screenshot(path=str(shot), full_page=True)
                log.warn(f"Screenshot salvo em: {shot}")
                log.warn(f"URL no momento da falha: {page.url}")
            except Exception:
                pass
            raise Exception(f"Nao foi possivel encontrar/abrir o processo {numero} apos buscar")

        await page.wait_for_load_state("networkidle", timeout=30000)

        # ── 2. Aguarda lista de tarefas ────────────────────────────────────
        TASK_SELS = (
            "text=Ciencia de novo processo, "
            "text=Ciência de novo processo, "
            "text=Ciencia, "
            "text=Escritorio Externo"
        )
        try:
            await page.wait_for_selector(TASK_SELS, timeout=20000)
        except Exception:
            log.warn("Texto da tarefa Ciencia nao encontrado em 20s, prosseguindo mesmo assim...")

        # ── 3. Clica no botao de confirmacao da tarefa ─────────────────────
        confirmado = await page.evaluate("""
            (() => {
                const TARGETS = [
                    'Ciencia de novo processo',
                    'Ciência de novo processo',
                    'Escritório Externo: Ciência de novo processo',
                    'Escritorio Externo: Ciencia de novo processo',
                    'Ciencia',
                    'novo processo',
                ];
                let taskRow = null;
                for (const target of TARGETS) {
                    for (const tr of document.querySelectorAll('tr')) {
                        if (tr.innerText.includes(target)) { taskRow = tr; break; }
                    }
                    if (taskRow) break;
                }
                if (!taskRow) return 'no-row';

                const btns = Array.from(taskRow.querySelectorAll('button, a.ui-button'));

                // Estrategia 1: title contem "confirm"
                for (const b of btns) {
                    const t = (b.title || b.getAttribute('aria-label') || '').toLowerCase();
                    if (t.includes('confirm')) { b.click(); return 'title-confirm'; }
                }
                // Estrategia 2: icone check
                for (const b of btns) {
                    const chk = b.querySelector('.pi-check, .fa-check, .ui-icon-check, [class*="check"]');
                    if (chk) {
                        const cls = chk.className.toLowerCase();
                        if (cls.includes('check') && !cls.includes('checkbox') && !cls.includes('checklist')) {
                            b.click(); return 'icon-check';
                        }
                    }
                }
                // Estrategia 3: posicao 3 (0=lupa, 1=editar, 2=CHECK)
                const submits = btns.filter(b => b.type === 'submit' || b.tagName === 'BUTTON');
                if (submits.length >= 3) { submits[2].click(); return 'pos-3'; }

                return 'not-found';
            })()
        """)
        log.info(f"Click no botao da tarefa Ciencia: {confirmado}")

        if confirmado in (False, "no-row", "not-found"):
            # fallback: tenta por texto da tarefa
            for target in TASK_TARGETS:
                try:
                    await page.locator(
                        f'tr:has-text("{target}") button'
                    ).nth(2).click(timeout=5000)
                    log.info(f"Fallback: clicou via tr:has-text({target!r})")
                    confirmado = "fallback"
                    break
                except Exception:
                    continue

        if confirmado in (False, "no-row", "not-found"):
            try:
                safe = re.sub(r"[^0-9a-zA-Z]", "_", numero)
                shot = LOGS / f"falhou_botao_{safe}.png"
                await page.screenshot(path=str(shot), full_page=True)
                log.warn(f"Screenshot salvo em: {shot}")
            except Exception:
                pass
            raise Exception(f"Nao consegui localizar botao de Ciencia para {numero}")

        # ── 4. Aguarda pagina de confirmacao ──────────────────────────────
        await page.wait_for_url("**/agendamentoContenciosoConfirm.elaw**", timeout=30000)
        await page.wait_for_load_state("networkidle", timeout=30000)
        log.info("Pagina de confirmacao carregada")

        # ── 5. Clica Confirmar e aguarda gravacao ─────────────────────────
        await page.wait_for_timeout(2000)
        url_antes = page.url
        import time as _t

        for tentativa_save in range(1, 3):
            log.info(f"Confirmando (tentativa {tentativa_save}/2)...")
            btn = page.locator('button:has-text("Confirmar")').last
            await btn.scroll_into_view_if_needed()
            await btn.click()
            log.info("Aguardando resposta do sistema (ate 90s)...")

            t_inicio = _t.time()

            try:
                sinal = await page.evaluate(f"""
                    async () => {{
                        const urlAntes = {json.dumps(url_antes)};
                        const inicio = Date.now();
                        const limite = 90000;
                        while ((Date.now() - inicio) < limite) {{
                            if (window.location.href !== urlAntes) {{
                                return {{tipo: 'url-mudou', url: window.location.href}};
                            }}
                            const erros = document.querySelectorAll(
                                '.ui-growl-message-error, .ui-messages-error, ' +
                                '.ui-growl-message.ui-state-error'
                            );
                            for (const e of erros) {{
                                const t = (e.innerText || '').trim();
                                if (t && !e.closest('.ui-helper-hidden')) {{
                                    return {{tipo: 'erro', msg: t.slice(0, 400)}};
                                }}
                            }}
                            const sucessos = document.querySelectorAll(
                                '.ui-growl-message-info, .ui-growl-message:not(.ui-state-error)'
                            );
                            for (const s of sucessos) {{
                                const t = (s.innerText || '').toLowerCase();
                                if (t.includes('sucesso') || t.includes('salvo') ||
                                    t.includes('confirmad') || t.includes('conclu') ||
                                    t.includes('agendad') || t.includes('ciencia')) {{
                                    return {{tipo: 'sucesso', msg: t.slice(0, 200)}};
                                }}
                            }}
                            await new Promise(r => setTimeout(r, 500));
                        }}
                        return {{tipo: 'timeout', url: window.location.href}};
                    }}
                """)
            except Exception as e:
                err_msg = str(e)
                if "Execution context was destroyed" in err_msg or "navigation" in err_msg.lower():
                    sinal = {"tipo": "url-mudou-via-nav", "url": page.url}
                else:
                    raise

            elapsed = _t.time() - t_inicio
            log.info(f"Confirmacao: sinal={sinal.get('tipo')} em {elapsed:.1f}s - {sinal.get('msg', sinal.get('url', ''))[:200]}")

            try:
                await page.wait_for_load_state("networkidle", timeout=15000)
            except Exception:
                pass
            await page.wait_for_timeout(2000)

            # verifica erros na pagina
            erros = await page.evaluate("""
                (() => {
                    const msgs = [];
                    document.querySelectorAll(
                        '.ui-growl-message-error, .ui-growl-item-container.ui-state-error, .ui-growl-message.ui-message-error'
                    ).forEach(e => {
                        const t = (e.innerText || '').trim();
                        if (t) msgs.push('GROWL: ' + t);
                    });
                    document.querySelectorAll(
                        '.ui-messages-error, .ui-message-error, .ui-messages-error-detail'
                    ).forEach(e => {
                        const t = (e.innerText || '').trim();
                        if (t) msgs.push('MSG: ' + t);
                    });
                    document.querySelectorAll('[role="alert"]').forEach(e => {
                        const t = (e.innerText || '').trim();
                        if (t) msgs.push('ALERT: ' + t);
                    });
                    return msgs;
                })()
            """)

            if erros:
                try:
                    safe = re.sub(r"[^0-9a-zA-Z]", "_", numero)
                    shot = LOGS / f"erro_save_{safe}.png"
                    await page.screenshot(path=str(shot), full_page=True)
                    log.warn(f"Screenshot do erro: {shot}")
                except Exception:
                    pass
                msg_curta = " | ".join(erros)[:300]
                if tentativa_save < 2:
                    log.warn(f"Confirmacao tentativa {tentativa_save} falhou ({msg_curta[:80]}) — retry em 5s...")
                    await page.wait_for_timeout(5000)
                    continue
                raise Exception(f"eLaw rejeitou a confirmacao: {msg_curta}")

            if page.url == url_antes:
                sucesso = await page.evaluate("""
                    (() => {
                        const candidatos = document.querySelectorAll(
                            '.ui-growl-message-info, .ui-growl-message:not(.ui-growl-message-error):not(.ui-state-error), ' +
                            '.ui-messages-info, .ui-message-info, ' +
                            '[class*="success"], [class*="ok"]'
                        );
                        for (const c of candidatos) {
                            const t = (c.innerText || '').toLowerCase();
                            if (t.includes('sucesso') || t.includes('salvo') ||
                                t.includes('confirmad') || t.includes('conclu') ||
                                t.includes('ciencia')) {
                                return t.slice(0, 200);
                            }
                        }
                        return null;
                    })()
                """)
                if sucesso:
                    log.success(f"Confirmacao por growl: {sucesso}")
                    break
                else:
                    if tentativa_save < 2:
                        try:
                            safe = re.sub(r"[^0-9a-zA-Z]", "_", numero)
                            shot = LOGS / f"sem_navegacao_{safe}.png"
                            await page.screenshot(path=str(shot), full_page=True)
                            log.warn(f"Sem navegacao (tentativa {tentativa_save}). Screenshot: {shot} — retry em 5s...")
                        except Exception:
                            pass
                        await page.wait_for_timeout(5000)
                        continue
                    try:
                        safe = re.sub(r"[^0-9a-zA-Z]", "_", numero)
                        shot = LOGS / f"sem_navegacao_{safe}.png"
                        await page.screenshot(path=str(shot), full_page=True)
                        log.warn(f"Sem navegacao. Screenshot: {shot}")
                    except Exception:
                        pass
                    raise Exception(f"Confirmacao nao funcionou - pagina nao mudou de {url_antes}")
            else:
                break  # pagina navegou = sucesso

        log.success(f"Concluido: {numero}", step="case_done", caso=numero)
        return True

    except Exception as e:
        msg_clara = mensagem_erro_amigavel(e)
        log.error(f"ERRO em {numero}: {msg_clara}", step="case_error", caso=numero)
        try:
            if log.log_file:
                with open(log.log_file, "a", encoding="utf-8") as f:
                    f.write(f"[{datetime.now().strftime('%H:%M:%S')}] [DEBUG] traceback completo:\n")
                    f.write(traceback.format_exc() + "\n")
        except Exception:
            pass
        return False


async def _login_microsoft_sso(page, usuario, senha, log):
    """Completa o fluxo de login Microsoft SSO (SAML2) apos redirect do eLaw."""

    async def _screenshot(nome):
        try:
            p = LOGS / nome
            await page.screenshot(path=str(p), full_page=True)
            log.info(f"Screenshot SSO: {p.name} (URL: {page.url})")
        except Exception:
            pass

    MS_EMAIL = 'input[type="email"], input[name="loginfmt"], #i0116'
    MS_PASS  = 'input[type="password"], input[name="passwd"], #i0118'
    MS_NEXT  = '#idSIButton9, input[type="submit"], button[type="submit"]'

    await page.wait_for_load_state("domcontentloaded", timeout=15000)
    await _screenshot("sso_01_inicial.png")

    # pick-account
    PICK_TILE = '[data-test-id="tile"]'
    try:
        await page.wait_for_selector(PICK_TILE, timeout=4000)
        log.info("SSO: pick-account detectado, clicando primeiro tile")
        await page.locator(PICK_TILE).first.click()
        await page.wait_for_timeout(2000)
        await _screenshot("sso_02_pos_pick.png")
    except Exception:
        pass

    # email
    try:
        await page.wait_for_selector(MS_EMAIL, timeout=6000)
        val = await page.locator(MS_EMAIL).first.input_value()
        if not val:
            log.info("SSO: campo email vazio, preenchendo")
            await page.locator(MS_EMAIL).first.fill(usuario)
        else:
            log.info(f"SSO: email pre-preenchido pelo SAML ({val}), mantendo")
        await page.locator(MS_NEXT).first.click()
        await page.wait_for_timeout(2500)
        await _screenshot("sso_03_pos_email.png")
    except Exception:
        log.info("SSO: campo email nao encontrado, indo para senha")

    # senha
    await page.wait_for_selector(MS_PASS, timeout=20000)
    await _screenshot("sso_04_senha.png")
    await page.locator(MS_PASS).first.fill(senha)
    await page.locator(MS_NEXT).first.click()
    log.info("SSO: senha preenchida, aguardando redirect...")
    await page.wait_for_timeout(3000)
    await _screenshot("sso_05_pos_senha.png")

    try:
        titulo = await page.title()
        log.info(f"SSO pos-senha: titulo='{titulo}' url={page.url[:80]}")
    except Exception:
        pass

    # MFA
    MFA_SEL = (
        '#idDiv_SAOTCAS_Title, #idDiv_SAOTCC_Section, '
        '[data-bind*="phoneConfirmation"], '
        'div:has-text("Microsoft Authenticator"), '
        'div:has-text("Verificar sua identidade"), '
        'div:has-text("Verify your identity")'
    )
    try:
        await page.wait_for_selector(MFA_SEL, timeout=3000)
        titulo = await page.title()
        raise RuntimeError(
            f"MFA ativado na conta — autenticacao em duas etapas detectada "
            f"(titulo: '{titulo}'). Use uma conta sem MFA."
        )
    except RuntimeError:
        raise
    except Exception:
        pass

    # KMSI
    KMSI_SEL = '#idSIButton9, #idBtn_Back, button:has-text("Sim"), button:has-text("Yes"), button:has-text("Nao"), button:has-text("No"), button:has-text("Não")'
    try:
        await page.wait_for_selector(KMSI_SEL, timeout=10000)
        titulo = await page.title()
        log.info(f"SSO: prompt pos-senha detectado (titulo: '{titulo}')")
        nao = page.locator('#idBtn_Back, button:has-text("Nao"), button:has-text("No"), button:has-text("Não")')
        if await nao.count() > 0:
            await nao.first.click()
            log.info("SSO: clicou Nao no prompt KMSI")
        else:
            await page.locator(KMSI_SEL).first.click()
            log.info("SSO: clicou botao no prompt pos-senha")
    except Exception:
        pass

    await page.wait_for_url("**/homePage.elaw", timeout=60000)
    log.success("Login SSO OK", step="login_ok")


async def run(cfg):
    log = Logger(cfg.get("log_file"))
    coluna_numero = cfg.get("coluna_numero") or None
    df = carregar_planilha(cfg["planilha"], coluna_numero)
    total = len(df)
    inicio = int(cfg.get("inicio") or 0)
    fim = int(cfg["fim"]) if cfg.get("fim") else total

    sucesso = erro = 0
    erros = []

    log.info("=" * 55, step="banner")
    log.info(f"RPA eLaw Anima - Ciencia de Novo Processo - {datetime.now().strftime('%d/%m/%Y %H:%M')}", step="banner")
    log.info(f"Linhas {inicio+1} -> {fim} (total: {total})", step="banner")
    log.info("=" * 55, step="banner")

    async with async_playwright() as p:
        browser = await p.chromium.launch(
            headless=bool(cfg.get("headless", False)),
            slow_mo=int(cfg.get("slow_mo", 600)),
            args=["--start-maximized"],
        )
        auth_path = BASE / "auth_state.json"
        auth_env = os.environ.get("AUTH_STATE_B64", "")
        storage_state = None

        if auth_env:
            import base64, json as _json, tempfile
            tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".json")
            tmp.write(base64.b64decode(auth_env))
            tmp.close()
            storage_state = tmp.name
            log.info("Auth state carregado da variavel de ambiente AUTH_STATE_B64")
        elif auth_path.exists():
            storage_state = str(auth_path)
            log.info(f"Auth state carregado de {auth_path.name}")

        context = await browser.new_context(
            viewport={"width": 1280, "height": 900},
            storage_state=storage_state,
        )
        page = await context.new_page()
        page.set_default_timeout(15000)

        log.info("Acessando eLaw...", step="login")
        await page.goto(f"{URL_BASE}/", wait_until="commit", timeout=45000)
        log.info(f"Pagina carregada: {page.url}")

        if storage_state:
            try:
                await page.wait_for_url("**/homePage.elaw", timeout=12000)
                log.success("Login via auth_state OK", step="login_ok")
            except Exception:
                log.warn("Auth state expirado ou invalido, tentando login normal...")
                storage_state = None

        if not storage_state or "homePage" not in page.url:
            for tentativa in range(1, 4):
                try:
                    if tentativa > 1:
                        await page.goto(f"{URL_BASE}/", wait_until="commit", timeout=45000)
                    await page.wait_for_selector("#username, #authKey", timeout=20000)
                    break
                except Exception as e:
                    if tentativa >= 3:
                        raise
                    log.warn(f"Tentativa {tentativa} falhou, retry em 3s... ({str(e)[:80]})")
                    await page.wait_for_timeout(3000)

            await page.fill("#username", cfg["usuario"])
            await page.fill("#authKey",  cfg["senha"])
            log.info("Campos #username / #authKey preenchidos")

            await page.evaluate("""() => {
                for (const f of document.querySelectorAll('form')) {
                    if (f.querySelector('#authKey, input[type="password"]')) {
                        const btn = f.querySelector('button.ui-button, button[type="submit"], button');
                        if (btn) { btn.setAttribute('data-rpa-login', 'true'); break; }
                    }
                }
            }""")
            await page.locator('[data-rpa-login="true"]').click(force=True, timeout=10000)
            log.info("Clique login: botao do form de credenciais")

            log.info("Aguardando autenticacao (networkidle, ate 30s)...")
            try:
                await page.wait_for_load_state("networkidle", timeout=30000)
            except Exception:
                pass
            try:
                await page.wait_for_load_state("load", timeout=15000)
            except Exception:
                pass

            cur_url = page.url
            log.info(f"URL pos-autenticacao: {cur_url}")

            if "homePage" in cur_url:
                log.success("Login OK", step="login_ok")
            elif "microsoftonline.com" in cur_url or "login.microsoft" in cur_url:
                log.info("Redirecionado para Microsoft SSO, preenchendo credenciais...")
                ms_email = cfg.get("microsoft_email") or cfg["usuario"]
                await _login_microsoft_sso(page, ms_email, cfg["senha"], log)
            else:
                err_msg = None
                try:
                    err_msg = await page.evaluate("""() => {
                        const sels = [
                            '.ui-messages-error', '.ui-message-error',
                            '.ui-growl-message-error', '[class*="error-msg"]',
                            '[id*="msgError"]', '.alert-danger', '.alert-error'
                        ];
                        for (const s of sels) {
                            for (const el of document.querySelectorAll(s)) {
                                const t = (el.innerText || '').trim();
                                if (t && t.length < 300) return t;
                            }
                        }
                        return null;
                    }""")
                except Exception:
                    pass
                try:
                    shot = LOGS / "login_falhou_final.png"
                    await page.screenshot(path=str(shot), full_page=True)
                    log.warn(f"Screenshot: {shot.name}")
                except Exception:
                    pass
                msg = err_msg or f"URL final: {cur_url}"
                raise Exception(f"Login falhou: {msg}")

        total_casos = fim - inicio
        for idx, row in df.iloc[inicio:fim].iterrows():
            ok = await processar(page, row, idx, total, log)
            if ok:
                sucesso += 1
            else:
                erro += 1
                erros.append(str(row["(Processo) Número"]).strip())
            feitos = sucesso + erro
            log.info(
                f"Progresso: {feitos}/{total_casos}",
                step="progress",
                done=feitos, total=total_casos,
                sucesso=sucesso, erro=erro,
            )
            await page.wait_for_timeout(5000)

        await browser.close()

    log.info("=" * 55, step="done")
    log.info(f"Sucesso: {sucesso} | Erros: {erro}",
             step="summary", sucesso=sucesso, erro=erro, erros=erros)
    if erros:
        log.warn("Processos com erro:")
        for e in erros:
            log.warn(f"  - {e}")


def main():
    if len(sys.argv) < 2:
        print("Uso: python rpa_ciencia.py config.json", file=sys.stderr)
        sys.exit(1)
    with open(sys.argv[1], "r", encoding="utf-8") as f:
        cfg = json.load(f)
    try:
        asyncio.run(run(cfg))
    except KeyboardInterrupt:
        print(json.dumps({"ts": datetime.now().isoformat(), "level": "warn",
                          "msg": "Interrompido", "step": "aborted"},
                         ensure_ascii=False), flush=True)
        sys.exit(130)
    except Exception as e:
        print(json.dumps({"ts": datetime.now().isoformat(), "level": "error",
                          "msg": f"Falha fatal: {e}", "step": "fatal",
                          "trace": traceback.format_exc()},
                         ensure_ascii=False), flush=True)
        sys.exit(1)


if __name__ == "__main__":
    main()
