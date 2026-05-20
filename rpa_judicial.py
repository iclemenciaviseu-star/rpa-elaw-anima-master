"""
RPA eLaw Anima - Complemento de Cadastro - Escritório (Judicial)
Modulo gerado a partir de rpa.py (copia identica com descricao atualizada).
"""
from __future__ import annotations

import asyncio
import glob
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

MAPA_PEDIDOS = {
    "Danos Morais":       "Dano Moral - Cível",
    "Danos Materiais":    "Dano Material - Cível",
    "Obrigação de Fazer": "Obrigação de Fazer - Cível",
    "Lucros Cessantes":   "Lucros Cessantes - Cível",
}

MAPA_TIPO_ACAO = {
    "Ação Indenizatória ":                          "Ação Indenizatória",
    "Ação Indenizatória":                           "Ação Indenizatória",
    "Ação Declaratória":                            "Ação Declaratória",
    "Ação Monitória ":                              "Ação Monitória",
    "Ação Monitória":                               "Ação Monitória",
    "Ação Revisional de Contrato":                  "Ação Revisional de Contrato",
    "Ação de Revisão de Contrato":                  "Ação Revisional de Contrato",
    "Ação de Obrigação de Fazer":                   "Ação de Obrigação de Fazer",
    "Mandado de Segurança ":                        "Mandado de Segurança",
    "Mandado de Segurança":                         "Mandado de Segurança",
    "Ação De Obrigação De Fazer c/c Indenizatória": "Ação de Obrigação de Fazer c/c Indenizatória",
    "Reclamação Pré-Processual ":                   "Reclamação Pré-Processual",
    "Reclamação Pré-Processual":                    "Reclamação Pré-Processual",
}


class Logger:
    def __init__(self, log_file=None):
        self.log_file = log_file
        if log_file:
            Path(log_file).parent.mkdir(parents=True, exist_ok=True)

    def _emit(self, level, msg, step=None, **extra):
        payload = {"ts": datetime.now().isoformat(), "level": level, "msg": msg}
        if step:
            payload["step"] = step
        payload.update(extra)
        try:
            print(json.dumps(payload, ensure_ascii=False), flush=True)
        except Exception:
            print(json.dumps(payload, ensure_ascii=True), flush=True)
        if self.log_file:
            tags = {"info": "[INFO]", "warn": "[WARN]", "error": "[ERR ]", "success": "[ OK ]"}
            tag = tags.get(level, "[INFO]")
            line = f"[{datetime.now().strftime('%H:%M:%S')}] {tag} {msg}"
            try:
                with open(self.log_file, "a", encoding="utf-8") as f:
                    f.write(line + "\n")
            except Exception:
                pass

    def info(self, msg, **kw):    self._emit("info", msg, **kw)
    def warn(self, msg, **kw):    self._emit("warn", msg, **kw)
    def error(self, msg, **kw):   self._emit("error", msg, **kw)
    def success(self, msg, **kw): self._emit("success", msg, **kw)


def carregar_planilha(caminho):
    c = caminho.lower()
    if c.endswith(".xls"):
        df = pd.read_excel(caminho, engine="xlrd")
    elif c.endswith(".csv"):
        df = pd.read_csv(caminho)
    else:
        df = pd.read_excel(caminho)
    df.columns = df.columns.str.strip()
    return df


def encontrar_arquivo(numero, pasta):
    n = numero.strip()
    for ext in ("", ".pdf", ".PDF", ".zip", ".ZIP"):
        p = os.path.join(pasta, n + ext)
        if os.path.exists(p):
            return p
    n_clean = re.sub(r"[.\-/]", "", n)
    for arq in glob.glob(os.path.join(pasta, "**", "*"), recursive=True):
        nome = os.path.basename(arq)
        nome_clean = re.sub(r"[.\-/\s_]", "", nome)
        if n_clean in nome_clean or n in nome:
            return arq
    return None


def normalizar_subsidios(valor):
    if pd.isna(valor):
        return "Não"
    s = str(valor).strip().lower()
    if s in ("sim", "s", "yes", "y", "true", "1"):
        return "Sim"
    return "Não"

def mensagem_erro_amigavel(exc):
    """Converte exception tecnica em mensagem clara pro usuario.
    Detecta validacao do eLaw, timeouts, PDF faltando, etc.
    """
    msg = str(exc)

    # Validacao do eLaw - extrai so a mensagem real do PrimeFaces
    if "eLaw rejeitou o save" in msg:
        # tira prefixos tecnicos: MSG:, ALERT:, "Erro " no inicio
        clean = msg.replace("eLaw rejeitou o save:", "").strip()
        # remove prefixos MSG: / ALERT: / Erro
        clean = re.sub(r'(MSG:|ALERT:)\s*', '', clean)
        clean = re.sub(r'^Erro\s+', '', clean)
        # divide pelos separadores
        partes = [p.strip().lstrip('Erro ').strip()
                  for p in clean.split('|') if p.strip()]
        # remove duplicatas mantendo ordem
        unicos = []
        for p in partes:
            p2 = p.replace('Erro ', '').strip()
            if p2 and p2 not in unicos:
                unicos.append(p2)
        if unicos:
            return "Validacao eLaw: " + " / ".join(unicos[:4])
        return "Validacao eLaw rejeitou o save"

    # PDF nao localizado - ja claro
    if "PDF nao localizado" in msg:
        return "PDF nao encontrado na pasta - anexar manualmente"

    # Botao Confirmar nao localizado
    if "Confirmar" in msg and ("nao localizar" in msg.lower() or "nao consegui" in msg.lower()):
        return "Botao 'Confirmar' da tarefa nao foi encontrado"

    # Falha em encontrar o processo apos buscar
    if "Nao foi possivel encontrar/abrir o processo" in msg:
        return "Processo nao foi encontrado no eLaw"

    # Timeout/wait_for_selector
    if "Timeout" in msg and "wait_for_selector" in msg:
        return "Tempo esgotado esperando elemento na pagina"
    if "TimeoutError" in msg or "Timeout" in msg:
        return "Tempo esgotado - eLaw demorou para responder"

    # NetworkError / conexao
    if "ERR_CONNECTION" in msg or "net::ERR" in msg:
        return "Falha de conexao com o eLaw"

    # Save aparentemente nao funcionou
    if "Save aparentemente nao funcionou" in msg:
        return "Save nao confirmou - pagina nao mudou"

    # Generico - primeira linha, ate 180 chars
    primeira = msg.splitlines()[0]
    if len(primeira) > 180:
        primeira = primeira[:177] + "..."
    return primeira



async def select_filled(page, name_fragment):
    return await page.evaluate(f"""
        (() => {{
            const s = document.querySelector('select[name*="{name_fragment}"], select[id*="{name_fragment}"]');
            if (!s) return false;
            const v = (s.value || '').trim();
            if (!v || v === '0' || v === '-1') return false;
            const w = s.closest('.ui-selectonemenu');
            if (w) {{
                const lbl = w.querySelector('.ui-selectonemenu-label');
                if (lbl) {{
                    const t = (lbl.textContent || '').trim().toLowerCase();
                    if (!t || t.startsWith('selecione') || t === '-') return false;
                }}
            }}
            return true;
        }})()
    """)


async def input_filled(page, selector):
    sel = selector.replace("\\", "\\\\").replace("'", "\\'")
    return await page.evaluate(f"""
        (() => {{
            const el = document.querySelector('{sel}');
            if (!el) return false;
            return (el.value || '').trim().length > 0;
        }})()
    """)


async def pedidos_filled(page):
    return await page.evaluate("""
        (() => {
            const tables = document.querySelectorAll('table[id*="objetoList"], table[id*="ObjetoList"], table[id*="objetoTable"]');
            for (const t of tables) {
                const tbody = t.querySelector('tbody');
                if (!tbody) continue;
                for (const r of tbody.querySelectorAll('tr')) {
                    if (r.classList.contains('ui-datatable-empty-message')) continue;
                    if ((r.textContent || '').trim()) return true;
                }
            }
            return false;
        })()
    """)


async def upload_filled(page):
    return await page.evaluate("""
        (() => {
            const tbody = document.querySelector('#j_id_6v_2_18_2_l_5_5d_1\\\\:gedEFileDataTable tbody');
            if (!tbody) return false;
            for (const r of tbody.querySelectorAll('tr')) {
                if (r.classList.contains('ui-datatable-empty-message')) continue;
                if ((r.textContent || '').trim()) return true;
            }
            return false;
        })()
    """)


async def selecionar_por_label(page, label_text, valor, log):
    """Seleciona valor num PrimeFaces selectonemenu pelo title= do widget."""
    ok = await page.evaluate(f"""
        (() => {{
            const target = {json.dumps(label_text.lower().strip())};
            // estrategia 1: title= no widget .ui-selectonemenu (mais estavel)
            const widgets = document.querySelectorAll('.ui-selectonemenu[title]');
            let sel = null;
            // normaliza removendo acentos pra comparar
            const norm = s => s.normalize('NFD').replace(/[\u0300-\u036f]/g, '').toLowerCase().trim();
            const targetN = norm(target);
            for (const w of widgets) {{
                const t = (w.getAttribute('title') || '').trim().toLowerCase();
                const tN = norm(t);
                if (t === target || tN === targetN || t.startsWith(target) || target.startsWith(t) || tN.startsWith(targetN) || targetN.startsWith(tN)) {{
                    sel = w.querySelector('select');
                    if (sel) break;
                }}
            }}
            // estrategia 2: span/label com texto
            if (!sel) {{
                const labels = Array.from(document.querySelectorAll(
                    'label, .ui-outputlabel, span'
                )).filter(e => e.children.length === 0);
                const lbl = labels.find(l => {{
                    const t = (l.textContent || '').trim().toLowerCase().replace(/[:\\s*()]+$/, '');
                    return t === target || t.startsWith(target);
                }});
                if (lbl) {{
                    let p = lbl.parentElement;
                    for (let i = 0; i < 8 && p && !sel; i++) {{
                        sel = p.querySelector('select');
                        if (!sel) p = p.parentElement;
                    }}
                }}
            }}
            if (!sel) return 'no-select';

            const wanted = {json.dumps(valor.lower().strip())};
            let opt = null;
            for (const o of sel.options) {{
                if ((o.textContent || '').trim().toLowerCase() === wanted) {{
                    opt = o; break;
                }}
            }}
            if (!opt) {{
                for (const o of sel.options) {{
                    if ((o.textContent || '').trim().toLowerCase().includes(wanted)) {{
                        opt = o; break;
                    }}
                }}
            }}
            if (!opt) {{
                const opts = Array.from(sel.options).map(o => o.textContent.trim()).slice(0, 8).join(' | ');
                return 'no-option:' + opts;
            }}
            sel.value = opt.value;
            sel.dispatchEvent(new Event('change', {{bubbles: true}}));
            const w = sel.closest('.ui-selectonemenu');
            if (w) {{
                const lblEl = w.querySelector('.ui-selectonemenu-label');
                if (lblEl) lblEl.textContent = opt.textContent;
            }}
            return 'ok';
        }})()
    """)
    if ok != 'ok':
        log.warn(f"selecionar_por_label '{label_text}' = '{valor}': {ok}")
        return False
    return True


async def _OLD_selecionar_por_label(page, label_text, valor, log):
    """Versao antiga - mantida por compat, nao usada."""
    ok = await page.evaluate(f"""
        (() => {{
            const target = {json.dumps(label_text.lower())};
            const labels = Array.from(document.querySelectorAll(
                'label, .ui-outputlabel, .ui-panelgrid-cell, .form-label, td, span, div'
            )).filter(e => e.children.length === 0 || e.tagName === 'LABEL' || e.classList.contains('ui-outputlabel'));
            const lbl = labels.find(l => {{
                const t = (l.textContent || '').trim().toLowerCase().replace(/[:\\s]+$/, '');
                return t === target || t.startsWith(target);
            }});
            if (!lbl) return 'no-label';
            // procura select associado: via for, ou irmao no DOM proximo
            let sel = null;
            const fid = lbl.getAttribute('for');
            if (fid) {{
                const el = document.getElementById(fid);
                if (el && el.tagName === 'SELECT') sel = el;
                else if (el) {{
                    const w = el.closest('.ui-selectonemenu');
                    if (w) sel = w.querySelector('select');
                }}
            }}
            if (!sel) {{
                let p = lbl.parentElement;
                for (let i = 0; i < 6 && p && !sel; i++) {{
                    sel = p.querySelector('select');
                    if (!sel) p = p.parentElement;
                }}
            }}
            if (!sel) return 'no-select';
            // procura option com label igual
            const wantedLow = {json.dumps(valor.lower().strip())};
            let chosen = null;
            for (const opt of sel.options) {{
                const t = (opt.textContent || '').trim().toLowerCase();
                if (t === wantedLow) {{ chosen = opt; break; }}
            }}
            if (!chosen) {{
                for (const opt of sel.options) {{
                    if ((opt.textContent || '').trim().toLowerCase().includes(wantedLow)) {{
                        chosen = opt; break;
                    }}
                }}
            }}
            if (!chosen) return 'no-option';
            sel.value = chosen.value;
            sel.dispatchEvent(new Event('change', {{bubbles: true}}));
            // atualiza visual do PrimeFaces
            const w = sel.closest('.ui-selectonemenu');
            if (w) {{
                const lblEl = w.querySelector('.ui-selectonemenu-label');
                if (lblEl) lblEl.textContent = chosen.textContent;
            }}
            return 'ok';
        }})()
    """)
    if ok != 'ok':
        log.warn(f"Select por label '{label_text}' = '{valor}': {ok}")
        return False
    return True

async def _find_field_by_span(page, label_text, kind):
    """Encontra select/input/textarea pelo span 'font-weight:bold' com o nome do campo.

    Retorna o ID do elemento encontrado, ou None.
    kind: 'select', 'input', 'textarea', 'date'
    """
    return await page.evaluate(f"""
        (() => {{
            const target = {json.dumps(label_text.lower())};
            const norm = s => s.normalize('NFD').replace(/[\u0300-\u036f]/g, '').toLowerCase().trim().replace(/[:?*\\s()]+$/, '');
            const targetN = norm(target);

            // procura span com font-weight:bold contendo o texto
            const spans = Array.from(document.querySelectorAll('span[style*="bold"], span[style*="font-weight"]'));
            const span = spans.find(s => {{
                const t = norm(s.textContent || '');
                return t === targetN || t.startsWith(targetN);
            }});
            if (!span) return null;

            // sobe ate <tr> e desce no proximo <td>
            let tr = span.closest('tr');
            // pode estar em tabela aninhada - sobe se necessario
            while (tr && tr.querySelectorAll('td').length < 2) {{
                const outer = tr.parentElement && tr.parentElement.closest('tr');
                if (outer === tr) break;
                tr = outer;
            }}
            if (!tr) return null;

            const tds = tr.querySelectorAll(':scope > td');
            // procura nos tds proximos (alem do que tem o span)
            for (let i = 0; i < tds.length; i++) {{
                const td = tds[i];
                if (td.contains(span)) continue;  // pula a celula da label
                const kind = {json.dumps(kind)};
                let el = null;
                if (kind === 'select') {{
                    el = td.querySelector('select');
                }} else if (kind === 'textarea') {{
                    el = td.querySelector('textarea');
                }} else if (kind === 'input') {{
                    el = td.querySelector('input:not([type="hidden"]):not([type="checkbox"])');
                }} else if (kind === 'date') {{
                    el = td.querySelector('input[id*="fieldDate"], input[id*="Date"], input.hasDatepicker, input[name*="Date"]');
                }}
                if (el) {{
                    if (!el.id) el.id = '_rpa_field_' + Math.random().toString(36).slice(2, 10);
                    return el.id;
                }}
            }}
            return null;
        }})()
    """)



async def selecionar_clicando_por_titulo(page, titulo, valor, log):
    """Seleciona valor simulando clique humano (abre painel + clica item)."""
    try:
        # 1. Acha o widget pelo title (estavel) - busca tambem por title parcial
        widget_loc = page.locator(f'.ui-selectonemenu[title="{titulo}"]').first
        n = await widget_loc.count()
        if n == 0:
            # tenta variacoes de capitalizacao
            found = False
            for variant in [titulo.lower(), titulo.capitalize(), titulo.title()]:
                widget_loc = page.locator(f'.ui-selectonemenu[title="{variant}"]').first
                if await widget_loc.count() > 0:
                    found = True
                    break
            if not found:
                # tenta com title CONTAINING o termo (case-insensitive via JS)
                widget_id = await page.evaluate(f"""
                    (() => {{
                        const target = {json.dumps(titulo.lower())};
                        const widgets = document.querySelectorAll('.ui-selectonemenu[title]');
                        const norm = s => s.normalize('NFD').replace(/[\u0300-\u036f]/g, '').toLowerCase().trim();
                        const targetN = norm(target);
                        for (const w of widgets) {{
                            const t = norm(w.getAttribute('title') || '');
                            if (t === targetN || t.startsWith(targetN) || targetN.startsWith(t)) {{
                                return w.id;
                            }}
                        }}
                        return null;
                    }})()
                """)
                if widget_id:
                    widget_loc = page.locator(f'#{widget_id}')
                    log.info(f"clicando '{titulo}': widget encontrado por norm match (id={widget_id})")
                else:
                    log.warn(f"clicando '{titulo}': widget NAO encontrado")
                    return False

        # 2. Pega o id do widget pra montar o id do painel
        widget_id = await widget_loc.get_attribute('id')
        if not widget_id:
            log.warn(f"clicando '{titulo}': widget sem id")
            return False

        # 3. Clica no label/trigger pra abrir o painel
        await widget_loc.scroll_into_view_if_needed()
        label = widget_loc.locator('.ui-selectonemenu-label, .ui-selectonemenu-trigger').first
        await label.click(timeout=5000)
        await page.wait_for_timeout(400)

        # 4. Espera o painel ficar visivel
        panel = page.locator(f'#{widget_id}_panel')
        try:
            await panel.wait_for(state='visible', timeout=5000)
        except Exception:
            log.warn(f"clicando '{titulo}': painel nao abriu")
            return False

        # 5. Aguarda os itens carregarem no painel (podem vir via AJAX)
        items_sel = f'#{widget_id}_panel li'  # PrimeFaces usa li sem classe
        try:
            await page.wait_for_selector(items_sel, state='visible', timeout=5000)
        except Exception:
            # itens nao apareceram — se tem filtro, digita para forcar carregamento
            filtro_loc = page.locator(f'#{widget_id}_filter')
            if await filtro_loc.count() > 0:
                try:
                    await filtro_loc.press_sequentially(valor[:8], delay=60)
                    await page.wait_for_selector(items_sel, state='visible', timeout=5000)
                except Exception:
                    pass

        # 6. Clica no item com o texto certo (case-insensitive, sem acentos)
        clicado = await page.evaluate(f"""
            (() => {{
                const norm = s => s.normalize('NFD').replace(/[\u0300-\u036f]/g, '').toLowerCase().trim();
                const wanted = norm({json.dumps(valor)});
                const panel = document.getElementById({json.dumps(widget_id + '_panel')});
                if (!panel) return 'no-panel';
                const items = panel.querySelectorAll('li');
                let exact = null, partial = null;
                for (const it of items) {{
                    const t = norm(it.textContent || '');
                    if (t === wanted) {{ exact = it; break; }}
                    if (!partial && t.includes(wanted)) partial = it;
                }}
                const target = exact || partial;
                if (!target) return 'no-item';
                target.click();
                return 'ok';
            }})()
        """)
        if clicado != 'ok':
            log.warn(f"clicando '{titulo}': {clicado}")
            # fecha painel se ainda aberto
            try:
                await page.keyboard.press("Escape")
            except Exception:
                pass
            return False

        # 6. Aguarda AJAX dispar (data-ajax=true do select)
        try:
            await page.wait_for_load_state("networkidle", timeout=8000)
        except Exception:
            pass
        return True

    except Exception as e:
        log.warn(f"clicando '{titulo}': {e}")
        return False



async def preencher_input_por_label(page, label_text, valor, log):
    """Preenche input/textarea encontrado pela label."""
    sel_id = await page.evaluate(f"""
        (() => {{
            const target = {json.dumps(label_text.lower())};
            const labels = Array.from(document.querySelectorAll(
                'label, .ui-outputlabel, .ui-panelgrid-cell, .form-label, td, span, div'
            )).filter(e => e.children.length === 0 || e.tagName === 'LABEL' || e.classList.contains('ui-outputlabel'));
            const lbl = labels.find(l => {{
                const t = (l.textContent || '').trim().toLowerCase().replace(/[:\\s]+$/, '');
                return t === target || t.startsWith(target);
            }});
            if (!lbl) return null;
            let el = null;
            const fid = lbl.getAttribute('for');
            if (fid) el = document.getElementById(fid);
            if (!el || !['INPUT','TEXTAREA'].includes(el.tagName)) {{
                let p = lbl.parentElement;
                for (let i = 0; i < 6 && p && !el; i++) {{
                    el = p.querySelector('input:not([type="hidden"]), textarea');
                    if (!el) p = p.parentElement;
                }}
            }}
            if (!el) return null;
            // garante um id pra acessar via locator
            if (!el.id) el.id = '_rpa_field_' + Math.random().toString(36).slice(2,8);
            return el.id;
        }})()
    """)
    if not sel_id:
        log.warn(f"Input por label '{label_text}' nao encontrado")
        return False
    try:
        loc = page.locator(f'#{sel_id}')
        await loc.fill(str(valor))
        return True
    except Exception as e:
        log.warn(f"Falha preenchendo '{label_text}': {e}")
        return False


async def input_filled_por_label(page, label_text):
    """Por enquanto sempre retorna False - SKIP estava bugado."""
    return False


async def _OLD_input_filled_por_label(page, label_text):
    return await page.evaluate(f"""
        (() => {{
            const target = {json.dumps(label_text.lower())};
            const labels = Array.from(document.querySelectorAll(
                'label, .ui-outputlabel, .ui-panelgrid-cell, .form-label, td, span, div'
            )).filter(e => e.children.length === 0 || e.tagName === 'LABEL' || e.classList.contains('ui-outputlabel'));
            const lbl = labels.find(l => {{
                const t = (l.textContent || '').trim().toLowerCase().replace(/[:\\s]+$/, '');
                return t === target || t.startsWith(target);
            }});
            if (!lbl) return false;
            let el = null;
            const fid = lbl.getAttribute('for');
            if (fid) el = document.getElementById(fid);
            if (!el) {{
                let p = lbl.parentElement;
                for (let i = 0; i < 6 && p && !el; i++) {{
                    el = p.querySelector('input:not([type="hidden"]), textarea');
                    if (!el) p = p.parentElement;
                }}
            }}
            if (!el) return false;
            return (el.value || '').trim().length > 0;
        }})()
    """)


async def select_filled_por_label(page, label_text):
    """Por enquanto sempre retorna False (forca preencher) - SKIP estava bugado."""
    return False


async def _OLD_select_filled_por_label(page, label_text):
    return await page.evaluate(f"""
        (() => {{
            const target = {json.dumps(label_text.lower())};
            const labels = Array.from(document.querySelectorAll(
                'label, .ui-outputlabel, .ui-panelgrid-cell, .form-label, td, span, div'
            )).filter(e => e.children.length === 0 || e.tagName === 'LABEL' || e.classList.contains('ui-outputlabel'));
            const lbl = labels.find(l => {{
                const t = (l.textContent || '').trim().toLowerCase().replace(/[:\\s]+$/, '');
                return t === target || t.startsWith(target);
            }});
            if (!lbl) return false;
            let sel = null;
            const fid = lbl.getAttribute('for');
            if (fid) {{
                const el = document.getElementById(fid);
                if (el && el.tagName === 'SELECT') sel = el;
                else if (el) {{
                    const w = el.closest('.ui-selectonemenu');
                    if (w) sel = w.querySelector('select');
                }}
            }}
            if (!sel) {{
                let p = lbl.parentElement;
                for (let i = 0; i < 6 && p && !sel; i++) {{
                    sel = p.querySelector('select');
                    if (!sel) p = p.parentElement;
                }}
            }}
            if (!sel) return false;
            const v = (sel.value || '').trim();
            if (!v || v === '0' || v === '-1') return false;
            const w = sel.closest('.ui-selectonemenu');
            if (w) {{
                const lblEl = w.querySelector('.ui-selectonemenu-label');
                if (lblEl) {{
                    const t = (lblEl.textContent || '').trim().toLowerCase();
                    if (!t || t.startsWith('selecione') || t === '-') return false;
                }}
            }}
            return true;
        }})()
    """)


async def dump_form_html(page, log_dir, suffix):
    """Salva o HTML completo do form para debug. So salva uma vez."""
    flag = log_dir / ".html_dumped"
    if flag.exists():
        return
    try:
        html = await page.content()
        out = log_dir / f"form_dump_{suffix}.html"
        out.write_text(html, encoding="utf-8")
        flag.write_text("done")
    except Exception:
        pass


async def selecionar_pf(page, name_fragment, valor, log):
    """Seleciona valor num PrimeFaces select hidden via JS puro (bypassa visibility)."""
    try:
        result = await page.evaluate(f"""
            (() => {{
                const sels = document.querySelectorAll(
                    'select[name*="{name_fragment}"], select[id*="{name_fragment}"]'
                );
                if (sels.length === 0) return 'no-select';
                const sel = sels[0];
                const wanted = {json.dumps(valor.lower().strip())};
                let opt = null;
                for (const o of sel.options) {{
                    if ((o.textContent || '').trim().toLowerCase() === wanted) {{
                        opt = o; break;
                    }}
                }}
                if (!opt) {{
                    for (const o of sel.options) {{
                        if ((o.textContent || '').trim().toLowerCase().includes(wanted)) {{
                            opt = o; break;
                        }}
                    }}
                }}
                if (!opt) return 'no-option:' + Array.from(sel.options).map(o => o.textContent.trim()).slice(0,5).join('|');
                sel.value = opt.value;
                sel.dispatchEvent(new Event('change', {{bubbles: true}}));
                const w = sel.closest('.ui-selectonemenu');
                if (w) {{
                    const lbl = w.querySelector('.ui-selectonemenu-label');
                    if (lbl) lbl.textContent = opt.textContent;
                }}
                return 'ok';
            }})()
        """)
        if result != 'ok':
            log.warn(f"selecionar_pf '{name_fragment}' = '{valor}': {result}")
            return False
        return True
    except Exception as e:
        log.warn(f"selecionar_pf '{name_fragment}' = '{valor}': {e}")
        return False


async def marcar_pedidos(page, pedidos_lista, log):
    try:
        # Abre o painel: tenta widget registry primeiro, senao clica no trigger visual
        wname = await page.evaluate("""
            (() => {
                if (!window.PrimeFaces || !PrimeFaces.widgets) return null;
                return Object.keys(PrimeFaces.widgets).find(k => k.toLowerCase().includes('objetolist')) || null;
            })()
        """)
        if wname:
            await page.evaluate(f"PrimeFaces.widgets[{json.dumps(wname)}].show()")
        else:
            trigger = page.locator(
                '.ui-selectcheckboxmenu .ui-selectcheckboxmenu-trigger, '
                '.ui-selectcheckboxmenu-trigger'
            ).first
            await trigger.click(timeout=5000)
        await page.wait_for_timeout(600)

        # Clica em cada pedido no painel aberto
        for pedido in pedidos_lista:
            ps = MAPA_PEDIDOS.get(pedido.strip(), pedido.strip() + " - Cível")
            clicado = await page.evaluate(f"""
                (() => {{
                    const norm = s => s.normalize('NFD').replace(/[\\u0300-\\u036f]/g, '').toLowerCase().trim();
                    const wanted = norm({json.dumps(ps)});
                    const panels = document.querySelectorAll(
                        '.ui-selectcheckboxmenu-panel, [id*="objetoList_panel"], [id*="ObjetoList_panel"]'
                    );
                    for (const panel of panels) {{
                        const items = panel.querySelectorAll('li.ui-selectcheckboxmenu-item');
                        for (const item of items) {{
                            const lbl = item.querySelector('label');
                            const t = norm(lbl ? lbl.textContent : item.textContent || '');
                            if (t === wanted || t.includes(wanted)) {{
                                const cb = item.querySelector('.ui-chkbox-box');
                                if (cb) cb.click();
                                else item.click();
                                return true;
                            }}
                        }}
                    }}
                    return false;
                }})()
            """)
            if clicado:
                log.success(f"Pedido marcado: {ps}")
            else:
                log.warn(f"Pedido nao encontrado no painel: {ps}")

        # Fecha o painel
        if wname:
            await page.evaluate(f"PrimeFaces.widgets[{json.dumps(wname)}].hide()")
        else:
            await page.keyboard.press("Escape")
        await page.wait_for_timeout(800)

        await page.locator('button[id*="addObjetos"]').first.click()
        await page.wait_for_timeout(1000)
        return True
    except Exception as e:
        log.warn(f"Pedidos: {e}")
        return False


async def fazer_upload(page, arquivo, log):
    try:
        # 1. seleciona o tipo de documento
        await selecionar_pf(page, "eFileTipoCombo", "Documentos Gerais", log)
        await page.wait_for_timeout(800)

        # 2. localiza o input de upload com multiplos seletores
        file_input = None
        for _sel in [
            'input[id*="uploadGedEFile"]',
            'input[type="file"][id*="GedEFile"]',
            'input[type="file"][name*="GedEFile"]',
            'input[type="file"]',
        ]:
            _loc = page.locator(_sel).first
            try:
                await _loc.wait_for(state='attached', timeout=3000)
                file_input = _loc
                log.info(f"Input upload encontrado: {_sel}")
                break
            except Exception:
                continue
        if file_input is None:
            log.warn("Input de upload nao encontrado")
            return False
        await file_input.set_input_files(arquivo)
        log.info(f"Anexando: {os.path.basename(arquivo)} - aguardando upload...")

        # 3. aguarda a row de upload aparecer em ui-fileupload-files
        try:
            await page.wait_for_selector(
                '.ui-fileupload-files .ui-fileupload-row, .ui-fileupload-files > div > div',
                timeout=10000, state='attached'
            )
        except Exception:
            log.info("(row de upload nao detectada, prosseguindo)")

        # 4. aguarda a row de upload SUMIR (= upload terminou no client side)
        log.info("Esperando barra de progresso terminar...")
        try:
            await page.wait_for_function(
                """() => {
                    const files = document.querySelector('.ui-fileupload-files');
                    if (!files) return true;
                    // verifica se tem alguma row ainda fazendo upload
                    const rows = files.querySelectorAll('.ui-fileupload-row, [class*="progress"], [role="progressbar"]');
                    return rows.length === 0;
                }""",
                timeout=120000  # 2 minutos pro upload
            )
        except Exception:
            log.warn("Timeout esperando upload terminar (pode ser arquivo grande)")

        # 5. aguarda o arquivo aparecer na tabela "Documentos"
        await page.wait_for_selector(
            '#j_id_6v_2_18_2_l_5_5d_1\\:gedEFileDataTable tbody tr',
            timeout=30000
        )

        # 6. networkidle - espera todos os requests AJAX terminarem
        try:
            await page.wait_for_load_state("networkidle", timeout=20000)
        except Exception:
            pass
        await page.wait_for_timeout(2000)
        log.success("Upload concluido (arquivo na tabela Documentos)")
        return True
    except PWTimeout:
        log.warn("Timeout no upload")
        return False
    except Exception as e:
        log.warn(f"Upload: {e}")
        return False

async def fill_select_robusto(page, nomes_label, valor, log, fallback_id_fragment=None):
    """Tenta preencher um PrimeFaces select usando varias estrategias:
    1. Click humano via title= no widget
    2. Span-based: acha <span>label:</span> e o select adjacente
    3. Fallback por id parcial
    """
    if isinstance(nomes_label, str):
        nomes_label = [nomes_label]

    # Estrategia 1: click via title
    for nome in nomes_label:
        try:
            if await selecionar_clicando_por_titulo(page, nome, valor, log):
                return True
        except Exception as e:
            log.warn(f"clicando '{nome}': {e}")

    # Estrategia 1.5: acha widget pelo title e seta o select hidden direto via JS
    # (mesmo mecanismo do selecionar_pf, mas sem depender de ID fixo)
    for nome in nomes_label:
        result = await page.evaluate(f"""
            (() => {{
                const norm = s => s.normalize('NFD').replace(/[\\u0300-\\u036f]/g, '').toLowerCase().trim();
                const targetN = norm({json.dumps(nome.lower())});
                const widgets = document.querySelectorAll('.ui-selectonemenu[title]');
                let widget = null;
                for (const w of widgets) {{
                    const t = norm(w.getAttribute('title') || '');
                    if (t === targetN || t.startsWith(targetN) || targetN.startsWith(t)) {{
                        widget = w; break;
                    }}
                }}
                if (!widget) return 'no-widget';
                const sel = widget.querySelector('select');
                if (!sel) return 'no-select';
                const wantedN = norm({json.dumps(valor)});
                let opt = null;
                const getText = o => (o.text || o.textContent || '').trim();
                for (const o of sel.options) {{
                    if (norm(getText(o)) === wantedN) {{ opt = o; break; }}
                }}
                if (!opt) for (const o of sel.options) {{
                    if (norm(getText(o)).includes(wantedN)) {{ opt = o; break; }}
                }}
                if (!opt) return 'no-opt:w=' + wantedN + ':' + Array.from(sel.options).map(o => norm(getText(o))).slice(0,10).join('|');
                sel.value = opt.value;
                sel.dispatchEvent(new Event('change', {{bubbles: true}}));
                if (window.jQuery) jQuery(sel).trigger('change');
                const lbl = widget.querySelector('.ui-selectonemenu-label');
                if (lbl) lbl.textContent = opt.textContent.trim();
                return 'ok';
            }})()
        """)
        if result == 'ok':
            log.info(f"  -> preenchido via title+select direto ({nome})")
            return True
        elif result not in ('no-widget', 'no-select'):
            log.warn(f"title+select direto '{nome}': {result}")

    # Estrategia 2: span-based
    for nome in nomes_label:
        sel_id = await _find_field_by_span(page, nome, 'select')
        if not sel_id:
            continue
        result = await page.evaluate(f"""
            (() => {{
                const sel = document.getElementById({json.dumps(sel_id)});
                if (!sel) return 'no-el';
                const wanted = {json.dumps(valor.lower().strip())};
                const norm = s => s.normalize('NFD').replace(/[\u0300-\u036f]/g, '').toLowerCase().trim();
                const wantedN = norm(wanted);
                let opt = null;
                for (const o of sel.options) {{
                    if (norm(o.textContent) === wantedN) {{ opt = o; break; }}
                }}
                if (!opt) {{
                    for (const o of sel.options) {{
                        if (norm(o.textContent).includes(wantedN)) {{ opt = o; break; }}
                    }}
                }}
                if (!opt) {{
                    const opts = Array.from(sel.options).map(o => o.textContent.trim()).slice(0, 6).join(' | ');
                    return 'no-opt:' + opts;
                }}
                sel.value = opt.value;
                sel.dispatchEvent(new Event('change', {{bubbles: true}}));
                const w = sel.closest('.ui-selectonemenu');
                if (w) {{
                    const lblEl = w.querySelector('.ui-selectonemenu-label');
                    if (lblEl) lblEl.textContent = opt.textContent;
                }}
                return 'ok';
            }})()
        """)
        if result == 'ok':
            log.info(f"  -> preenchido via span-based ({nome})")
            return True
        else:
            log.warn(f"span-based '{nome}': {result}")

    # Estrategia 3: fallback id
    if fallback_id_fragment:
        return await selecionar_pf(page, fallback_id_fragment, valor, log)
    return False


async def fill_input_robusto(page, nomes_label, valor, log, fallback_selector=None, kind='input'):
    """Preenche input/textarea por label com fallback."""
    if isinstance(nomes_label, str):
        nomes_label = [nomes_label]

    for nome in nomes_label:
        sel_id = await _find_field_by_span(page, nome, kind)
        if sel_id:
            try:
                # usa locator com escape de :
                escaped_id = sel_id.replace(':', '\\:')
                loc = page.locator(f'#{escaped_id}')
                await loc.fill(str(valor))
                return True, sel_id
            except Exception as e:
                log.warn(f"fill por span '{nome}': {e}")

    # fallback
    if fallback_selector:
        try:
            await page.fill(fallback_selector, str(valor))
            return True, None
        except Exception as e:
            log.warn(f"fill fallback: {e}")
    return False, None



async def fill_typeahead_juiz(page, valor, log):
    """Preenche o campo Juiz (PrimeFaces AutoComplete).
    Nao sobrepoe se o campo ja estiver preenchido.
    Retorna True se preenchido/mantido, False se falhou, None se campo nao encontrado.
    """
    # verifica se o input existe — usa ID estavel do componente
    field_id = await page.evaluate("""
        (() => {
            const el = document.getElementById('autocompleteJuiz_input');
            if (el) return 'autocompleteJuiz_input';
            // fallback: procura por role=combobox com aria-controls contendo 'Juiz'
            const combo = document.querySelector('input[aria-controls*="Juiz"], input[id*="Juiz"][type="text"]');
            if (combo) {
                if (!combo.id) combo.id = '_rpa_juiz_' + Math.random().toString(36).slice(2, 8);
                return combo.id;
            }
            return null;
        })()
    """)
    if not field_id:
        log.warn("Campo Juiz: nao encontrado no formulario")
        return None

    # le o valor atual via JS — NAO clica (click handler do PF limpa o campo)
    current = await page.evaluate(f"""
        (() => {{
            const el = document.getElementById({json.dumps(field_id)});
            return el ? (el.value || '').trim() : '';
        }})()
    """)
    if current:
        log.info(f"Campo Juiz ja preenchido ('{current}') - mantendo")
        return True

    loc = page.locator(f'#{field_id.replace(":", "\\:")}')

    try:
        await loc.click()
        await page.wait_for_timeout(200)
        await loc.press_sequentially(valor, delay=80)
        await page.wait_for_timeout(600)
    except Exception as e:
        log.warn(f"Campo Juiz: erro ao digitar - {e}")
        return False

    # aguarda o painel do autocomplete — usa ID estavel quando disponivel
    try:
        painel_loc = page.locator('#autocompleteJuiz_panel li, #autocompleteJuiz_panel .ui-autocomplete-item').first
        try:
            await painel_loc.wait_for(state='visible', timeout=8000)
        except Exception:
            # fallback para outros seletores de painel
            painel_loc = page.locator(
                '.ui-autocomplete-panel li, .ui-autocomplete-items li'
            ).first
            await painel_loc.wait_for(state='visible', timeout=5000)

        clicado = await page.evaluate(f"""
            (() => {{
                const norm = s => s.normalize('NFD').replace(/[\\u0300-\\u036f]/g, '').toLowerCase().trim();
                const wanted = norm({json.dumps(valor)});
                const items = document.querySelectorAll('#autocompleteJuiz_panel li, #autocompleteJuiz_panel .ui-autocomplete-item, .ui-autocomplete-panel li, .ui-autocomplete-items li');
                let exact = null, partial = null;
                for (const it of items) {{
                    const t = norm(it.textContent || '');
                    if (t === wanted) {{ exact = it; break; }}
                    if (!partial && t.includes(wanted)) partial = it;
                }}
                const target = exact || partial;
                if (target) {{ target.click(); return true; }}
                return false;
            }})()
        """)

        if clicado:
            log.info(f"Juiz: '{valor}' preenchido via typeahead")
        else:
            await painel_loc.click()
            log.info(f"Juiz: sem match exato para '{valor}', clicou no primeiro item disponivel")
        return True

    except Exception as e:
        log.warn(f"Campo Juiz: autocomplete nao apareceu ou falhou - {e}")
        try:
            await page.keyboard.press("Escape")
            await page.wait_for_timeout(300)
        except Exception:
            pass
        return False


async def preencher_formulario(page, row, pasta_pdfs, log, preservar=True):
    numero       = str(row["(Processo) Número"]).strip()
    tipo_acao    = str(row["Tipo de Ação"]).strip()
    procedimento = str(row["Procedimento"]).strip()
    fase         = str(row["Fase processual"]).strip()
    resumo       = str(row["Resumo da Ação"]).strip()
    pedidos      = str(row["Pedidos"]).strip()
    valor        = row["Valor da Causa"]
    prazo        = row["Prazo para Defesa"]

    subsidios_raw = None
    for col in ("Deseja solicitar subsídios?", "Subsídios", "Subsidios"):
        if col in row.index:
            subsidios_raw = row[col]
            break
    subsidios = normalizar_subsidios(subsidios_raw)

    juiz_raw = None
    for col in ("Juiz", "juiz", "Juíza", "juiza", "Magistrado", "magistrado"):
        if col in row.index:
            juiz_raw = row[col]
            break
    _juiz_str = str(juiz_raw).strip() if pd.notna(juiz_raw) else ""
    _juiz_invalidos = {"", "nan", "n/a", "-", "none", "s/n"}
    juiz = _juiz_str if _juiz_str.lower() not in _juiz_invalidos else None

    tipo_sistema = MAPA_TIPO_ACAO.get(tipo_acao, tipo_acao)

    # diagnostico: dump HTML do form na primeira vez
    log_dir = Path(__file__).resolve().parent / "logs"
    log_dir.mkdir(exist_ok=True)
    safe = re.sub(r'[^0-9a-zA-Z]', '_', numero)
    await dump_form_html(page, log_dir, safe)

    # ==== Tipo de Audiencia Inicial (condicional — le Data diretamente da pagina) ====
    # Nao vem da planilha: a RPA verifica se o campo ja esta preenchido pelo sistema.
    # Se houver data → seleciona "Conciliacao e Mediacao - Civel" no Tipo de Audiencia.
    data_audiencia = await page.evaluate("""
        (() => {
            const norm = s => (s || '').normalize('NFD').replace(/[\\u0300-\\u036f]/g, '').toLowerCase().trim();
            const all = Array.from(document.querySelectorAll('label, td, th, span, div'))
                .filter(el => el.children.length === 0
                    && norm(el.textContent).includes('audiencia')
                    && norm(el.textContent).includes('inicial'));
            for (const lbl of all) {
                let p = lbl.parentElement;
                for (let i = 0; i < 8 && p; i++) {
                    const inputs = Array.from(p.querySelectorAll('input[type="text"], input[type="date"]'))
                        .filter(inp => {
                            const id = (inp.name || inp.id || '').toLowerCase();
                            return !id.includes('tipo') && inp.value && inp.value.trim();
                        });
                    if (inputs.length > 0) return inputs[0].value.trim();
                    p = p.parentElement;
                }
            }
            return null;
        })()
    """)
    if data_audiencia:
        log.info(f"Data Audiencia Inicial detectada: '{data_audiencia}' → preenchendo Tipo de Audiencia")
        await fill_select_robusto(
            page,
            ["Tipo de Audiência Inicial", "Tipo de Audiencia Inicial", "Tipo de Audiência"],
            "Conciliação e Mediação - Cível", log,
            fallback_id_fragment=None
        )
        await page.wait_for_timeout(800)
    else:
        log.info("Data Audiencia Inicial vazia → Tipo de Audiencia nao preenchido")

    # ==== Juiz ====
    if juiz:
        log.info(f"Juiz: {juiz}")
        await fill_typeahead_juiz(page, juiz, log)
        await page.wait_for_timeout(800)
    else:
        log.info("Juiz: coluna ausente ou vazia na planilha - pulando")

    # ==== Tipo de Acao ====
    log.info(f"Tipo de Acao: {tipo_sistema}")
    await fill_select_robusto(
        page,
        ["Tipo de ação", "Tipo de Ação", "Tipo de acao"],
        tipo_sistema, log,
        fallback_id_fragment="j_id_6v_2_18_2_f_5_2_1_2_1"
    )
    await page.wait_for_timeout(1500)

    # ==== Justiça (planilha "Procedimento") ====
    log.info(f"Justiça (da coluna Procedimento): {procedimento}")
    await fill_select_robusto(
        page,
        ["Justiça", "Justica"],
        procedimento, log,
        fallback_id_fragment="LocationType"
    )
    await page.wait_for_timeout(1500)

    # ==== Fase processual ====
    log.info(f"Fase processual: {fase}")
    ok_fase = await fill_select_robusto(
        page,
        ["Fase processual", "Fase"],
        fase, log,
        fallback_id_fragment="j_id_6v_2_18_2_h_5_3a_1"
    )
    if not ok_fase:
        ok_fase = await page.evaluate(f"""
            (() => {{
                const norm = s => s.normalize('NFD').replace(/[\\u0300-\\u036f]/g, '').toLowerCase().trim().replace(/[:*()\\s]+$/, '');
                const target = 'fase processual';
                const wanted = norm({json.dumps(fase)});
                const all = Array.from(document.querySelectorAll('td, th, label, span, div'))
                    .filter(e => e.children.length === 0 && norm(e.textContent || '').startsWith('fase'));
                for (const el of all) {{
                    let p = el.parentElement;
                    let sel = null;
                    for (let i = 0; i < 8 && p && !sel; i++) {{
                        sel = p.querySelector('select');
                        if (!sel) p = p.parentElement;
                    }}
                    if (!sel) continue;
                    let opt = null;
                    for (const o of sel.options) {{
                        if (norm(o.textContent) === wanted) {{ opt = o; break; }}
                    }}
                    if (!opt) for (const o of sel.options) {{
                        if (norm(o.textContent).includes(wanted)) {{ opt = o; break; }}
                    }}
                    if (!opt) continue;
                    sel.value = opt.value;
                    sel.dispatchEvent(new Event('change', {{bubbles: true}}));
                    const w = sel.closest('.ui-selectonemenu');
                    if (w) {{
                        const lbl = w.querySelector('.ui-selectonemenu-label');
                        if (lbl) lbl.textContent = opt.textContent;
                    }}
                    return true;
                }}
                return false;
            }})()
        """)
        if ok_fase:
            log.info("Fase processual preenchida via varredura ampla")
        else:
            log.warn("Fase processual nao preenchida!")
    await page.wait_for_timeout(1200)

    # ==== Resumo da Acao (textarea) ====
    log.info("Preenchendo Resumo")
    ok, _ = await fill_input_robusto(
        page,
        ["Resumo da ação", "Resumo da Ação", "Resumo da acao"],
        resumo, log,
        fallback_selector='textarea[name*="i_5_45"]',
        kind='textarea'
    )
    if not ok:
        log.warn("Resumo nao preenchido!")

    # ==== Pedidos ====
    log.info(f"Pedidos: {pedidos}")
    await marcar_pedidos(page, [p.strip() for p in pedidos.split(",")], log)

    # ==== Valor da Causa ====
    if pd.notna(valor):
        valor_str = f"{float(valor):.2f}".replace(".", ",")
        log.info(f"Valor: {valor_str}")
        ok, _ = await fill_input_robusto(
            page,
            ["Valor da causa", "Valor da Causa"],
            valor_str, log,
            fallback_selector='input[name*="k_5_30"][name$="_input"]',
            kind='input'
        )
        if not ok:
            log.warn("Valor nao preenchido!")

    # ==== Prazo para Defesa ====
    if pd.notna(prazo):
        data_str = pd.Timestamp(prazo).strftime("%d/%m/%Y")
        log.info(f"Prazo: {data_str}")
        ok, sel_id = await fill_input_robusto(
            page,
            ["Prazo para Defesa", "Prazo de Defesa", "Prazo"],
            data_str, log,
            fallback_selector='input[id*="fieldDate"][type="text"], input[name*="fieldDate"]',
            kind='date'
        )
        if ok:
            try:
                await page.keyboard.press("Tab")
            except Exception:
                pass
        else:
            log.warn("Prazo nao preenchido!")
        await page.wait_for_timeout(500)

    # ==== Subsidios ====
    if subsidios_raw is None:
        log.warn("Coluna 'Deseja solicitar subsidios?' nao encontrada - usando 'Nao'")
    log.info(f"Subsidios: {subsidios}")
    await fill_select_robusto(
        page,
        ["Deseja solicitar subsídios?", "Subsídios", "Subsidios", "Deseja solicitar"],
        subsidios, log,
        fallback_id_fragment="pgTypeSelectFie"
    )
    await page.wait_for_timeout(800)

    # ==== Upload de documento ====
    arquivo = encontrar_arquivo(numero, pasta_pdfs)
    if arquivo:
        log.info(f"Upload: {os.path.basename(arquivo)}")
        await fazer_upload(page, arquivo, log)
    else:
        # arquivo nao existe na pasta - levanta erro claro antes de salvar
        log.warn(f"PDF nao localizado na pasta para o processo {numero}")
        raise Exception(f"PDF nao localizado na pasta para o processo {numero} - anexar manualmente")



async def _abrir_processo_da_pagina(page, numero, log):
    """Tenta clicar no processo na pagina atual (resultados ou dashboard).
    Retorna True se navegou para processoView.
    """
    # prefixo numerico: "1234567" — parte antes do primeiro "-"
    # evita problemas de formatacao (separadores, espacos)
    prefixo = numero.split("-")[0].strip()

    estrategias = [
        # link direto para processoView
        '[href*="processoView"]',
        # row com numero completo
        f'tr:has-text("{numero}") a',
        f'a:has-text("{numero}")',
        f'td:has-text("{numero}")',
        # row com prefixo (mais tolerante a formatacao)
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


async def processar(page, row, idx, total, pasta_pdfs, log, preservar=True):
    numero = str(row["(Processo) Número"]).strip()
    log.info(f"--- [{idx+1}/{total}] {numero} ---", step="start_case", caso=numero)
    try:
        await page.goto(f"{URL_BASE}/contenciosoDashboard.elaw", wait_until="domcontentloaded", timeout=45000)

        # aguarda campo de busca estar disponivel
        campo = page.locator('input[placeholder="Pesquise por aqui!"]')
        await campo.wait_for(state='visible', timeout=15000)
        await campo.click()
        # limpa com Ctrl+A + Delete (mais confiavel que fill(""))
        await page.keyboard.press("Control+a")
        await page.keyboard.press("Delete")
        await page.wait_for_timeout(300)
        # digita com delay para disparar eventos do autocomplete
        await campo.press_sequentially(numero, delay=50)
        await page.wait_for_timeout(1000)

        clicou = False

        # Estrategia 1: autocomplete ficou visivel — clica primeiro item
        # Render e mais lento — aguarda ate 10s pelo autocomplete
        try:
            ac = page.locator(
                '.ui-autocomplete-panel li, .ui-autocomplete-items li'
            ).first
            await ac.wait_for(state='visible', timeout=10000)
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

        # Estrategia 2: Enter — eLaw pode ir direto pro processo
        # ou mostrar pagina de resultados intermediaria
        if not clicou:
            await campo.focus()
            await campo.press("Enter")
            # aguarda navegacao — Render pode ser lento
            try:
                await page.wait_for_load_state("load", timeout=30000)
            except Exception:
                pass
            await page.wait_for_timeout(3000)

            if "processoView.elaw" in page.url:
                clicou = True
                log.info("Processo aberto via Enter (direto)")
            else:
                # Enter navegou para pagina de resultados — tenta clicar no processo
                log.info(f"Pagina apos Enter: {page.url} — buscando processo na pagina")
                clicou = await _abrir_processo_da_pagina(page, numero, log)

        # Estrategia 3: tenta na pagina atual sem Enter adicional
        if not clicou:
            log.warn(f"Estrategias 1 e 2 falharam — tentando localizar na pagina (URL: {page.url})")
            clicou = await _abrir_processo_da_pagina(page, numero, log)

        # Estrategia 4: aguarda mais um pouco — navegacao pode ter iniciado mas nao completado
        if not clicou:
            log.info("Aguardando possivel navegacao em andamento (15s)...")
            try:
                await page.wait_for_url("**/processoView.elaw**", timeout=15000)
                clicou = True
                log.info("Processo aberto (navegacao tardia detectada)")
            except Exception:
                pass

        if not clicou:
            # diagnostico: salva screenshot + URL atual
            try:
                safe = re.sub(r'[^0-9a-zA-Z]', '_', numero)
                log_dir = Path(__file__).resolve().parent / "logs"
                log_dir.mkdir(exist_ok=True)
                shot = log_dir / f"falhou_busca_{safe}.png"
                await page.screenshot(path=str(shot), full_page=True)
                log.warn(f"Screenshot salvo em: {shot}")
                log.warn(f"URL no momento da falha: {page.url}")
            except Exception as se:
                log.warn(f"Falha ao salvar diagnostico: {se}")
            raise Exception(
                f"Nao foi possivel encontrar/abrir o processo {numero} apos buscar"
            )

        await page.wait_for_load_state("networkidle", timeout=30000)

        # Aguarda a lista de tarefas carregar — aceita PT e EN, nao bloqueia se ausente
        TASK_SELS = (
            'text=Complemento de cadastro, '
            'text=Complementary registration, '
            'text=Complementary, '
            'text=Cadastro'
        )
        try:
            await page.wait_for_selector(TASK_SELS, timeout=20000)
        except Exception:
            log.warn("Texto da tarefa nao encontrado em 20s, prosseguindo mesmo assim...")

        # Identifica o botao correto da tarefa "Complemento de cadastro - Escritorio (Judicial)"
        # Na linha tem ~8 botoes: lupa, editar, CHECK, calendario, info, X, ordenar, download
        # O que abre o formulario e' o CHECK (titulo "Confirmar" ou icone pi-check/fa-check)
        confirmado = await page.evaluate("""
            (() => {
                // aceita PT e EN
                const TARGETS = [
                    'Complemento de cadastro - Escritório (Judicial)',
                    'Complemento de cadastro',
                    'Complementary registration',
                    'Complementary',
                ];
                let row = null;
                for (const target of TARGETS) {
                    for (const tr of document.querySelectorAll('tr')) {
                        if (tr.innerText.includes(target)) { row = tr; break; }
                    }
                    if (row) break;
                }
                if (!row) return 'no-row';

                const btns = Array.from(row.querySelectorAll('button, a.ui-button'));

                // Estrategia 1: title contem "confirm"
                for (const b of btns) {
                    const t = (b.title || b.getAttribute('aria-label') || '').toLowerCase();
                    if (t.includes('confirm')) { b.click(); return 'title-confirm'; }
                }
                // Estrategia 2: icone check
                for (const b of btns) {
                    if (b.querySelector('.pi-check, .fa-check, .ui-icon-check, [class*="check"]')) {
                        const cls = b.querySelector('[class*="check"]').className.toLowerCase();
                        // exclui icones de "checkbox" ou "checklist"
                        if (cls.includes('check') && !cls.includes('checkbox') && !cls.includes('checklist')) {
                            b.click(); return 'icon-check';
                        }
                    }
                }
                // Estrategia 3: a posicao 3 (0=lupa, 1=editar, 2=CHECK, 3=calendario)
                const submits = btns.filter(b => b.type === 'submit' || b.tagName === 'BUTTON');
                if (submits.length >= 3) { submits[2].click(); return 'pos-3'; }

                return 'not-found';
            })()
        """)
        log.info(f"Click no botao Confirmar: {confirmado}")
        if confirmado in (False, 'no-row', 'not-found'):
            # fallback final: tenta clicar em qualquer botao da row
            try:
                await page.locator(
                    'tr:has-text("Complemento de cadastro - Escritório (Judicial)") button'
                ).nth(2).click(timeout=5000)
            except Exception:
                # ultimo recurso: salva diagnostico
                try:
                    safe = re.sub(r'[^0-9a-zA-Z]', '_', numero)
                    log_dir = Path(__file__).resolve().parent / "logs"
                    log_dir.mkdir(exist_ok=True)
                    shot = log_dir / f"falhou_botao_{safe}.png"
                    await page.screenshot(path=str(shot), full_page=True)
                    log.warn(f"Screenshot salvo em: {shot}")
                except Exception:
                    pass
                raise Exception(f"Nao consegui localizar botao de Confirmar para {numero}")

        await page.wait_for_url("**/agendamentoContenciosoConfirm.elaw**", timeout=30000)
        await page.wait_for_load_state("networkidle", timeout=30000)

        await preencher_formulario(page, row, pasta_pdfs, log, preservar=preservar)

        # espera tudo se acomodar antes de salvar (especialmente upload)
        await page.wait_for_timeout(3000)
        url_antes = page.url
        import time as _t

        for tentativa_save in range(1, 3):  # 2 tentativas
            log.info(f"Salvando (tentativa {tentativa_save}/2)...")
            btn = page.locator('button:has-text("Confirmar")').last
            await btn.scroll_into_view_if_needed()
            await btn.click()
            log.info("Aguardando confirmacao do save (URL mudar OU growl aparecer, ate 90s)...")

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
                                    t.includes('agendad')) {{
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
            log.info(f"Save: sinal={sinal.get('tipo')} em {elapsed:.1f}s - {sinal.get('msg', sinal.get('url', ''))[:200]}")

            try:
                await page.wait_for_load_state("networkidle", timeout=15000)
            except Exception:
                pass
            await page.wait_for_timeout(2000)

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
                    safe = re.sub(r'[^0-9a-zA-Z]', '_', numero)
                    log_dir = Path(__file__).resolve().parent / "logs"
                    log_dir.mkdir(exist_ok=True)
                    shot = log_dir / f"erro_save_{safe}.png"
                    await page.screenshot(path=str(shot), full_page=True)
                    log.warn(f"Screenshot do erro: {shot}")
                except Exception:
                    pass
                msg_curta = " | ".join(erros)[:300]
                if tentativa_save < 2:
                    log.warn(f"Save tentativa {tentativa_save} falhou ({msg_curta[:80]}) — aguardando 5s para retry...")
                    await page.wait_for_timeout(5000)
                    continue
                raise Exception(f"eLaw rejeitou o save: {msg_curta}")

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
                            if (t.includes('sucesso') || t.includes('salvo') || t.includes('confirmad') || t.includes('conclu')) {
                                return t.slice(0, 200);
                            }
                        }
                        return null;
                    })()
                """)
                if sucesso:
                    log.success(f"Save confirmado por growl: {sucesso}")
                    break
                else:
                    if tentativa_save < 2:
                        try:
                            safe = re.sub(r'[^0-9a-zA-Z]', '_', numero)
                            log_dir = Path(__file__).resolve().parent / "logs"
                            log_dir.mkdir(exist_ok=True)
                            shot = log_dir / f"sem_navegacao_{safe}.png"
                            await page.screenshot(path=str(shot), full_page=True)
                            log.warn(f"Save nao navegou (tentativa {tentativa_save}). Screenshot: {shot} — aguardando 5s para retry...")
                        except Exception:
                            pass
                        await page.wait_for_timeout(5000)
                        continue
                    try:
                        safe = re.sub(r'[^0-9a-zA-Z]', '_', numero)
                        log_dir = Path(__file__).resolve().parent / "logs"
                        log_dir.mkdir(exist_ok=True)
                        shot = log_dir / f"sem_navegacao_{safe}.png"
                        await page.screenshot(path=str(shot), full_page=True)
                        log.warn(f"Save nao navegou. Screenshot: {shot}")
                    except Exception:
                        pass
                    raise Exception(
                        f"Save aparentemente nao funcionou - pagina nao mudou de {url_antes}"
                    )
            else:
                break  # pagina navegou = sucesso

        log.success(f"Concluido: {numero}", step="case_done", caso=numero)
        return True
    except Exception as e:
        msg_clara = mensagem_erro_amigavel(e)
        log.error(f"ERRO em {numero}: {msg_clara}", step="case_error", caso=numero)
        # log do traceback completo so no arquivo (nao no JSON)
        try:
            if log.log_file:
                with open(log.log_file, "a", encoding="utf-8") as f:
                    f.write(f"[{datetime.now().strftime('%H:%M:%S')}] [DEBUG] traceback completo:\n")
                    f.write(traceback.format_exc() + "\n")
        except Exception:
            pass
        return False


async def _login_microsoft_sso(page, usuario, senha, log):
    """Completa o fluxo de login Microsoft SSO (SAML2) após redirect do eLaw."""

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
    NO_SEL   = '#idBtn_Back, button:has-text("Não"), button:has-text("No"), button:has-text("Nao")'

    await page.wait_for_load_state("domcontentloaded", timeout=15000)
    await _screenshot("sso_01_inicial.png")

    # — Passo 1: pick-account ——————————————————————————————————————————
    # Usa seletor preciso [data-test-id="tile"] — jamais clicar "Usar outra conta"
    # pois isso quebra o contexto SAML e roteia para MSA (login.live.com)
    PICK_TILE = '[data-test-id="tile"]'
    try:
        await page.wait_for_selector(PICK_TILE, timeout=4000)
        log.info("SSO: pick-account detectado, clicando primeiro tile")
        await page.locator(PICK_TILE).first.click()
        await page.wait_for_timeout(2000)
        await _screenshot("sso_02_pos_pick.png")
    except Exception:
        pass

    # — Passo 2: campo de email ——————————————————————————————————————
    # Se SAML pre-preencheu o valor, preserva. Só preenche se vazio.
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

    # — Passo 3: campo de senha ——————————————————————————————————————
    await page.wait_for_selector(MS_PASS, timeout=20000)
    await _screenshot("sso_04_senha.png")
    await page.locator(MS_PASS).first.fill(senha)
    await page.locator(MS_NEXT).first.click()
    log.info("SSO: senha preenchida, aguardando redirect...")
    await page.wait_for_timeout(3000)
    await _screenshot("sso_05_pos_senha.png")

    # loga titulo da pagina para diagnostico (MFA / KMSI / erro / redirect)
    try:
        titulo = await page.title()
        log.info(f"SSO pos-senha: titulo='{titulo}' url={page.url[:80]}")
    except Exception:
        pass

    # detecta MFA explicitamente antes de continuar
    MFA_SEL = (
        '#idDiv_SAOTCAS_Title, '
        '#idDiv_SAOTCC_Section, '
        '[data-bind*="phoneConfirmation"], '
        'div:has-text("Microsoft Authenticator"), '
        'div:has-text("Authenticator"), '
        'div:has-text("Verificar sua identidade"), '
        'div:has-text("Verify your identity"), '
        'div:has-text("Aprovar solicitacao"), '
        'div:has-text("Approve sign-in")'
    )
    try:
        await page.wait_for_selector(MFA_SEL, timeout=3000)
        titulo = await page.title()
        raise RuntimeError(
            f"MFA ativado na conta — autenticacao em duas etapas detectada "
            f"(titulo: '{titulo}'). Use uma conta sem MFA ou faca login manual "
            f"e importe o storage_state."
        )
    except RuntimeError:
        raise
    except Exception:
        pass

    # — Passo 4: prompt "Continuar conectado?" / KMSI ————————————————
    KMSI_SEL = '#idSIButton9, #idBtn_Back, button:has-text("Sim"), button:has-text("Yes"), button:has-text("Nao"), button:has-text("No"), button:has-text("Não")'
    try:
        await page.wait_for_selector(KMSI_SEL, timeout=10000)
        titulo = await page.title()
        log.info(f"SSO: prompt pos-senha detectado (titulo: '{titulo}')")
        # clica "Nao" / "No" / "idBtn_Back" para nao persistir sessao
        nao = page.locator('#idBtn_Back, button:has-text("Nao"), button:has-text("No"), button:has-text("Não")')
        if await nao.count() > 0:
            await nao.first.click()
            log.info("SSO: clicou Nao no prompt KMSI")
        else:
            # se so tiver "Sim"/"Yes", clica para avançar
            await page.locator(KMSI_SEL).first.click()
            log.info("SSO: clicou botao no prompt pos-senha")
    except Exception:
        pass

    # — Passo 5: aguarda retorno ao eLaw ————————————————————————————
    await page.wait_for_url("**/homePage.elaw", timeout=60000)
    log.success("Login SSO OK", step="login_ok")


async def run(cfg):
    log = Logger(cfg.get("log_file"))
    df = carregar_planilha(cfg["planilha"])
    total = len(df)
    inicio = int(cfg.get("inicio") or 0)
    fim = int(cfg["fim"]) if cfg.get("fim") else total

    sucesso = erro = 0
    erros = []

    log.info("=" * 55, step="banner")
    log.info(f"RPA eLaw Anima - {datetime.now().strftime('%d/%m/%Y %H:%M')}", step="banner")
    log.info(f"Linhas {inicio+1} -> {fim} (total: {total})", step="banner")
    log.info("=" * 55, step="banner")

    async with async_playwright() as p:
        browser = await p.chromium.launch(
            headless=bool(cfg.get("headless", False)),
            slow_mo=int(cfg.get("slow_mo", 600)),
            args=["--start-maximized"],
        )
        # carrega auth_state.json se disponivel (pula SSO/MFA completamente)
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

        # se auth_state valido → redireciona direto para homePage sem login
        if storage_state:
            try:
                await page.wait_for_url("**/homePage.elaw", timeout=12000)
                log.success("Login via auth_state OK", step="login_ok")
            except Exception:
                log.warn("Auth state expirado ou invalido, tentando login normal...")
                storage_state = None  # flag para log

        if not storage_state or "homePage" not in page.url:
            # ── Login direto com #username / #authKey ──────────────────────────
            # Mesmo padrão do RPA Carrefour (preposto-rpa) que funciona no Render
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

            # fill() — mesmo padrão do Carrefour RPA (funcionou no run das 15:09)
            await page.fill("#username", cfg["usuario"])
            await page.fill("#authKey",  cfg["senha"])
            log.info("Campos #username / #authKey preenchidos")

            # Marca o botão dentro do form com #authKey (não o botão SSO)
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

            # networkidle 30s — mesmo timeout que funcionou
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
                # Login falhou — captura msg de erro
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
            ok = await processar(
                page, row, idx, total, cfg["pasta_pdfs"], log,
                preservar=bool(cfg.get("preservar_campos", True))
            )
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
        print("Uso: python rpa.py config.json", file=sys.stderr)
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
