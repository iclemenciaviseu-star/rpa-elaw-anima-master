"""Script temporario para substituir o bloco Objeto/CausaRaiz em rpa_administrativo.py."""
import sys

path = r"C:\Users\User\Desktop\rpa-elaw-anima-master-main\rpa_administrativo.py"

with open(path, "r", encoding="utf-8-sig") as f:
    content = f.read()

# ── marca inicio e fim do bloco a substituir ──────────────────────────────────
START = "    # ==== Objeto da Acao (ID estavel: comboSelect2Level1) ===="
END   = "    else:\n        log.warn(\"Causa Raiz: vazia na planilha - pulando\")"

i1 = content.find(START)
i2 = content.find(END, i1)
if i1 == -1 or i2 == -1:
    sys.exit(f"Marcadores nao encontrados: i1={i1}, i2={i2}")

# inclui o END no bloco a remover
i2 += len(END)

# ── novo bloco ─────────────────────────────────────────────────────────────────
NEW = r"""    # ==== Objeto da Acao ====
    # NOTA: comboSelect2Level1 e' o seletor de INSTITUICAO (AGES/ANIMA/EBRADI),
    # nao o Objeto da Acao. O Objeto e' preenchido via label proximity.
    # O value numerico varia por instituicao — usar TEXTO para decisao dos campos extras.
    import unicodedata as _uc

    obj_tipo = "desconhecido"   # debito | debito_valor | reclamacao | nenhum | desconhecido
    obj_text = ""               # texto da opcao selecionada
    causa_raiz_sel_id = None    # ID cacheado do select Causa Raiz (encontrado no poll)

    if objeto:
        log.info(f"Objeto da Acao: {objeto}")
        # Estrategia 1: label proximity — funciona nesta pagina
        ok_obj = await _selecionar_dropdown_filtrado(page, "Objeto da ação", objeto, log)
        if not ok_obj:
            ok_obj = await fill_select_robusto(
                page,
                ["Objeto da ação", "Objeto da Ação", "Objeto de ação", "Objeto"],
                objeto, log,
                fallback_id_fragment=None
            )
        if not ok_obj:
            log.warn("Objeto da Acao: todas as estrategias falharam")

        # Le o texto da opcao selecionada para determinar o tipo de campos extras.
        # Identifica o select pelos marcadores de opcoes (ex: 'questoes academicas').
        obj_text = await page.evaluate("""
            (() => {
                const norm = s => s.normalize('NFD').replace(/[\\u0300-\\u036f]/g,'').toLowerCase().trim();
                const MARKERS = ['questoes academicas', 'questoes financeiras', 'alteracao de grade',
                                 'covid', 'financiamentos', 'nao recorrentes', 'atividades complementares'];
                for (const sel of document.querySelectorAll('select')) {
                    const opts = Array.from(sel.options).map(o => norm(o.text || o.textContent));
                    if (opts.some(t => MARKERS.some(m => t.includes(m)))) {
                        const idx = sel.selectedIndex;
                        return idx >= 0
                            ? (sel.options[idx].text || sel.options[idx].textContent || '').trim()
                            : '';
                    }
                }
                return '';
            })()
        """)

        # Determina tipo de campos extras pelo TEXTO (value numerico muda por instituicao)
        _on = ''.join(
            c for c in _uc.normalize('NFD', (obj_text or '').lower())
            if _uc.category(c) != 'Mn'
        )
        if 'academicas' in _on or 'academica' in _on:
            obj_tipo = 'reclamacao'
        elif 'financeiras' in _on:
            obj_tipo = 'debito_valor'
        elif 'covid' in _on or ('nao' in _on and 'recorrentes' in _on):
            obj_tipo = 'nenhum'
        elif _on and _on not in ('', 'selecione'):
            obj_tipo = 'debito'   # Alteracao grade, Atividades comp, Financiamentos priv/pub
        log.info(f"Objeto: '{obj_text}' tipo={obj_tipo}")

        # Timeout AJAX adaptativo: Questoes Academicas e Financiamentos Bolsas carregam em ~12s
        _LENTOS = ('financiamentos e bolsas', 'academicas')
        ajax_timeout_ms = 18000 if any(x in _on for x in _LENTOS) else 10000
        log.info(f"Aguardando AJAX cascata (timeout={ajax_timeout_ms // 1000}s)...")
        try:
            await page.wait_for_load_state("networkidle", timeout=ajax_timeout_ms)
        except Exception:
            pass

        # Poll: aguarda Causa Raiz ter >1 opcao E cacheia o ID do select.
        # Exclui: Objeto da Acao (marcadores de opcao), Instituicao (siglas curtas <=8 chars),
        #         grid extras, combo de upload, combo de subsidios.
        _JS_CR_FIND = """
            (() => {
                const norm = s => s.normalize('NFD').replace(/[\\u0300-\\u036f]/g,'').toLowerCase().trim();
                const OBJETO_MK = ['questoes academicas', 'questoes financeiras', 'alteracao de grade',
                                   'covid', 'financiamentos', 'nao recorrentes', 'atividades complementares'];
                const grid = document.querySelector('[id*="Level1ChildGrid"]');
                for (const s of document.querySelectorAll('select')) {
                    if (s.id.includes('eFileTipoCombo')) continue;
                    if (s.id.includes('pgTypeSelectFie')) continue;
                    if (grid && grid.contains(s))         continue;
                    if (s.options.length <= 1)            continue;
                    const opts = Array.from(s.options).map(o => norm(o.text || o.textContent));
                    // Pula Objeto da Acao (reconhecido pelos marcadores)
                    if (opts.some(t => OBJETO_MK.some(m => t.includes(m)))) continue;
                    // Pula Instituicao: todas as opcoes sao siglas curtas (<=8 chars)
                    if (opts.filter(t => t.length > 8).length === 0) continue;
                    // Este e' o Causa Raiz!
                    if (!s.id) s.id = '_rpa_cr_' + Date.now();
                    return {id: s.id, count: s.options.length};
                }
                return null;
            })()
        """
        poll_max = 36 if ajax_timeout_ms >= 15000 else 20
        for _poll in range(poll_max):
            info = await page.evaluate(_JS_CR_FIND)
            if info and isinstance(info, dict) and info.get('count', 0) > 1:
                causa_raiz_sel_id = info['id']
                log.info(
                    f"Causa Raiz select: id='{causa_raiz_sel_id}' "
                    f"({info['count']} opcoes) apos poll {_poll + 1}"
                )
                break
            if _poll == 0:
                log.info(f"Aguardando AJAX Causa Raiz (tipo={obj_tipo})...")
            await page.wait_for_timeout(500)
    else:
        log.warn("Objeto da Acao: vazio na planilha - pulando")

    # ==== Campos condicionais por tipo de Objeto (grid por indice) ====
    # Ficha Tecnica — pgTypeSelectField2LevelLevel1ChildGrid:
    #   grid[select 0] = Periodo do Debito            (debito / debito_valor)
    #                  = Periodo da Reclamacao do Aluno (reclamacao)
    #   grid[select 1] = Negativacao Sim/Nao           (debito / debito_valor)
    #   grid[input  0] = Valor do Debito (texto livre) (debito_valor apenas)
    _v_pd  = _col(row, "Período do Débito", "Periodo do Debito", "Período do débito")
    _v_neg = _col(row, "Negativação", "Negativacao", "Negação", "Negacao")
    _v_vd  = _col(row, "Valor do Débito", "Valor do Debito", "Valor do débito")
    _v_pr  = _col(row, "Período da Reclamação do Aluno", "Periodo da Reclamacao do Aluno",
                  "Período da reclamação do aluno", "Periodo da reclamacao")

    periodo_debito     = "" if (_v_pd  is None or str(_v_pd ).strip().lower() in ("nan","none","")) else str(_v_pd ).strip()
    negativacao        = "" if (_v_neg is None or str(_v_neg).strip().lower() in ("nan","none","")) else str(_v_neg).strip()
    valor_debito_raw   = _v_vd if (_v_vd is not None and pd.notna(_v_vd)) else None
    periodo_reclamacao = "" if (_v_pr  is None or str(_v_pr ).strip().lower() in ("nan","none","")) else str(_v_pr ).strip()

    if obj_tipo in ('debito', 'debito_valor'):
        if periodo_debito:
            log.info(f"Periodo do Debito (grid[0]): {periodo_debito}")
            ok_pd = await _preencher_campo_grid(page, 0, periodo_debito, 'select', log, "Periodo do Debito")
            if not ok_pd:
                await fill_select_robusto(
                    page, ["Período do débito", "Período do Débito", "Periodo do debito"],
                    periodo_debito, log
                )
            await page.wait_for_timeout(300)

        if negativacao:
            log.info(f"Negativacao (grid[1]): {negativacao}")
            ok_neg = await _preencher_campo_grid(page, 1, negativacao, 'select', log, "Negativacao")
            if not ok_neg:
                await fill_select_robusto(
                    page, ["Negativação", "Negação", "Negativacao", "Negacao"],
                    negativacao, log
                )
            await page.wait_for_timeout(300)

        if obj_tipo == 'debito_valor' and valor_debito_raw is not None:
            try:
                valor_debito_str = f"{float(valor_debito_raw):.2f}".replace(".", ",")
            except (ValueError, TypeError):
                valor_debito_str = str(valor_debito_raw).strip()
            if valor_debito_str:
                log.info(f"Valor do Debito (grid input): {valor_debito_str}")
                ok_vd = await _preencher_campo_grid(page, 0, valor_debito_str, 'input', log, "Valor do Debito")
                if not ok_vd:
                    ok_vd, _ = await fill_input_robusto(
                        page, ["Valor do débito", "Valor do Débito", "Valor do debito"],
                        valor_debito_str, log, kind='input'
                    )
                if not ok_vd:
                    log.warn("Valor do Debito nao preenchido")

    elif obj_tipo == 'reclamacao':
        if periodo_reclamacao:
            log.info(f"Periodo da Reclamacao do Aluno (grid[0]): {periodo_reclamacao}")
            ok_pr = await _preencher_campo_grid(page, 0, periodo_reclamacao, 'select', log, "Periodo Reclamacao")
            if not ok_pr:
                await fill_select_robusto(
                    page, ["Período da reclamação do aluno", "Período da Reclamação do Aluno",
                           "Periodo da reclamacao do aluno", "Periodo da reclamacao"],
                    periodo_reclamacao, log
                )
            await page.wait_for_timeout(300)

    elif obj_tipo == 'nenhum':
        log.info("Objeto sem campos extras (Covid-19 / Nao recorrentes)")

    elif objeto:
        log.warn(f"obj_tipo desconhecido ('{obj_text}') — tentando campos por label")
        if periodo_debito:
            await fill_select_robusto(page, ["Período do débito", "Período do Débito"], periodo_debito, log)
        if negativacao:
            await fill_select_robusto(page, ["Negativação", "Negação", "Negativacao"], negativacao, log)
        if periodo_reclamacao:
            await fill_select_robusto(
                page, ["Período da reclamação do aluno", "Período da Reclamação do Aluno"],
                periodo_reclamacao, log
            )
        if valor_debito_raw is not None:
            try:
                valor_debito_str = f"{float(valor_debito_raw):.2f}".replace(".", ",")
            except Exception:
                valor_debito_str = str(valor_debito_raw).strip()
            if valor_debito_str:
                await fill_input_robusto(
                    page, ["Valor do débito", "Valor do Débito"], valor_debito_str, log, kind='input'
                )

    # ==== Causa Raiz (cascata — populada via AJAX apos Objeto da Acao) ====
    # O ID do select foi cacheado durante o poll acima (estrategia de exclusao).
    if causa_raiz:
        log.info(f"Causa Raiz: {causa_raiz}")
        ok_cr = False

        # Estrategia 1: ID cacheado durante o poll (encontra o select certo diretamente)
        if causa_raiz_sel_id:
            result = await page.evaluate(f"""
                (() => {{
                    const norm = s => s.normalize('NFD').replace(/[\\u0300-\\u036f]/g,'').toLowerCase().trim();
                    const sel = document.getElementById({json.dumps(causa_raiz_sel_id)});
                    if (!sel) return 'no-sel';
                    const wantedN = norm({json.dumps(causa_raiz)});
                    let opt = null;
                    for (const o of sel.options) {{
                        if (norm(o.text || o.textContent) === wantedN) {{ opt = o; break; }}
                    }}
                    if (!opt) {{
                        for (const o of sel.options) {{
                            const t = norm(o.text || o.textContent);
                            if (t.includes(wantedN) || wantedN.includes(t)) {{ opt = o; break; }}
                        }}
                    }}
                    if (!opt) {{
                        const avail = Array.from(sel.options)
                            .map(o => (o.text || o.textContent).trim()).slice(0, 8).join('|');
                        return 'no-opt:' + avail;
                    }}
                    sel.value = opt.value;
                    const widget = sel.closest('.ui-selectonemenu');
                    if (widget) {{
                        const lbl = widget.querySelector('.ui-selectonemenu-label');
                        if (lbl) lbl.textContent = (opt.text || opt.textContent).trim();
                    }}
                    sel.dispatchEvent(new Event('change', {{bubbles: true}}));
                    if (window.jQuery) jQuery(sel).trigger('change');
                    if (window.PrimeFaces && PrimeFaces.widgets && widget) {{
                        const wid = widget.id;
                        for (const pw of Object.values(PrimeFaces.widgets)) {{
                            if (!pw) continue;
                            if ((wid && pw.id === wid) ||
                                (pw.input && (pw.input[0] === sel ||
                                 (typeof pw.input.is === 'function' && pw.input.is(sel))))) {{
                                try {{ if (typeof pw.callBehavior === 'function') pw.callBehavior('change'); }} catch(e) {{}}
                                break;
                            }}
                        }}
                    }}
                    return 'ok';
                }})()
            """)
            if result == 'ok':
                log.info("Causa Raiz preenchida via ID cacheado do poll")
                ok_cr = True
            else:
                log.warn(f"Causa Raiz (ID cacheado): {result}")

        # Estrategia 2: fill_select_robusto
        if not ok_cr:
            ok_cr = await fill_select_robusto(
                page, ["Causa raiz", "Causa Raiz"], causa_raiz, log, fallback_id_fragment=None
            )

        # Estrategia 3: _selecionar_dropdown_filtrado
        if not ok_cr:
            log.warn("Causa Raiz: tentando click direto (fallback final)...")
            try:
                ok_cr = await _selecionar_dropdown_filtrado(page, "Causa raiz", causa_raiz, log)
            except Exception as e:
                log.warn(f"Causa Raiz fallback: {e}")

        await page.wait_for_timeout(1000)
    else:
        log.warn("Causa Raiz: vazia na planilha - pulando")"""

content_new = content[:i1] + NEW + content[i2:]

with open(path, "w", encoding="utf-8") as f:
    f.write(content_new)

print(f"OK — substituicao aplicada ({i2-i1} chars removidos, {len(NEW)} chars inseridos)")

# Verifica sintaxe
import py_compile
try:
    py_compile.compile(path, doraise=True)
    print("Sintaxe Python OK")
except py_compile.PyCompileError as e:
    print(f"ERRO SINTAXE: {e}")
