"""
Elaw — Cumprimento de Tarefa: Externo Inserir Relatório de Audiência

Fluxo por processo:
  1. Elaw VISEU   → ler dados preenchidos pelo correspondente parceiro
  2. Elaw Carrefour → preencher o formulário "Externo: Inserir Relatório da Audiência"

Notas técnicas (JSF/PrimeFaces):
- Navegação direta por URL não funciona — usar busca global
- Botões via JavaScript (podem estar fora do viewport)
- page.evaluate com string NÃO aceita return no top-level — usar IIFEs
"""

import os
import re
import time
from pathlib import Path
from datetime import datetime
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font

if os.environ.get("RENDER"):
    os.environ["PLAYWRIGHT_BROWSERS_PATH"] = "/opt/render/project/src/.playwright-browsers"

from playwright.sync_api import sync_playwright, TimeoutError as PWTimeout

ELAW_VISEU_URL      = os.environ.get("ELAW_VISEU_URL",      "https://viseuadv.elawio.com.br")
ELAW_CARREFOUR_URL  = os.environ.get("ELAW_CARREFOUR_URL",  "https://carrefour.elaw.com.br")

LOGIN_TIMEOUT  = 120_000
PAGE_TIMEOUT   = 40_000
POLL_ATTEMPTS  = 12
POLL_WAIT      = 2.0

IS_SERVER = bool(os.environ.get("RENDER") or os.environ.get("IS_SERVER"))


# ── Entry point ───────────────────────────────────────────────────────────────

def run_automation(rows: list[dict], log, report_path: Path, state: dict | None = None):
    results = []

    with sync_playwright() as p:
        if IS_SERVER:
            browser = _launch_server(p)
            ctx = browser.new_context(viewport={"width": 1280, "height": 900})
        else:
            chrome_profile = str(Path(__file__).parent / "chrome_profile")
            ctx = p.chromium.launch_persistent_context(
                chrome_profile,
                headless=False,
                viewport={"width": 1280, "height": 900},
                slow_mo=120,
            )

        page_viseu     = ctx.new_page()
        page_carrefour = ctx.new_page()

        log("Abrindo Elaw Viseu...")
        _setup_page(page_viseu, ELAW_VISEU_URL, "VISEU", log)

        log("Abrindo Elaw Carrefour...")
        _setup_page(page_carrefour, ELAW_CARREFOUR_URL, "CARREFOUR", log)

        total = len(rows)
        for i, row in enumerate(rows, 1):
            if state and state.get("paused"):
                log("⏸ Pausado — aguardando retomada...", "warn")
                while state.get("paused"):
                    time.sleep(1)
                log("▶️ Retomando...", "info")

            numero         = str(row.get("numero_processo", "")).strip()
            data_audiencia = str(row.get("data_audiencia", "")).strip()

            log(f"[{i}/{total}] {numero} — {data_audiencia}...")

            for attempt in range(2):
                try:
                    status, detail, observacao = _process_row(
                        page_viseu, page_carrefour, numero, data_audiencia, log
                    )
                    break
                except Exception as e:
                    err_str = str(e)
                    if "Execution context was destroyed" in err_str and attempt == 0:
                        log("  🔁 Contexto destruído, tentando novamente...", "warn")
                        _recover_page(page_viseu,     ELAW_VISEU_URL,     log)
                        _recover_page(page_carrefour, ELAW_CARREFOUR_URL, log)
                    else:
                        status, detail, observacao = "ERRO", err_str[:300], ""
                        _recover_page(page_viseu,     ELAW_VISEU_URL,     log)
                        _recover_page(page_carrefour, ELAW_CARREFOUR_URL, log)
                        break

            row_result = {
                "numero_processo": numero,
                "data_audiencia":  data_audiencia,
                "status":          status,
                "detalhe":         detail,
                "observacao":      observacao,
                "horario":         datetime.now().strftime("%H:%M:%S"),
            }
            results.append(row_result)
            if state is not None:
                state["results"].append(row_result)

            icons = {"OK": "✅", "JÁ CUMPRIDO": "ℹ️", "ERRO": "❌", "OBSERVAÇÃO": "⚠️"}
            css   = {"OK": "ok", "JÁ CUMPRIDO": "ok", "ERRO": "error", "OBSERVAÇÃO": "warn"}
            log(f"  {icons.get(status, '❌')} {status}: {detail}", css.get(status, "error"))
            if observacao:
                log(f"  ⚠️  Observação/Intercorrência detectada — verificação manual necessária", "warn")

        if IS_SERVER:
            browser.close()
        else:
            ctx.close()

    _build_report(results, report_path)
    ok = sum(1 for r in results if r["status"] in ("OK", "JÁ CUMPRIDO", "OBSERVAÇÃO"))
    obs = sum(1 for r in results if r["status"] == "OBSERVAÇÃO")
    log(
        f"Concluído: {ok}/{len(results)} processos com sucesso"
        + (f" — {obs} com observação para revisão manual" if obs else "") + ".",
        "done",
    )


# ── Browser helpers ───────────────────────────────────────────────────────────

def _launch_server(p):
    return p.chromium.launch(
        headless=True,
        args=[
            "--no-sandbox",
            "--disable-dev-shm-usage",
            "--disable-gpu",
            "--disable-setuid-sandbox",
            "--single-process",
        ],
    )


def _setup_page(page, base_url: str, system: str, log):
    """Abre o sistema, faz login se necessário e aguarda a tela principal."""
    page.goto(base_url, wait_until="networkidle", timeout=30_000)

    if _is_login_page(page):
        if IS_SERVER:
            log(f"Fazendo login automático no {system}...", "info")
            _auto_login(page, system, log)
        else:
            log(f"⚠️ Sessão {system} expirada — faça login no browser aberto.", "warn")
            page.wait_for_url(f"**{base_url}/**", timeout=LOGIN_TIMEOUT)
            page.wait_for_load_state("networkidle", timeout=PAGE_TIMEOUT)
            log(f"✅ Login {system} detectado.")
    else:
        log(f"✅ Sessão {system} ativa.")

    page.goto(f"{base_url}/processoList.elaw", wait_until="networkidle", timeout=PAGE_TIMEOUT)
    page.wait_for_selector('[id*="globaSearchAutocomplete_input"]', state="visible", timeout=PAGE_TIMEOUT)


def _auto_login(page, system: str, log):
    if system == "VISEU":
        user = os.environ.get("ELAW_VISEU_USER", "")
        pwd  = os.environ.get("ELAW_VISEU_PASS", "")
        env_vars = "ELAW_VISEU_USER e ELAW_VISEU_PASS"
    else:
        user = os.environ.get("ELAW_CARREFOUR_USER", "")
        pwd  = os.environ.get("ELAW_CARREFOUR_PASS", "")
        env_vars = "ELAW_CARREFOUR_USER e ELAW_CARREFOUR_PASS"

    if not user or not pwd:
        raise Exception(
            f"Variáveis {env_vars} não configuradas. "
            "Adicione-as nas variáveis de ambiente."
        )

    page.wait_for_selector("#username", state="visible", timeout=PAGE_TIMEOUT)
    page.wait_for_selector("#authKey",  state="visible", timeout=PAGE_TIMEOUT)

    log(f"Preenchendo credenciais {system}...")
    page.fill("#username", user, timeout=PAGE_TIMEOUT)
    page.fill("#authKey",  pwd,  timeout=PAGE_TIMEOUT)

    try:
        page.locator("button.ui-button").first.click(force=True, timeout=PAGE_TIMEOUT)
    except Exception:
        page.evaluate("""(() => {
            const btn = document.querySelector('button.ui-button') ||
                        Array.from(document.querySelectorAll('button'))
                            .find(b => b.textContent.trim().includes('Acessar'));
            if (btn) btn.click();
            else { const f = document.querySelector('form'); if (f) f.submit(); }
        })()""")

    page.wait_for_load_state("networkidle", timeout=30_000)

    if _is_login_page(page):
        raise Exception(
            f"Login {system} falhou — verifique {env_vars}."
        )
    log(f"✅ Login {system} concluído.")


def _is_login_page(page) -> bool:
    url = page.url.lower()
    if any(k in url for k in ("login", "signin", "sso", "auth", "microsoftonline")):
        return True
    try:
        return bool(page.evaluate("""(() => {
            const isVisible = el => el && el.offsetParent !== null
                && getComputedStyle(el).display !== 'none'
                && getComputedStyle(el).visibility !== 'hidden';
            return isVisible(document.getElementById('username'))
                || isVisible(document.getElementById('authKey'));
        })()"""))
    except Exception:
        return False


def _recover_page(page, base_url: str, log):
    try:
        page.goto(f"{base_url}/processoList.elaw", wait_until="networkidle", timeout=PAGE_TIMEOUT)
        page.wait_for_selector(
            '[id*="globaSearchAutocomplete_input"]',
            state="visible",
            timeout=PAGE_TIMEOUT,
        )
    except Exception as e:
        log(f"  ⚠️ Recovery falhou: {str(e)[:120]}", "warn")


# ── Fluxo por processo ────────────────────────────────────────────────────────

def _process_row(page_viseu, page_carrefour, numero, data_audiencia, log):
    log("  📖 Lendo dados no Elaw Viseu...")
    viseu_data = _read_viseu_data(page_viseu, numero, data_audiencia)

    log("  ✏️ Preenchendo no Elaw Carrefour...")
    status = _write_carrefour_data(page_carrefour, numero, data_audiencia, viseu_data, log)

    observacao = viseu_data.get("observacao") or ""
    observacao = observacao.strip()

    if status == "JÁ CUMPRIDO":
        return "JÁ CUMPRIDO", "Tarefa já estava cumprida anteriormente", observacao

    if observacao:
        return "OBSERVAÇÃO", "Relatório inserido — verificar observação/intercorrência", observacao

    return "OK", "Relatório inserido com sucesso", observacao


# ── Elaw VISEU — Leitura ──────────────────────────────────────────────────────

def _read_viseu_data(page, numero: str, data_audiencia: str) -> dict:
    _navigate_to_process(page, numero, ELAW_VISEU_URL)
    _click_prazos(page)
    _click_enviar_relatorio_viseu(page, data_audiencia)
    return _extract_viseu_form(page)


def _click_prazos(page):
    """Clica na aba/botão 'Prazos' da página do processo no Elaw Viseu."""
    result = page.evaluate("""(() => {
        // Tenta por texto exato em vários tipos de elemento
        const candidates = Array.from(
            document.querySelectorAll('a, button, li > a, [role="tab"], .ui-menuitem-link')
        );
        const el = candidates.find(e =>
            e.textContent.trim().toLowerCase() === 'prazos'
        );
        if (el) { el.click(); return 'ok'; }

        // Fallback: qualquer elemento cujo texto contém apenas "prazos"
        const fallback = Array.from(document.querySelectorAll('*')).find(e =>
            e.children.length === 0 &&
            e.textContent.trim().toLowerCase() === 'prazos'
        );
        if (fallback) { fallback.click(); return 'ok'; }
        return 'nao_encontrado';
    })()""")
    if result != "ok":
        raise Exception("Botão/aba 'Prazos' não encontrado na página do processo Viseu")
    time.sleep(1.5)
    page.wait_for_load_state("networkidle", timeout=PAGE_TIMEOUT)


def _click_enviar_relatorio_viseu(page, data_audiencia: str):
    """Abre o prazo 'Enviar relatório da audiência' preenchido pelo parceiro."""
    # Normaliza para comparação: pega só DD/MM/YYYY
    date_part = data_audiencia.strip()[:10]

    result = page.evaluate(f"""(() => {{
        const dateHint = '{date_part}';
        const rows = Array.from(document.querySelectorAll('tr, li, .prazo-row, [class*="prazo"]'));

        // Filtra linhas que mencionam "enviar relat" (ignore case)
        const candidates = rows.filter(r =>
            r.textContent.toLowerCase().includes('enviar relat')
        );

        if (candidates.length === 0) return 'nao_encontrado';

        // Se há mais de um, tenta casar pela data da audiência
        let target = candidates[0];
        if (candidates.length > 1 && dateHint) {{
            const byDate = candidates.find(r => r.textContent.includes(dateHint));
            if (byDate) target = byDate;
        }}

        // Tenta clicar num botão de "ver/abrir" dentro da linha;
        // se não houver, clica na própria linha
        const btn = target.querySelector(
            'button[id*="btn"], a[id*="btn"], [title*="Ver"], [title*="Detalhe"], [class*="lupa"]'
        ) || target.querySelector('a, button');

        if (btn) {{ btn.click(); return 'ok_btn'; }}
        target.click();
        return 'ok_row';
    }})()""")

    if result == "nao_encontrado":
        raise Exception(
            "Prazo 'Enviar relatório da audiência' não encontrado no Elaw Viseu. "
            "Verifique se o correspondente já preencheu."
        )
    time.sleep(1.5)
    page.wait_for_load_state("networkidle", timeout=PAGE_TIMEOUT)


def _extract_viseu_form(page) -> dict:
    """Extrai os campos do formulário preenchido pelo correspondente no Viseu."""
    return page.evaluate("""(() => {
        // Auxiliar: lê valor de um campo (input/select/textarea/div readonly)
        function readField(el) {
            if (!el) return null;
            if (el.tagName === 'SELECT') {
                const opt = el.options[el.selectedIndex];
                return opt ? opt.text.trim() : el.value.trim();
            }
            if (el.tagName === 'INPUT' || el.tagName === 'TEXTAREA') {
                return el.value.trim();
            }
            // PrimeFaces: o valor visível fica num span .ui-selectonemenu-label ou similar
            const label = el.querySelector('.ui-selectonemenu-label, .ui-chkbox-label, [class*="label"]');
            if (label) return label.textContent.trim();
            return el.textContent.trim();
        }

        // Tenta encontrar valor pelo texto parcial do label/th
        function findByLabel(keywords) {
            const allLabels = Array.from(
                document.querySelectorAll('label, th, td.label, .ui-outputlabel, [class*="label"]')
            );
            for (const lbl of allLabels) {
                const txt = lbl.textContent.trim().toLowerCase();
                if (keywords.every(k => txt.includes(k))) {
                    // Tenta campo associado via 'for'
                    const forId = lbl.getAttribute('for');
                    if (forId) {
                        const field = document.getElementById(forId);
                        if (field) return readField(field);
                    }
                    // Tenta elemento seguinte na mesma linha de tabela
                    const tr = lbl.closest('tr');
                    if (tr) {
                        const cells = Array.from(tr.querySelectorAll('td'));
                        const lblCell = cells.indexOf(lbl.closest('td'));
                        if (lblCell >= 0 && cells[lblCell + 1]) {
                            const inp = cells[lblCell + 1].querySelector(
                                'input:not([type=hidden]), select, textarea, .ui-selectonemenu, div[id]'
                            );
                            if (inp) return readField(inp);
                        }
                    }
                    // Tenta container pai
                    const container = lbl.closest('.ui-grid-col, .field, .form-group, fieldset');
                    if (container) {
                        const inp = container.querySelector(
                            'input:not([type=hidden]), select, textarea, .ui-selectonemenu'
                        );
                        if (inp && inp !== lbl) return readField(inp);
                    }
                }
            }
            return null;
        }

        // Lê checkboxes Sim/Não pelo texto do label
        function findSimNao(keywords) {
            const val = findByLabel(keywords);
            if (val) return val;

            // Fallback: procura radio/checkbox marcado próximo ao label
            const allInputs = Array.from(document.querySelectorAll('input[type=radio]:checked, input[type=checkbox]:checked'));
            for (const inp of allInputs) {
                const lbl = document.querySelector(`label[for="${inp.id}"]`);
                if (!lbl) continue;
                const row = inp.closest('tr, .field, .form-group');
                if (!row) continue;
                const rowTxt = row.textContent.toLowerCase();
                if (keywords.every(k => rowTxt.includes(k))) {
                    return lbl.textContent.trim() || inp.value;
                }
            }
            return null;
        }

        const dados_representante = findByLabel(['dados', 'representante'])
                                 || findByLabel(['representante']);
        const teve_testemunha     = findSimNao(['testemunha'])
                                 || findByLabel(['testemunha']);
        const resultado_audiencia = findByLabel(['resultado', 'audi'])
                                 || findByLabel(['resultado']);
        const observacao          = findByLabel(['observa'])
                                 || findByLabel(['intercorr']);

        // Campos de testemunha (nome, RG, etc.) — captura todos os campos
        // que aparecem após "Testemunha" quando teve_testemunha === 'Sim'
        let testemunha_nome = null;
        let testemunha_documento = null;
        if (teve_testemunha && teve_testemunha.toLowerCase().includes('sim')) {
            testemunha_nome      = findByLabel(['nome', 'testemunha'])
                                || findByLabel(['testemunha', 'nome']);
            testemunha_documento = findByLabel(['rg', 'testemunha'])
                                || findByLabel(['documento', 'testemunha'])
                                || findByLabel(['cpf', 'testemunha']);
        }

        return {
            dados_representante,
            teve_testemunha,
            testemunha_nome,
            testemunha_documento,
            resultado_audiencia,
            observacao,
        };
    })()""")


# ── Elaw Carrefour — Escrita ──────────────────────────────────────────────────

def _write_carrefour_data(page, numero: str, data_audiencia: str, viseu_data: dict, log) -> str:
    _navigate_to_process(page, numero, ELAW_CARREFOUR_URL)
    _click_pauta_andamento(page)

    task_status = _find_verify_and_open_task(page, data_audiencia, log)
    if task_status == "ja_cumprido":
        return "JÁ CUMPRIDO"

    page.wait_for_load_state("networkidle", timeout=PAGE_TIMEOUT)
    _fill_carrefour_form(page, viseu_data)
    _confirm_task(page)
    return "OK"


def _click_pauta_andamento(page):
    """Clica na aba 'Pauta e Andamento' do processo no Carrefour."""
    result = page.evaluate("""(() => {
        const candidates = Array.from(
            document.querySelectorAll('a, button, li > a, [role="tab"], .ui-menuitem-link')
        );
        const el = candidates.find(e => {
            const t = e.textContent.trim().toLowerCase();
            return t.includes('pauta') && t.includes('andamento');
        });
        if (el) { el.click(); return 'ok'; }
        return 'nao_encontrado';
    })()""")
    if result != "ok":
        raise Exception("Aba 'Pauta e Andamento' não encontrada na página do processo Carrefour")
    time.sleep(1.5)
    page.wait_for_load_state("networkidle", timeout=PAGE_TIMEOUT)


def _find_verify_and_open_task(page, data_audiencia: str, log) -> str:
    """
    Localiza o prazo 'Externo: Inserir Relatório da Audiência' correspondente
    à audiência em questão. Clica na lupa para verificar a data e depois
    no ícone check para abrir o formulário de cumprimento.

    Retorna 'ok' ou 'ja_cumprido'.
    """
    date_part = data_audiencia.strip()[:10]  # DD/MM/YYYY

    # Conta quantos prazos deste tipo existem na tela
    task_count = page.evaluate("""(() => {
        const rows = Array.from(document.querySelectorAll('tr'));
        return rows.filter(r =>
            r.textContent.toLowerCase().includes('externo') &&
            r.textContent.toLowerCase().includes('relat') &&
            r.textContent.toLowerCase().includes('audi')
        ).length;
    })()""")

    if task_count == 0:
        return "ja_cumprido"

    matched_index = None

    if task_count == 1:
        # Apenas um prazo — verifica a data por precaução mas usa o único existente
        ok = _verify_task_date(page, 0, date_part, log)
        if ok or not date_part:
            matched_index = 0
        else:
            raise Exception(
                f"O único prazo 'Externo: Inserir Relatório' encontrado não corresponde "
                f"à data {data_audiencia}. Verifique a planilha."
            )
    else:
        # Múltiplos prazos — itera para casar pela data
        for i in range(task_count):
            if _verify_task_date(page, i, date_part, log):
                matched_index = i
                break
        if matched_index is None:
            raise Exception(
                f"Nenhum dos {task_count} prazos 'Externo: Inserir Relatório' "
                f"corresponde à data {data_audiencia}."
            )

    _click_check_icon(page, matched_index)
    return "ok"


def _verify_task_date(page, row_index: int, date_part: str, log) -> bool:
    """
    Clica na lupa da linha row_index, lê a data/hora do popup,
    fecha o popup e retorna True se a data bater.
    """
    # Clica na lupa
    lupa_result = page.evaluate(f"""(() => {{
        const rows = Array.from(document.querySelectorAll('tr')).filter(r =>
            r.textContent.toLowerCase().includes('externo') &&
            r.textContent.toLowerCase().includes('relat') &&
            r.textContent.toLowerCase().includes('audi')
        );
        const row = rows[{row_index}];
        if (!row) return 'sem_row';

        // Procura botão de lupa/busca/detalhe na linha
        const buttons = Array.from(row.querySelectorAll('button, a'));
        // Primeiro tenta pelo ícone de lupa (PrimeFaces usa class "ui-icon-search" ou "fa-search")
        let btn = buttons.find(b =>
            b.innerHTML.toLowerCase().includes('search') ||
            b.innerHTML.toLowerCase().includes('lupa') ||
            b.innerHTML.toLowerCase().includes('zoom') ||
            b.innerHTML.toLowerCase().includes('eye') ||
            (b.querySelector && b.querySelector('[class*="search"], [class*="lupa"]'))
        );
        // Fallback: último botão que não seja o check
        if (!btn) {{
            btn = buttons.find(b => !b.innerHTML.toLowerCase().includes('check') &&
                                    !b.innerHTML.toLowerCase().includes('tick'));
        }}
        if (!btn && buttons.length > 0) btn = buttons[0];
        if (!btn) return 'sem_lupa';
        btn.click();
        return 'ok';
    }})()""")

    if lupa_result != "ok":
        log(f"  ⚠️ Lupa não encontrada na linha {row_index}: {lupa_result}", "warn")
        return True  # Na dúvida, assume que é o correto se for o único

    time.sleep(1.2)

    # Lê a data do popup/dialog
    popup_text = page.evaluate("""(() => {
        const selectors = [
            '.ui-dialog',
            '.ui-overlaypanel',
            '.ui-tooltip-text',
            '[class*="popup"]',
            '[class*="modal"]',
            '[class*="overlay"]',
        ];
        for (const sel of selectors) {
            const el = document.querySelector(sel);
            if (el && el.offsetParent !== null) return el.textContent;
        }
        // Fallback: conteúdo do body para debug
        return null;
    })()""")

    # Fecha o popup
    page.evaluate("""(() => {
        const closeSelectors = [
            '.ui-dialog-titlebar-close',
            'button[aria-label="Close"]',
        ];
        for (const sel of closeSelectors) {
            const btn = document.querySelector(sel);
            if (btn) { btn.click(); return; }
        }
        // Tenta botão "Fechar" por texto
        const btns = Array.from(document.querySelectorAll('button, a'));
        const fechar = btns.find(b => b.textContent.trim().toLowerCase() === 'fechar');
        if (fechar) fechar.click();
    })()""")
    time.sleep(0.5)

    if popup_text and date_part:
        return date_part in popup_text

    return True  # Sem popup ou sem data para comparar — assume correto


def _click_check_icon(page, row_index: int):
    """Clica no ícone de check (confirmar/cumprir) da linha especificada."""
    result = page.evaluate(f"""(() => {{
        const rows = Array.from(document.querySelectorAll('tr')).filter(r =>
            r.textContent.toLowerCase().includes('externo') &&
            r.textContent.toLowerCase().includes('relat') &&
            r.textContent.toLowerCase().includes('audi')
        );
        const row = rows[{row_index}];
        if (!row) return 'sem_row';

        const buttons = Array.from(row.querySelectorAll('button, a'));
        // Procura pelo check/tick/confirm
        let btn = buttons.find(b =>
            b.innerHTML.toLowerCase().includes('check') ||
            b.innerHTML.toLowerCase().includes('tick') ||
            b.innerHTML.toLowerCase().includes('confirm') ||
            b.id.toLowerCase().includes('confirm') ||
            b.id.toLowerCase().includes('cumprir') ||
            (b.querySelector && b.querySelector('[class*="check"], [class*="confirm"]'))
        );
        // Fallback: o último botão da linha (geralmente o de ação principal)
        if (!btn && buttons.length > 0) btn = buttons[buttons.length - 1];
        if (!btn) return 'sem_check';
        btn.click();
        return btn.id || 'clicado';
    }})()""")

    if result in ("sem_row", "sem_check"):
        raise Exception(f"Botão de check/confirmar não encontrado na linha {row_index}: {result}")

    try:
        page.wait_for_url("**agendamentoContenciosoConfirm.elaw**", timeout=PAGE_TIMEOUT)
    except PWTimeout:
        # Algumas versões do Elaw carregam inline sem mudar URL
        try:
            page.wait_for_selector(
                '[id*="confirmar"], [id*="btnConfirm"], [id*="form"], .ui-dialog',
                state="visible",
                timeout=PAGE_TIMEOUT,
            )
        except PWTimeout:
            raise Exception(
                f"Formulário de cumprimento não carregou após clicar no check "
                f"(URL atual: {page.url})"
            )


# ── Preenchimento do formulário Carrefour ─────────────────────────────────────

def _fill_carrefour_form(page, viseu_data: dict):
    """Preenche os campos do formulário 'Externo: Inserir Relatório da Audiência'."""

    # 1. Preposto designado compareceu → Sim (se há dados do representante no Viseu)
    preposto_compareceu = "Sim" if viseu_data.get("dados_representante") else "Não"
    _select_radio_or_dropdown(page, ["preposto", "compareceu"], preposto_compareceu)
    time.sleep(0.5)

    # 2. Teve Testemunha?
    teve_testemunha = viseu_data.get("teve_testemunha") or "Não"
    _select_radio_or_dropdown(page, ["testemunha"], teve_testemunha)
    time.sleep(0.5)

    # 3. Dados da testemunha (se houve)
    if "sim" in str(teve_testemunha).lower():
        if viseu_data.get("testemunha_nome"):
            _fill_text_field(page, ["nome", "testemunha"], viseu_data["testemunha_nome"])
        if viseu_data.get("testemunha_documento"):
            _fill_text_field(page, ["rg", "testemunha"], viseu_data["testemunha_documento"])
        time.sleep(0.5)

    # 4. Externo – Resultado da Audiência
    resultado = viseu_data.get("resultado_audiencia") or ""
    if resultado:
        _select_radio_or_dropdown(page, ["resultado", "audi"], resultado)
        time.sleep(0.5)


def _select_radio_or_dropdown(page, label_keywords: list[str], value: str):
    """
    Seleciona um valor em radio, checkbox ou dropdown PrimeFaces.
    label_keywords: lista de palavras que identificam o campo pelo texto do label.
    value: valor a selecionar (ex: 'Sim', 'Não', 'Conciliado').
    """
    keywords_js  = str(label_keywords)
    value_lower  = value.strip().lower()

    page.evaluate(f"""(() => {{
        const keywords = {keywords_js};
        const valueLower = '{value_lower}';

        // Encontra o campo pelo label
        function findContainer(keywords) {{
            const allLabels = Array.from(
                document.querySelectorAll('label, th, .ui-outputlabel, [class*="label"]')
            );
            for (const lbl of allLabels) {{
                const txt = lbl.textContent.trim().toLowerCase();
                if (keywords.every(k => txt.includes(k))) {{
                    const row = lbl.closest('tr, .ui-grid-row, .field, .form-group, fieldset, .ui-panelgrid-cell');
                    return row || lbl.parentElement;
                }}
            }}
            return null;
        }}

        const container = findContainer(keywords);
        if (!container) return;

        // Tenta select nativo
        const sel = container.querySelector('select');
        if (sel) {{
            const opt = Array.from(sel.options).find(o =>
                o.text.trim().toLowerCase().includes(valueLower) ||
                o.value.trim().toLowerCase().includes(valueLower)
            );
            if (opt) {{
                sel.value = opt.value;
                sel.dispatchEvent(new Event('change', {{ bubbles: true }}));
            }}
            return;
        }}

        // Tenta radio buttons
        const radios = Array.from(container.querySelectorAll('input[type=radio]'));
        for (const radio of radios) {{
            const lbl = document.querySelector('label[for="' + radio.id + '"]');
            const labelTxt = (lbl ? lbl.textContent : radio.value).trim().toLowerCase();
            if (labelTxt.includes(valueLower)) {{
                radio.click();
                return;
            }}
        }}

        // Tenta PrimeFaces dropdown (.ui-selectonemenu)
        const pfDrop = container.querySelector('.ui-selectonemenu');
        if (pfDrop) {{
            // Abre o dropdown
            const trigger = pfDrop.querySelector('.ui-selectonemenu-trigger');
            if (trigger) trigger.click();
            setTimeout(() => {{
                const panel = document.querySelector('.ui-selectonemenu-panel:not([style*="display: none"])');
                if (!panel) return;
                const items = Array.from(panel.querySelectorAll('li.ui-selectonemenu-item'));
                const target = items.find(i => i.textContent.trim().toLowerCase().includes(valueLower));
                if (target) target.click();
            }}, 400);
        }}
    }})()""")


def _fill_text_field(page, label_keywords: list[str], value: str):
    """Preenche um campo de texto identificado pelo label."""
    keywords_js = str(label_keywords)
    page.evaluate(f"""(() => {{
        const keywords = {keywords_js};
        const value = {repr(value)};

        const allLabels = Array.from(
            document.querySelectorAll('label, th, .ui-outputlabel')
        );
        for (const lbl of allLabels) {{
            const txt = lbl.textContent.trim().toLowerCase();
            if (keywords.every(k => txt.includes(k))) {{
                const forId = lbl.getAttribute('for');
                const field = forId ? document.getElementById(forId) : null;
                const container = lbl.closest('tr, .field, .form-group');
                const inp = field || (container && container.querySelector('input[type=text], textarea'));
                if (inp) {{
                    inp.value = value;
                    inp.dispatchEvent(new Event('input',  {{ bubbles: true }}));
                    inp.dispatchEvent(new Event('change', {{ bubbles: true }}));
                    return;
                }}
            }}
        }}
    }})()""")


def _confirm_task(page):
    """Rola a página e clica em Confirmar."""
    page.evaluate("window.scrollTo(0, document.body.scrollHeight)")
    time.sleep(0.5)

    result = page.evaluate("""(() => {
        // Tenta pelo ID exato do preposto-rpa (mesmo sistema)
        const byId = document.getElementById('btnConfirmaSim');
        if (byId) { byId.click(); return 'ok'; }

        // Tenta qualquer botão "Confirmar"
        const btns = Array.from(document.querySelectorAll('button, input[type="submit"]'));
        const confirmar = btns.find(b =>
            b.textContent.trim().toLowerCase().includes('confirmar') ||
            (b.value && b.value.trim().toLowerCase().includes('confirmar'))
        );
        if (confirmar) { confirmar.click(); return 'ok'; }
        return 'nao_encontrado';
    })()""")

    if result != "ok":
        raise Exception("Botão 'Confirmar' não encontrado no formulário")

    time.sleep(2)
    page.wait_for_load_state("networkidle", timeout=PAGE_TIMEOUT)


# ── Navegação de processo (comum aos dois sistemas) ───────────────────────────

def _navigate_to_process(page, numero: str, base_url: str):
    """Navega até a página do processo usando o autocomplete global."""
    page.wait_for_selector('[id*="globaSearchAutocomplete_input"]', timeout=PAGE_TIMEOUT)

    clicked = False
    for attempt in range(2):
        page.evaluate("""
            const el = document.querySelector('[id*="globaSearchAutocomplete_input"]');
            el.value = '';
            el.focus();
            el.click();
        """)
        time.sleep(0.4)
        page.keyboard.type(numero, delay=65)

        for _ in range(POLL_ATTEMPTS):
            time.sleep(POLL_WAIT)
            result = page.evaluate("""(() => {
                const panel = document.querySelector('[id$="globaSearchAutocomplete_panel"]');
                const items = panel ? panel.querySelectorAll('li') : [];
                if (items.length > 0) { items[0].click(); return 'clicado'; }
                return 'vazio';
            })()""")
            if result == "clicado":
                clicked = True
                break

        if clicked:
            break

    if not clicked:
        raise Exception(f"Autocomplete da busca não abriu para o processo {numero}")

    try:
        page.wait_for_url("**/processoView.elaw**", timeout=PAGE_TIMEOUT)
    except PWTimeout:
        raise Exception(f"Processo {numero} não encontrado no sistema")

    page.wait_for_load_state("networkidle", timeout=PAGE_TIMEOUT)


# ── Relatório Excel ───────────────────────────────────────────────────────────

def _build_report(results: list[dict], path: Path):
    wb = Workbook()
    ws = wb.active
    ws.title = "Resultado"

    headers = ["Número Processo", "Data Audiência", "Status", "Detalhe", "Observação", "Horário"]
    ws.append(headers)
    for cell in ws[1]:
        cell.fill = PatternFill("solid", fgColor="1F4E79")
        cell.font = Font(bold=True, color="FFFFFF")

    fills = {
        "OK":          PatternFill("solid", fgColor="C6EFCE"),
        "JÁ CUMPRIDO": PatternFill("solid", fgColor="DDEBF7"),
        "OBSERVAÇÃO":  PatternFill("solid", fgColor="FFEB9C"),
        "ERRO":        PatternFill("solid", fgColor="FFC7CE"),
    }

    for r in results:
        ws.append([
            r["numero_processo"],
            r["data_audiencia"],
            r["status"],
            r["detalhe"],
            r["observacao"],
            r["horario"],
        ])
        fill = fills.get(r["status"], fills["ERRO"])
        for cell in ws[ws.max_row]:
            cell.fill = fill

    for col, w in zip("ABCDEF", [22, 16, 16, 45, 60, 10]):
        ws.column_dimensions[col].width = w

    wb.save(path)
