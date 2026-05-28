"""
=============================================================
  AUTOMAÇÃO - PREENCHER USINAS NO ACESSO ENERGIA
  Versão para rodar diretamente no PyCharm
=============================================================

COMO USAR:
  1. Instale as dependências (uma vez só):
       pip install playwright pandas openpyxl
       python -m playwright install chromium

  2. Edite a seção CONFIG abaixo com:
       - suas credenciais
       - caminho da planilha
       - opções de execução

  3. Clique em Run (Shift+F10)
=============================================================
"""

# ============================================================
#  CONFIG — edite aqui antes de rodar
# ============================================================

EMAIL    = "lucas.castro@thopenergy.com.br"
PASSWORD = "krkbdoya"

# Caminho completo para a planilha .xlsx
PLANILHA = "C:/Users/LucasSilvaCastro/OneDrive - Thopen/Área de Trabalho/Dados para ACESSO _ Portfólio.xlsx"

# None = processa todas as linhas
# 1    = apenas a primeira linha
# 2    = apenas a segunda linha
APENAS_LINHA = 1

# 1 = começa da linha 1 (padrão)
# 2 = pula a linha 1 e começa da 2
DE_LINHA = 1

# True  = abre janela do navegador (você vê o que está acontecendo)
# False = roda invisível (mais rápido)
MOSTRAR_NAVEGADOR = True

# True  = salva screenshots na pasta ./screenshots/
SALVAR_SCREENSHOTS = True

# True  = simula sem enviar dados ao site (para testar)
DRY_RUN = False

# ============================================================
#  Empresa padrão (campo obrigatório não presente na planilha)
# ============================================================
EMPRESA_BUSCA = "THOPEN ENERGIA S.A."
EMPRESA_CNPJ  = "0001-48"

# ============================================================
#  NÃO EDITE ABAIXO DESTA LINHA
# ============================================================

import os
import sys
import time
import pandas as pd
from playwright.sync_api import sync_playwright

BASE_URL   = "https://rzk.acessoenergia.com.br"
NEW_URL    = f"{BASE_URL}/page/crm/usinas/new"
LOGIN_URL  = f"{BASE_URL}/auth/login"
EXCEL_SHEET = "Acesso"

SELECTS = {
    "projeto":               "Usina (Projeto)*",
    "finalidadeUsina":       "Finalidade",
    "clientePadrao":         "Cliente",
    "reembolsoTUSD":         "Reembolso TUSD",
    "situacaoUsina":         "Situação do Projeto/Usina",
    "conectada":             "Conectada",
    "construida":            None,
    "idAreaResponsavelUsina": None,
}

DEFAULTS_SELECTS = {
    "construida":            "Sim",
    "idAreaResponsavelUsina": "Portfólio",
}

INPUTS_BY_NAME = {
    "nomeUsina":       "Unnamed: 7",
    "nomeUsina2":      "DETALHE UG/ APELIDO",
    "unidadeGeradora": "Num. UC/UG",
}

INPUTS_BY_PLACEHOLDER = {
    "MWm":           "MWm",
    "MWac":          "Mwac",
    "MWp":           "MWp",
    "kWh":           "Geração Estimada (KWh)",
    "Digite a qtde": "Energia Anual*",
}

DATE_FIELDS = {
    "dataCod":      ("COD", "mm/aaaa"),
    "dataHandover": ("Data estimada de Internalização do ativo", "dd/mm/aaaa"),
}


# --- Utilitários ---

def vazio(v):
    return v is None or (isinstance(v, float) and pd.isna(v)) or str(v).strip() in ("", "nan")

def normalizar_sim_nao(v):
    if vazio(v): return None
    s = str(v).strip().lower()
    if s in ("sim", "yes", "s", "true", "1"): return "Sim"
    if s in ("não", "nao", "no", "n", "false", "0"): return "Não"
    return str(v).strip()

def formatar_data(v, fmt):
    if vazio(v): return None
    try:
        dt = pd.to_datetime(str(v).strip())
        return dt.strftime("%m/%Y") if fmt == "mm/aaaa" else dt.strftime("%d/%m/%Y")
    except Exception:
        return str(v).strip()


# --- Preenchimento ---

def select_por_texto(page, name, texto, descricao=""):
    if vazio(texto): return
    texto_str = str(texto).strip()
    el = page.query_selector(f"select[name='{name}']")
    if not el:
        print(f"  ✗ {descricao or name}: não encontrado"); return
    if el.is_disabled():
        atual = el.evaluate("el => el.options[el.selectedIndex]?.text?.trim() || ''")
        print(f"  ↷ {descricao or name}: autopreenchido '{atual}'"); return
    opts = el.evaluate("el => Array.from(el.options).map(o => ({value:o.value, text:o.text.trim()}))")
    for nivel, fn in [
        ("", lambda a, b: a == b),
        (" (ci)", lambda a, b: a.lower() == b.lower()),
        (" (parcial)", lambda a, b: b.lower() in a.lower() or a.lower() in b.lower()),
    ]:
        for o in opts:
            if fn(o["text"], texto_str):
                el.select_option(value=o["value"])
                print(f"  ✓ {descricao or name}{nivel}: {o['text']}"); return
    print(f"  ✗ {descricao or name}: '{texto_str}' não encontrado")

def react_fill(page, selector, valor, descricao=""):
    if vazio(valor): return
    el = page.query_selector(selector)
    if not el or el.is_disabled():
        if not el: print(f"  ✗ {descricao}: campo não encontrado"); return
        return
    page.evaluate("""(args) => {
        const el = document.querySelector(args.sel);
        if (!el) return;
        const s = Object.getOwnPropertyDescriptor(window.HTMLInputElement.prototype,'value').set;
        s.call(el, args.val);
        el.dispatchEvent(new Event('input',  {bubbles:true}));
        el.dispatchEvent(new Event('change', {bubbles:true}));
    }""", {"sel": selector, "val": str(valor).strip()})
    print(f"  ✓ {descricao}: {valor}")

def keyboard_fill(page, selector, valor, descricao="", casas=4):
    if vazio(valor): return
    el = page.query_selector(selector)
    if not el or el.is_disabled():
        if not el: print(f"  ✗ {descricao}: campo não encontrado"); return
        return
    try:
        num = round(float(str(valor)), casas)
        val_str = f"{num:.{casas}f}".rstrip("0").rstrip(".")
    except Exception:
        val_str = str(valor).strip()
    el.click()
    page.keyboard.press("Control+a")
    page.keyboard.press("Delete")
    page.wait_for_timeout(100)
    page.keyboard.type(val_str, delay=40)
    page.keyboard.press("Tab")
    page.wait_for_timeout(150)
    print(f"  ✓ {descricao}: {val_str}")

def selecionar_empresa(page):
    el = page.query_selector("select[name='idEmpresa']")
    if not el: return
    opts = el.evaluate("el => Array.from(el.options).map(o => ({value:o.value, text:o.text.trim()}))")
    cands = [o for o in opts if EMPRESA_BUSCA.lower() in o["text"].lower()]
    if not cands:
        print(f"  ✗ Empresa '{EMPRESA_BUSCA}' não encontrada"); return
    match = next((c for c in cands if EMPRESA_CNPJ in c["text"]), cands[0])
    el.select_option(value=match["value"])
    print(f"  ✓ Empresa: {match['text']}")

def aguardar_autopreenchimento(page):
    for _ in range(40):
        page.wait_for_timeout(300)
        sel = page.query_selector("select[name='projeto']")
        if sel and sel.is_disabled(): return
    page.wait_for_timeout(1000)

def fill_form(page, row, screenshots_dir, linha_num):
    step = [0]
    def ss(nome):
        if screenshots_dir:
            step[0] += 1
            page.screenshot(path=f"{screenshots_dir}/{linha_num:02d}-{step[0]:02d}-{nome}.png", full_page=False)

    # 1. Projeto
    proj = row.get("Usina (Projeto)*")
    if not vazio(proj):
        select_por_texto(page, "projeto", proj, "Usina (Projeto)*")
        aguardar_autopreenchimento(page)
        ss("apos-projeto")

    # 2. Selects
    for field, col in SELECTS.items():
        if field == "projeto": continue
        valor = DEFAULTS_SELECTS.get(field) if col is None else row.get(col)
        if field in ("conectada", "reembolsoTUSD", "construida"):
            valor = normalizar_sim_nao(valor) if not vazio(valor) else valor
        select_por_texto(page, field, valor, col or field)

    # 3. Empresa
    selecionar_empresa(page)

    # 4. Inputs texto
    for field, col in INPUTS_BY_NAME.items():
        react_fill(page, f"input[name='{field}']", row.get(col), col)

    # 5. Inputs numéricos com máscara
    for placeholder, col in INPUTS_BY_PLACEHOLDER.items():
        keyboard_fill(page, f"input[placeholder='{placeholder}']", row.get(col), col)

    # 6. Datas
    for field, (col, fmt) in DATE_FIELDS.items():
        react_fill(page, f"input[name='{field}']", formatar_data(row.get(col), fmt), col)

    ss("formulario-preenchido")


# --- Fluxo principal ---

def processar_linha(page, row, linha_num, screenshots_dir):
    usina = row.get("Unnamed: 7") or row.get("UG*") or f"Linha {linha_num}"
    print(f"\n{'='*60}")
    print(f"Linha {linha_num}: {usina}")
    print(f"{'='*60}")

    if DRY_RUN:
        print("  [DRY RUN] dados:")
        for k, v in row.items():
            if not vazio(v): print(f"    {k}: {v}")
        return True

    page.goto(NEW_URL, wait_until="networkidle")
    page.wait_for_timeout(2000)

    if screenshots_dir:
        page.screenshot(path=f"{screenshots_dir}/{linha_num:02d}-00-vazio.png")

    fill_form(page, row, screenshots_dir, linha_num)
    page.wait_for_timeout(500)

    btn = page.query_selector("button:has-text('Salvar')")
    if not btn or not btn.is_visible():
        print("  ✗ Botão 'Salvar' não encontrado"); return False

    api_calls = []
    page.on("request", lambda r: api_calls.append(r))
    btn.click()
    page.wait_for_timeout(4000)

    if screenshots_dir:
        page.screenshot(path=f"{screenshots_dir}/{linha_num:02d}-99-resultado.png")

    hits = [r for r in api_calls if not any(
        x in r.url for x in [".js", ".css", ".png", ".ico", ".woff", "clarity", "fonts"]
    )]
    if hits:
        print(f"  → API: {hits[0].method} {hits[0].url}")

    if "/new" not in page.url and "/crm/usinas" in page.url:
        print(f"  ✓ '{usina}' adicionada com sucesso!")
        return True

    erros = page.evaluate("""
        () => Array.from(document.querySelectorAll('[class*=error],[class*=Error]'))
              .map(e => e.textContent.trim()).filter(t => t && t.length < 300)
    """)
    if erros:
        print(f"  ✗ Erro: {erros[:2]}"); return False

    print(f"  ⚠ Verificar manualmente se foi salvo")
    return True


# --- Entry point ---

def main():
    print("=" * 60)
    print("  AUTOMAÇÃO - ACESSO ENERGIA - PREENCHIMENTO DE USINAS")
    print("=" * 60)
    print(f"  Planilha : {PLANILHA}")
    print(f"  Navegador: {'visível' if MOSTRAR_NAVEGADOR else 'headless'}")
    print(f"  Modo     : {'DRY RUN (simulação)' if DRY_RUN else 'REAL (vai cadastrar)'}")
    print("=" * 60)

    if not os.path.exists(PLANILHA):
        print(f"\nERRO: Planilha não encontrada: {PLANILHA}")
        print("Verifique o caminho na variável PLANILHA no início do script.")
        return

    print(f"\n→ Lendo planilha...")
    df = pd.read_excel(PLANILHA, sheet_name=EXCEL_SHEET, header=0)
    df = df.dropna(how="all")
    print(f"  ✓ {len(df)} linhas encontradas na aba '{EXCEL_SHEET}'")

    if APENAS_LINHA:
        df = df.iloc[[APENAS_LINHA - 1]]
        print(f"  → Apenas linha {APENAS_LINHA}")
    elif DE_LINHA > 1:
        df = df.iloc[DE_LINHA - 1:]
        print(f"  → A partir da linha {DE_LINHA}")

    if DRY_RUN:
        start = DE_LINHA if not APENAS_LINHA else 1
        for i, (_, row) in enumerate(df.iterrows(), start=start):
            processar_linha(None, row.to_dict(), i, None)
        return

    screenshots_dir = None
    if SALVAR_SCREENSHOTS:
        screenshots_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "screenshots")
        os.makedirs(screenshots_dir, exist_ok=True)
        print(f"  📸 Screenshots em: {screenshots_dir}")

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=not MOSTRAR_NAVEGADOR, slow_mo=80)
        context = browser.new_context(
            ignore_https_errors=True,
            viewport={"width": 1280, "height": 900},
        )
        page = context.new_page()
        page.set_default_timeout(20000)

        try:
            # Login
            print(f"\n→ Fazendo login como {EMAIL}...")
            page.goto(LOGIN_URL, wait_until="networkidle")
            page.fill("#email", EMAIL)
            page.fill("#password", PASSWORD)
            page.click("button[type='button']")
            page.wait_for_timeout(4000)

            if "/auth/login" in page.url:
                print("\nERRO: Login falhou — verifique EMAIL e PASSWORD no topo do script.")
                return

            print(f"  ✓ Login OK")

            start = DE_LINHA if not APENAS_LINHA else 1
            ok = falhou = 0
            for i, (_, row) in enumerate(df.iterrows(), start=start):
                sucesso = processar_linha(page, row.to_dict(), i, screenshots_dir)
                if sucesso: ok += 1
                else:       falhou += 1
                time.sleep(1)

            print(f"\n{'='*60}")
            print(f"CONCLUÍDO: {ok} adicionada(s), {falhou} falha(s).")
            print(f"{'='*60}")

        except KeyboardInterrupt:
            print("\nInterrompido pelo usuário.")
        finally:
            if MOSTRAR_NAVEGADOR:
                input("\nPressione Enter para fechar o navegador...")
            browser.close()


if __name__ == "__main__":
    main()
