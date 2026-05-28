#!/usr/bin/env python3
"""
Automação para preenchimento de usinas no Acesso Energia.

Lê a aba 'Acesso' da planilha Excel e preenche o formulário em:
https://rzk.acessoenergia.com.br/page/crm/usinas

Uso:
  export EMAIL=seu@email.com
  export PASSWORD=suasenha
  python3 preencher_usinas.py planilha.xlsx
  python3 preencher_usinas.py planilha.xlsx --dry-run
  python3 preencher_usinas.py planilha.xlsx --linha 1
  python3 preencher_usinas.py --inspect
  python3 preencher_usinas.py planilha.xlsx --screenshots  # tira prints de cada etapa
"""

import os
import sys
import time
import argparse
import pandas as pd
from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeout

BASE_URL = "https://rzk.acessoenergia.com.br"
USINAS_URL = f"{BASE_URL}/page/crm/usinas"
NEW_URL = f"{BASE_URL}/page/crm/usinas/new"
LOGIN_URL = f"{BASE_URL}/auth/login"
EXCEL_SHEET = "Acesso"

# Empresa padrão quando não especificada na planilha
# Ajuste o nome conforme necessário (busca parcial, case-insensitive)
EMPRESA_PADRAO_BUSCA = "THOPEN ENERGIA S.A."
EMPRESA_PADRAO_CNPJ = "0001-48"  # Preferência de CNPJ para desambiguar


# ---------------------------------------------------------------------------
# Mapeamento de campos (confirmado via inspeção do formulário)
# ---------------------------------------------------------------------------

# select[name] → coluna da planilha
# Campos autopreenchidos pelo projeto ficam disabled; select_por_texto ignora-os.
SELECTS = {
    "projeto":              "Usina (Projeto)*",     # autopreenchimento: distribuidora, fonte, tipo
    "finalidadeUsina":      "Finalidade",
    "clientePadrao":        "Cliente",
    "reembolsoTUSD":        "Reembolso TUSD",       # Sim/Não
    "situacaoUsina":        "Situação do Projeto/Usina",
    "conectada":            "Conectada",             # Sim/Não
    "construida":           None,                   # sempre "Sim" — ver DEFAULTS_SELECTS
    "idAreaResponsavelUsina": None,                 # ver DEFAULTS_SELECTS
}

# Valores padrão para selects sem coluna na planilha
DEFAULTS_SELECTS = {
    "construida": "Sim",
    "idAreaResponsavelUsina": "Portfólio",
}

# input[name] → coluna da planilha
# nomeUsina  = "Usina (UC/UG)*"       → coluna "Unnamed: 7"  (ex: UFV Goytacazes 5.13)
# nomeUsina2 = "Usina (UC/UG) Apelido"→ coluna "DETALHE UG/ APELIDO"
INPUTS_BY_NAME = {
    "nomeUsina":        "Unnamed: 7",           # Nome combinado projeto.UG
    "nomeUsina2":       "DETALHE UG/ APELIDO",  # Apelido (ex: Atua_ACER_UFV Vila Nova 1.13_Enel RJ)
    "unidadeGeradora":  "Num. UC/UG",           # Número da UC
}

# input[placeholder] → coluna da planilha (confirmado via inspeção)
INPUTS_BY_PLACEHOLDER = {
    "MWm":          "MWm",
    "MWac":         "Mwac",
    "MWp":          "MWp",
    "kWh":          "Geração Estimada (KWh)",   # "Geração Estimada Mensal (kWh)*"
    "Digite a qtde": "Energia Anual*",           # "Energia Anual (kWh)*"
}

# input[name] → (coluna, formato da data)
DATE_FIELDS = {
    "dataCod":      ("COD", "mm/aaaa"),
    "dataHandover": ("Data estimada de Internalização do ativo", "dd/mm/aaaa"),
}


# ---------------------------------------------------------------------------
# Utilitários gerais
# ---------------------------------------------------------------------------

def normalizar_sim_nao(valor):
    if valor is None or (isinstance(valor, float) and pd.isna(valor)):
        return None
    v = str(valor).strip().lower()
    if v in ("sim", "yes", "s", "true", "1"):
        return "Sim"
    if v in ("não", "nao", "no", "n", "false", "0"):
        return "Não"
    return str(valor).strip()


def formatar_data(valor, fmt):
    if valor is None or (isinstance(valor, float) and pd.isna(valor)):
        return None
    s = str(valor).strip()
    if not s or s == "nan":
        return None
    try:
        dt = pd.to_datetime(s)
        return dt.strftime("%m/%Y") if fmt == "mm/aaaa" else dt.strftime("%d/%m/%Y")
    except Exception:
        return s


def vazio(valor):
    return valor is None or (isinstance(valor, float) and pd.isna(valor)) or str(valor).strip() in ("", "nan")


# ---------------------------------------------------------------------------
# Preenchimento de campos individuais
# ---------------------------------------------------------------------------

def select_por_texto(page, name, texto, descricao=""):
    """Seleciona option em <select name='...'> por texto. Pula disabled."""
    if vazio(texto):
        return
    texto_str = str(texto).strip()

    sel_el = page.query_selector(f"select[name='{name}']")
    if not sel_el:
        print(f"  ✗ {descricao or name}: campo não encontrado")
        return

    if sel_el.is_disabled():
        current = sel_el.evaluate("el => el.options[el.selectedIndex]?.text?.trim() || ''")
        print(f"  ↷ {descricao or name}: autopreenchido '{current}'")
        return

    options = sel_el.evaluate(
        "el => Array.from(el.options).map(o => ({value: o.value, text: o.text.trim()}))"
    )

    # Correspondência: exata → case-insensitive → parcial
    for nivel, comparar in [
        ("exato", lambda a, b: a == b),
        ("ci", lambda a, b: a.lower() == b.lower()),
        ("parcial", lambda a, b: b.lower() in a.lower() or a.lower() in b.lower()),
    ]:
        for opt in options:
            if comparar(opt["text"], texto_str):
                sel_el.select_option(value=opt["value"])
                label = "" if nivel == "exato" else f" ({nivel})"
                print(f"  ✓ {descricao or name}{label}: {opt['text']}")
                return

    print(f"  ✗ {descricao or name}: opção '{texto_str}' não encontrada")


def react_fill(page, selector, valor, descricao=""):
    """
    Preenche input de formulário React usando o setter nativo + evento input.
    Funciona para campos de texto simples (nomeUsina, unidadeGeradora, datas).
    Para campos com máscara numérica use keyboard_fill.
    """
    if vazio(valor):
        return
    valor_str = str(valor).strip()

    el = page.query_selector(selector)
    if not el:
        print(f"  ✗ {descricao}: campo não encontrado ({selector})")
        return
    if el.is_disabled():
        return

    page.evaluate(
        """(args) => {
            const el = document.querySelector(args.sel);
            if (!el) return;
            const setter = Object.getOwnPropertyDescriptor(
                window.HTMLInputElement.prototype, 'value'
            ).set;
            setter.call(el, args.val);
            el.dispatchEvent(new Event('input',  { bubbles: true }));
            el.dispatchEvent(new Event('change', { bubbles: true }));
        }""",
        {"sel": selector, "val": valor_str},
    )
    print(f"  ✓ {descricao}: {valor_str}")


def keyboard_fill(page, selector, valor, descricao="", casas_decimais=4):
    """
    Preenche input com máscara numérica via teclado (simulação de digitação).
    Necessário para campos como MWm, MWac, MWp, kWh que usam máscara.
    Arredonda para casas_decimais casas decimais antes de digitar.
    """
    if vazio(valor):
        return

    # Arredonda o valor numérico para evitar casa decimais em excesso
    try:
        num = round(float(str(valor)), casas_decimais)
        valor_str = f"{num:.{casas_decimais}f}".rstrip("0").rstrip(".")
        if "." not in valor_str:
            valor_str = valor_str  # inteiro, ok
    except (ValueError, TypeError):
        valor_str = str(valor).strip()

    el = page.query_selector(selector)
    if not el:
        print(f"  ✗ {descricao}: campo não encontrado ({selector})")
        return
    if el.is_disabled():
        return

    el.click()
    page.keyboard.press("Control+a")
    page.keyboard.press("Delete")
    page.wait_for_timeout(100)
    page.keyboard.type(valor_str, delay=40)
    page.keyboard.press("Tab")
    page.wait_for_timeout(150)
    print(f"  ✓ {descricao}: {valor_str}")


def selecionar_empresa(page, busca=EMPRESA_PADRAO_BUSCA, cnpj_pref=EMPRESA_PADRAO_CNPJ):
    """
    Seleciona a Empresa (CNPJ)* buscando por texto parcial.
    Prefere o CNPJ indicado em cnpj_pref para desambiguar múltiplos resultados.
    """
    sel_el = page.query_selector("select[name='idEmpresa']")
    if not sel_el:
        return

    options = sel_el.evaluate(
        "el => Array.from(el.options).map(o => ({value: o.value, text: o.text.trim()}))"
    )

    candidatos = [o for o in options if busca.lower() in o["text"].lower()]
    if not candidatos:
        print(f"  ✗ Empresa: '{busca}' não encontrada")
        return

    # Prefere o com CNPJ preferido
    match = next((c for c in candidatos if cnpj_pref in c["text"]), candidatos[0])
    sel_el.select_option(value=match["value"])
    print(f"  ✓ Empresa: {match['text']}")


# ---------------------------------------------------------------------------
# Login
# ---------------------------------------------------------------------------

def fazer_login(page, email, password):
    print(f"\n→ Fazendo login como {email}...")
    page.goto(LOGIN_URL, wait_until="networkidle")
    page.fill("#email", email)
    page.fill("#password", password)
    page.click("button[type='button']")
    page.wait_for_timeout(4000)

    if "/auth/login" in page.url:
        raise RuntimeError(
            "Login falhou — verifique EMAIL e PASSWORD nas variáveis de ambiente."
        )
    print(f"  ✓ Login OK (URL: {page.url})")


# ---------------------------------------------------------------------------
# Modo inspeção
# ---------------------------------------------------------------------------

def inspecionar_formulario(page):
    print("\n=== MODO INSPEÇÃO ===")
    page.goto(USINAS_URL, wait_until="networkidle")
    page.wait_for_timeout(2000)

    btn = _encontrar_botao_adicionar(page)
    if btn:
        btn.click()
        page.wait_for_timeout(2000)

    print("\n--- Inputs ---")
    for el in page.query_selector_all("input, select, textarea"):
        tag = el.evaluate("e => e.tagName").lower()
        name = el.get_attribute("name") or ""
        placeholder = el.get_attribute("placeholder") or ""
        type_ = el.get_attribute("type") or ""
        print(f"  <{tag}> name={name!r} type={type_!r} placeholder={placeholder!r}")

    print("\n--- Select options ---")
    for sel in page.query_selector_all("select"):
        name = sel.get_attribute("name") or "?"
        opts = sel.evaluate("el => Array.from(el.options).map(o => o.text.trim()).filter(t => t)")
        print(f"  select[name={name!r}]: {opts[:8]}")

    print("\n=== FIM DA INSPEÇÃO ===")


# ---------------------------------------------------------------------------
# Preenchimento do formulário
# ---------------------------------------------------------------------------

def aguardar_autopreenchimento(page):
    """Espera o projeto ficar disabled (sinal que o auto-fill terminou)."""
    for _ in range(40):
        page.wait_for_timeout(300)
        sel = page.query_selector("select[name='projeto']")
        if sel and sel.is_disabled():
            return
    page.wait_for_timeout(1000)


def fill_form(page, row, screenshots_dir=None, num=0):
    step = 0

    def ss(nome):
        nonlocal step
        if screenshots_dir:
            step += 1
            page.screenshot(
                path=f"{screenshots_dir}/{num:02d}-{step:02d}-{nome}.png",
                full_page=False,
            )

    # 1. Projeto (autopreenchimento dispara aqui)
    projeto_val = row.get("Usina (Projeto)*")
    if not vazio(projeto_val):
        select_por_texto(page, "projeto", projeto_val, "Usina (Projeto)*")
        aguardar_autopreenchimento(page)
        ss("apos-projeto")

    # 2. Demais selects com mapeamento de coluna
    for field_name, col in SELECTS.items():
        if field_name == "projeto":
            continue
        if col is None:
            # Usa valor padrão se definido
            valor = DEFAULTS_SELECTS.get(field_name)
        else:
            valor = row.get(col)

        if field_name in ("conectada", "reembolsoTUSD", "construida"):
            valor = normalizar_sim_nao(valor) if not vazio(valor) else valor

        select_por_texto(page, field_name, valor, col or field_name)

    # 3. Empresa (campo obrigatório, não está na planilha)
    selecionar_empresa(page)

    # 4. Inputs por name
    for field_name, col in INPUTS_BY_NAME.items():
        react_fill(page, f"input[name='{field_name}']", row.get(col), col)

    # 5. Inputs por placeholder (campos com máscara numérica — usar teclado)
    for placeholder, col in INPUTS_BY_PLACEHOLDER.items():
        keyboard_fill(page, f"input[placeholder='{placeholder}']", row.get(col), col)

    # 6. Datas
    for field_name, (col, fmt) in DATE_FIELDS.items():
        data_fmt = formatar_data(row.get(col), fmt)
        react_fill(page, f"input[name='{field_name}']", data_fmt, col)

    ss("formulario-preenchido")


# ---------------------------------------------------------------------------
# Fluxo por linha
# ---------------------------------------------------------------------------

def processar_linha(page, row, linha_num, dry_run=False, screenshots_dir=None):
    usina = row.get("Unnamed: 7") or row.get("UG*") or f"Linha {linha_num}"
    print(f"\n{'='*60}")
    print(f"Linha {linha_num}: {usina}")
    print(f"{'='*60}")

    if dry_run:
        print("  [DRY RUN] dados:")
        for k, v in row.items():
            if not vazio(v):
                print(f"    {k}: {v}")
        return True

    # Navega para o formulário
    page.goto(NEW_URL, wait_until="networkidle")
    page.wait_for_timeout(2000)

    if screenshots_dir:
        page.screenshot(path=f"{screenshots_dir}/{linha_num:02d}-00-formulario-vazio.png")

    fill_form(page, row, screenshots_dir=screenshots_dir, num=linha_num)
    page.wait_for_timeout(500)

    # Submete
    btn = page.query_selector("button:has-text('Salvar')")
    if not btn or not btn.is_visible():
        print("  ✗ Botão 'Salvar' não encontrado")
        return False

    # Monitora requests para detectar a chamada de API
    api_calls = []
    page.on("request", lambda r: api_calls.append(r))

    btn.click()
    page.wait_for_timeout(4000)

    if screenshots_dir:
        page.screenshot(path=f"{screenshots_dir}/{linha_num:02d}-99-apos-salvar.png")

    # Verifica se houve chamada de API
    api_hits = [r for r in api_calls if not any(
        ext in r.url for ext in [".js", ".css", ".png", ".ico", ".woff", "clarity", "fonts"]
    )]
    if api_hits:
        print(f"  → API chamada: {api_hits[0].method} {api_hits[0].url}")

    # Verifica redirecionamento (sucesso) ou permanência (erro)
    if "/new" not in page.url and "/crm/usinas" in page.url:
        print(f"  ✓ '{usina}' adicionada com sucesso! (URL: {page.url})")
        return True

    # Checa mensagens de erro na página
    erros = page.evaluate("""
        () => Array.from(document.querySelectorAll('[class*=error],[class*=Error]'))
              .map(e => e.textContent.trim())
              .filter(t => t && t.length < 300)
    """)
    if erros:
        print(f"  ✗ Erro(s): {erros[:3]}")
        return False

    # Permaneceu em /new mas sem erro detectado
    print(f"  ⚠ URL ainda em /new — verificar manualmente se foi salvo")
    return True


# ---------------------------------------------------------------------------
# Leitura da planilha
# ---------------------------------------------------------------------------

def ler_planilha(caminho):
    print(f"\n→ Lendo: {caminho}")
    df = pd.read_excel(caminho, sheet_name=EXCEL_SHEET, header=0)
    df = df.dropna(how="all")
    print(f"  ✓ {len(df)} linha(s) na aba '{EXCEL_SHEET}'")
    return df


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main():
    parser = argparse.ArgumentParser(
        description="Preenche usinas no Acesso Energia a partir de planilha Excel"
    )
    parser.add_argument("planilha", nargs="?", help="Caminho para o .xlsx")
    parser.add_argument("--inspect",     action="store_true", help="Inspeciona campos do formulário")
    parser.add_argument("--dry-run",     action="store_true", help="Simula sem acessar o site")
    parser.add_argument("--linha",       type=int,            help="Processa apenas a linha N (1-based)")
    parser.add_argument("--show-browser",action="store_true", help="Abre janela visível (xvfb necessário)")
    parser.add_argument("--screenshots", action="store_true", help="Salva prints em ./screenshots/")
    args = parser.parse_args()

    if not args.inspect and not args.planilha:
        parser.print_help()
        sys.exit(1)

    # Dry-run: não precisa de browser
    if args.dry_run and args.planilha:
        df = ler_planilha(args.planilha)
        if args.linha:
            df = df.iloc[[args.linha - 1]]
        for i, (_, row) in enumerate(df.iterrows(), start=1):
            processar_linha(None, row.to_dict(), i, dry_run=True)
        return

    email = os.environ.get("EMAIL", "")
    password = os.environ.get("PASSWORD", "")
    if not email or not password:
        print(
            "ERRO: defina EMAIL e PASSWORD como variáveis de ambiente.\n"
            "  export EMAIL=seu@email.com\n"
            "  export PASSWORD=suasenha"
        )
        sys.exit(1)

    screenshots_dir = None
    if args.screenshots:
        screenshots_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "screenshots")
        os.makedirs(screenshots_dir, exist_ok=True)
        print(f"  📸 Screenshots em: {screenshots_dir}")

    headless = not args.show_browser
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=headless, slow_mo=80)
        context = browser.new_context(
            ignore_https_errors=True,
            viewport={"width": 1280, "height": 900},
        )
        page = context.new_page()
        page.set_default_timeout(20000)

        try:
            fazer_login(page, email, password)

            if args.inspect:
                inspecionar_formulario(page)
                return

            df = ler_planilha(args.planilha)
            if args.linha:
                df = df.iloc[[args.linha - 1]]
                print(f"  → Apenas linha {args.linha}")

            ok = falhou = 0
            for i, (_, row) in enumerate(df.iterrows(), start=1):
                sucesso = processar_linha(
                    page, row.to_dict(), i,
                    screenshots_dir=screenshots_dir,
                )
                if sucesso:
                    ok += 1
                else:
                    falhou += 1
                time.sleep(1)

            print(f"\n{'='*60}")
            print(f"Resultado: {ok} OK, {falhou} falha(s).")

        except RuntimeError as e:
            print(f"\nERRO: {e}")
            sys.exit(1)
        except KeyboardInterrupt:
            print("\nInterrompido.")
        finally:
            browser.close()


if __name__ == "__main__":
    main()
