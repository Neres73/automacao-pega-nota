"""
Coupa - Recebimento de Requisições

Lê números de requisição da planilha Excel e realiza o recebimento no Coupa.

  Primeiro uso (login manual):   python coupa_recebimento.py --login
  Uso normal (já logado):        python coupa_recebimento.py

Configure as variáveis no arquivo .env (veja .env.example)
"""

import os
import re
import sys
from datetime import date
from pathlib import Path

import openpyxl
from dotenv import load_dotenv
from playwright.sync_api import (
    sync_playwright,
    expect,
    Page,
    BrowserContext,
    TimeoutError as PlaywrightTimeoutError,
)

load_dotenv()

COUPA_URL     = os.getenv("COUPA_URL", "https://gpabr.coupahost.com/user/home")
EXCEL_PATH    = Path(r"C:\Users\628563\OneDrive\Onedrive - GPA\Área de Trabalho\Pasta1.xlsx")
COLUNA_REQ    = int(os.getenv("COLUNA_REQ", "6"))
USER_DATA_DIR = Path(os.getenv("USER_DATA_DIR", "./perfil_recebimento")).resolve()
USER_DATA_DIR.mkdir(exist_ok=True, parents=True)

# Timeout padrão (ms) para esperas por condição — substitui os cooldowns fixos
ESPERA_ENVIO = 15_000


# ============================================================================
# LEITURA DA PLANILHA
# ============================================================================

def ler_requisicoes_excel() -> list[str]:
    wb = openpyxl.load_workbook(EXCEL_PATH, data_only=True)
    ws = wb.active

    requisicoes = []
    for row in ws.iter_rows(min_row=2):
        celula = row[COLUNA_REQ - 1]
        if celula.value:
            requisicoes.append(str(celula.value).strip())

    vistas = set()
    unicas = []
    for r in requisicoes:
        if r not in vistas:
            vistas.add(r)
            unicas.append(r)

    print(f"{len(requisicoes)} linha(s) lida(s) → {len(unicas)} requisição(ões) única(s) a processar.")
    return unicas


# ============================================================================
# ETAPA 1 — Navegar para Requisições
# ============================================================================

def abrir_requisicoes(page: Page) -> None:
    print("Abrindo Coupa...")
    page.goto(COUPA_URL, wait_until="load")

    print("Clicando em Requisições / Procurement...")
    icone = page.locator("img[alt='Requests / Procurement']")
    icone.wait_for(timeout=15_000)
    icone.click()
    page.wait_for_load_state("load")


# ============================================================================
# ETAPA 2 — Processar cada requisição
# ============================================================================

def selecionar_max_itens_por_pagina(page: Page) -> None:
    """Seleciona a maior opção de itens por página, se houver paginação."""
    opcoes = page.locator("a[aria-label*='por página']")
    if opcoes.count() == 0:
        return

    maior_valor = -1
    maior_label = ""
    for i in range(opcoes.count()):
        label = opcoes.nth(i).get_attribute("aria-label") or ""
        nums = re.findall(r"\d+", label)
        if nums:
            val = int(nums[0])
            if val > maior_valor:
                maior_valor = val
                maior_label = label

    if maior_label:
        maior_opcao = page.locator(f"a[aria-label='{maior_label}']").first
        if "selected" not in (maior_opcao.get_attribute("class") or ""):
            print(f"  Selecionando {maior_valor} itens por página...")
            maior_opcao.click()
            page.wait_for_load_state("networkidle")


def coletar_line_ids(page: Page) -> list[str]:
    """Coleta os IDs de linha a partir dos campos de data presentes na tela."""
    line_ids = []
    for el in page.locator("input[name*='receipt_date']").all():
        name = el.get_attribute("name") or ""
        m = re.search(r"\[(\d+)\]", name)
        if m:
            line_ids.append(m.group(1))
    return line_ids


def processar_linha(page: Page, line_id: str, hoje: str) -> str | None:
    """
    Processa uma única linha de pedido. Data e checkbox usam seletores
    únicos por line_id; o número do pedido é lido do <tr> ancestral mais
    próximo do checkbox — nunca por índice global, que fica desatualizado
    quando o Coupa re-renderiza a tabela após cada envio.

    Retorna o número do pedido enviado, ou None se a linha já estava aceita.
    """
    checkbox = page.locator(f"#order_line_{line_id}_receive")
    if checkbox.count() == 0:
        print("     ⏭️  Já aceito, pulando.")
        return None

    # Linha da tabela: o <tr> ancestral MAIS PRÓXIMO do checkbox.
    # (Não usar page.locator("tr", has=...) — tabelas aninhadas do Coupa
    # fazem esse locator casar com vários <tr> e o modo estrito quebra.)
    linha = checkbox.locator("xpath=ancestor::tr[1]")

    # 1) Data de recebimento — seletor único por line_id, como no original
    campo_data = page.locator(f"input[name='order_line[{line_id}][receipt_date]']")
    campo_data.fill(hoje, force=True)
    campo_data.press("Tab")
    page.wait_for_timeout(300)

    # 2) Checkbox de recebimento
    if not checkbox.is_checked():
        checkbox.check(force=True)
    expect(checkbox).to_be_checked(timeout=5_000)

    # 3) Número do pedido — lido da própria linha, não por nth() global
    numero = linha.locator("span.dt_open_link").first.inner_text().strip()
    print(f"     Nº pedido: {numero}")

    # 4) Número da nota — tenta dentro da linha; se não existir ali
    #    (campo em painel/modal), usa o primeiro da página
    campo_nota = linha.locator("input[aria-label='Número da nota']")
    if campo_nota.count() == 0:
        campo_nota = page.locator("input[aria-label='Número da nota']")
    campo_nota.first.fill(numero)
    page.wait_for_timeout(300)

    # 5) Enviar — mesmo esquema: prioriza o botão da linha
    botao_enviar = linha.locator("span:text-is('Enviar')")
    if botao_enviar.count() == 0:
        botao_enviar = page.locator("span:text-is('Enviar')")
    botao_enviar.first.click(force=True)

    # 6) Espera DETERMINÍSTICA pela conclusão do envio: o checkbox da
    #    linha sai do DOM (ou a linha inteira é re-renderizada) quando o
    #    Coupa confirma o recebimento. Substitui o wait_for_timeout(1000).
    try:
        checkbox.wait_for(state="detached", timeout=ESPERA_ENVIO)
    except PlaywrightTimeoutError:
        # Fallback: alguns layouts apenas desabilitam a linha em vez de
        # removê-la; nesse caso, aguarda a rede estabilizar.
        page.wait_for_load_state("networkidle")

    print("     ✅ Enviado.")
    return numero


def processar_requisicao(page: Page, numero_req: str) -> tuple[list, list]:
    hoje        = date.today().strftime("%d/%m/%Y")
    pedidos_ok  = []
    pedidos_err = []

    print(f"  Pesquisando requisição {numero_req}...")
    campo_busca = page.locator("input#sf_requisition_header")
    campo_busca.clear()
    campo_busca.fill(numero_req)
    campo_busca.press("Enter")
    page.wait_for_load_state("load")

    print(f"  Clicando em 'Receber contra requisição {numero_req}'...")
    botao_receber = page.locator(f"img[id='receive_requisition_{numero_req}']")
    botao_receber.wait_for(timeout=10_000)
    botao_receber.click()
    page.wait_for_load_state("networkidle")

    selecionar_max_itens_por_pagina(page)

    line_ids = coletar_line_ids(page)
    total_pedidos = len(line_ids)
    print(f"  {total_pedidos} pedido(s) encontrado(s) nessa requisição.")

    for idx, line_id in enumerate(line_ids):
        print(f"  → Pedido {idx + 1}/{total_pedidos} (linha {line_id})")
        try:
            numero = processar_linha(page, line_id, hoje)
            if numero:
                pedidos_ok.append(numero)
        except Exception as e:
            print(f"     ❌ Erro no pedido (linha {line_id}): {e}")
            pedidos_err.append((line_id, str(e)))

    print(f"  ✅ Requisição {numero_req} concluída.")
    return pedidos_ok, pedidos_err


# ============================================================================
# EXECUÇÃO PRINCIPAL
# ============================================================================

def fazer_login_manual(context: BrowserContext) -> None:
    page = context.new_page()
    page.goto(COUPA_URL)
    print("\n>>> Faça login no Coupa nessa janela do Edge.")
    input("\nDepois de logar, volte aqui e aperte Enter para salvar a sessão...")
    print("Sessão salva em:", USER_DATA_DIR)


def main() -> None:
    modo_login = "--login" in sys.argv

    with sync_playwright() as p:
        context = p.chromium.launch_persistent_context(
            user_data_dir=str(USER_DATA_DIR),
            channel="msedge",
            headless=False,
        )

        if modo_login:
            fazer_login_manual(context)
            context.close()
            return

        requisicoes = ler_requisicoes_excel()
        if not requisicoes:
            print("Nenhuma requisição encontrada na planilha. Encerrando.")
            context.close()
            return

        page = context.pages[0] if context.pages else context.new_page()
        abrir_requisicoes(page)

        todos_ok  = []
        todos_err = []

        for numero_req in requisicoes:
            print(f"\nProcessando requisição {numero_req}...")
            try:
                ok, err = processar_requisicao(page, numero_req)
                todos_ok.extend(ok)
                todos_err.extend(err)
            except Exception as e:
                print(f"  ❌ Erro geral na requisição {numero_req}: {e}")
                todos_err.append((f"req:{numero_req}", str(e)))

            # Volta para a lista de requisições antes da próxima — sempre
            abrir_requisicoes(page)

        total = len(todos_ok) + len(todos_err)
        print("\n" + "=" * 52)
        print("📋  RELATÓRIO FINAL — RECEBIMENTO")
        print("=" * 52)
        print(f"  Total de pedidos processados : {total}")
        print(f"  ✅ Enviados com sucesso       : {len(todos_ok)}")
        print(f"  ❌ Com erro                  : {len(todos_err)}")
        if todos_ok:
            print("\n  Enviados:")
            for num in todos_ok:
                print(f"    • {num}")
        if todos_err:
            print("\n  Erros:")
            for ref, motivo in todos_err:
                print(f"    • {ref}  →  {motivo}")
        print("=" * 52)

        input("\nPressione ENTER para fechar o browser...")
        context.close()


if __name__ == "__main__":
    main()