from playwright.sync_api import expect
import os

def abrir_navegador(pw):
    browser = pw.chromium.launch_persistent_context(
    user_data_dir=r"C:\playwright-edge-profile", # Carrega o perfil criado
    channel="msedge", # Abre o Navegador Edge 
    headless=True, # False = Mostra o processo | True = Não mostra
    accept_downloads=True # Permite download no navegador
    )
    return browser

def pagina_aging(browser):
    aging = browser.pages[0]

    # Acessa a url do .env
    url = os.getenv("PBI_URL")

    # Navegar para uma página
    aging.goto(url)
    return aging

def filtros_bi(aging, data_formatada):
    # Data
    data = aging.get_by_role("textbox", name="Data de término. Intervalo de")
    expect(data).to_be_visible
    data.click(click_count=3)
    data.fill(data_formatada)

    # Convênios
    convenios = aging.get_by_role("combobox", name="ConveniosGrupos")
    expect(convenios).to_be_visible
    convenios.click()
    
    aging.get_by_text("Selecionar tudo").click()

    lista_convenios = ["IAMSPE", "IAMSPE - (FATURAMENTO)", "IAMSPE - ENDOSCOPIA (FATURAMENTO)",
                        "IAMSPE-ENDOSCOPIA", "SOCIAL"]

    aging.get_by_role("combobox", name="ConveniosGrupos")
    topo = aging.get_by_text("Selecionar tudo", exact=True)

    for nome in lista_convenios:
        topo.hover()

        for _ in range(20):
            opcao = aging.locator(f':text("{nome}"):visible')

            if opcao.count() > 0:
                opcao.first.click()
                aging.wait_for_timeout(30)
                aging.keyboard.press("Home")
                aging.wait_for_timeout(200)
                break
            
            aging.mouse.wheel(0, 100)
            aging.wait_for_timeout(100)
        else:
            raise Exception(f"Convênio {nome} não encontrado.")
        
    # Conta ADM
    aging.get_by_role("combobox", name="IND_FECHAMENTO_ADM").click()
    aging.get_by_role("option", name="N", exact=True).click()

    # Pegando Data
    aging.get_by_role("button", name="Aging Faturamento").click()
    data_bi = aging.locator(".popper-item label").nth(1).inner_text()

    aging.keyboard.press('Escape')
    return data_bi

def exportar_dados(aging, data_aging):
    # 1. Trocando de aba
    aging.get_by_role("tab", name="Aging analítico").click()

    # 2. Aba de exportar
    aging.get_by_role("group").filter(has_text="Pressionar Enter para explorar os dados Rolar para cima Rolar para baixo Role").get_by_test_id("visual-style").click()
    aging.locator("visual-container-options-menu visual-header-item-container").click()
    aging.get_by_test_id("pbimenu-item.Exportar dados").click()

    #3. Baixar arquivo e salvar na pasta
    with aging.expect_download() as download_info:
        aging.get_by_test_id("export-btn").click()

        download = download_info.value
        download.save_as(r"data\modelo\export.xlsx")