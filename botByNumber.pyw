from playwright.sync_api import sync_playwright
import pandas as pd

nome_arquivo = "Lista CT-e.xlsx"
df = pd.read_excel(nome_arquivo)

with sync_playwright() as p:
    def waitTimeout(seconds):
        pageBling.wait_for_timeout(seconds)
    browser   = p.chromium.launch(headless=False)
    pageBling = browser.new_page()
    pageBling.goto("https://www.bling.com.br/ctes.php#list")
    waitTimeout(2000)
    pageBling.fill("input[id='username']", "") # Lembrar sempre de mascarar a senha e o e-mail pro GitHub...
    waitTimeout(1000)
    pageBling.fill("input[id='senha']", "")
    waitTimeout(1000)
    pageBling.click("button[name='enviar']")
    waitTimeout(1000)
    # warning_full_memory_delete = pageBling.locator("xpath=//html/body/div[5]/div[1]/div/i")
    # if(warning_full_memory_delete):
    #     warning_full_memory_delete.click()
    waitTimeout(2000)
    clear_period        = pageBling.locator("xpath=//html/body/div[6]/div[4]/div[2]/div[1]/div[2]/span[2]/span/span[2]")
    clear_period.click()
    waitTimeout(1000)
    for index, row in df.iterrows():
        input_pesquisa  = pageBling.locator("xpath=//html/body/div[6]/div[4]/div[2]/div[1]/div[1]/div/div[1]/input")
        waitTimeout(2000)
        input_pesquisa.fill(str(row["CT-e"]))
        waitTimeout(2000)
        button_search   = pageBling.locator("xpath=//html/body/div[6]/div[4]/div[2]/div[1]/div[1]/div/div[1]/div/div/div") 
        if (button_search):
            button_search.click()
        waitTimeout(2000)
        try:
            row_item =    pageBling.locator("xpath=//html/body/div[6]/div[4]/div[2]/div[2]/table/tbody/tr")
            row_item.click()
        except:
            row_item = pageBling.locator("xpath=//html/body/div[6]/div[4]/div[2]/div[2]/table/tbody/tr[1]")
            row_item.click()
        waitTimeout(1000)
        ufEmit       =    pageBling.input_value("select[id=cte_emit_UF]")
        ufRem        =    pageBling.input_value("select[id=cte_rem_UF]")
        cfop         =    pageBling.input_value("input[id='cte_CFOP']")
        if (ufEmit == "ES" and ufRem == "ES"):
            waitTimeout(2000)
            pageBling.fill("input[id='cte_CFOP']", "1353")
            pageBling.wait_for_timeout(3000)
            try:
                pageBling.click("button[id='botaoSalvar']")
            except:
                pageBling.click("button[id='botaoCancelar']")
            waitTimeout(3000)
        elif (ufEmit != "ES"):
            pageBling.fill("input[id='cte_CFOP']", "2353")
            pageBling.wait_for_timeout(3000)
            try:
                pageBling.click("button[id='botaoSalvar']")
            except:
                pageBling.click("button[id='botaoCancelar']")
            pageBling.wait_for_timeout(3000)
        elif (ufRem != "ES"):
            pageBling.fill("input[id='cte_CFOP']", "2353")
            pageBling.wait_for_timeout(3000)
            try:
                pageBling.click("button[id='botaoSalvar']")
            except:
                pageBling.click("button[id='botaoCancelar']")
            pageBling.wait_for_timeout(3000)
        else:
            pageBling.click("button[id='botaoCancelar']")
            pageBling.wait_for_timeout(3000)
        waitTimeout(3000)



