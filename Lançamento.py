import pandas as pd
import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager

# ==============================
# CONFIGURAÇÕES
# ==============================

EXCEL_PATH = r"C:\Users\vinia\OneDrive\Desktop\TESTE\Pendencia_formatado.xlsx"
URL_LOGIN = "https://login.freitasleiloeiro.com.br/Home/Login"

SELECTORS = {
    "campo_pesquisa": (By.ID, "txtPesquisarVeiculos"),
    "botao_pesquisar": (By.ID, "btnPesquisarVeiculos"),
    "aba_leilao": (By.ID, "leilao-tab"),
    "campo_emissao": (By.ID, "ATPVEmitida"),
    "campo_data": (By.ID, "DataEmitidaATPV"),
    "botao_salvar": (By.XPATH, "//button[@onclick='atualizarEmissaoATPVForm();']"),
    "popup_sim": (By.XPATH, "//button[contains(text(),'Sim')]"),
    "popup_ok": (By.ID, "modalJsAlertOk"),  # ID do botão OK
}

# ==============================
# FUNÇÃO PRINCIPAL
# ==============================

def main():
    # 1. Ler planilha
    df = pd.read_excel(EXCEL_PATH)

    # 2. Abrir navegador
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()))
    driver.maximize_window()
    driver.get(URL_LOGIN)

    # 3. Login manual
    input("🚨 Faça login manualmente e pressione ENTER para continuar...")

    # 4. Seleção manual da área
    input("🚨 Após selecionar a área desejada, pressione ENTER para continuar...")

    # 5. Alterna para a nova aba/janela aberta pelo site
    time.sleep(2)
    driver.switch_to.window(driver.window_handles[-1])
    print("✅ Mudou para a nova aba com a área selecionada.")

    # Espera extra para garantir que a página carregou
    time.sleep(5)

    # 6. Loop pelas placas
    for i, row in df.iterrows():
        placa = str(row["placa"]).strip()
        data_emissao_raw = row.get("data_emissao", "")

        # Formata a data se existir
        if pd.notna(data_emissao_raw) and data_emissao_raw != "":
            data_emissao = pd.to_datetime(data_emissao_raw).strftime("%d/%m/%Y")
        else:
            data_emissao = ""

        print(f"➡️ Processando placa: {placa}")

        # Espera pelo campo de pesquisa
        campo = WebDriverWait(driver, 40).until(
            EC.visibility_of_element_located(SELECTORS["campo_pesquisa"])
        )
        campo.clear()
        campo.send_keys(placa)

        driver.find_element(*SELECTORS["botao_pesquisar"]).click()

        # Clica no link correto usando a placa no href
        link_id = WebDriverWait(driver, 40).until(
            EC.element_to_be_clickable(
                (By.XPATH, f"//a[contains(@href,'{placa}')]")
            )
        )
        link_id.click()

        # Aba Leilão
        WebDriverWait(driver, 20).until(
            EC.element_to_be_clickable(SELECTORS["aba_leilao"])
        ).click()

        # Espera extra para a aba carregar totalmente
        time.sleep(2)

        # Emissão ATPV = "Sim" usando Select (dropdown)
        campo_emissao = WebDriverWait(driver, 20).until(
            EC.presence_of_element_located(SELECTORS["campo_emissao"])
        )
        WebDriverWait(driver, 10).until(lambda d: campo_emissao.is_enabled())
        select_emissao = Select(campo_emissao)
        select_emissao.select_by_visible_text("Sim")

        # Data de emissão
        campo_data = driver.find_element(*SELECTORS["campo_data"])
        campo_data.clear()
        campo_data.send_keys(data_emissao)

        # Salvar
        driver.find_element(*SELECTORS["botao_salvar"]).click()

        # Pop-ups
        WebDriverWait(driver, 20).until(
            EC.element_to_be_clickable(SELECTORS["popup_sim"])
        ).click()

        # Pop-up OK usando o ID
        WebDriverWait(driver, 30).until(
            EC.element_to_be_clickable(SELECTORS["popup_ok"])
        ).click()

        # Pequeno delay antes de continuar
        time.sleep(1)

    print("✅ Todos os documentos foram processados!")
    driver.quit()


if __name__ == "__main__":
    main()
