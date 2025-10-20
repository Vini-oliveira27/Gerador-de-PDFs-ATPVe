from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import pandas as pd
import time
import os

def automatizar_ecrv_aguardar_campos(nome_planilha, pasta_download=None):
    """
    Sistema que AGUARDA os campos carregarem após navegação
    """
    try:
        pasta_script = os.path.dirname(os.path.abspath(__file__))
        caminho_planilha = os.path.join(pasta_script, nome_planilha)
        
        if pasta_download is None:
            pasta_download = os.path.join(pasta_script, "PDFs_CRV")
        
        df = pd.read_excel(caminho_planilha)
        print(f"📊 Veículos na planilha: {len(df)}")
        
        options = webdriver.ChromeOptions()
        prefs = {
            "download.default_directory": pasta_download,
            "download.prompt_for_download": False,
            "plugins.always_open_pdf_externally": True
        }
        options.add_experimental_option("prefs", prefs)
        
        driver = webdriver.Chrome(options=options)
        
        try:
            driver.get("https://www.e-crvsp.sp.gov.br/")
            input("✅ Faça o login COMPLETO e pressione ENTER...")
            
            # Entrar no frame
            driver.switch_to.frame("body")
            print("✅ No frame 'body'")
            
            # 🔥 NAVEGAÇÃO (já sabemos que funciona)
            print("🔍 Navegando para ATPV...")
            
            # Clicar no menu
            menu = driver.find_element(By.XPATH, "//a[contains(., 'ATPVe')]")
            menu.click()
            print("✅ Menu clicado!")
            time.sleep(3)
            
            # Clicar no submenu
            submenu = driver.find_element(By.XPATH, "//a[contains(text(), 'imprimir ATPV')]")
            submenu.click()
            print("✅ Submenu clicado!")
            
            # 🔥 AGORA A MUDANÇA CRÍTICA: Aguardar os campos carregarem
            print("⏳ AGUARDANDO campos carregarem (até 30 segundos)...")
            
            # Aguardar até 30 segundos pelos campos
            campo_renavam = WebDriverWait(driver, 30).until(
                EC.presence_of_element_located((By.ID, "renavam"))
            )
            campo_placa = WebDriverWait(driver, 30).until(
                EC.presence_of_element_located((By.ID, "placa"))
            )
            
            print("✅✅✅ CAMPOS ENCONTRADOS! Sistema pronto!")
            
            # 🔥 PROCESSAR VEÍCULOS
            for index, veiculo in df.iterrows():
                renavam = str(veiculo['renavam']).strip()
                placa = str(veiculo['placa']).strip()
                
                print(f"\n🚗 Processando {index + 1}/58 - {placa}")
                
                try:
                    # Preencher campos (já temos as referências)
                    campo_renavam.clear()
                    campo_renavam.send_keys(renavam)
                    campo_placa.clear()
                    campo_placa.send_keys(placa)
                    
                    # Clicar em IMPRIMIR
                    btn_imprimir = WebDriverWait(driver, 10).until(
                        EC.element_to_be_clickable((By.XPATH, "//*[contains(text(), 'IMPRIMIR')]"))
                    )
                    btn_imprimir.click()
                    
                    print(f"✅ PDF gerado! - {placa}")
                    time.sleep(6)  # Aguardar download
                    
                    # Limpar campos para próximo
                    campo_renavam.clear()
                    campo_placa.clear()
                    time.sleep(1)
                    
                except Exception as e:
                    print(f"❌ Erro no veículo {placa}: {e}")
                    continue
            
            print("\n🎉 TODOS OS VEÍCULOS PROCESSADOS!")
            
        finally:
            driver.quit()
            
    except Exception as e:
        print(f"❌ Erro: {e}")

# 🎯 TESTE RÁPIDO - APENAS PARA VERIFICAR SE CAMPOS CARREGAM
def teste_aguardar_campos():
    """
    Teste apenas para ver se os campos carregam após navegação
    """
    options = webdriver.ChromeOptions()
    driver = webdriver.Chrome(options=options)
    
    try:
        driver.get("https://www.e-crvsp.sp.gov.br/")
        input("✅ Faça o login e pressione ENTER...")
        
        driver.switch_to.frame("body")
        print("✅ No frame")
        
        # Navegação
        menu = driver.find_element(By.XPATH, "//a[contains(., 'ATPVe')]")
        menu.click()
        print("✅ Menu clicado")
        time.sleep(3)
        
        submenu = driver.find_element(By.XPATH, "//a[contains(text(), 'imprimir ATPV')]")
        submenu.click()
        print("✅ Submenu clicado")
        
        print("⏳ AGUARDANDO CAMPOS... (máximo 30 segundos)")
        
        # Aguardar campos com timeout longo
        try:
            campo_renavam = WebDriverWait(driver, 30).until(
                EC.presence_of_element_located((By.ID, "renavam"))
            )
            campo_placa = WebDriverWait(driver, 30).until(
                EC.presence_of_element_located((By.ID, "placa"))
            )
            
            print("✅✅✅ SUCESSO! Campos carregaram!")
            print(f"   RENAVAM: {campo_renavam.get_attribute('id')}")
            print(f"   PLACA: {campo_placa.get_attribute('id')}")
            
            # Testar preenchimento
            campo_renavam.send_keys("12345678901")
            campo_placa.send_keys("TEST123")
            print("✅ Campos podem ser preenchidos!")
            
            input("✅✅✅ TUDO FUNCIONANDO! ENTER para fechar...")
            
        except Exception as e:
            print(f"❌ Campos não carregaram após 30 segundos: {e}")
            print("📋 Elementos na página atual:")
            elementos = driver.find_elements(By.XPATH, "//*")
            print(f"Total de elementos: {len(elementos)}")
            
            # Mostrar inputs
            inputs = driver.find_elements(By.TAG_NAME, "input")
            print(f"Inputs encontrados: {len(inputs)}")
            for inp in inputs:
                print(f"   Input: id='{inp.get_attribute('id')}', name='{inp.get_attribute('name')}'")
            
            input("ENTER para fechar...")
            
    finally:
        driver.quit()

# 🎯 EXECUTAR
if __name__ == "__main__":
    print("=" * 60)
    print("🤖 SISTEMA - AGUARDAR CAMPOS CARREGAREM")
    print("=" * 60)
    
    print("Problema resolvido: Navegação funciona, mas campos demoram para carregar")
    print("Solução: Aguardar até 30 segundos pelos campos")
    
    escolha = input("\nEscolha:\n1. Teste rápido (apenas verificar campos)\n2. Sistema completo (processar todos)\nDigite 1 ou 2: ")
    
    if escolha == "1":
        teste_aguardar_campos()
    else:
        nome_planilha = "planilha_veiculos.xlsx"
        automatizar_ecrv_aguardar_campos(nome_planilha)
    
    print("\n📍 Processamento concluído!")