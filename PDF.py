# Primeiro instale: pip install undetected-chromedriver

import undetected_chromedriver as uc
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import pandas as pd
import time
import os
import shutil

def automatizar_ecrv_com_undetected(nome_planilha, pasta_base=None):
    '''
    SISTEMA SEM PAUSAS - APENAS LIMITE 70 DOWNLOADS/MINUTO
    '''
    try:
        pasta_script = os.path.dirname(os.path.abspath(__file__))
        caminho_planilha = os.path.join(pasta_script, nome_planilha)
        
        if pasta_base is None:
            pasta_base = os.path.join(pasta_script, "PDFs_ORGANIZADOS")
        
        # Ler planilha
        df = pd.read_excel(caminho_planilha)
        total_veiculos = len(df)
        
        # Verificar coluna COMITENTE
        if 'comitente' not in df.columns and 'COMITENTE' not in df.columns:
            print("❌ ERRO: Planilha não tem coluna 'COMITENTE'")
            return
        
        coluna_comitente = 'COMITENTE' if 'COMITENTE' in df.columns else 'comitente'
        
        # Criar pastas
        pasta_temp = os.path.join(pasta_script, "TEMP_DOWNLOADS")
        os.makedirs(pasta_temp, exist_ok=True)
        
        comitentes = df[coluna_comitente].unique()
        print(f"📁 COMITENTES ENCONTRADOS: {len(comitentes)}")
        
        for comitente in comitentes:
            if pd.notna(comitente):
                pasta_comitente = os.path.join(pasta_base, str(comitente).strip())
                os.makedirs(pasta_comitente, exist_ok=True)
        
        print(f"📊 TOTAL DE VEÍCULOS: {total_veiculos}")
        
        # 🔥 VARIÁVEIS DE CONTROLE DE DOWNLOADS
        downloads_por_minuto = 0
        timestamp_minuto_atual = time.time()
        sucessos = 0
        falhas = 0
        
        # 🔥 CONFIGURAÇÃO UNDETECTED CHROMEDRIVER
        print("\n" + "="*70)
        print("🚀 MODO RÁPIDO - SEM PAUSAS (70 downloads/minuto)")
        print("="*70)
        print("⚡ Máxima velocidade mantendo limite do sistema")
        print("="*70 + "\n")
        
        profile_dir = os.path.join(pasta_script, "chrome_profile_undetected")
        os.makedirs(profile_dir, exist_ok=True)
        
        print(f"📁 Perfil do Chrome será salvo em: {profile_dir}")
        
        # Configurações do Chrome
        options = uc.ChromeOptions()
        options.add_argument(f"--user-data-dir={profile_dir}")
        
        # Configurações de download
        prefs = {
            "download.default_directory": pasta_temp,
            "download.prompt_for_download": False,
            "plugins.always_open_pdf_externally": True,
            "profile.default_content_setting_values.automatic_downloads": 1,
        }
        options.add_experimental_option("prefs", prefs)
        
        # Argumentos anti-detecção
        options.add_argument("--disable-blink-features=AutomationControlled")
        options.add_argument("--disable-features=VizDisplayCompositor")
        options.add_argument("--no-first-run")
        options.add_argument("--no-default-browser-check")
        options.add_argument("--disable-popup-blocking")
        options.add_argument("--start-maximized")
        options.add_argument("--no-sandbox")
        options.add_argument("--disable-dev-shm-usage")
        
        print("🌐 Iniciando Chrome...")
        driver = uc.Chrome(options=options)
        
        try:
            # ABRIR SISTEMA
            print("\n🌐 Abrindo sistema...")
            driver.get("https://www.e-crvsp.sp.gov.br/")
            
            print("\n" + "="*70)
            print("🔐 AGORA FAÇA O LOGIN MANUALMENTE")
            print("="*70)
            print("\n1️⃣ Faça o login completo")
            print("2️⃣ Navegue até a página de IMPRIMIR ATPV")
            print("="*70)
            
            input("\n✅ PRESSIONE ENTER QUANDO ESTIVER NA PÁGINA DE IMPRIMIR ATPV...")
            
            # ABA PRINCIPAL
            aba_principal = driver.current_window_handle
            
            print("\n🚀 INICIANDO PROCESSAMENTO RÁPIDO...")
            
            for index, veiculo in df.iterrows():
                renavam = str(veiculo['renavam']).strip()
                placa = str(veiculo['placa']).strip()
                comitente = str(veiculo[coluna_comitente]).strip() if pd.notna(veiculo[coluna_comitente]) else "SEM_COMITENTE"
                
                print(f"\n🚗 [{index + 1}/{total_veiculos}] {placa} - Comitente: {comitente}")
                
                # 🔄 VERIFICAR LIMITE DE DOWNLOADS POR MINUTO
                tempo_atual = time.time()
                if tempo_atual - timestamp_minuto_atual >= 60:
                    downloads_por_minuto = 0
                    timestamp_minuto_atual = tempo_atual
                
                if downloads_por_minuto >= 70:
                    tempo_espera = 60 - (tempo_atual - timestamp_minuto_atual)
                    print(f"   ⏳ Limite 4/min - aguardando {tempo_espera:.0f}s")
                    time.sleep(tempo_espera)
                    downloads_por_minuto = 0
                    timestamp_minuto_atual = time.time()
                
                try:
                    # Voltar para contexto principal
                    driver.switch_to.window(aba_principal)
                    driver.switch_to.default_content()
                    
                    # Encontrar frame
                    try:
                        driver.switch_to.frame("body")
                    except:
                        frames = driver.find_elements(By.TAG_NAME, "frame")
                        if frames:
                            driver.switch_to.frame(0)
                    
                    # Encontrar e preencher campos
                    campo_renavam = WebDriverWait(driver, 15).until(
                        EC.presence_of_element_located((By.ID, "renavam"))
                    )
                    campo_placa = WebDriverWait(driver, 15).until(
                        EC.presence_of_element_located((By.ID, "placa"))
                    )
                    
                    campo_renavam.clear()
                    campo_renavam.send_keys(renavam)
                    campo_placa.clear()
                    campo_placa.send_keys(placa)
                    
                    # Clicar em IMPRIMIR
                    try:
                        btn_imprimir = WebDriverWait(driver, 10).until(
                            EC.element_to_be_clickable((By.XPATH, "//*[contains(text(), 'IMPRIMIR')]"))
                        )
                    except:
                        btn_imprimir = driver.find_element(By.XPATH, "//input[@value='IMPRIMIR']")
                    
                    btn_imprimir.click()
                    
                    print("   ⏳ Aguardando download...")
                    
                    # Aguardar download
                    arquivo_baixado = None
                    tempo_espera = 0
                    while tempo_espera < 30:
                        time.sleep(1)
                        arquivos = os.listdir(pasta_temp)
                        if arquivos:
                            arquivos_pdf = [f for f in arquivos if f.lower().endswith('.pdf')]
                            if arquivos_pdf:
                                caminho_completo = os.path.join(pasta_temp, arquivos_pdf[0])
                                if os.path.exists(caminho_completo) and os.path.getsize(caminho_completo) > 10000:
                                    arquivo_baixado = arquivos_pdf[0]
                                    break
                        tempo_espera += 1
                    
                    if arquivo_baixado:
                        caminho_origem = os.path.join(pasta_temp, arquivo_baixado)
                        
                        if os.path.exists(caminho_origem) and os.path.getsize(caminho_origem) > 0:
                            # Mover e renomear
                            pasta_destino = os.path.join(pasta_base, comitente)
                            nome_novo = f"{placa}.pdf"
                            caminho_destino = os.path.join(pasta_destino, nome_novo)
                            
                            contador = 1
                            while os.path.exists(caminho_destino):
                                nome_novo = f"{placa}_{contador}.pdf"
                                caminho_destino = os.path.join(pasta_destino, nome_novo)
                                contador += 1
                            
                            shutil.move(caminho_origem, caminho_destino)
                            print(f"   ✅ Download: {nome_novo}")
                            
                            # 🔢 INCREMENTAR CONTADOR
                            downloads_por_minuto += 1
                            df.at[index, 'processado'] = 'Sim'
                            df.at[index, 'data_processamento'] = time.strftime('%Y-%m-%d %H:%M:%S')
                            sucessos += 1
                            
                        else:
                            print("   ⚠️  Arquivo vazio")
                            falhas += 1
                    else:
                        print("   ⚠️  Download não detectado")
                        falhas += 1
                        
                        if falhas > 0 and index < total_veiculos - 1:
                            continuar = input("❓ Continuar? (s/n): ").lower()
                            if continuar != 's':
                                break
                    
                except Exception as e:
                    print(f"   ❌ Erro: {e}")
                    falhas += 1
                    
                    if index < total_veiculos - 1:
                        continuar = input("❓ Continuar? (s/n): ").lower()
                        if continuar != 's':
                            break
            
            # Salvar relatório
            relatorio_path = os.path.join(pasta_script, "RELATORIO.xlsx")
            df.to_excel(relatorio_path, index=False)
            print(f"\n📊 Relatório salvo em: {relatorio_path}")
            
            print(f"\n📊 RESUMO: {sucessos} sucessos, {falhas} falhas")
            
        finally:
            input("\nPressione ENTER para fechar...")
            driver.quit()
            
            try:
                shutil.rmtree(pasta_temp)
                print("🧹 Pasta temp limpa")
            except:
                pass
            
    except Exception as e:
        print(f"❌ Erro: {e}")
        input("\nPressione ENTER para sair...")

if __name__ == "__main__":
    nome_planilha = "planilha_veiculos.xlsx"
    
    if not os.path.exists(nome_planilha):
        print("❌ Planilha não encontrada!")
        print("Crie com colunas: renavam, placa, COMITENTE")
        
        dados_exemplo = {
            'renavam': ['12345678901', '98765432109'],
            'placa': ['ABC1D23', 'XYZ9W87'],
            'COMITENTE': ['COMITENTE_A', 'COMITENTE_B']
        }
        df = pd.DataFrame(dados_exemplo)
        df.to_excel(nome_planilha, index=False)
        print(f"✅ Planilha exemplo criada: {nome_planilha}")
        exit()
    
    automatizar_ecrv_com_undetected(nome_planilha)
