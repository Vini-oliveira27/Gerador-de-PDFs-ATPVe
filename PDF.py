from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
import pandas as pd
import time
import os
import shutil

def automatizar_ecrv_com_comitentes(nome_planilha, pasta_base=None):
    """
    SISTEMA QUE ORGANIZA PDFs POR COMITENTE
    """
    try:
        pasta_script = os.path.dirname(os.path.abspath(__file__))
        caminho_planilha = os.path.join(pasta_script, nome_planilha)
        
        if pasta_base is None:
            pasta_base = os.path.join(pasta_script, "PDFs_ORGANIZADOS")
        
        # Ler planilha
        df = pd.read_excel(caminho_planilha)
        total_veiculos = len(df)
        
        # 🔥 VERIFICAR SE TEM COLUNA COMITENTE
        if 'comitente' not in df.columns and 'COMITENTE' not in df.columns:
            print("❌ ERRO: Planilha não tem coluna 'COMITENTE'")
            print("📋 Colunas encontradas:", list(df.columns))
            return
        
        # Padronizar nome da coluna
        coluna_comitente = 'COMITENTE' if 'COMITENTE' in df.columns else 'comitente'
        
        # 🔥 CRIAR PASTA TEMPORÁRIA PARA DOWNLOADS
        pasta_temp = os.path.join(pasta_script, "TEMP_DOWNLOADS")
        os.makedirs(pasta_temp, exist_ok=True)
        
        # 🔥 CRIAR PASTAS PARA CADA COMITENTE
        comitentes = df[coluna_comitente].unique()
        print(f"📁 COMITENTES ENCONTRADOS: {len(comitentes)}")
        
        for comitente in comitentes:
            if pd.notna(comitente):  # Ignorar valores NaN
                pasta_comitente = os.path.join(pasta_base, str(comitente).strip())
                os.makedirs(pasta_comitente, exist_ok=True)
                print(f"   ✅ Pasta criada: {comitente}")
        
        print(f"📊 TOTAL DE VEÍCULOS: {total_veiculos}")
        
        # CONFIGURAÇÕES CHROME
        options = Options()
        prefs = {
            "download.default_directory": pasta_temp,  # 🔥 Download para pasta temporária
            "download.prompt_for_download": False,
            "plugins.always_open_pdf_externally": True,
        }
        options.add_experimental_option("prefs", prefs)
        
        driver = webdriver.Chrome(options=options)
        
        try:
            print("🌐 Acessando sistema...")
            driver.get("https://www.e-crvsp.sp.gov.br/")
            
            input("✅ Faça o login e pressione ENTER para começar...")
            
            # Entrar no frame
            driver.switch_to.frame("body")
            print("✅ Sistema carregado")
            
            # NAVEGAÇÃO
            print("📝 Navegando para impressão ATPV...")
            
            menu = driver.find_element(By.XPATH, "//a[contains(., 'ATPVe')]")
            menu.click()
            time.sleep(3)
            
            submenu = driver.find_element(By.XPATH, "//a[contains(text(), 'imprimir ATPV')]")
            submenu.click()
            
            # AGUARDAR CAMPOS
            print("⏳ Aguardando formulário carregar...")
            campo_renavam = WebDriverWait(driver, 30).until(
                EC.presence_of_element_located((By.ID, "renavam"))
            )
            campo_placa = WebDriverWait(driver, 30).until(
                EC.presence_of_element_located((By.ID, "placa"))
            )
            
            print("✅✅✅ FORMULÁRIO PRONTO!")
            
            # ABA PRINCIPAL
            aba_principal = driver.current_window_handle
            
            # 🔥 PROCESSAR TODOS OS VEÍCULOS
            for index, veiculo in df.iterrows():
                renavam = str(veiculo['renavam']).strip()
                placa = str(veiculo['placa']).strip()
                comitente = str(veiculo[coluna_comitente]).strip() if pd.notna(veiculo[coluna_comitente]) else "SEM_COMITENTE"
                
                print(f"\n🚗 [{index + 1}/{total_veiculos}] {placa} - Comitente: {comitente}")
                
                try:
                    # VOLTAR PARA ABA/FORMA PRINCIPAL
                    driver.switch_to.window(aba_principal)
                    driver.switch_to.frame("body")
                    
                    # Atualizar referências dos campos
                    campo_renavam = driver.find_element(By.ID, "renavam")
                    campo_placa = driver.find_element(By.ID, "placa")
                    
                    # Preencher campos
                    campo_renavam.clear()
                    campo_renavam.send_keys(renavam)
                    campo_placa.clear()
                    campo_placa.send_keys(placa)
                    
                    # 🔥 LIMPAR PASTA TEMPORÁRIA ANTES DE BAIXAR
                    for arquivo in os.listdir(pasta_temp):
                        caminho_arquivo = os.path.join(pasta_temp, arquivo)
                        try:
                            if os.path.isfile(caminho_arquivo):
                                os.remove(caminho_arquivo)
                        except:
                            pass
                    
                    # Clicar em IMPRIMIR
                    btn_imprimir = WebDriverWait(driver, 10).until(
                        EC.element_to_be_clickable((By.XPATH, "//*[contains(text(), 'IMPRIMIR')]"))
                    )
                    btn_imprimir.click()
                    
                    print(f"   ⏳ Baixando PDF...")
                    
                    # 🔥 AGUARDAR E CAPTURAR O ARQUIVO BAIXADO
                    arquivo_baixado = None
                    for _ in range(20):  # Tentar por 20 segundos
                        time.sleep(1)
                        arquivos = os.listdir(pasta_temp)
                        if arquivos:
                            arquivo_baixado = arquivos[0]
                            break
                    
                    if arquivo_baixado:
                        # 🔥 MOVER ARQUIVO PARA PASTA DO COMITENTE
                        pasta_destino = os.path.join(pasta_base, comitente)
                        caminho_origem = os.path.join(pasta_temp, arquivo_baixado)
                        
                        # 🔥 RENOMEAR ARQUIVO COM PLACA
                        nome_novo = f"{placa}_{arquivo_baixado}"
                        caminho_destino = os.path.join(pasta_destino, nome_novo)
                        
                        shutil.move(caminho_origem, caminho_destino)
                        print(f"   ✅ PDF ORGANIZADO: {nome_novo}")
                    else:
                        print(f"   ⚠️  PDF não foi baixado")
                    
                    # Limpar campos para próximo
                    campo_renavam.clear()
                    campo_placa.clear()
                    time.sleep(1)
                    
                    # Marcar como processado
                    if 'processado' not in df.columns:
                        df['processado'] = ''
                    df.at[index, 'processado'] = 'Sim'
                    df.at[index, 'data_processamento'] = time.strftime('%Y-%m-%d %H:%M:%S')
                    
                except Exception as e:
                    print(f"   ❌ Erro: {e}")
                    if 'processado' not in df.columns:
                        df['processado'] = ''
                    df.at[index, 'processado'] = 'Erro'
                    df.at[index, 'erro'] = str(e)
                    
                    # Recuperação
                    try:
                        driver.switch_to.window(aba_principal)
                        driver.switch_to.frame("body")
                    except:
                        pass
                    continue
            
            # 💾 SALVAR RELATÓRIO FINAL
            nome_arquivo_saida = os.path.join(pasta_script, "RELATORIO_COMITENTES.xlsx")
            df.to_excel(nome_arquivo_saida, index=False)
            
            # 📊 RELATÓRIO FINAL POR COMITENTE
            print(f"\n{'='*60}")
            print("🎉 PROCESSAMENTO CONCLUÍDO!")
            print(f"{'='*60}")
            
            # Estatísticas por comitente
            for comitente in comitentes:
                if pd.notna(comitente):
                    comitente_str = str(comitente).strip()
                    pasta_comitente = os.path.join(pasta_base, comitente_str)
                    qtd_pdfs = len([f for f in os.listdir(pasta_comitente) if f.endswith('.pdf')])
                    qtd_veiculos = df[df[coluna_comitente] == comitente].shape[0]
                    print(f"📁 {comitente_str}: {qtd_pdfs}/{qtd_veiculos} PDFs")
            
            sucessos = df[df['processado'] == 'Sim'].shape[0]
            erros = df[df['processado'] == 'Erro'].shape[0]
            
            print(f"\n📊 RESUMO GERAL:")
            print(f"   ✅ Sucessos: {sucessos} veículos")
            print(f"   ❌ Erros: {erros} veículos")
            print(f"   📁 Pasta organizada: {pasta_base}")
            print(f"   📋 Relatório: RELATORIO_COMITENTES.xlsx")
            
        except Exception as e:
            print(f"❌ Erro durante automação: {e}")
        finally:
            driver.quit()
            
            # 🔥 LIMPAR PASTA TEMPORÁRIA
            try:
                shutil.rmtree(pasta_temp)
            except:
                pass
            
    except Exception as e:
        print(f"❌ Erro: {e}")

# 🎯 CRIAR PLANILHA EXEMPLO
def criar_planilha_exemplo():
    """
    Cria uma planilha de exemplo com comitentes
    """
    dados_exemplo = {
        'renavam': ['12345678901', '98765432109', '55544433322', '11122233344'],
        'placa': ['ABC1D23', 'XYZ9W87', 'TEST123', 'SAMPLE99'],
        'COMITENTE': ['COMITENTE_A', 'COMITENTE_B', 'COMITENTE_A', 'COMITENTE_C']
    }
    
    df = pd.DataFrame(dados_exemplo)
    df.to_excel('planilha_comitentes_exemplo.xlsx', index=False)
    print("📋 Planilha exemplo criada: planilha_comitentes_exemplo.xlsx")

# 🚀 EXECUTAR
if __name__ == "__main__":
    print("=" * 60)
    print("🤖 SISTEMA ORGANIZADOR POR COMITENTE")
    print("=" * 60)
    
    # Verificar se planilha existe
    nome_planilha = "planilha_veiculos.xlsx"
    caminho_planilha = os.path.join(os.getcwd(), nome_planilha)
    
    if not os.path.exists(caminho_planilha):
        print(f"❌ Planilha {nome_planilha} não encontrada")
        criar = input("Criar planilha exemplo? (s/n): ")
        if criar.lower() == 's':
            criar_planilha_exemplo()
        exit()
    
    # Verificar coluna COMITENTE
    df = pd.read_excel(caminho_planilha)
    if 'comitente' not in df.columns and 'COMITENTE' not in df.columns:
        print("❌ Planilha não tem coluna 'COMITENTE'")
        print("📋 Colunas encontradas:", list(df.columns))
        print("\n💡 A planilha deve ter colunas: renavam, placa, COMITENTE")
        exit()
    
    print("✅ Planilha com comitentes encontrada!")
    
    input("\n🚀 Pressione ENTER para iniciar organização por comitentes...")
    
    automatizar_ecrv_com_comitentes(nome_planilha)
    
    print(f"\n⭐ ORGANIZAÇÃO CONCLUÍDA!")
    print("⭐ PDFs organizados em pastas por comitente!")