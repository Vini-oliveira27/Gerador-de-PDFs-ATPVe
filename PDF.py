from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.edge.options import Options
from selenium.webdriver.edge.service import Service  # ⬅ IMPORTANTE
from pathlib import Path
import pandas as pd
import time
import os
import shutil

# 🔥 CAMINHO FIXO DO DRIVER NA OUTRA MÁQUINA
EDGE_DRIVER_PATH = r"C:\Users\jrwil\OneDrive\Desktop\GERADOR DE PDFs\Gerador-de-PDFs-ATPVe\msedgedriver.exe"


def automatizar_ecrv_com_comitentes(nome_planilha, pasta_base=None):
    """
    SISTEMA QUE ORGANIZA PDFs POR COMITENTE - VERSÃO PORTÁVEL
    """
    try:
        # 🔥 USAR PATHLIB PARA CAMINHOS MULTIPLATAFORMA
        pasta_script = Path(__file__).parent.absolute()
        caminho_planilha = pasta_script / nome_planilha

        if pasta_base is None:
            pasta_base = pasta_script / "PDFs_ORGANIZADOS"
        else:
            pasta_base = Path(pasta_base)  # Converter string em Path se necessário

        # Garantir que Path é objeto Path (não string)
        if isinstance(caminho_planilha, str):
            caminho_planilha = Path(caminho_planilha)
        if isinstance(pasta_base, str):
            pasta_base = Path(pasta_base)

        # Ler planilha
        df = pd.read_excel(str(caminho_planilha))  # ← Converter Path para string para pandas
        total_veiculos = len(df)

        # 🔥 VERIFICAR SE TEM COLUNA COMITENTE
        if 'comitente' not in df.columns and 'COMITENTE' not in df.columns:
            print("❌ ERRO: Planilha não tem coluna 'COMITENTE'")
            print("📋 Colunas encontradas:", list(df.columns))
            return

        # Padronizar nome da coluna
        coluna_comitente = 'COMITENTE' if 'COMITENTE' in df.columns else 'comitente'

        # 🔥 CRIAR PASTA TEMPORÁRIA COM PATHLIB
        pasta_temp = pasta_script / "TEMP_DOWNLOADS"
        pasta_temp.mkdir(parents=True, exist_ok=True)

        # 🔥 CRIAR PASTAS PARA CADA COMITENTE
        comitentes = df[coluna_comitente].unique()
        print(f"📁 COMITENTES ENCONTRADOS: {len(comitentes)}")

        for comitente in comitentes:
            if pd.notna(comitente):
                pasta_comitente = pasta_base / str(comitente).strip()
                pasta_comitente.mkdir(parents=True, exist_ok=True)
                print(f"   ✅ Pasta criada: {comitente}")

        print(f"📊 TOTAL DE VEÍCULOS: {total_veiculos}")

        # CONFIGURAÇÕES EDGE
        options = Options()

        # Configurar preferências de download
        prefs = {
            "download.default_directory": str(pasta_temp),
            "download.prompt_for_download": False,
            "plugins.always_open_pdf_externally": True,
            "download.directory_upgrade": True,
            "safebrowsing.enabled": False
        }

        options.use_chromium = True

        try:
            options.add_experimental_option("prefs", prefs)
        except Exception:
            try:
                options.set_capability("ms:edgeOptions", {"prefs": prefs})
            except Exception:
                print("⚠️  Usando configuração básica do Edge")

        try:
            options.add_argument("--inprivate")
        except AttributeError:
            print("⚠️  Argumentos não suportados nesta versão, continuando...")

        # 🚀 INICIALIZAR EDGE COM SERVICE E CAMINHO FIXO DO DRIVER
        try:
            service = Service(executable_path=EDGE_DRIVER_PATH)
            driver = webdriver.Edge(service=service, options=options)
            print(f"✅ Edge WebDriver iniciado com driver em: {EDGE_DRIVER_PATH}")
        except Exception as e:
            print(f"❌ ERRO ao inicializar Edge com driver fixo: {e}")
            raise

        # Configurar download via DevTools Protocol (se disponível)
        try:
            driver.execute_cdp_cmd('Page.setDownloadBehavior', {
                'behavior': 'allow',
                'downloadPath': str(pasta_temp)
            })
        except Exception:
            print("⚠️  CDP commands não disponíveis, usando configuração padrão")

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
                    for arquivo in pasta_temp.iterdir():
                        try:
                            if arquivo.is_file():
                                arquivo.unlink()
                        except Exception as e:
                            print(f"   ⚠️  Erro ao limpar {arquivo}: {e}")

                    # Clicar em IMPRIMIR
                    btn_imprimir = WebDriverWait(driver, 10).until(
                        EC.element_to_be_clickable((By.XPATH, "//*[contains(text(), 'IMPRIMIR')]"))
                    )
                    btn_imprimir.click()

                    print(f"   ⏳ Baixando PDF...")

                    # 🔥 AGUARDAR E CAPTURAR O ARQUIVO BAIXADO
                    arquivo_baixado = None
                    tempo_inicio = time.time()

                    while time.time() - tempo_inicio < 30:
                        time.sleep(1)
                        arquivos = list(pasta_temp.glob('*'))

                        # Verificar se há algum arquivo na pasta
                        if arquivos:
                            for arquivo in arquivos:
                                # Ignorar arquivos temporários
                                if not arquivo.name.endswith('.tmp') and not arquivo.name.endswith('.crdownload'):
                                    arquivo_baixado = arquivo
                                    break

                        if arquivo_baixado:
                            # Verificar se o arquivo está completamente baixado
                            if arquivo_baixado.exists():
                                try:
                                    tamanho_atual = arquivo_baixado.stat().st_size
                                    time.sleep(1)
                                    tamanho_depois = arquivo_baixado.stat().st_size
                                    if tamanho_atual == tamanho_depois and tamanho_atual > 0:
                                        break
                                except Exception:
                                    break

                    if arquivo_baixado:
                        # 🔥 MOVER ARQUIVO PARA PASTA DO COMITENTE COM PATHLIB
                        pasta_destino = pasta_base / comitente
                        pasta_destino.mkdir(parents=True, exist_ok=True)

                        # 🔥 RENOMEAR ARQUIVO COM PLACA
                        extensao = arquivo_baixado.suffix or '.pdf'
                        nome_novo = f"{placa}{extensao}"
                        caminho_destino = pasta_destino / nome_novo

                        # Esperar um pouco e tentar mover
                        time.sleep(2)

                        tentativas = 0
                        while tentativas < 3:
                            try:
                                arquivo_baixado.rename(caminho_destino)
                                print(f"   ✅ PDF ORGANIZADO: {nome_novo}")
                                break
                            except Exception as e:
                                tentativas += 1
                                time.sleep(1)
                                if tentativas == 3:
                                    print(f"   ⚠️  Erro ao mover arquivo: {e}")
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
                    except Exception:
                        pass
                    continue

            # 💾 SALVAR RELATÓRIO FINAL
            nome_arquivo_saida = pasta_script / "RELATORIO_COMITENTES.xlsx"
            df.to_excel(str(nome_arquivo_saida), index=False)

            # 📊 RELATÓRIO FINAL POR COMITENTE
            print(f"\n{'='*60}")
            print("🎉 PROCESSAMENTO CONCLUÍDO!")
            print(f"{'='*60}")

            # Estatísticas por comitente
            for comitente in comitentes:
                if pd.notna(comitente):
                    comitente_str = str(comitente).strip()
                    pasta_comitente = pasta_base / comitente_str
                    if pasta_comitente.exists():
                        qtd_pdfs = len(list(pasta_comitente.glob('*.pdf')))
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
            import traceback
            traceback.print_exc()
        finally:
            try:
                driver.quit()
            except Exception:
                pass

            # 🔥 LIMPAR PASTA TEMPORÁRIA COM PATHLIB
            try:
                shutil.rmtree(str(pasta_temp))
                print(f"🧹 Pasta temporária limpa: {pasta_temp}")
            except Exception as e:
                print(f"⚠️  Erro ao limpar pasta temp: {e}")

    except Exception as e:
        print(f"❌ Erro: {e}")
        import traceback
        traceback.print_exc()


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

    # Usar pathlib para salvar
    arquivo_saida = Path(__file__).parent / 'planilha_comitentes_exemplo.xlsx'
    df.to_excel(str(arquivo_saida), index=False)
    print(f"📋 Planilha exemplo criada: {arquivo_saida}")


# 🚀 EXECUTAR
if __name__ == "__main__":
    print("=" * 60)
    print("🤖 SISTEMA ORGANIZADOR POR COMITENTE")
    print("=" * 60)
    print("🚀 CONFIGURADO PARA MICROSOFT EDGE (VERSÃO PORTÁVEL)")
    print("=" * 60)

    # Usar pathlib para verificar arquivos
    nome_planilha = "planilha_veiculos.xlsx"
    pasta_script = Path(__file__).parent.absolute()
    caminho_planilha = pasta_script / nome_planilha

    if not caminho_planilha.exists():
        print(f"❌ Planilha {nome_planilha} não encontrada em {pasta_script}")
        criar = input("Criar planilha exemplo? (s/n): ")
        if criar.lower() == 's':
            criar_planilha_exemplo()
        exit()

    # Verificar coluna COMITENTE
    df = pd.read_excel(str(caminho_planilha))
    if 'comitente' not in df.columns and 'COMITENTE' not in df.columns:
        print("❌ Planilha não tem coluna 'COMITENTE'")
        print("📋 Colunas encontradas:", list(df.columns))
        print("\n💡 A planilha deve ter colunas: renavam, placa, COMITENTE")
        exit()

    print("✅ Planilha com comitentes encontrada!")

    input("\n🚀 Pressione ENTER para iniciar organização por comitentes...")

    automatizar_ecrv_com_comitentes(str(caminho_planilha))

    print(f"\n⭐ ORGANIZAÇÃO CONCLUÍDA!")
    print("⭐ PDFs organizados em pastas por comitente!")
