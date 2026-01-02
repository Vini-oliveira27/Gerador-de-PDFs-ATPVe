import pandas as pd
import time
import os
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
import logging

# ==============================
# CONFIGURAÇÕES
# ==============================

# Caminho relativo ou verificação de existência
EXCEL_PATH = r"C:\Users\vinia\OneDrive\Desktop\FREITAS\Pendencia_formatado.xlsx"
# Alternative: usar caminho relativo
# EXCEL_PATH = "Pendencia_formatado.xlsx"

SELECTORS = {
    "campo_pesquisa": (By.ID, "txtPesquisarVeiculos"),
    "botao_pesquisar": (By.ID, "btnPesquisarVeiculos"),
    "aba_leilao": (By.ID, "leilao-tab"),
    "campo_emissao": (By.ID, "ATPVEmitida"),
    "campo_data": (By.ID, "DataEmitidaATPV"),
    "botao_salvar": (By.XPATH, "//button[@onclick='atualizarEmissaoATPVForm();']"),
    "popup_sim": (By.XPATH, "//button[contains(text(),'Sim')]"),
    "popup_ok": (By.ID, "modalJsAlertOk"),
}

# Configurar logging sem emojis para evitar problemas de encoding
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('automacao.log', encoding='utf-8'),
        logging.StreamHandler()
    ]
)

class AutomatizadorFreitas:
    def __init__(self):
        self.driver = None
        self.wait = None
    
    def verificar_arquivo_excel(self):
        """Verifica se o arquivo Excel existe e é válido"""
        if not os.path.exists(EXCEL_PATH):
            logging.error(f"Arquivo não encontrado: {EXCEL_PATH}")
            logging.info("Por favor, verifique:")
            logging.info("1. O caminho do arquivo está correto")
            logging.info("2. O arquivo existe no local especificado")
            logging.info("3. O nome do arquivo está correto (incluindo extensão .xlsx)")
            return False
        
        try:
            # Testa se consegue ler o arquivo
            df = pd.read_excel(EXCEL_PATH)
            if df.empty:
                logging.error("O arquivo Excel está vazio")
                return False
            
            # Verifica se as colunas necessárias existem
            colunas_necessarias = ['placa']
            colunas_faltantes = [col for col in colunas_necessarias if col not in df.columns]
            
            if colunas_faltantes:
                logging.error(f"Colunas faltantes no Excel: {colunas_faltantes}")
                logging.info(f"Colunas encontradas: {list(df.columns)}")
                return False
            
            logging.info(f"Arquivo Excel validado com sucesso. {len(df)} registros encontrados.")
            return True
            
        except Exception as e:
            logging.error(f"Erro ao ler arquivo Excel: {e}")
            return False
    
    def iniciar_navegador(self):
        """Inicializa o navegador Chrome"""
        try:
            service = Service(ChromeDriverManager().install())
            options = webdriver.ChromeOptions()
            options.add_argument('--start-maximized')
            options.add_argument('--disable-blink-features=AutomationControlled')
            options.add_experimental_option("excludeSwitches", ["enable-automation"])
            options.add_experimental_option('useAutomationExtension', False)
            
            self.driver = webdriver.Chrome(service=service, options=options)
            self.driver.execute_script("Object.defineProperty(navigator, 'webdriver', {get: () => undefined})")
            self.wait = WebDriverWait(self.driver, 30)
            
            logging.info("Navegador iniciado com sucesso")
            return True
        except Exception as e:
            logging.error(f"Erro ao iniciar navegador: {e}")
            return False
    
    def fazer_login_manual(self):
        """Aguarda login manual do usuário"""
        URL_LOGIN = "https://login.freitasleiloeiro.com.br/Home/Login"
        self.driver.get(URL_LOGIN)
        input("Faça login manualmente e pressione ENTER para continuar...")
        
        # Aguarda mudança para nova aba após seleção da área
        input("Após selecionar a área desejada, pressione ENTER para continuar...")
        
        # Espera explícita por nova aba
        self.wait.until(lambda d: len(d.window_handles) > 1)
        self.driver.switch_to.window(self.driver.window_handles[-1])
        logging.info("Mudou para a nova aba com a área selecionada")
        
        # Espera adicional para carregamento
        time.sleep(3)
    
    def processar_placa(self, placa, data_emissao):
        """Processa uma placa individual"""
        try:
            logging.info(f"Processando placa: {placa}")
            
            # Pesquisar placa
            campo = self.wait.until(EC.visibility_of_element_located(SELECTORS["campo_pesquisa"]))
            campo.clear()
            campo.send_keys(placa)
            
            self.driver.find_element(*SELECTORS["botao_pesquisar"]).click()
            
            # Clicar no link da placa
            link_placa = self.wait.until(
                EC.element_to_be_clickable((By.XPATH, f"//a[contains(@href,'{placa}')]"))
            )
            link_placa.click()
            
            # Aba Leilão
            aba_leilao = self.wait.until(EC.element_to_be_clickable(SELECTORS["aba_leilao"]))
            aba_leilao.click()
            time.sleep(2)  # Espera para carregamento da aba
            
            # Emissão ATPV = "Sim"
            campo_emissao = self.wait.until(EC.presence_of_element_located(SELECTORS["campo_emissao"]))
            self.wait.until(lambda d: campo_emissao.is_enabled())
            select_emissao = Select(campo_emissao)
            select_emissao.select_by_visible_text("Sim")
            
            # Data de emissão
            if data_emissao:
                campo_data = self.driver.find_element(*SELECTORS["campo_data"])
                campo_data.clear()
                campo_data.send_keys(data_emissao)
            
            # Salvar
            self.driver.find_element(*SELECTORS["botao_salvar"]).click()
            
            # Pop-ups de confirmação
            self.wait.until(EC.element_to_be_clickable(SELECTORS["popup_sim"])).click()
            self.wait.until(EC.element_to_be_clickable(SELECTORS["popup_ok"])).click()
            
            # Pequeno delay entre processamentos
            time.sleep(1)
            
            logging.info(f"Placa {placa} processada com sucesso")
            return True
            
        except Exception as e:
            logging.error(f"Erro ao processar placa {placa}: {str(e)}")
            return False
    
    def executar(self):
        """Função principal de execução"""
        try:
            # Verificar arquivo primeiro
            if not self.verificar_arquivo_excel():
                return
            
            # Ler planilha
            df = pd.read_excel(EXCEL_PATH)
            total_placas = len(df)
            sucessos = 0
            
            logging.info(f"Iniciando processamento de {total_placas} placas")
            
            # Iniciar navegador e login
            if not self.iniciar_navegador():
                return
            
            self.fazer_login_manual()
            
            # Processar cada placa
            for i, row in df.iterrows():
                placa = str(row["placa"]).strip()
                data_emissao_raw = row.get("data_emissao", "")
                
                # Formatar data
                if pd.notna(data_emissao_raw) and data_emissao_raw != "":
                    try:
                        data_emissao = pd.to_datetime(data_emissao_raw).strftime("%d/%m/%Y")
                    except:
                        data_emissao = ""
                        logging.warning(f"Data inválida para placa {placa}: {data_emissao_raw}")
                else:
                    data_emissao = ""
                
                if self.processar_placa(placa, data_emissao):
                    sucessos += 1
                
                # Progresso
                progresso = (i + 1) / total_placas * 100
                logging.info(f"Progresso: {i + 1}/{total_placas} ({progresso:.1f}%)")
            
            # Relatório final
            logging.info(f"Processamento concluído! Sucessos: {sucessos}/{total_placas}")
            
        except Exception as e:
            logging.error(f"Erro geral na execução: {e}")
        
        finally:
            if self.driver:
                self.driver.quit()
                logging.info("Navegador fechado")

def main():
    automatizador = AutomatizadorFreitas()
    automatizador.executar()

if __name__ == "__main__":
    main()