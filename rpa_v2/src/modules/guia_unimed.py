import os
import time
import pandas as pd
from tkinter import messagebox
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
from dotenv import load_dotenv

from src.core.browser_factory import BrowserFactory
from src.core.logger import log_message
from src.modules.base import BaseModule

load_dotenv()

class GuiaUnimedModule(BaseModule):
    def __init__(self):
        super().__init__(nome="Guia Unimed")

    def get_unique_exames(self, file_path: str) -> list:
        try:
            df = pd.read_excel(file_path)
            unique_exames = df['Exame'].dropna().unique().tolist()
            return unique_exames
        except Exception as e:
            raise ValueError(f"Erro ao ler o Excel: {e}")

    def run(self, params: dict):
        username = params.get("username")
        password = params.get("password")
        cancel_flag = params.get("cancel_flag")
        headless_mode = params.get("headless_mode")
        excel_file = params.get("excel_file")

        url = os.getenv("SYSTEM_URL", "https://pathoweb.com.br/login/auth")
        driver = BrowserFactory.create_chrome(headless=headless_mode)
        wait = WebDriverWait(driver, 15)

        try:
            log_message("Iniciando automação de Guia Unimed...", "INFO")

            # Carregar exames do Excel
            if not excel_file or not os.path.exists(excel_file):
                messagebox.showerror("Erro", "Arquivo Excel não informado ou não encontrado.")
                return
            try:
                exames_unicos = self.get_unique_exames(excel_file)
            except Exception as e:
                messagebox.showerror("Erro", str(e))
                return
            if not exames_unicos:
                messagebox.showerror("Erro", "Nenhum exame encontrado no arquivo.")
                return

            # Login
            driver.get(url)
            wait.until(EC.presence_of_element_located((By.ID, "username"))).send_keys(username)
            driver.find_element(By.ID, "password").send_keys(password)
            driver.find_element(By.CSS_SELECTOR, "button[type='submit']").click()

            log_message("Verificando se precisa navegar para módulo de faturamento...", "INFO")
            current_url = driver.current_url

            if current_url == "https://pathoweb.com.br/" or "trocarModulo" in current_url:
                log_message("Detectada tela de seleção de módulos - navegando para módulo de faturamento...", "INFO")
                try:
                    modulo_link = wait.until(
                        EC.element_to_be_clickable((By.CSS_SELECTOR, "a[href='/site/trocarModulo?modulo=2']")))
                    modulo_link.click()
                    time.sleep(2)
                    log_message("✅ Navegação para módulo de faturamento realizada", "SUCCESS")
                except Exception as e:
                    log_message(f"⚠️ Erro ao navegar para módulo: {e}", "WARNING")
                    driver.get("https://pathoweb.com.br/moduloFaturamento/index")
                    time.sleep(2)
                    log_message("🔄 Navegação direta para módulo realizada", "INFO")

            elif "moduloFaturamento" in current_url:
                log_message("✅ Já está no módulo de faturamento - pulando navegação", "SUCCESS")
            else:
                log_message(f"⚠️ URL inesperada detectada: {current_url}", "WARNING")
                driver.get("https://pathoweb.com.br/moduloFaturamento/index")
                time.sleep(2)
                log_message("🔄 Navegação direta para módulo realizada (fallback)", "INFO")

            try:
                modal_close_button = driver.find_element(By.CSS_SELECTOR, "#mensagemParaClienteModal .modal-footer button")
                if modal_close_button.is_displayed():
                    modal_close_button.click()
                    time.sleep(1)
            except Exception:
                pass

            # Acessar explicitamente a página do módulo de faturamento
            log_message("Acessando módulo de faturamento via URL...", "INFO")
            driver.get("https://pathoweb.com.br/moduloFaturamento/index")

            # Clicar no botão "Preparar exames para fatura"
            log_message("Clicando em 'Preparar exames para fatura'...", "INFO")
            try:
                preparar_btn = wait.until(EC.element_to_be_clickable((
                    By.CSS_SELECTOR,
                    "a.btn.btn-danger.chamadaAjax.setupAjax[data-url='/moduloFaturamento/preFaturamento']"
                )))
                preparar_btn.click()
            except Exception:
                preparar_btn = wait.until(EC.element_to_be_clickable((
                    By.XPATH,
                    "//a[contains(@class, 'setupAjax') and contains(text(), 'Preparar exames para fatura')]"
                )))
                preparar_btn.click()

            # Aguardar possível spinner/modal carregar
            try:
                WebDriverWait(driver, 3).until(EC.presence_of_element_located((By.ID, "spinner")))
                log_message("🔄 Modal de carregamento detectado, aguardando...", "INFO")
                WebDriverWait(driver, 30).until(EC.invisibility_of_element_located((By.ID, "spinner")))
                log_message("✅ Modal de carregamento fechado", "INFO")
            except Exception:
                time.sleep(1)

            log_message("Tela de Pré Faturamento aberta.", "SUCCESS")

            # Digitar cada exame no campo numeroExame
            for exame in exames_unicos:
                if cancel_flag and cancel_flag.is_set():
                    log_message("Execução cancelada pelo usuário.", "WARNING")
                    break
                try:
                    log_message(f"➡️ Digitando exame: {exame}", "INFO")
                    campo_exame = wait.until(EC.presence_of_element_located((By.ID, "numeroExame")))
                    campo_exame.clear()
                    campo_exame.send_keys(str(exame))
                    time.sleep(0.5)
                except Exception as e:
                    log_message(f"❌ Erro ao preencher exame {exame}: {e}", "ERROR")
                    continue

            log_message("Preenchimento dos exames concluído.", "SUCCESS")

        except Exception as e:
            log_message(f"❌ Erro durante a automação: {e}", "ERROR")
            messagebox.showerror("Erro", f"❌ Erro durante a automação:\n{e}")
        finally:
            driver.quit()


def run(params: dict):
    module = GuiaUnimedModule()
    module.run(params)
