import os
import pandas as pd
import time
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

class PreparacaoLoteModule(BaseModule):
    def __init__(self):
        super().__init__(nome="Preparação de Lote")

    def get_unique_exames(self, file_path: str, modo_busca: str) -> list:
        if modo_busca == "exame":
            #Caso precise buscar por tab do excel, usar o sheet_name
            #df = pd.read_excel(file_path, sheet_name=2)
            df = pd.read_excel(file_path)
            unique_exames = df['Exame'].dropna().unique().tolist()
        elif modo_busca == "guia":
            df = pd.read_excel(file_path, header=None)
            unique_exames = df.iloc[:, 0].dropna().unique().tolist()
        else:
            raise ValueError("Modo de busca inválido. Use 'exame' ou 'guia'.")
        return unique_exames

    def run(self, params: dict):
        username = params.get("username")
        password = params.get("password")
        excel_file = params.get("excel_file")
        modo_busca = params.get("modo_busca", "exame")  # padrão: exame
        cancel_flag = params.get("cancel_flag")
        gera_xml_tiss = params.get("gera_xml_tiss", "sim")
        headless_mode = params.get("headless_mode")
        try:
            exames_unicos = self.get_unique_exames(excel_file, modo_busca)
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao ler o Excel: {e}")
            return
        if not exames_unicos:
            messagebox.showerror("Erro", "Nenhum exame encontrado no arquivo.")
            return
        url = os.getenv("SYSTEM_URL", "https://pathoweb.com.br/login/auth")
        driver = BrowserFactory.create_chrome(headless=headless_mode)
        wait = WebDriverWait(driver, 15)
        resultados = []
        try:
            log_message("Iniciando automação de preparação de exames...", "INFO")
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
                    # Tentar navegar diretamente pela URL como fallback
                    driver.get("https://pathoweb.com.br/moduloFaturamento/index")
                    time.sleep(2)
                    log_message("🔄 Navegação direta para módulo realizada", "INFO")

            elif "moduloFaturamento" in current_url:
                log_message("✅ Já está no módulo de faturamento - pulando navegação", "SUCCESS")
            else:
                log_message(f"⚠️ URL inesperada detectada: {current_url}", "WARNING")
                # Tentar navegar diretamente como fallback
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
            wait.until(EC.element_to_be_clickable((
                By.XPATH, "//a[contains(@class, 'setupAjax') and contains(text(), 'Preparar exames para fatura')]"
            ))).click()
            time.sleep(1)
            for exame in exames_unicos:
                if cancel_flag and cancel_flag.is_set():
                    log_message("Execução cancelada pelo usuário.", "WARNING")
                    break
                try:
                    log_message(f"➡️ Processando {modo_busca}: {exame}", "INFO")
                    campo_id = "numeroGuia" if modo_busca == "guia" else "numeroExame"
                    campo_exame = wait.until(EC.presence_of_element_located((By.ID, campo_id)))
                    campo_exame.clear()
                    campo_exame.send_keys(exame)
                    wait.until(EC.element_to_be_clickable((By.ID, "pesquisaFaturamento"))).click()
                    try:
                        modal_carregando = driver.find_element(By.XPATH,
                            "//div[contains(@class,'modal-body') and contains(., 'Carregando')]")
                        if modal_carregando.is_displayed():
                            log_message("🔄 Modal de carregamento detectado, aguardando...", "INFO")
                            WebDriverWait(driver, 30).until(EC.invisibility_of_element_located((By.ID, "spinner")))
                            log_message("✅ Modal de carregamento fechado", "INFO")
                    except Exception:
                        log_message("ℹ️ Modal não detectado. Prosseguindo...", "INFO")
                    time.sleep(1)
                    try:
                        tbody_rows = driver.find_elements(By.CSS_SELECTOR, "#tabelaPreFaturamentoTbody tr")
                        if len(tbody_rows) == 0:
                            log_message(f"⚠️ Nenhum resultado encontrado para {exame}. Pulando.", "WARNING")
                            resultados.append({"exame": exame, "status": "sem_resultados"})
                            continue
                    except Exception as e:
                        log_message(f"⚠️ Erro ao validar resultados da tabela: {e}", "WARNING")
                        resultados.append({"exame": exame, "status": "erro_validacao", "erro": str(e)})
                        continue
                    wait.until(EC.element_to_be_clickable((By.ID, "checkTodosPreFaturar"))).click()
                    acoes_btn = wait.until(EC.element_to_be_clickable((
                        By.XPATH, "//a[contains(@class, 'toggleMaisDeUm') and contains(., 'Ações')]"
                    )))
                    acoes_btn.click()
                    time.sleep(1)
                    driver.execute_script("""
                        const onlineBtn = document.querySelector("a[data-url*='statusConferido=O']");
                        if (onlineBtn) { onlineBtn.click(); }
                    """)
                    time.sleep(1)
                    if modo_busca == "guia":
                        try:
                            WebDriverWait(driver, 3).until(EC.presence_of_element_located((By.ID, "spinner")))
                            log_message("🔄 Modal de carregamento detectado, aguardando...", "INFO")
                            WebDriverWait(driver, 30).until(EC.invisibility_of_element_located((By.ID, "spinner")))
                            log_message("✅ Modal de carregamento fechado", "INFO")
                        except Exception:
                            log_message("ℹ️ Modal não detectado. Prosseguindo...", "INFO")
                            time.sleep(1)
                    resultados.append({"exame": exame, "status": "sucesso"})
                    log_message(f"✅ {modo_busca.title()} {exame} processado com sucesso.", "SUCCESS")
                except Exception as e:
                    resultados.append({"exame": exame, "status": "erro", "erro": str(e)})
                    log_message(f"❌ Erro ao processar {exame}: {e}", "ERROR")
            total = len(resultados)
            sucesso = [r for r in resultados if r["status"] == "sucesso"]
            erro = [r for r in resultados if r["status"] == "erro"]
            sem_resultados = [r for r in resultados if r["status"] == "sem_resultados"]
            erro_validacao = [r for r in resultados if r["status"] == "erro_validacao"]
            log_message("\nResumo do processamento:", "INFO")
            log_message(f"Total: {total}", "INFO")
            log_message(f"Sucesso: {len(sucesso)}", "SUCCESS")
            log_message(f"Sem resultados: {len(sem_resultados)}", "WARNING")
            log_message(f"Erro validação: {len(erro_validacao)}", "WARNING")
            log_message(f"Erro processamento: {len(erro)}", "ERROR")
            messagebox.showinfo("Sucesso",
                f"✅ Processamento finalizado!\n"
                f"Total: {total}\n"
                f"Sucesso: {len(sucesso)}\n"
                f"Sem resultados: {len(sem_resultados)}\n"
                f"Erros: {len(erro) + len(erro_validacao)}"
            )

            if gera_xml_tiss == "nao":
                log_message("Executando rotina de geração de lote manual...", "INFO")
                try:
                    campo_id = "numeroGuia" if modo_busca == "guia" else "numeroExame"
                    campo_exame = wait.until(EC.presence_of_element_located((By.ID, campo_id)))
                    campo_exame.clear()
                    time.sleep(0.5)

                    select2_container = wait.until(
                        EC.element_to_be_clickable(
                            (By.CSS_SELECTOR, ".select2-selection[aria-labelledby*='convenioId']"))
                    )
                    select2_container.click()
                    time.sleep(1)
                    opcao_unimed = wait.until(
                        EC.element_to_be_clickable(
                            (By.XPATH,
                             "//li[contains(@class, 'select2-results__option') and text()='UNIMED (LONDRINA)']"))
                    )
                    opcao_unimed.click()
                    time.sleep(1)

                    select_element = wait.until(EC.presence_of_element_located((By.ID, "conferido")))
                    select_conferido = Select(select_element)
                    select_conferido.select_by_value("O")
                    time.sleep(1)

                    # Executar pesquisa
                    botao_pesquisar = wait.until(EC.element_to_be_clickable((By.ID, "pesquisaFaturamento")))
                    botao_pesquisar.click()
                    time.sleep(2)
                    # Aguardar finalização da pesquisa
                    tempo_maximo = time.time() + 60
                    while time.time() < tempo_maximo:
                        try:
                            modal_carregando = driver.find_element(By.XPATH,
                                "//div[contains(@class,'modal-body') and contains(., 'Carregando')]")
                            if modal_carregando.is_displayed():
                                time.sleep(1)
                            else:
                                break
                        except Exception:
                            break
                    try:
                        gerar_tiss_checkbox = driver.find_element(By.ID, "gerarArquivoTiss")
                        if gerar_tiss_checkbox.is_selected():
                            gerar_tiss_checkbox.click()
                            log_message("Checkbox 'gerarArquivoTiss' desmarcado.", "INFO")
                        time.sleep(1)
                    except Exception as e:
                        log_message(f"Não foi possível desmarcar gerarArquivoTiss: {e}", "WARNING")
                    # Clicar no botão de situação de faturamento
                    try:
                        botao_situacao = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "a.btn.btn-danger[onclick*='modalFaturamento']")))
                        botao_situacao.click()
                        log_message("Botão de situação de faturamento clicado.", "INFO")

                        modal_carregando = driver.find_element(By.XPATH, "//div[contains(@class,'modal-body') and contains(., 'Carregando')]")
                        if modal_carregando.is_displayed():
                            time.sleep(1)

                    except Exception as e:
                        log_message(f"Não foi possível clicar no botão de situação de faturamento: {e}", "ERROR")
                except Exception as e:
                    log_message(f"Erro na rotina de geração de XML TISS manual: {e}", "ERROR")
        except Exception as e:
            log_message(f"❌ Erro durante a automação: {e}", "ERROR")
            messagebox.showerror("Erro", f"❌ Erro durante a automação:\n{e}")
        finally:
            driver.quit()

def run(params: dict):
    module = PreparacaoLoteModule()
    module.run(params)
