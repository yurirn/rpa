import os
import pandas as pd
import time
from tkinter import messagebox
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from dotenv import load_dotenv

from src.core.browser_factory import BrowserFactory
from src.core.logger import log_message

load_dotenv()

def get_unique_exames(file_path: str, modo_busca: str) -> list:
    if modo_busca == "exame":
        df = pd.read_excel(file_path)
        unique_exames = df['Exame'].dropna().unique().tolist()
    elif modo_busca == "guia":
        df = pd.read_excel(file_path, header=None)
        unique_exames = df.iloc[:, 0].dropna().unique().tolist()
    else:
        raise ValueError("Modo de busca inválido. Use 'exame' ou 'guia'.")
    return unique_exames

def run(params: dict):
    username = params.get("username")
    password = params.get("password")
    excel_file = params.get("excel_file")
    modo_busca = params.get("modo_busca", "exame")  # padrão: exame

    if not username or not password or not excel_file:
        messagebox.showwarning("Campos vazios", "Preencha usuário, senha e selecione o arquivo Excel.")
        return

    try:
        exames_unicos = get_unique_exames(excel_file, modo_busca)
    except Exception as e:
        messagebox.showerror("Erro", f"Erro ao ler o Excel: {e}")
        return

    if not exames_unicos:
        messagebox.showerror("Erro", "Nenhum exame encontrado no arquivo.")
        return

    url = os.getenv("SYSTEM_URL", "https://pathoweb.com.br/login/auth")
    driver = BrowserFactory.create_chrome()
    wait = WebDriverWait(driver, 15)
    resultados = []

    try:
        log_message("Iniciando automação de preparação de exames...", "INFO")
        driver.get(url)
        wait.until(EC.presence_of_element_located((By.ID, "username"))).send_keys(username)
        driver.find_element(By.ID, "password").send_keys(password)
        driver.find_element(By.CSS_SELECTOR, "button[type='submit']").click()
        wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "a[href='/site/trocarModulo?modulo=2']"))).click()
        time.sleep(4)

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
        time.sleep(2)

        for exame in exames_unicos:
            try:
                log_message(f"➡️ Processando {modo_busca}: {exame}", "INFO")
                campo_id = "numeroGuia" if modo_busca == "guia" else "numeroExame"
                campo_exame = wait.until(EC.presence_of_element_located((By.ID, campo_id)))
                campo_exame.clear()
                campo_exame.send_keys(exame)

                wait.until(EC.element_to_be_clickable((By.ID, "pesquisaFaturamento"))).click()
                time.sleep(2)

                if modo_busca == "guia":
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

    except Exception as e:
        log_message(f"❌ Erro durante a automação: {e}", "ERROR")
        messagebox.showerror("Erro", f"❌ Erro durante a automação:\n{e}")
    finally:
        driver.quit()
