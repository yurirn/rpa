import os
import time
import pandas as pd
from datetime import datetime
from tkinter import messagebox
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
from selenium.common.exceptions import TimeoutException, NoSuchElementException, StaleElementReferenceException
from selenium.webdriver.common.action_chains import ActionChains

from src.core.browser_factory import BrowserFactory
from src.core.logger import log_message
from src.modules.base import BaseModule

# Configurações padrão (podem ser sobrescritas via params)
EXCEL_FILE_DEFAULT = "recursos.xlsx"
TIMEOUT_PADRAO = 20


class BaixaRecursoModule(BaseModule):
    def __init__(self):
        super().__init__(nome="Baixa de Recurso")

    # ========================== Funções auxiliares ==========================

    def _esperar_elemento_clicavel(self, driver, by, value, timeout=TIMEOUT_PADRAO, descricao="elemento"):
        try:
            wait = WebDriverWait(driver, timeout)
            elemento = wait.until(EC.element_to_be_clickable((by, value)))
            return elemento
        except TimeoutException:
            log_message(f"⏱️ Timeout ao aguardar {descricao} estar clicável", "WARNING")
            return None
        except Exception as e:
            log_message(f"❌ Erro ao aguardar {descricao}: {type(e).__name__}", "ERROR")
            return None

    def _esperar_elemento_presente(self, driver, by, value, timeout=TIMEOUT_PADRAO, descricao="elemento"):
        try:
            wait = WebDriverWait(driver, timeout)
            elemento = wait.until(EC.presence_of_element_located((by, value)))
            return elemento
        except TimeoutException:
            log_message(f"⏱️ Timeout ao aguardar {descricao} estar presente", "WARNING")
            return None
        except Exception as e:
            log_message(f"❌ Erro ao aguardar {descricao}: {type(e).__name__}", "ERROR")
            return None

    def _esperar_elemento_visivel(self, driver, by, value, timeout=TIMEOUT_PADRAO, descricao="elemento"):
        try:
            wait = WebDriverWait(driver, timeout)
            elemento = wait.until(EC.visibility_of_element_located((by, value)))
            return elemento
        except TimeoutException:
            log_message(f"⏱️ Timeout ao aguardar {descricao} estar visível", "WARNING")
            return None
        except Exception as e:
            log_message(f"❌ Erro ao aguardar {descricao}: {type(e).__name__}", "ERROR")
            return None

    def _scroll_to_element(self, driver, element):
        try:
            driver.execute_script("arguments[0].scrollIntoView({behavior: 'smooth', block: 'center'});", element)
            time.sleep(0.5)
            return True
        except Exception as e:
            log_message(f"⚠️ Erro ao rolar para elemento: {e}", "WARNING")
            return False

    def _click_element_safe(self, driver, element, descricao="elemento", tentar_js=True):
        tentativas = [
            ("ActionChains", lambda: ActionChains(driver).move_to_element(element).click().perform()),
            ("click() direto", lambda: element.click()),
        ]
        if tentar_js:
            tentativas.append(("JavaScript", lambda: driver.execute_script("arguments[0].click();", element)))

        for metodo, acao in tentativas:
            try:
                self._scroll_to_element(driver, element)
                acao()
                time.sleep(0.3)
                return True
            except StaleElementReferenceException:
                log_message(f"⚠️ Elemento stale no método {metodo}, tentando próximo...", "WARNING")
                continue
            except Exception:
                if metodo == tentativas[-1][0]:
                    log_message(f"❌ Falha ao clicar em {descricao} após {len(tentativas)} tentativas", "ERROR")
                    return False
                continue
        return False

    def _fechar_modais_abertos(self, driver):
        try:
            botoes_fechar = driver.find_elements(
                By.CSS_SELECTOR,
                ".modal .close, .modal button[data-dismiss='modal'], .modal .btn-close",
            )
            for btn in botoes_fechar:
                try:
                    if btn.is_displayed():
                        btn.click()
                        time.sleep(0.3)
                except Exception:
                    pass

            driver.execute_script(
                """
                var modals = document.querySelectorAll('.modal');
                modals.forEach(function(modal) {
                    modal.style.display = 'none';
                    modal.classList.remove('show');
                });
                var backdrops = document.querySelectorAll('.modal-backdrop');
                backdrops.forEach(function(backdrop) { backdrop.remove(); });
                document.body.classList.remove('modal-open');
                document.body.style.overflow = '';
                document.body.style.paddingRight = '';
                """
            )
            time.sleep(0.3)
            return True
        except Exception as e:
            log_message(f"⚠️ Erro ao fechar modais: {e}", "WARNING")
            return False

    def _preencher_campo_valor(self, driver, campo_selector, valor, tentativas=5):
        if pd.isna(valor) or valor == "" or valor == 0:
            return True
        try:
            valor_float = float(valor)
            valor_formatado = f"{valor_float:.2f}".replace(".", ",")
        except (ValueError, TypeError):
            log_message(f"⚠️ Valor inválido: {valor}", "WARNING")
            return False

        try:
            wait = WebDriverWait(driver, 15)
            campo = None
            try:
                campo = wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR, campo_selector)))
            except Exception:
                try:
                    campo = wait.until(EC.visibility_of_element_located((By.NAME, "valorAcatado")))
                except Exception:
                    campo = wait.until(EC.visibility_of_element_located((By.XPATH, "//input[contains(@name, 'valor')]")))
            wait.until(EC.element_to_be_clickable(campo))
            time.sleep(0.3)
        except TimeoutException:
            log_message("⏱️ Campo de valor não ficou visível/clicável a tempo", "WARNING")
            return False

        for tentativa in range(tentativas):
            try:
                driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", campo)
                time.sleep(0.2)
                campo.click()
                time.sleep(0.2)
                campo.clear()
                time.sleep(0.2)
                driver.execute_script("arguments[0].value = '';", campo)
                time.sleep(0.2)
                campo.send_keys(valor_formatado)
                time.sleep(0.5)
                for event in ["input", "change", "blur"]:
                    driver.execute_script(
                        "arguments[0].dispatchEvent(new Event(arguments[1], { bubbles: true }));",
                        campo,
                        event,
                    )
                time.sleep(0.3)
                valor_atual = campo.get_attribute("value")
                if (
                    valor_formatado in valor_atual
                    or valor_atual.replace(".", "").replace(",", ".")
                    == f"{valor_float:.2f}"
                ):
                    log_message(f"Valor preenchido: R$ {valor_formatado}", "INFO")
                    return True
                if tentativa < tentativas - 1:
                    log_message("Tentativa de preenchimento de valor falhou, tentando novamente...", "WARNING")
            except Exception as e:
                log_message(f"Tentativa {tentativa + 1} falhou: {type(e).__name__}", "WARNING")
            if tentativa < tentativas - 1:
                time.sleep(0.8)
        log_message("Campo de valor não foi preenchido corretamente após múltiplas tentativas", "ERROR")
        return False

    def _preencher_campo_select(self, driver, campo_selector, valor, descricao="select", por_label=False):
        if pd.isna(valor) or str(valor).strip() == "":
            return True
        try:
            wait = WebDriverWait(driver, 15)
            select_element = None
            try:
                select_element = wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR, campo_selector)))
            except Exception:
                try:
                    select_element = wait.until(
                        EC.visibility_of_element_located((By.ID, campo_selector.replace("#", "")))
                    )
                except Exception:
                    select_element = wait.until(EC.visibility_of_element_located((By.NAME, campo_selector)))
            wait.until(EC.element_to_be_clickable(select_element))
            time.sleep(0.3)
            driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", select_element)
            time.sleep(0.2)
            select = Select(select_element)
            valor_str = str(valor).strip()
            if not por_label:
                try:
                    select.select_by_value(valor_str)
                    log_message(f"{descricao} selecionado por value: {valor_str}", "INFO")
                    time.sleep(0.3)
                    return True
                except NoSuchElementException:
                    log_message(f"Value '{valor_str}' não encontrado, tentando por texto...", "WARNING")
            try:
                select.select_by_visible_text(valor_str)
                log_message(f"{descricao} selecionado por texto: {valor_str}", "INFO")
                time.sleep(0.3)
                return True
            except NoSuchElementException:
                pass
            try:
                for option in select.options:
                    if valor_str.lower() in option.text.lower():
                        select.select_by_visible_text(option.text)
                        log_message(f"{descricao} selecionado (parcial): {option.text}", "INFO")
                        time.sleep(0.3)
                        return True
            except Exception:
                pass
            log_message(f"Tentando selecionar {descricao} via JavaScript...", "WARNING")
            script = f"""
                const select = document.querySelector('{campo_selector}');
                if (select) {{
                    let opcao = Array.from(select.options).find(opt => opt.value === '{valor_str}');
                    if (!opcao) {{
                        opcao = Array.from(select.options).find(opt =>
                            opt.text.toLowerCase().includes('{valor_str}'.toLowerCase())
                        );
                    }}
                    if (opcao) {{
                        select.value = opcao.value;
                        select.dispatchEvent(new Event('change', {{ bubbles: true }}));
                        return true;
                    }}
                }}
                return false;
            """
            try:
                resultado = driver.execute_script(script)
            except Exception:
                resultado = False
            if resultado:
                log_message(f"{descricao} selecionado via JS", "INFO")
                time.sleep(0.3)
                return True
            log_message(f"Valor '{valor_str}' não encontrado no {descricao}", "ERROR")
            return False
        except TimeoutException:
            log_message(f"Timeout: {descricao} não ficou visível", "WARNING")
            return False
        except Exception as e:
            log_message(f"Erro ao preencher {descricao}: {type(e).__name__} - {e}", "ERROR")
            return False

    def _preencher_campo_texto(self, driver, campo_selector, valor, descricao="campo"):
        if pd.isna(valor) or str(valor).strip() == "":
            return True
        try:
            wait = WebDriverWait(driver, 15)
            campo = None
            try:
                campo = wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR, campo_selector)))
            except Exception:
                try:
                    campo = wait.until(
                        EC.visibility_of_element_located((By.ID, campo_selector.replace("#", "")))
                    )
                except Exception:
                    campo = wait.until(EC.visibility_of_element_located((By.NAME, campo_selector)))
            wait.until(EC.element_to_be_clickable(campo))
            time.sleep(0.3)
            driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", campo)
            time.sleep(0.2)
            campo.click()
            time.sleep(0.2)
            campo.clear()
            time.sleep(0.2)
            valor_str = str(valor).strip()
            campo.send_keys(valor_str)
            time.sleep(0.3)
            valor_atual = campo.get_attribute("value")
            if valor_str in valor_atual or valor_atual in valor_str:
                log_message(f"{descricao} preenchido: {valor_str}", "INFO")
                return True
            log_message(f"{descricao} pode não ter sido preenchido corretamente", "WARNING")
            return True
        except TimeoutException:
            log_message(f"Timeout: {descricao} não ficou visível", "WARNING")
            return False
        except Exception as e:
            log_message(f"Erro ao preencher {descricao}: {type(e).__name__}", "ERROR")
            return False

    def _preencher_campo_textarea(self, driver, campo_selector, valor, descricao="textarea"):
        if pd.isna(valor) or str(valor).strip() == "":
            return True
        try:
            wait = WebDriverWait(driver, 15)
            campo = wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR, campo_selector)))
            wait.until(EC.element_to_be_clickable(campo))
            time.sleep(0.3)
            driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", campo)
            time.sleep(0.2)
            campo.click()
            time.sleep(0.2)
            campo.clear()
            time.sleep(0.2)
            valor_str = str(valor).strip()
            campo.send_keys(valor_str)
            time.sleep(0.3)
            log_message(f"{descricao} preenchido", "INFO")
            return True
        except TimeoutException:
            log_message(f"Timeout: {descricao} não ficou visível", "WARNING")
            return False
        except Exception as e:
            log_message(f"Erro ao preencher {descricao}: {type(e).__name__}", "ERROR")
            return False

    def _preencher_data(self, driver, campo_selector, valor):
        if not valor or pd.isna(valor):
            return True
        try:
            if isinstance(valor, pd.Timestamp):
                valor_formatado = valor.strftime("%Y-%m-%d")
            else:
                valor_formatado = pd.to_datetime(valor, dayfirst=True).strftime("%Y-%m-%d")
        except Exception as e:
            log_message(f"Erro ao formatar data {valor}: {e}", "ERROR")
            return False
        try:
            wait = WebDriverWait(driver, 10)
            campo = wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR, campo_selector)))
            wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, campo_selector)))
            time.sleep(0.3)
            driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", campo)
            time.sleep(0.2)
            script = """
                const el = document.querySelector(arguments[0]);
                if (!el) return;
                el.value = arguments[1];
                el.dispatchEvent(new Event('input', { bubbles: true }));
                el.dispatchEvent(new Event('change', { bubbles: true }));
            """
            driver.execute_script(script, campo_selector, valor_formatado)
            time.sleep(0.5)
            valor_atual = campo.get_attribute("value")
            if valor_formatado in valor_atual:
                log_message(f"Data preenchida: {valor_formatado}", "INFO")
                return True
            log_message(
                f"Data preenchida divergente: esperado={valor_formatado}, atual={valor_atual}",
                "WARNING",
            )
            return True
        except TimeoutException:
            log_message("Timeout: Campo de data não ficou visível", "WARNING")
            return False
        except Exception as e:
            log_message(f"Erro ao preencher data: {type(e).__name__} - {e}", "ERROR")
            return False

    # ========================== Fluxo de tela ===============================

    def _fazer_login(self, driver, username, password, url_login):
        log_message("Fazendo login no Pathoweb...", "INFO")
        try:
            driver.get(url_login)
            campo_email = self._esperar_elemento_presente(driver, By.ID, "username", descricao="campo de usuário")
            if not campo_email:
                campo_email = self._esperar_elemento_presente(
                    driver,
                    By.XPATH,
                    "//input[@type='email' or @name='username']",
                    descricao="campo de usuário",
                )
            if not campo_email:
                return False
            campo_email.clear()
            campo_email.send_keys(username)
            campo_senha = driver.find_element(By.ID, "password")
            campo_senha.clear()
            campo_senha.send_keys(password)
            btn_entrar = driver.find_element(By.CSS_SELECTOR, "button[type='submit']")
            btn_entrar.click()
            log_message("Credenciais enviadas, aguardando sistema...", "INFO")
            time.sleep(5)
            return True
        except Exception as e:
            log_message(f"Erro no login: {type(e).__name__} - {e}", "ERROR")
            return False

    def _acessar_faturamento(self, driver):
        log_message("Acessando módulo Faturamento...", "INFO")
        try:
            xpath_faturamento = (
                "//a[contains(@href, 'modulo=2')] | "
                "//a[.//h2[contains(text(), 'Faturamento')]] | "
                "//a[contains(text(), 'Faturamento')]"
            )
            link_faturamento = self._esperar_elemento_clicavel(
                driver, By.XPATH, xpath_faturamento, descricao="link Faturamento"
            )
            if not link_faturamento:
                return False
            link_faturamento.click()
            time.sleep(3)
            try:
                modal_close = driver.find_element(By.CSS_SELECTOR, "#mensagemParaClienteModal button")
                if modal_close.is_displayed():
                    modal_close.click()
                    time.sleep(1)
            except Exception:
                pass
            log_message("Módulo Faturamento acessado", "SUCCESS")
            return True
        except Exception as e:
            log_message(f"Erro ao acessar Faturamento: {type(e).__name__} - {e}", "ERROR")
            return False

    def _acessar_recurso(self, driver):
        log_message("Acessando tela de Recurso...", "INFO")
        try:
            time.sleep(2)
            estrategias = [
                ("data-url recurso", "//button[@data-url='/moduloFaturamento/recurso']"),
                ("chamadaAjax + recurso", "//button[contains(@class, 'chamadaAjax') and contains(@data-url, 'recurso')]]"),
                ("texto normalizado", "//button[contains(normalize-space(.), 'Recurso')]]"),
                ("texto + chamadaAjax", "//button[contains(@class, 'chamadaAjax') and contains(text(), 'Recurso')]]"),
                ("svg + texto", "//button[.//svg and contains(text(), 'Recurso')]]"),
            ]
            btn_recurso = None
            for nome_estrategia, xpath in estrategias:
                try:
                    log_message(f"Tentando localizar botão Recurso: {nome_estrategia}", "INFO")
                    elementos = driver.find_elements(By.XPATH, xpath)
                    if not elementos:
                        continue
                    for elem in elementos:
                        try:
                            if elem.is_displayed():
                                btn_recurso = elem
                                break
                        except Exception:
                            continue
                    if btn_recurso:
                        break
                except Exception:
                    continue
            if not btn_recurso:
                log_message("Botão 'Recurso' não encontrado", "ERROR")
                try:
                    screenshot_path = f"debug_botao_recurso_{int(time.time())}.png"
                    driver.save_screenshot(screenshot_path)
                    log_message(f"Screenshot salvo: {screenshot_path}", "INFO")
                except Exception:
                    pass
                return False
            try:
                wait = WebDriverWait(driver, 10)
                wait.until(EC.element_to_be_clickable(btn_recurso))
                time.sleep(0.3)
            except TimeoutException:
                log_message("Botão Recurso demorou para ficar clicável, tentando assim mesmo...", "WARNING")
            if not self._click_element_safe(driver, btn_recurso, "botão Recurso"):
                return False
            time.sleep(2)
            log_message("Tela de Recurso acessada", "SUCCESS")
            return True
        except Exception as e:
            log_message(f"Erro ao acessar Recurso: {type(e).__name__} - {e}", "ERROR")
            try:
                screenshot_path = f"erro_acessar_recurso_{int(time.time())}.png"
                driver.save_screenshot(screenshot_path)
                log_message(f"Screenshot salvo: {screenshot_path}", "INFO")
            except Exception:
                pass
            return False

    def _buscar_exame(self, driver, numero_exame):
        log_message(f"Buscando exame {numero_exame}...", "INFO")
        try:
            campo_pesquisa = self._esperar_elemento_visivel(
                driver,
                By.XPATH,
                "//input[@placeholder='Número do exame' or @name='numeroExame' or contains(@id, 'exame')]",
                timeout=10,
                descricao="campo de pesquisa de exame",
            )
            if not campo_pesquisa:
                return False
            campo_pesquisa.clear()
            time.sleep(0.2)
            campo_pesquisa.send_keys(str(numero_exame))
            time.sleep(0.5)
            btn_pesquisar = self._esperar_elemento_clicavel(
                driver,
                By.XPATH,
                "//button[contains(text(), 'Pesquisar')] | //a[contains(text(), 'Pesquisar')]",
                timeout=5,
                descricao="botão Pesquisar",
            )
            if not btn_pesquisar:
                return False
            if not self._click_element_safe(driver, btn_pesquisar, "botão Pesquisar"):
                return False
            time.sleep(2)
            if not self._esperar_elemento_presente(
                driver,
                By.CSS_SELECTOR,
                "table tbody tr",
                timeout=10,
                descricao="tabela de resultados",
            ):
                log_message("Nenhum resultado encontrado para o exame", "WARNING")
                return False
            return True
        except Exception as e:
            log_message(f"Erro ao buscar exame: {type(e).__name__} - {e}", "ERROR")
            return False

    def _encontrar_e_clicar_procedimento(self, driver, codigo_procedimento):
        log_message(f"Localizando procedimento {codigo_procedimento}...", "INFO")
        try:
            time.sleep(1)
            linhas = driver.find_elements(By.CSS_SELECTOR, "table tbody tr")
            if not linhas:
                log_message("Nenhuma linha encontrada na tabela", "ERROR")
                return False
            codigo_str = str(codigo_procedimento).strip()
            for i, linha in enumerate(linhas):
                try:
                    texto_linha = linha.text
                    if codigo_str not in texto_linha:
                        continue
                    log_message(f"Procedimento {codigo_str} encontrado na linha {i + 1}", "INFO")
                    seletores = [
                        "a.btn.btn-success svg.fa-money-bill",
                        "a.btn.btn-success i.fa-money-bill",
                        "a.btn-success[title*='recurso' i]",
                        "a.btn-success",
                    ]
                    botao = None
                    for seletor in seletores:
                        try:
                            elementos = linha.find_elements(By.CSS_SELECTOR, seletor)
                            if elementos:
                                if "svg" in seletor or "i." in seletor:
                                    botao = elementos[0].find_element(By.XPATH, "./ancestor::a")
                                else:
                                    botao = elementos[0]
                                break
                        except Exception:
                            continue
                    if not botao:
                        log_message(
                            f"Botão de recurso não encontrado na linha {i + 1} para o procedimento",
                            "WARNING",
                        )
                        continue
                    self._scroll_to_element(driver, botao)
                    if not self._click_element_safe(
                        driver, botao, f"botão do procedimento {codigo_str}"
                    ):
                        continue
                    time.sleep(2)
                    return True
                except StaleElementReferenceException:
                    log_message(f"Elemento stale na linha {i + 1}, continuando...", "WARNING")
                    continue
                except Exception as e:
                    log_message(f"Erro na linha {i + 1}: {type(e).__name__}", "WARNING")
                    continue
            log_message(f"Código {codigo_str} não encontrado em nenhuma linha", "ERROR")
            return False
        except Exception as e:
            log_message(f"Erro ao localizar procedimento: {type(e).__name__} - {e}", "ERROR")
            return False

    def _preencher_formulario_recurso(self, driver, row):
        log_message("Preenchendo formulário de recurso...", "INFO")
        try:
            wait = WebDriverWait(driver, 15)
            try:
                wait.until(
                    EC.visibility_of_element_located(
                        (By.CSS_SELECTOR, "#dataRecebimento, #conta\\.id, #tipoPagamento")
                    )
                )
            except TimeoutException:
                log_message("Formulário demorou para carregar", "WARNING")
            time.sleep(1)
            sucesso_total = True
            data_receb = row.get("Data de recebimento")
            if pd.notna(data_receb):
                if not self._preencher_data(driver, "#dataRecebimento", data_receb):
                    sucesso_total = False
            conta_id = row.get("Conta")
            if pd.notna(conta_id) and str(conta_id).strip() != "":
                if not self._preencher_campo_select(
                    driver, "#conta\\.id", conta_id, "Conta", por_label=False
                ):
                    sucesso_total = False
            nota = row.get("Nota")
            if pd.notna(nota) and str(nota).strip() != "":
                if not self._preencher_campo_texto(driver, "#numeroDocumento", nota, "Nota"):
                    sucesso_total = False
            tipo_pag = row.get("Tipo de pagamento")
            if pd.notna(tipo_pag) and str(tipo_pag).strip() != "":
                if not self._preencher_campo_select(
                    driver,
                    "#tipoPagamento",
                    tipo_pag,
                    "Tipo de pagamento",
                    por_label=True,
                ):
                    sucesso_total = False
            justificativa = row.get("Justificativa recurso")
            if pd.notna(justificativa) and str(justificativa).strip() != "":
                if not self._preencher_campo_textarea(
                    driver,
                    "#justificativaRecurso",
                    justificativa,
                    "Justificativa",
                ):
                    sucesso_total = False
            valor = row.get("O valor a ser recebido é")
            if pd.notna(valor) and valor != 0:
                if not self._preencher_campo_valor(
                    driver, 'input[name="valorAcatado"]', valor
                ):
                    sucesso_total = False
            if sucesso_total:
                log_message("Formulário preenchido com sucesso", "SUCCESS")
            else:
                log_message("Formulário preenchido com algumas falhas", "WARNING")
            return True
        except Exception as e:
            log_message(f"Erro ao preencher formulário: {type(e).__name__} - {e}", "ERROR")
            return False

    def _executar_acao(self, driver, acao):
        log_message(f"Executando ação: {acao}", "INFO")
        acao_lower = acao.lower().strip()
        acoes_mapping = {
            "glosar": {
                "data_url": "/moduloFaturamento/glosarRecursoDefinitivo",
                "classe": "btn-danger",
            },
            "receber": {
                "data_url": "/moduloFaturamento/receberRecurso",
                "classe": "btn-primary",
            },
            "receber e gerar novo": {
                "data_url": "/moduloFaturamento/receberRecursoGerarDiferenca",
                "classe": "btn-success",
            },
        }
        if acao_lower not in acoes_mapping:
            log_message(
                f"Ação '{acao}' não reconhecida. Válidas: {', '.join(acoes_mapping.keys())}",
                "ERROR",
            )
            return False
        acao_config = acoes_mapping[acao_lower]
        try:
            estrategias = [
                ("data-url específico", f"//a[@data-url='{acao_config['data_url']}']"),
                (
                    "data-url + classe",
                    f"//a[@data-url='{acao_config['data_url']}' and contains(@class, '{acao_config['classe']}')]",
                ),
                (
                    "classe + chamadaAjax",
                    f"//a[contains(@class, '{acao_config['classe']}') and contains(@class, 'chamadaAjax')]",
                ),
            ]
            btn_acao = None
            for nome_estrategia, xpath in estrategias:
                try:
                    log_message(f"Tentando localizar botão ação: {nome_estrategia}", "INFO")
                    elementos = driver.find_elements(By.XPATH, xpath)
                    if not elementos:
                        continue
                    for elem in elementos:
                        try:
                            if elem.is_displayed():
                                btn_acao = elem
                                break
                        except Exception:
                            continue
                    if btn_acao:
                        break
                except Exception:
                    continue
            if not btn_acao:
                log_message(f"Botão para ação '{acao}' não encontrado", "ERROR")
                return False
            try:
                wait = WebDriverWait(driver, 10)
                wait.until(EC.element_to_be_clickable(btn_acao))
                time.sleep(0.3)
            except TimeoutException:
                log_message("Botão de ação demorou para ficar clicável, tentando assim mesmo...", "WARNING")
            if not self._click_element_safe(driver, btn_acao, f"botão '{acao}'"):
                return False
            time.sleep(1)
            return True
        except Exception as e:
            log_message(f"Erro ao executar ação: {type(e).__name__} - {e}", "ERROR")
            return False

    def _fechar_modal_sucesso(self, driver):
        log_message("Fechando modal de sucesso...", "INFO")
        try:
            xpaths = [
                "//button[contains(text(), 'Fechar')]",
                "//button[@data-dismiss='modal']",
                "//button[contains(@class, 'close')]",
            ]
            for xpath in xpaths:
                btn_fechar = self._esperar_elemento_clicavel(
                    driver, By.XPATH, xpath, timeout=5, descricao="botão Fechar"
                )
                if btn_fechar and self._click_element_safe(driver, btn_fechar, "botão Fechar"):
                    time.sleep(1)
                    return True
            self._fechar_modais_abertos(driver)
            return True
        except Exception as e:
            log_message(f"Erro ao fechar modal: {type(e).__name__}", "WARNING")
            self._fechar_modais_abertos(driver)
            return True

    # ========================== Processamento ===============================

    def _carregar_planilha(self, excel_file: str) -> pd.DataFrame | None:
        try:
            if not os.path.exists(excel_file):
                log_message(f"Arquivo de recursos não encontrado: {excel_file}", "ERROR")
                return None
            df = pd.read_excel(excel_file)
            df["Status"] = ""
            df.columns = df.columns.str.strip()
            colunas_necessarias = [
                "Número do exame",
                "Código do procedimento",
                "Ação",
            ]
            faltantes = [c for c in colunas_necessarias if c not in df.columns]
            if faltantes:
                log_message(f"Colunas faltantes na planilha: {faltantes}", "ERROR")
                return None
            log_message(f"Planilha carregada: {len(df)} recursos", "INFO")
            return df
        except Exception as e:
            log_message(f"Erro ao carregar planilha: {type(e).__name__} - {e}", "ERROR")
            return None

    def _salvar_resultados(self, df: pd.DataFrame, excel_file: str) -> str | None:
        try:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            df_falhas = df[df["Status"] != "OK"].copy()
            if len(df_falhas) == 0:
                log_message("Todos os recursos foram completados com sucesso. Nenhum arquivo gerado.", "SUCCESS")
                return None
            output_path = (
                os.path.splitext(excel_file)[0]
                + f"_NAO_COMPLETADOS_{timestamp}.xlsx"
            )
            df_falhas.to_excel(output_path, index=False)
            log_message(
                f"Planilha de recursos não completados salva em: {output_path}",
                "INFO",
            )
            return output_path
        except Exception as e:
            log_message(f"Erro ao salvar planilha de resultados: {e}", "ERROR")
            return None

    def _processar_recurso(self, driver, row, index, total):
        numero_exame = row.get("Número do exame")
        codigo_proc = row.get("Código do procedimento")
        acao = row.get("Ação", "").strip().lower()
        log_message(
            f"Processando recurso {index + 1}/{total} - Exame {numero_exame} | Proc {codigo_proc} | Ação {acao}",
            "INFO",
        )
        self._fechar_modais_abertos(driver)
        time.sleep(0.5)
        try:
            acoes_validas = ["glosar", "receber", "receber e gerar novo"]
            if acao not in acoes_validas:
                return False, f"Ação '{acao}' não reconhecida"
            if not self._buscar_exame(driver, numero_exame):
                return False, "Falha ao buscar exame"
            if not self._encontrar_e_clicar_procedimento(driver, codigo_proc):
                return False, f"Procedimento {codigo_proc} não encontrado"
            if not self._preencher_formulario_recurso(driver, row):
                return False, "Falha ao preencher formulário"
            if not self._executar_acao(driver, acao):
                return False, f"Falha ao executar ação '{acao}'"
            self._fechar_modal_sucesso(driver)
            return True, "OK"
        except Exception as e:
            erro = f"{type(e).__name__} - {str(e)[:100]}"
            log_message(f"Erro ao processar recurso: {erro}", "ERROR")
            try:
                screenshot_path = f"erro_recurso_{index}_{int(time.time())}.png"
                driver.save_screenshot(screenshot_path)
                log_message(f"Screenshot salvo: {screenshot_path}", "INFO")
            except Exception:
                pass
            return False, erro

    def _exibir_resumo(self, df: pd.DataFrame):
        total = len(df)
        sucesso = len(df[df["Status"] == "OK"])
        falhas = total - sucesso
        log_message(f"Total de recursos: {total}", "INFO")
        log_message(f"Completados com sucesso: {sucesso}", "SUCCESS")
        log_message(f"Não completados: {falhas}", "WARNING")
        if falhas > 0:
            df_falhas = df[df["Status"] != "OK"].copy()
            for _, row in df_falhas.iterrows():
                log_message(
                    f"Exame {row.get('Número do exame')} | Proc {row.get('Código do procedimento')} | Motivo: {row['Status']}",
                    "ERROR",
                )
        messagebox.showinfo(
            "Processamento Concluído",
            f"Total: {total}\nSucesso: {sucesso}\nFalhas: {falhas}",
        )

    # ========================== Interface módulo ============================

    def run(self, params: dict):
        username = params.get("username")
        password = params.get("password")
        excel_file = params.get("excel_file", EXCEL_FILE_DEFAULT)
        cancel_flag = params.get("cancel_flag")
        headless_mode = params.get("headless_mode", False)
        url = params.get(
            "url_login",
            os.getenv("SYSTEM_URL", "https://dap.pathoweb.com.br/login/auth"),
        )
        df = self._carregar_planilha(excel_file)
        if df is None:
            messagebox.showerror(
                "Erro",
                "Não foi possível carregar a planilha de recursos. Verifique o arquivo e as colunas.",
            )
            return
        driver = None
        try:
            driver = BrowserFactory.create_chrome(headless=headless_mode)
            log_message("Navegador inicializado para baixa de recurso", "INFO")
            if not self._fazer_login(driver, username, password, url):
                messagebox.showerror("Erro", "Falha no login no Pathoweb.")
                return
            if not self._acessar_faturamento(driver):
                messagebox.showerror(
                    "Erro", "Falha ao acessar módulo de Faturamento no Pathoweb."
                )
                return
            if not self._acessar_recurso(driver):
                messagebox.showerror("Erro", "Falha ao acessar tela de Recurso.")
                return
            total = len(df)
            log_message(f"Iniciando processamento de {total} recursos...", "INFO")
            for index, row in df.iterrows():
                if cancel_flag and cancel_flag.is_set():
                    log_message("Execução cancelada pelo usuário.", "WARNING")
                    break
                sucesso, mensagem = self._processar_recurso(driver, row, index, total)
                df.at[index, "Status"] = mensagem
                if index < total - 1:
                    time.sleep(1)
            self._salvar_resultados(df, excel_file)
            self._exibir_resumo(df)
        except Exception as e:
            log_message(f"Erro crítico durante a automação de baixa de recurso: {e}", "ERROR")
            messagebox.showerror(
                "Erro",
                f"Erro crítico durante a automação de baixa de recurso:\n{str(e)[:200]}",
            )
        finally:
            if driver:
                try:
                    driver.quit()
                except Exception:
                    pass


def run(params: dict):
    module = BaixaRecursoModule()
    module.run(params)

