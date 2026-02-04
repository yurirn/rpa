# src/modules/financeiro/baixa_lote.py
"""
Módulo de Baixa de Lote - Processa pagamentos parciais de exames por lote.
Lê uma planilha Excel com informações de lotes e procedimentos,
acessa o sistema de faturamento e registra os pagamentos parciais.
"""

import os
import time
import pandas as pd
from datetime import datetime
from tkinter import messagebox
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException, StaleElementReferenceException
from selenium.webdriver.common.action_chains import ActionChains

from src.core.browser_factory import BrowserFactory
from src.core.logger import log_message
from src.modules.base import BaseModule

# Configurações padrão
TIMEOUT_PADRAO = 20


class BaixaLoteModule(BaseModule):
    def __init__(self):
        super().__init__(nome="Baixa de Lote")

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

    def _salvar_screenshot(self, driver, nome="erro"):
        try:
            filename = f"{nome}_{int(time.time())}.png"
            driver.save_screenshot(filename)
            log_message(f"📸 Screenshot salvo: {filename}", "INFO")
            return filename
        except Exception as e:
            log_message(f"⚠️ Não foi possível salvar screenshot: {e}", "WARNING")
            return None

    def _salvar_html(self, driver, nome="debug"):
        try:
            filename = f"{nome}_{int(time.time())}.html"
            with open(filename, "w", encoding="utf-8") as f:
                f.write(driver.page_source)
            log_message(f"📄 HTML salvo: {filename}", "INFO")
            return filename
        except Exception as e:
            log_message(f"⚠️ Não foi possível salvar HTML: {e}", "WARNING")
            return None

    # ========================== Fluxo de tela ===============================

    def _fazer_login(self, driver, username, password, url_login):
        log_message("🔐 Iniciando login no Pathoweb...", "INFO")
        try:
            driver.get(url_login)
            time.sleep(3)
            
            campo_email = self._esperar_elemento_presente(
                driver, By.NAME, "j_username", timeout=10, descricao="campo de usuário"
            )
            if not campo_email:
                self._salvar_screenshot(driver, "erro_login")
                self._salvar_html(driver, "erro_login")
                return False
            
            campo_senha = driver.find_element(By.NAME, "j_password")
            
            campo_email.clear()
            campo_email.send_keys(username)
            campo_senha.clear()
            campo_senha.send_keys(password)
            time.sleep(0.5)
            
            # Botão login - tentar múltiplos seletores
            seletores_botao = [
                (By.XPATH, "//button[normalize-space()='Entrar']"),
                (By.XPATH, "//button[contains(text(), 'Entrar')]"),
                (By.XPATH, "//button[contains(text(), 'Login')]"),
                (By.XPATH, "//button[@type='submit']"),
                (By.XPATH, "//input[@type='submit']"),
                (By.CSS_SELECTOR, "button[type='submit']"),
                (By.CSS_SELECTOR, "input[type='submit']"),
                (By.XPATH, "//button[contains(@class, 'btn-primary')]"),
            ]
            
            btn_entrar = None
            for by, selector in seletores_botao:
                try:
                    btn_entrar = driver.find_element(by, selector)
                    break
                except NoSuchElementException:
                    continue
            
            if not btn_entrar:
                self._salvar_screenshot(driver, "erro_botao_login")
                log_message("❌ Botão de login não encontrado", "ERROR")
                return False
            
            btn_entrar.click()
            time.sleep(5)
            log_message("✓ Login realizado com sucesso", "SUCCESS")
            return True
            
        except Exception as e:
            log_message(f"❌ Erro no login: {type(e).__name__} - {e}", "ERROR")
            self._salvar_screenshot(driver, "erro_login")
            return False

    def _acessar_faturamento(self, driver):
        log_message("📊 Acessando módulo de Faturamento...", "INFO")
        try:
            menu_faturamento = self._esperar_elemento_clicavel(
                driver,
                By.XPATH,
                "//a[contains(@href, '/site/trocarModulo?modulo=2') and .//h2[contains(text(), 'Faturamento')]]",
                timeout=20,
                descricao="menu Faturamento"
            )
            if not menu_faturamento:
                return False
            
            menu_faturamento.click()
            time.sleep(3)
            log_message("✓ Módulo de Faturamento acessado", "SUCCESS")
            return True
            
        except Exception as e:
            log_message(f"❌ Erro ao acessar Faturamento: {type(e).__name__}", "ERROR")
            self._salvar_screenshot(driver, "erro_menu_faturamento")
            return False

    def _fechar_modal_inicial(self, driver):
        """Fecha o modal de mensagem que pode aparecer após acessar o módulo"""
        try:
            modal_btn = WebDriverWait(driver, 5).until(
                EC.element_to_be_clickable((
                    By.CSS_SELECTOR,
                    "#mensagemParaClienteModal button[data-dismiss='modal']"
                ))
            )
            modal_btn.click()
            time.sleep(1)
            log_message("✓ Modal inicial fechado", "INFO")
        except TimeoutException:
            pass

    def _acessar_faturas_enviadas(self, driver):
        log_message("📄 Acessando 'Faturas enviadas e recebimento'...", "INFO")
        try:
            botao_faturas = self._esperar_elemento_clicavel(
                driver,
                By.XPATH,
                "//button[contains(@class, 'btn-cabecalho') and contains(., 'Faturas enviadas e recebimento')]",
                timeout=20,
                descricao="botão Faturas enviadas"
            )
            if not botao_faturas:
                return False
            
            botao_faturas.click()
            time.sleep(3)
            log_message("✓ Tela de faturas enviadas acessada", "SUCCESS")
            return True
            
        except Exception as e:
            log_message(f"❌ Erro ao acessar faturas: {type(e).__name__}", "ERROR")
            self._salvar_screenshot(driver, "erro_botao_faturas")
            return False

    def _retornar_tela_busca(self, driver):
        """Retorna à tela de busca clicando em 'Faturas enviadas e recebimento'"""
        log_message("🔄 Retornando à tela de busca...", "INFO")
        
        try:
            time.sleep(2)
            
            # Scroll para o topo da página
            driver.execute_script("window.scrollTo(0, 0);")
            time.sleep(1)
            
            # Fechar todos os modais e overlays
            self._fechar_modais_abertos(driver)
            
            # Buscar o botão usando JavaScript diretamente
            try:
                driver.execute_script("""
                    var botoes = document.querySelectorAll('button.btn-cabecalho');
                    for (var i = 0; i < botoes.length; i++) {
                        if (botoes[i].textContent.includes('Faturas enviadas e recebimento') || 
                            botoes[i].textContent.includes('Faturas enviadas')) {
                            botoes[i].click();
                            return true;
                        }
                    }
                    return false;
                """)
                log_message("✓ Botão 'Faturas enviadas e recebimento' clicado", "INFO")
                time.sleep(3)
            except Exception as e:
                log_message(f"⚠️ Erro ao clicar via JavaScript puro: {e}", "WARNING")
                
                # Fallback: tentar encontrar o elemento e clicar via JavaScript
                botao_faturas = self._esperar_elemento_presente(
                    driver,
                    By.XPATH,
                    "//button[contains(@class, 'btn-cabecalho') and contains(., 'Faturas enviadas e recebimento')]",
                    descricao="botão Faturas enviadas"
                )
                if botao_faturas:
                    driver.execute_script("arguments[0].click();", botao_faturas)
                    log_message("✓ Botão clicado via fallback", "INFO")
                    time.sleep(3)
            
            # Verificar se o campo de busca está disponível
            self._esperar_elemento_presente(driver, By.ID, "numeroLote", descricao="campo de busca de lote")
            log_message("✓ Tela de busca carregada", "SUCCESS")
            return True
            
        except Exception as e:
            log_message(f"❌ Erro ao retornar à tela de busca: {e}", "ERROR")
            self._salvar_screenshot(driver, "erro_retornar_busca")
            return False

    def _buscar_lote(self, driver, lote_str):
        log_message(f"🔍 Buscando lote: {lote_str}", "INFO")
        try:
            input_lote = self._esperar_elemento_presente(
                driver, By.ID, "numeroLote", timeout=20, descricao="campo de busca de lote"
            )
            if not input_lote:
                return False
            
            input_lote.click()
            input_lote.clear()
            time.sleep(0.5)
            input_lote.send_keys(str(lote_str).strip())
            log_message(f"✓ Lote '{lote_str}' inserido no campo de busca", "INFO")
            
            btn_pesquisar = self._esperar_elemento_clicavel(
                driver, By.ID, "pesquisaFaturamento", timeout=10, descricao="botão Pesquisar"
            )
            if not btn_pesquisar:
                return False
            
            btn_pesquisar.click()
            log_message("✓ Botão 'Pesquisar' clicado", "INFO")
            time.sleep(3)
            return True
            
        except Exception as e:
            log_message(f"❌ Erro ao buscar lote: {e}", "ERROR")
            self._salvar_screenshot(driver, "erro_buscar_lote")
            return False

    def _selecionar_checkbox_por_lote(self, driver, lote_str):
        log_message(f"📋 Selecionando checkbox para lote: {lote_str}", "INFO")
        try:
            tabela = self._esperar_elemento_presente(
                driver, By.XPATH, "//table/tbody", timeout=20, descricao="tabela de lotes"
            )
            if not tabela:
                return False
            
            linhas = tabela.find_elements(By.TAG_NAME, "tr")
            
            for linha in linhas:
                try:
                    cols = linha.find_elements(By.TAG_NAME, "td")
                    if len(cols) < 2:
                        continue
                    
                    texto_lote = cols[1].text.strip()
                    
                    if texto_lote == lote_str:
                        checkbox = linha.find_element(
                            By.XPATH, ".//input[@type='checkbox' and contains(@id, 'checkboxSelecao')]"
                        )
                        if not checkbox.is_selected():
                            driver.execute_script("arguments[0].click();", checkbox)
                            log_message(f"✓ Checkbox do lote {lote_str} marcado", "SUCCESS")
                        else:
                            log_message(f"ℹ️ Checkbox do lote {lote_str} já estava marcado", "INFO")
                        time.sleep(1)
                        return True
                except Exception:
                    continue
            
            log_message(f"❌ Checkbox para lote {lote_str} não encontrado", "ERROR")
            return False
            
        except Exception as e:
            log_message(f"❌ Erro ao selecionar checkbox: {e}", "ERROR")
            self._salvar_screenshot(driver, "erro_selecionar_checkbox")
            return False

    def _clicar_botao_receber(self, driver):
        log_message("🔘 Clicando no botão 'Receber'...", "INFO")
        try:
            btn_receber = self._esperar_elemento_clicavel(
                driver,
                By.XPATH,
                "//a[contains(@class, 'btn') and contains(text(), 'Receber')]",
                timeout=10,
                descricao="botão Receber"
            )
            if not btn_receber:
                return False
            
            btn_receber.click()
            log_message("✓ Botão 'Receber' clicado", "SUCCESS")
            time.sleep(3)
            return True
            
        except Exception as e:
            log_message(f"❌ Botão 'Receber' não encontrado: {e}", "ERROR")
            self._salvar_screenshot(driver, "erro_botao_receber")
            return False

    def _encontrar_tabela_procedimentos(self, driver, lote_str):
        """Tenta encontrar a tabela de procedimentos usando múltiplas estratégias"""
        log_message(f"🔍 Procurando tabela de procedimentos para lote {lote_str}...", "INFO")
        
        # Estratégia 1: Buscar por ID específico do lote
        try:
            tabela = driver.find_element(By.ID, f"tdBodyConvenioExame_{lote_str}")
            log_message(f"✓ Tabela encontrada por ID: tdBodyConvenioExame_{lote_str}", "SUCCESS")
            return tabela
        except:
            pass
        
        # Estratégia 2: Buscar tbody visível que contenha botões de pagamento
        try:
            tabelas = driver.find_elements(By.TAG_NAME, "tbody")
            for tabela in tabelas:
                if tabela.is_displayed():
                    linhas = tabela.find_elements(By.TAG_NAME, "tr")
                    if len(linhas) > 0:
                        try:
                            # Verifica se tem o botão de pagamento parcial
                            tabela.find_element(By.XPATH, ".//a[@title='Registrar pagamento parcial do exame']")
                            log_message("✓ Tabela encontrada por tbody visível", "SUCCESS")
                            return tabela
                        except:
                            continue
        except:
            pass
        
        # Estratégia 3: Buscar em divs expandidas
        try:
            divs_expandidos = driver.find_elements(
                By.XPATH, 
                "//div[contains(@style, 'display: block') or contains(@style, 'display: table')]//tbody"
            )
            for div_tbody in divs_expandidos:
                if div_tbody.is_displayed():
                    linhas = div_tbody.find_elements(By.TAG_NAME, "tr")
                    if len(linhas) > 0:
                        log_message("✓ Tabela encontrada em div expandido", "SUCCESS")
                        return div_tbody
        except:
            pass
        
        log_message("❌ Não foi possível encontrar a tabela de procedimentos", "ERROR")
        self._salvar_screenshot(driver, f"erro_tabela_procedimentos_{lote_str}")
        return None

    def _preencher_modal(self, driver, valor_pago, justificativa, recurso):
        """
        Preenche o modal de pagamento parcial que abre após clicar no botão de dólar
        e clica no link correto de salvar conforme recurso ('sim' ou 'não').
        """
        log_message("📝 Preenchendo modal de pagamento parcial...", "INFO")

        try:
            time.sleep(2)

            # Campo valor pago
            seletores_valor = [
                (By.CSS_SELECTOR, 'input[id^="valorPago"]'),
                (By.XPATH, "//input[contains(@id, 'valorPago')]"),
                (By.XPATH, "//input[@type='text' and contains(@name, 'valor')]"),
                (By.XPATH, "//label[contains(text(), 'Valor')]/following::input[1]"),
            ]

            campo_valor = None
            for by, selector in seletores_valor:
                try:
                    campo_valor = WebDriverWait(driver, 5).until(
                        EC.presence_of_element_located((by, selector))
                    )
                    break
                except:
                    continue

            if not campo_valor:
                log_message("❌ Campo de valor não encontrado", "ERROR")
                self._salvar_screenshot(driver, "erro_campo_valor")
                self._salvar_html(driver, "erro_campo_valor")
                return False

            campo_valor.clear()
            time.sleep(0.3)
            
            # Formatar valor
            try:
                valor_float = float(valor_pago)
                valor_formatado = f"{valor_float:.2f}".replace(".", ",")
            except (ValueError, TypeError):
                valor_formatado = str(valor_pago).replace(".", ",")
            
            campo_valor.send_keys(valor_formatado)
            log_message(f"💰 Valor preenchido: R$ {valor_formatado}", "INFO")
            time.sleep(0.5)

            # Campo justificativa
            seletores_just = [
                (By.NAME, "Justificativa"),
                (By.XPATH, "//textarea[@name='Justificativa']"),
                (By.XPATH, "//textarea[@placeholder='Justificativa']"),
                (By.XPATH, "//textarea[contains(@id, 'justificativa')]"),
                (By.CSS_SELECTOR, "textarea"),
            ]

            campo_just = None
            for by, selector in seletores_just:
                try:
                    campo_just = driver.find_element(by, selector)
                    break
                except:
                    continue

            if campo_just:
                campo_just.clear()
                campo_just.send_keys(str(justificativa) if pd.notna(justificativa) else "")
                log_message(f"📄 Justificativa preenchida: {str(justificativa)[:50]}...", "INFO")
            else:
                log_message("⚠️ Campo de justificativa não encontrado", "WARNING")

            time.sleep(0.5)

            recurso_normalizado = str(recurso).strip().lower()

            if recurso_normalizado == "sim":
                # Clicar no <a> com id salvarGerar
                try:
                    btn = WebDriverWait(driver, 5).until(
                        EC.element_to_be_clickable((By.ID, "salvarGerar"))
                    )
                    btn.click()
                    log_message("✓ Clicado em 'Salvar e gerar recurso'", "SUCCESS")
                except Exception as e:
                    log_message(f"❌ Erro ao clicar em 'Salvar e gerar recurso': {e}", "ERROR")
                    self._salvar_screenshot(driver, "erro_botao_salvarGerar")
                    return False
            else:
                # Clicar no <a> com id salvarGlosar
                try:
                    btn = WebDriverWait(driver, 5).until(
                        EC.element_to_be_clickable((By.ID, "salvarGlosar"))
                    )
                    btn.click()
                    log_message("✓ Clicado em 'Salvar e glosar os valores'", "SUCCESS")
                except Exception as e:
                    log_message(f"❌ Erro ao clicar em 'Salvar e glosar os valores': {e}", "ERROR")
                    self._salvar_screenshot(driver, "erro_botao_salvarGlosar")
                    return False

            time.sleep(2)
            return True

        except Exception as e:
            log_message(f"❌ Erro ao preencher modal: {e}", "ERROR")
            self._salvar_screenshot(driver, "erro_preencher_modal")
            return False

    def _processar_procedimento(self, driver, lote_str, row):
        """Processa um procedimento individual dentro de um lote"""
        paciente = str(row["Paciente"]).strip()
        exame = str(row["Número"]).strip()
        procedimento = str(row["Procedimento"]).strip()
        recurso = str(row["Recurso?"]).strip().lower() if pd.notna(row.get("Recurso?")) else "não"
        valor_pago = row["Valor Pago"]
        justificativa = row["Justificativa"] if pd.notna(row.get("Justificativa")) else ""

        log_message(f"👤 Paciente: {paciente}", "INFO")
        log_message(f"🔬 Exame: {exame} | Procedimento: {procedimento}", "INFO")

        try:
            tabela_interna = self._encontrar_tabela_procedimentos(driver, lote_str)

            if not tabela_interna:
                log_message("❌ Tabela de procedimentos não encontrada", "ERROR")
                self._salvar_screenshot(driver, f"erro_tabela_nao_encontrada_{exame}")
                return False, "Tabela não encontrada"

            linhas_int = tabela_interna.find_elements(By.TAG_NAME, "tr")
            log_message(f"📋 Encontradas {len(linhas_int)} linhas na tabela", "INFO")

            linha_encontrada = False

            for i, linha_int in enumerate(linhas_int):
                try:
                    if not linha_int.is_displayed():
                        continue

                    cols_int = linha_int.find_elements(By.TAG_NAME, "td")
                    if len(cols_int) < 3:
                        continue

                    textos_colunas = [col.text.strip() for col in cols_int]
                    texto_completo = " ".join(textos_colunas)

                    numero_match = exame in texto_completo
                    procedimento_match = procedimento in texto_completo
                    paciente_match = paciente.upper() in texto_completo.upper()

                    if numero_match and procedimento_match and paciente_match:
                        log_message(f"✓ Linha encontrada (índice {i})", "INFO")

                        seletores_botao_dollar = [
                            (By.XPATH, ".//a[@title='Registrar pagamento parcial do exame']"),
                            (By.XPATH, ".//a[contains(@data-url, 'registrarPagamentoParcialModalAjax')]"),
                            (By.XPATH, ".//a[@class='btn btn-default btn-xs chamadaAjax setupAjax' and contains(@data-url, 'registrarPagamentoParcial')]"),
                            (By.XPATH, ".//a[.//svg[contains(@data-icon, 'search-dollar')]]"),
                        ]

                        botao_dollar = None
                        for by, selector in seletores_botao_dollar:
                            try:
                                botao_dollar = linha_int.find_element(by, selector)
                                break
                            except:
                                continue

                        if not botao_dollar:
                            log_message(f"❌ Botão de pagamento parcial não encontrado na linha {i}", "WARNING")
                            self._salvar_screenshot(driver, f"erro_botao_dollar_{exame}")
                            continue

                        driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", botao_dollar)
                        time.sleep(0.5)
                        driver.execute_script("arguments[0].click();", botao_dollar)
                        log_message("✓ Botão de pagamento parcial clicado", "INFO")

                        linha_encontrada = True
                        break

                except Exception as e:
                    log_message(f"⚠️ Erro ao processar linha {i}: {e}", "WARNING")
                    continue

            if not linha_encontrada:
                log_message(f"❌ Linha não encontrada para: {paciente} | {exame} | {procedimento}", "ERROR")
                self._salvar_screenshot(driver, f"erro_linha_nao_encontrada_{exame}")
                return False, "Linha não encontrada"

            sucesso = self._preencher_modal(driver, valor_pago, justificativa, recurso)

            if sucesso:
                log_message("✅ Procedimento processado com sucesso", "SUCCESS")
                return True, "OK"
            else:
                log_message("❌ Falha ao processar procedimento", "ERROR")
                return False, "Falha no modal"

        except Exception as e:
            log_message(f"❌ Erro ao processar procedimento: {e}", "ERROR")
            self._salvar_screenshot(driver, f"erro_procedimento_{exame}")
            return False, str(e)

    # ========================== Processamento ===============================

    def _carregar_planilha(self, excel_file: str) -> pd.DataFrame | None:
        try:
            if not os.path.exists(excel_file):
                log_message(f"❌ Arquivo não encontrado: {excel_file}", "ERROR")
                return None
            
            df = pd.read_excel(excel_file)
            df["Status"] = ""
            df.columns = df.columns.str.strip()
            
            colunas_necessarias = [
                "Paciente", "Número", "Procedimento", "Lote", "Valor Pago",
                "Justificativa", "Recurso?"
            ]
            faltantes = [c for c in colunas_necessarias if c not in df.columns]
            if faltantes:
                log_message(f"❌ Colunas faltantes na planilha: {faltantes}", "ERROR")
                return None
            
            log_message(f"✓ Planilha carregada: {len(df)} registros", "SUCCESS")
            return df
            
        except Exception as e:
            log_message(f"❌ Erro ao carregar planilha: {type(e).__name__} - {e}", "ERROR")
            return None

    def _salvar_resultados(self, df: pd.DataFrame, excel_file: str) -> str | None:
        try:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            df_falhas = df[df["Status"] != "OK"].copy()
            
            if len(df_falhas) == 0:
                log_message("✅ Todos os procedimentos foram processados com sucesso!", "SUCCESS")
                return None
            
            output_path = (
                os.path.splitext(excel_file)[0]
                + f"_NAO_COMPLETADOS_{timestamp}.xlsx"
            )
            df_falhas.to_excel(output_path, index=False)
            log_message(f"📁 Planilha de não completados salva: {output_path}", "INFO")
            return output_path
            
        except Exception as e:
            log_message(f"❌ Erro ao salvar planilha de resultados: {e}", "ERROR")
            return None

    def _exibir_resumo(self, df: pd.DataFrame):
        total = len(df)
        sucesso = len(df[df["Status"] == "OK"])
        falhas = total - sucesso
        
        log_message(f"📊 Total de procedimentos: {total}", "INFO")
        log_message(f"✅ Completados com sucesso: {sucesso}", "SUCCESS")
        log_message(f"❌ Não completados: {falhas}", "WARNING" if falhas > 0 else "INFO")
        
        if falhas > 0:
            df_falhas = df[df["Status"] != "OK"].copy()
            for _, row in df_falhas.iterrows():
                log_message(
                    f"  - Exame {row.get('Número')} | Proc {row.get('Procedimento')} | Motivo: {row['Status']}",
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
        excel_file = params.get("excel_file")
        cancel_flag = params.get("cancel_flag")
        headless_mode = params.get("headless_mode", False)
        url = params.get(
            "url_login",
            os.getenv("SYSTEM_URL", "https://dap.pathoweb.com.br/login/auth?format="),
        )

        df = self._carregar_planilha(excel_file)
        if df is None:
            messagebox.showerror(
                "Erro",
                "Não foi possível carregar a planilha. Verifique o arquivo e as colunas necessárias:\n"
                "Paciente, Número, Procedimento, Lote, Valor Pago, Justificativa, Recurso?",
            )
            return

        driver = None
        try:
            driver = BrowserFactory.create_chrome(headless=headless_mode)
            log_message("🚀 Navegador inicializado para Baixa de Lote", "INFO")

            if not self._fazer_login(driver, username, password, url):
                messagebox.showerror("Erro", "Falha no login no Pathoweb.")
                return

            if not self._acessar_faturamento(driver):
                messagebox.showerror("Erro", "Falha ao acessar módulo de Faturamento.")
                return

            self._fechar_modal_inicial(driver)

            if not self._acessar_faturas_enviadas(driver):
                messagebox.showerror("Erro", "Falha ao acessar 'Faturas enviadas e recebimento'.")
                return

            # Processar lotes únicos
            lotes = df["Lote"].dropna().unique()
            log_message(f"📦 Total de lotes a processar: {len(lotes)}", "INFO")

            for idx_lote, lote in enumerate(lotes):
                if cancel_flag and cancel_flag.is_set():
                    log_message("⚠️ Execução cancelada pelo usuário.", "WARNING")
                    break

                lote_str = str(lote).strip()
                
                log_message(f"\n{'='*60}", "INFO")
                log_message(f"🔍 Processando lote {idx_lote+1}/{len(lotes)}: {lote_str}", "INFO")
                log_message(f"{'='*60}", "INFO")
                
                # Se não é o primeiro lote, retornar à tela de busca
                if idx_lote > 0:
                    if not self._retornar_tela_busca(driver):
                        log_message(f"⚠️ Não foi possível retornar à tela de busca, pulando lote {lote_str}", "WARNING")
                        continue
                
                if not self._buscar_lote(driver, lote_str):
                    log_message(f"⚠️ Pulando lote {lote_str} - não foi possível buscar", "WARNING")
                    # Marcar todos os procedimentos deste lote como falha
                    df.loc[df["Lote"] == lote, "Status"] = "Erro na busca do lote"
                    continue

                if not self._selecionar_checkbox_por_lote(driver, lote_str):
                    log_message(f"⚠️ Pulando lote {lote_str} - checkbox não selecionado", "WARNING")
                    df.loc[df["Lote"] == lote, "Status"] = "Checkbox não encontrado"
                    continue

                if not self._clicar_botao_receber(driver):
                    log_message(f"⚠️ Pulando lote {lote_str} - não foi possível clicar em Receber", "WARNING")
                    df.loc[df["Lote"] == lote, "Status"] = "Erro ao clicar em Receber"
                    continue

                # Processar procedimentos do lote
                linhas_lote = df[df["Lote"] == lote]
                for idx_proc, (df_index, row) in enumerate(linhas_lote.iterrows()):
                    if cancel_flag and cancel_flag.is_set():
                        log_message("⚠️ Execução cancelada pelo usuário.", "WARNING")
                        break

                    log_message(f"\n📌 Procedimento {idx_proc+1}/{len(linhas_lote)}", "INFO")
                    sucesso, mensagem = self._processar_procedimento(driver, lote_str, row)
                    df.at[df_index, "Status"] = mensagem
                    time.sleep(1)

                time.sleep(3)

            log_message("\n✅ Processamento de todos os lotes finalizado!", "SUCCESS")
            self._salvar_resultados(df, excel_file)
            self._exibir_resumo(df)

        except Exception as e:
            log_message(f"❌ Erro crítico durante a automação: {e}", "ERROR")
            messagebox.showerror(
                "Erro",
                f"Erro crítico durante a automação:\n{str(e)[:200]}",
            )
        finally:
            if driver:
                try:
                    driver.quit()
                except Exception:
                    pass


def run(params: dict):
    module = BaixaLoteModule()
    module.run(params)
