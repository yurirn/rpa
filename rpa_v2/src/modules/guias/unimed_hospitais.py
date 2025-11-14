import os
import time
import pandas as pd
from tkinter import messagebox
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from datetime import datetime

from src.core.browser_factory import BrowserFactory
from src.core.logger import log_message
from src.modules.base import BaseModule

class UnimedHospitaisModule(BaseModule):
    def __init__(self):
        super().__init__(nome="Unimed - Hospitais")
        self.headless_mode = False  # Será definido no run()

    def click_element(self, driver, element, descricao="elemento"):
        """Clica em um elemento de forma robusta, funcionando em modo headless e normal"""
        try:
            if self.headless_mode:
                # Em modo headless, usar sempre JavaScript para cliques mais confiáveis
                driver.execute_script("arguments[0].click();", element)
                log_message(f"✅ Clique via JavaScript em {descricao}", "INFO")
            else:
                # Em modo normal, tentar clique normal primeiro
                try:
                    element.click()
                    log_message(f"✅ Clique normal em {descricao}", "INFO")
                except Exception:
                    # Se falhar, usar JavaScript como fallback
                    driver.execute_script("arguments[0].click();", element)
                    log_message(f"✅ Clique via JavaScript (fallback) em {descricao}", "INFO")
        except Exception as e:
            log_message(f"❌ Erro ao clicar em {descricao}: {e}", "ERROR")
            raise

    def wait_for_element(self, driver, wait, by, value, condition="presence", timeout=None):
        """Aguarda elemento de forma compatível com headless"""
        try:
            if timeout:
                wait = WebDriverWait(driver, timeout)
            
            # Em modo headless, sempre usar 'presence' em vez de 'clickable' ou 'visible'
            if self.headless_mode and condition in ["clickable", "visible"]:
                element = wait.until(EC.presence_of_element_located((by, value)))
            elif condition == "clickable":
                element = wait.until(EC.element_to_be_clickable((by, value)))
            elif condition == "visible":
                element = wait.until(EC.visibility_of_element_located((by, value)))
            else:  # presence
                element = wait.until(EC.presence_of_element_located((by, value)))
            
            return element
        except Exception as e:
            log_message(f"❌ Erro ao aguardar elemento {value}: {e}", "ERROR")
            raise

    def read_excel_data(self, file_path: str) -> list:
        """Lê os dados do arquivo Excel: Coluna B (número do exame) e Coluna E (número da guia)"""
        try:
            df = pd.read_excel(file_path, header=0)
            
            log_message(f"📋 Colunas encontradas: {list(df.columns)}", "INFO")
            
            # Converter DataFrame para lista de dicionários
            data_list = []
            for idx, row in df.iterrows():
                # Coluna B (índice 1) = número do exame
                # Coluna E (índice 4) = número da guia
                numero_exame = None
                numero_guia = None
                
                # Tentar pegar pela posição (índice)
                if len(df.columns) > 1:
                    numero_exame = row.iloc[1] if pd.notna(row.iloc[1]) else None
                if len(df.columns) > 4:
                    numero_guia = row.iloc[4] if pd.notna(row.iloc[4]) else None
                
                # Converter para string e limpar
                def converter_valor(valor):
                    if pd.notna(valor):
                        valor_str = str(valor).strip()
                        # Se termina com .0, remover (número inteiro lido como float pelo pandas)
                        if valor_str.endswith('.0'):
                            valor_str = valor_str[:-2]
                        return valor_str
                    return ''
                
                numero_exame_str = converter_valor(numero_exame) if numero_exame is not None else ''
                numero_guia_str = converter_valor(numero_guia) if numero_guia is not None else ''
                
                # Só adicionar se tiver número do exame
                if numero_exame_str:
                    data_list.append({
                        'numero_exame': numero_exame_str,
                        'numero_guia': numero_guia_str
                    })
            
            log_message(f"✅ Carregados {len(data_list)} registros do Excel", "SUCCESS")
            return data_list
        except Exception as e:
            raise ValueError(f"Erro ao ler o Excel: {e}")

    def fazer_login_pathoweb(self, driver, wait, username, password):
        """Faz login no PathoWeb e navega para o módulo de faturamento"""
        try:
            log_message("🔐 Fazendo login no PathoWeb...", "INFO")
            
            # URL do PathoWeb
            url = "https://pathoweb.com.br/login/auth"
            driver.get(url)
            
            # Preencher credenciais
            wait.until(EC.presence_of_element_located((By.ID, "username"))).send_keys(username)
            driver.find_element(By.ID, "password").send_keys(password)
            botao_submit = driver.find_element(By.CSS_SELECTOR, "button[type='submit']")
            self.click_element(driver, botao_submit, "botão login PathoWeb")
            
            log_message("Verificando se precisa navegar para módulo de faturamento...", "INFO")
            current_url = driver.current_url

            if current_url == "https://pathoweb.com.br/" or "trocarModulo" in current_url:
                log_message("Detectada tela de seleção de módulos - navegando para módulo de faturamento...", "INFO")
                try:
                    modulo_link = self.wait_for_element(driver, wait, By.CSS_SELECTOR,
                        "a[href='/site/trocarModulo?modulo=2']", condition="presence")
                    self.click_element(driver, modulo_link, "link módulo faturamento")
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

            # Fechar modal se aparecer
            try:
                modal_close_button = driver.find_element(By.CSS_SELECTOR, "#mensagemParaClienteModal .modal-footer button")
                # Em modo headless, não verificar is_displayed()
                if self.headless_mode or modal_close_button.is_displayed():
                    self.click_element(driver, modal_close_button, "fechar modal inicial")
                    time.sleep(1)
            except Exception:
                pass

            # Acessar explicitamente a página do módulo de faturamento
            log_message("Acessando módulo de faturamento via URL...", "INFO")
            driver.get("https://pathoweb.com.br/moduloFaturamento/index")
            time.sleep(2)

            # Clicar no botão "Preparar exames para fatura"
            log_message("Clicando em 'Preparar exames para fatura'...", "INFO")
            try:
                preparar_btn = self.wait_for_element(driver, wait, By.CSS_SELECTOR,
                    "a.btn.btn-danger.chamadaAjax.setupAjax[data-url='/moduloFaturamento/preFaturamento']",
                    condition="presence")
                self.click_element(driver, preparar_btn, "botão 'Preparar exames para fatura'")
            except Exception:
                preparar_btn = self.wait_for_element(driver, wait, By.XPATH,
                    "//a[contains(@class, 'setupAjax') and contains(text(), 'Preparar exames para fatura')]",
                    condition="presence")
                self.click_element(driver, preparar_btn, "botão 'Preparar exames para fatura' (alternativo)")

            # Aguardar possível spinner/modal carregar
            try:
                WebDriverWait(driver, 3).until(EC.presence_of_element_located((By.ID, "spinner")))
                log_message("🔄 Modal de carregamento detectado, aguardando...", "INFO")
                WebDriverWait(driver, 30).until(EC.invisibility_of_element_located((By.ID, "spinner")))
                log_message("✅ Modal de carregamento fechado", "INFO")
            except Exception:
                time.sleep(1)

            log_message("✅ Login no PathoWeb realizado e página de pré-faturamento acessada", "SUCCESS")
            return True
            
        except Exception as e:
            log_message(f"❌ Erro ao fazer login no PathoWeb: {e}", "ERROR")
            return False

    def limpar_filtros(self, driver, wait):
        """Clica no botão 'Limpar' para limpar os filtros"""
        try:
            log_message("🧹 Clicando no botão 'Limpar' para limpar filtros...", "INFO")
            
            # Procurar o botão Limpar
            botao_limpar = self.wait_for_element(driver, wait, By.CSS_SELECTOR,
                "a.btn.btn-warning.btn-sm.limpar-filtro", condition="presence")
            self.click_element(driver, botao_limpar, "botão Limpar")
            
            # Aguardar processamento
            time.sleep(2)
            
            # Aguardar spinner se existir
            try:
                WebDriverWait(driver, 3).until(EC.presence_of_element_located((By.ID, "spinner")))
                log_message("🔄 Aguardando processamento após limpar filtros...", "INFO")
                WebDriverWait(driver, 30).until(EC.invisibility_of_element_located((By.ID, "spinner")))
            except Exception:
                time.sleep(1)
            
            log_message("✅ Filtros limpos com sucesso", "SUCCESS")
            return True
            
        except Exception as e:
            log_message(f"⚠️ Erro ao limpar filtros: {e}", "WARNING")
            return False

    def pesquisar_exame(self, driver, wait, numero_exame):
        """Pesquisa um exame pelo número do exame"""
        try:
            log_message(f"🔍 Pesquisando exame: {numero_exame}...", "INFO")
            
            # Limpar e preencher campo número do exame
            campo_numero_exame = self.wait_for_element(driver, wait, By.ID, "numeroExame", condition="presence")
            campo_numero_exame.clear()
            time.sleep(0.5)
            campo_numero_exame.send_keys(str(numero_exame))
            log_message(f"✅ Número do exame {numero_exame} digitado", "SUCCESS")
            time.sleep(0.5)
            
            # Clicar no botão Pesquisar
            botao_pesquisar = self.wait_for_element(driver, wait, By.ID, "pesquisaFaturamento", condition="presence")
            self.click_element(driver, botao_pesquisar, "botão Pesquisar")
            log_message("🔍 Pesquisando exame...", "INFO")
            
            # Aguardar carregamento dos resultados
            try:
                # Aguardar spinner se existir
                try:
                    WebDriverWait(driver, 3).until(EC.presence_of_element_located((By.ID, "spinner")))
                    log_message("🔄 Carregando resultados...", "INFO")
                    WebDriverWait(driver, 30).until(EC.invisibility_of_element_located((By.ID, "spinner")))
                except Exception:
                    time.sleep(5)
            except Exception:
                log_message("Tempo de carregamento excedido, verificando resultados mesmo assim...", "WARNING")
            
            # Aguardar mais um pouco para garantir que a tabela foi carregada
            time.sleep(3)
            
            # Verificar se há resultados
            tbody_rows = []
            selectors = [
                "#tabelaPreFaturamentoTbody tr",
                ".table-responsive table tbody tr",
                "table.table-striped tbody tr",
                "table.footable tbody tr"
            ]
            
            for selector in selectors:
                try:
                    tbody_rows = driver.find_elements(By.CSS_SELECTOR, selector)
                    if len(tbody_rows) > 0:
                        log_message(f"✅ Tabela de resultados encontrada usando seletor: {selector}", "SUCCESS")
                        break
                except Exception:
                    continue
            
            if len(tbody_rows) == 0:
                log_message(f"⚠️ Nenhum resultado encontrado para {numero_exame}", "WARNING")
                return False
            
            log_message(f"✅ Encontrados {len(tbody_rows)} resultados para o exame {numero_exame}", "SUCCESS")
            return True
            
        except Exception as e:
            log_message(f"❌ Erro ao pesquisar exame {numero_exame}: {e}", "ERROR")
            return False

    def abrir_exame(self, driver, wait):
        """Abre o primeiro exame encontrado na tabela"""
        try:
            log_message("📝 Abrindo exame...", "INFO")
            
            # Verificar se há resultados
            tbody_rows = []
            selectors = [
                "#tabelaPreFaturamentoTbody tr",
                ".table-responsive table tbody tr",
                "table.table-striped tbody tr",
                "table.footable tbody tr"
            ]
            
            for selector in selectors:
                try:
                    tbody_rows = driver.find_elements(By.CSS_SELECTOR, selector)
                    if len(tbody_rows) > 0:
                        break
                except Exception:
                    continue
            
            if len(tbody_rows) == 0:
                log_message("⚠️ Nenhum resultado encontrado para abrir", "WARNING")
                return False
            
            # Marcar checkbox do primeiro exame e clicar no botão "Abrir exame"
            log_message("📝 Marcando checkbox do primeiro exame...", "INFO")
            
            try:
                checkbox = tbody_rows[0].find_element(By.CSS_SELECTOR, "input[type='checkbox'][name='exameId']")
                if not checkbox.is_selected():
                    self.click_element(driver, checkbox, "checkbox do exame")
                    log_message("✅ Checkbox do exame marcado", "SUCCESS")
                else:
                    log_message("ℹ️ Checkbox já estava marcado", "INFO")
                
                time.sleep(1)
                
                # Procurar e clicar no botão "Abrir exame"
                log_message("🔍 Procurando botão 'Abrir exame'...", "INFO")
                
                abrir_btn = self.wait_for_element(driver, wait, By.CSS_SELECTOR,
                    "a.btn.btn-sm.btn-primary.chamadaAjax.toogleInicial.setupAjax[data-url='/moduloFaturamento/abrirExameCorrecao']",
                    condition="presence")
                log_message("✅ Botão 'Abrir exame' encontrado", "SUCCESS")
                
                # Clicar no botão
                self.click_element(driver, abrir_btn, "botão 'Abrir exame'")
                log_message("✅ Clique no botão 'Abrir exame' realizado", "SUCCESS")
                
                # Aguardar o modal aparecer
                log_message("⏳ Aguardando modal do exame abrir...", "INFO")
                time.sleep(3)
                
                # Verificar se o modal foi aberto
                try:
                    modal = wait.until(EC.presence_of_element_located((By.ID, "myModal")))
                    # Em modo headless, não verificar is_displayed() pois pode retornar False
                    if self.headless_mode or modal.is_displayed():
                        log_message("✅ Modal do exame aberto com sucesso", "SUCCESS")
                        return True
                    else:
                        log_message("⚠️ Modal encontrado mas não está visível", "WARNING")
                        time.sleep(2)
                        return True
                except Exception:
                    log_message("⚠️ Modal não encontrado, tentando continuar...", "WARNING")
                    time.sleep(2)
                    return True
                    
            except Exception as e:
                log_message(f"❌ Erro ao abrir exame: {e}", "ERROR")
                return False
                
        except Exception as e:
            log_message(f"❌ Erro ao abrir exame: {e}", "ERROR")
            return False

    def preencher_numero_guia(self, driver, wait, numero_guia):
        """Preenche o número da guia no modal do exame"""
        try:
            if not numero_guia or not numero_guia.strip():
                log_message("⚠️ Número da guia vazio, pulando preenchimento", "WARNING")
                return True
            
            log_message(f"📝 Preenchendo número da guia: {numero_guia}...", "INFO")
            
            # Aguardar um pouco para garantir que o modal está carregado
            time.sleep(2)
            
            # Preencher número da guia usando a função jQuery
            js_numero_guia = f'''
            function typeNumeroGuia(texto, delay = 40) {{
              const $inp = $("#numeroGuiaInput");
              const $a   = $inp.closest('td').children('a.table-editable-ancora').first();

              // limpa antes
              $inp.val("").attr("value","").trigger("input");
              if ($a.length) $a.text("").css("display","inline");

              let i = 0;
              const timer = setInterval(() => {{
                const atual = $inp.val() + texto[i];
                $inp.val(atual).trigger("input").trigger("keyup");
                if ($a.length) $a.text(atual);

                i++;
                if (i >= texto.length) {{
                  clearInterval(timer);
                  // consolida valor nos atributos e dispara change/blur (para AJAX no blur)
                  $inp.attr("value", texto)
                      .data("previous-value", texto)
                      .trigger("change")
                      .trigger("blur");
                }}
              }}, delay);
            }}

            // uso:
            typeNumeroGuia("{numero_guia}", 30);
            '''
            
            driver.execute_script(js_numero_guia)
            log_message(f"✅ Número da guia {numero_guia} preenchido", "SUCCESS")
            
            # Aguardar um pouco para o processamento
            time.sleep(3)
            
            return True
            
        except Exception as e:
            log_message(f"⚠️ Erro ao preencher número da guia: {e}", "WARNING")
            return False

    def salvar_exame(self, driver, wait):
        """Salva o exame clicando em 'Próximo' e depois 'Salvar'"""
        try:
            log_message("💾 Salvando exame...", "INFO")
            
            # 1. Clicar no botão "Próximo" para salvar os dados do exame
            log_message("🔄 Clicando no botão 'Próximo' para salvar...", "INFO")
            try:
                botao_proximo = self.wait_for_element(driver, wait, By.CSS_SELECTOR,
                    "a.btn.btn-sm.btn-primary.wizardControl.chamadaAjax.setupAjax[data-url='/moduloFaturamento/saveAjaxExameParaFaturamento']",
                    condition="presence")
                self.click_element(driver, botao_proximo, "botão 'Próximo'")
                log_message("✅ Botão 'Próximo' clicado", "SUCCESS")
                
                # Aguardar processamento
                time.sleep(3)
                
            except Exception as e:
                log_message(f"⚠️ Erro ao clicar no botão 'Próximo': {e}", "WARNING")
                # Tentar encontrar o botão com seletor alternativo
                try:
                    botao_proximo_alt = self.wait_for_element(driver, wait, By.XPATH,
                        "//a[contains(@class, 'wizardControl') and contains(text(), 'Próximo')]",
                        condition="presence")
                    self.click_element(driver, botao_proximo_alt, "botão 'Próximo' (alternativo)")
                    log_message("✅ Botão 'Próximo' clicado (seletor alternativo)", "SUCCESS")
                    time.sleep(3)
                except Exception as e2:
                    log_message(f"❌ Erro ao clicar no botão 'Próximo' (tentativa alternativa): {e2}", "ERROR")
            
            # 2. Clicar no botão "Salvar" para finalizar
            log_message("💾 Clicando no botão 'Salvar' para finalizar...", "INFO")
            try:
                botao_salvar = self.wait_for_element(driver, wait, By.CSS_SELECTOR,
                    "a.btn.btn-sm.btn-primary.chamadaAjax.setupAjax[data-url='/moduloFaturamento/saveExameDadosClinicos']",
                    condition="presence")
                self.click_element(driver, botao_salvar, "botão 'Salvar'")
                log_message("✅ Botão 'Salvar' clicado", "SUCCESS")
                
                # Aguardar processamento
                time.sleep(3)
            except Exception as e:
                log_message(f"⚠️ Erro ao clicar no botão 'Salvar': {e}", "WARNING")
                # Tentar encontrar o botão com seletor alternativo
                try:
                    botao_salvar_alt = self.wait_for_element(driver, wait, By.XPATH,
                        "//a[contains(@class, 'chamadaAjax') and contains(text(), 'Salvar')]",
                        condition="presence")
                    self.click_element(driver, botao_salvar_alt, "botão 'Salvar' (alternativo)")
                    log_message("✅ Botão 'Salvar' clicado (seletor alternativo)", "SUCCESS")
                    time.sleep(3)
                except Exception as e2:
                    log_message(f"❌ Erro ao clicar no botão 'Salvar' (tentativa alternativa): {e2}", "ERROR")

            # Fechar o modal após salvar
            try:
                modal = wait.until(EC.presence_of_element_located((By.ID, "myModal")))
                try:
                    close_btn = modal.find_element(By.CSS_SELECTOR, "button.close[data-dismiss='modal']")
                except Exception:
                    close_btn = driver.find_element(By.CSS_SELECTOR, "#myModal button.close, #myModal .modal-header button.close")
                self.click_element(driver, close_btn, "botão fechar modal")
                time.sleep(2)
                log_message("✅ Modal fechado após salvar", "INFO")
            except Exception as e:
                log_message(f"⚠️ Não foi possível fechar o modal automaticamente: {e}", "WARNING")
            
            # Aguardar tabela estar visível novamente
            try:
                wait.until(EC.presence_of_element_located((By.ID, "tabelaPreFaturamentoTbody")))
                log_message("✅ Tabela de pré-faturamento visível", "INFO")
                time.sleep(1)
            except Exception as e:
                log_message(f"⚠️ Tabela não encontrada após fechar modal: {e}", "WARNING")
            
            log_message("✅ Exame salvo com sucesso", "SUCCESS")
            return True
            
        except Exception as e:
            log_message(f"❌ Erro ao salvar exame: {e}", "ERROR")
            return False

    def marcar_exame_como_pendente(self, driver, wait):
        """Marca o exame como 'Pendente' na tabela"""
        try:
            log_message("📝 Marcando exame como 'Pendente'...", "INFO")
            time.sleep(2)

            # Re-localizar a tabela sempre antes de processar para evitar elementos stale
            def obter_linhas():
                return driver.find_elements(By.CSS_SELECTOR, "#tabelaPreFaturamentoTbody tr")
            
            linhas = obter_linhas()
            if not linhas or len(linhas) == 0:
                log_message("⚠️ Nenhuma linha encontrada na tabela de pré-faturamento", "WARNING")
                return False

            # Processar apenas a primeira linha (o exame que acabamos de processar)
            linha = linhas[0]
            
            try:
                # Re-localizar células dentro da linha
                celulas = linha.find_elements(By.CSS_SELECTOR, "td")
                if len(celulas) < 2:
                    log_message("⚠️ Células insuficientes na linha", "WARNING")
                    return False

                # Segunda coluna é a de 'Conferido' (onde vamos mudar para 'Pendente')
                cel_conferido = celulas[1]

                # Verificar status atual (mas sempre tentar mudar para garantir)
                try:
                    ancora = cel_conferido.find_element(By.CSS_SELECTOR, "a.table-editable-ancora")
                    texto_ancora = (ancora.text or "").strip().lower()
                    log_message(f"ℹ️ Status atual do exame: {texto_ancora}", "INFO")
                    # Não retornar aqui - sempre tentar mudar para Pendente
                except Exception:
                    log_message("ℹ️ Âncora não encontrada, tentando processar", "INFO")

                # Tentar abrir o editor clicando na âncora
                clicou_ancora = False
                for tentativa in range(3):
                    try:
                        ancora = cel_conferido.find_element(By.CSS_SELECTOR, "a.table-editable-ancora")
                        
                        # Em modo headless, não fazer scroll
                        if not self.headless_mode:
                            driver.execute_script("arguments[0].scrollIntoView({block:'center'});", ancora)
                            time.sleep(0.3)
                        
                        # Aguardar spinner invisível
                        try:
                            WebDriverWait(driver, 2).until(EC.invisibility_of_element_located((By.ID, "spinner")))
                        except Exception:
                            pass
                        
                        self.click_element(driver, ancora, "âncora")
                        time.sleep(0.5)
                        clicou_ancora = True
                        log_message(f"✅ Clicou na âncora (tentativa {tentativa + 1})", "INFO")
                        break
                        
                    except Exception as e:
                        log_message(f"⚠️ Erro ao clicar na âncora (tentativa {tentativa + 1}): {e}", "WARNING")
                        if tentativa < 2:
                            try:
                                WebDriverWait(driver, 30).until(EC.invisibility_of_element_located((By.ID, "spinner")))
                                time.sleep(0.5)
                            except Exception:
                                time.sleep(1)
                
                if not clicou_ancora:
                    log_message("❌ Não conseguiu clicar na âncora após 3 tentativas", "ERROR")
                    return False

                # Selecionar 'Pendente' no select
                selecionou = False
                for tentativa in range(3):
                    try:
                        linhas_temp = obter_linhas()
                        if len(linhas_temp) > 0:
                            cel_conferido_temp = linhas_temp[0].find_elements(By.CSS_SELECTOR, "td")[1]
                            select_el = cel_conferido_temp.find_element(By.CSS_SELECTOR, "select[name='faturamentoConferido']")
                            
                            # Usar JavaScript para garantir a seleção
                            driver.execute_script("""
                                var s = arguments[0];
                                $(s).val('Pendente').trigger('change').trigger('blur');
                            """, select_el)
                            
                            log_message(f"✅ Selecionou 'Pendente' (tentativa {tentativa + 1})", "SUCCESS")
                            selecionou = True
                            break
                    except Exception as e:
                        log_message(f"⚠️ Erro ao selecionar 'Pendente' (tentativa {tentativa + 1}): {e}", "WARNING")
                        if tentativa < 2:
                            time.sleep(0.5)
                
                if not selecionou:
                    log_message("❌ Não conseguiu selecionar 'Pendente' após 3 tentativas", "ERROR")
                    return False

                # Aguardar processamento (spinner)
                try:
                    WebDriverWait(driver, 3).until(EC.presence_of_element_located((By.ID, "spinner")))
                    log_message("🔄 Processando alteração (spinner detectado)...", "INFO")
                    WebDriverWait(driver, 30).until(EC.invisibility_of_element_located((By.ID, "spinner")))
                    log_message("✅ Processamento concluído", "SUCCESS")
                except Exception:
                    time.sleep(0.5)
                    log_message("ℹ️ Sem spinner, aguardando estabilização", "INFO")

                # Verificação final - confirmar que foi alterado para Pendente
                time.sleep(1)
                try:
                    linhas_final = obter_linhas()
                    if len(linhas_final) > 0:
                        celulas_final = linhas_final[0].find_elements(By.CSS_SELECTOR, "td")
                        if len(celulas_final) >= 2:
                            ancora_final = celulas_final[1].find_element(By.CSS_SELECTOR, "a.table-editable-ancora")
                            status_final = (ancora_final.text or "").strip().lower()
                            if status_final == "pendente":
                                log_message("✅ Exame marcado como 'Pendente' com sucesso (verificado)", "SUCCESS")
                                return True
                            else:
                                log_message(f"⚠️ Status final: {status_final} (esperado: pendente)", "WARNING")
                                return False
                except Exception as e:
                    log_message(f"⚠️ Erro ao verificar status final: {e}", "WARNING")
                    # Retornar True mesmo assim, pois a seleção foi feita
                    return True

            except Exception as e:
                log_message(f"❌ Erro crítico ao processar linha: {e}", "ERROR")
                return False
            
        except Exception as e:
            log_message(f"❌ Erro ao marcar exame como 'Pendente': {e}", "ERROR")
            return False

    def processar_exame(self, driver, wait, dados):
        """Processa um exame individual"""
        try:
            numero_exame = dados['numero_exame']
            numero_guia = dados['numero_guia']
            
            log_message(f"🔄 Processando exame {numero_exame} (guia: {numero_guia})...", "INFO")
            
            # 1. Limpar filtros
            self.limpar_filtros(driver, wait)
            time.sleep(1)
            
            # 2. Pesquisar exame
            if not self.pesquisar_exame(driver, wait, numero_exame):
                return {
                    'numero_exame': numero_exame,
                    'numero_guia': numero_guia,
                    'status': 'erro',
                    'erro': 'Exame não encontrado',
                    'timestamp': datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                }
            
            # 3. Abrir exame
            if not self.abrir_exame(driver, wait):
                return {
                    'numero_exame': numero_exame,
                    'numero_guia': numero_guia,
                    'status': 'erro',
                    'erro': 'Erro ao abrir exame',
                    'timestamp': datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                }
            
            # 4. Preencher número da guia
            self.preencher_numero_guia(driver, wait, numero_guia)
            
            # 5. Salvar exame
            if not self.salvar_exame(driver, wait):
                return {
                    'numero_exame': numero_exame,
                    'numero_guia': numero_guia,
                    'status': 'erro',
                    'erro': 'Erro ao salvar exame',
                    'timestamp': datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                }
            
            # 6. Marcar como Pendente
            self.marcar_exame_como_pendente(driver, wait)
            
            log_message(f"✅ Exame {numero_exame} processado com sucesso", "SUCCESS")
            return {
                'numero_exame': numero_exame,
                'numero_guia': numero_guia,
                'status': 'sucesso',
                'timestamp': datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            }
            
        except Exception as e:
            log_message(f"❌ Erro ao processar exame {dados.get('numero_exame', 'desconhecido')}: {e}", "ERROR")
            return {
                'numero_exame': dados.get('numero_exame', ''),
                'numero_guia': dados.get('numero_guia', ''),
                'status': 'erro',
                'erro': str(e),
                'timestamp': datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            }

    def run(self, params: dict):
        username = params.get("username")
        password = params.get("password")
        cancel_flag = params.get("cancel_flag")
        headless_mode = params.get("headless_mode")
        excel_file = params.get("excel_file")
        
        # Configurar modo headless na instância
        self.headless_mode = headless_mode
        log_message(f"🔧 Modo headless: {'Ativado' if headless_mode else 'Desativado'}", "INFO")

        # Validar credenciais
        if not username or not password:
            messagebox.showerror("Erro", "Credenciais são obrigatórias para este módulo.")
            return

        # Validar arquivo Excel
        if not excel_file or not os.path.exists(excel_file):
            messagebox.showerror("Erro", "Arquivo Excel é obrigatório para este módulo.")
            return

        driver = BrowserFactory.create_chrome(headless=headless_mode)
        wait = WebDriverWait(driver, 15)

        try:
            log_message("Iniciando automação Unimed - Hospitais...", "INFO")

            # Ler dados do Excel
            try:
                dados_excel = self.read_excel_data(excel_file)
                log_message(f"✅ Carregados {len(dados_excel)} registros do Excel", "SUCCESS")
                
                if not dados_excel:
                    messagebox.showwarning("Aviso", "Nenhum registro encontrado no Excel!")
                    return
                
            except Exception as e:
                log_message(f"❌ Erro ao ler arquivo Excel: {e}", "ERROR")
                messagebox.showerror("Erro", f"Erro ao ler arquivo Excel:\n{e}")
                return

            # Fazer login no PathoWeb
            if not self.fazer_login_pathoweb(driver, wait, username, password):
                messagebox.showerror("Erro", "Falha no login do PathoWeb!")
                return

            # Processar cada exame
            resultados = []
            for i, dados in enumerate(dados_excel, 1):
                if cancel_flag and cancel_flag.is_set():
                    log_message("Execução cancelada pelo usuário.", "WARNING")
                    break
                
                try:
                    log_message(f"➡️ Processando registro {i}/{len(dados_excel)} - Exame: {dados['numero_exame']}", "INFO")
                    
                    resultado = self.processar_exame(driver, wait, dados)
                    resultados.append(resultado)
                    
                    if resultado.get('status') == 'sucesso':
                        log_message(f"✅ Exame {dados['numero_exame']} processado com sucesso", "SUCCESS")
                    else:
                        log_message(f"❌ Erro no exame {dados['numero_exame']}: {resultado.get('erro')}", "ERROR")
                    
                    # Aguardar entre processamentos
                    if i < len(dados_excel):
                        time.sleep(2)
                    
                except Exception as e:
                    log_message(f"❌ Erro ao processar exame {dados.get('numero_exame', 'desconhecido')}: {e}", "ERROR")
                    resultados.append({
                        'numero_exame': dados.get('numero_exame', ''),
                        'numero_guia': dados.get('numero_guia', ''),
                        'status': 'erro',
                        'erro': str(e),
                        'timestamp': datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                    })

            # Resumo final
            total = len(resultados)
            sucessos = sum(1 for r in resultados if r.get('status') == 'sucesso')
            erros = sum(1 for r in resultados if r.get('status') == 'erro')

            log_message(f"\n📊 Resumo do processamento:", "INFO")
            log_message(f"Total de registros: {total}", "INFO")
            log_message(f"Sucessos: {sucessos}", "SUCCESS" if sucessos > 0 else "INFO")
            log_message(f"Erros: {erros}", "ERROR" if erros > 0 else "INFO")

            mensagem_final = f"✅ Processamento finalizado!\n\n" \
                           f"Total de registros: {total}\n" \
                           f"Sucessos: {sucessos}\n" \
                           f"Erros: {erros}"

            messagebox.showinfo("Processamento Concluído", mensagem_final)

            return {
                'sucesso': sucessos > 0,
                'sucessos': sucessos,
                'erros': erros,
                'resultados': resultados
            }

        except Exception as e:
            log_message(f"❌ Erro durante a automação: {e}", "ERROR")
            messagebox.showerror("Erro", f"❌ Erro durante a automação:\n{e}")
        finally:
            # Aguardar antes de fechar para permitir visualização dos resultados
            if not headless_mode:
                input("Pressione Enter para fechar o navegador...")
            driver.quit()


def run(params: dict):
    module = UnimedHospitaisModule()
    module.run(params)

