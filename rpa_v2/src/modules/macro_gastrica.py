import os
import time
from tkinter import messagebox
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
from dotenv import load_dotenv
from openpyxl import load_workbook

from src.core.browser_factory import BrowserFactory
from src.core.logger import log_message
from src.modules.base import BaseModule

load_dotenv()

class MacroGastricaModule(BaseModule):
    def __init__(self):
        super().__init__(nome="Macro Gástrica")

    def get_dados_exames(self, file_path: str) -> list:
        try:
            workbook = load_workbook(file_path)
            sheet = workbook.active
            dados = []
            ultima_mascara = None
            data_fixacao = None

            # Lê da linha 2 em diante (linha 1 é cabeçalho)
            for row in range(2, sheet.max_row + 1):
                codigo = sheet[f'A{row}'].value
                mascara = sheet[f'B{row}'].value
                campo_d = sheet[f'D{row}'].value  # Campo 182 (número de fragmentos)
                campo_e = sheet[f'E{row}'].value  # Campo 200 (medida 1)
                campo_f = sheet[f'F{row}'].value  # Campo 209 (medida 2)
                campo_g = sheet[f'G{row}'].value  # Campo 218 (medida 3)
                data_col = sheet[f'L{row}'].value  # Coluna L: data de fixação

                if row == 2 and data_col:
                    data_fixacao = str(data_col).strip()

                if codigo is not None:
                    codigo = str(codigo).strip()
                    
                    # Se não tem máscara, usa a última válida
                    if mascara is not None and str(mascara).strip():
                        mascara = str(mascara).strip()
                        ultima_mascara = mascara
                    else:
                        mascara = ultima_mascara

                    # Regra: se campo_d for 'mult', usar 6
                    if campo_d is not None and str(campo_d).strip().lower() == 'mult':
                        campo_d_valor = '6'
                    else:
                        campo_d_valor = str(campo_d).strip() if campo_d is not None else ""

                    dados.append({
                        'codigo': codigo,
                        'mascara': mascara,
                        'campo_d': campo_d_valor,
                        'campo_e': str(campo_e).strip() if campo_e is not None else "",
                        'campo_f': str(campo_f).strip() if campo_f is not None else "",
                        'campo_g': str(campo_g).strip() if campo_g is not None else "",
                        'data_fixacao': data_fixacao
                    })
            
            workbook.close()
            return dados
        except Exception as e:
            raise Exception(f"Erro ao ler planilha: {e}")

    def verificar_sessao_browser(self, driver) -> bool:
        """Verifica se a sessão do browser ainda está ativa"""
        try:
            driver.current_url
            return True
        except Exception as e:
            if "invalid session id" in str(e).lower():
                log_message("❌ Sessão do browser perdida", "ERROR")
                return False
            return True

    def selecionar_responsavel_macroscopia(self, driver, wait):
        """Seleciona 'Nathalia Fernanda da Silva Lopes' como responsável pela macroscopia"""
        # Aguardar o componente Select2 estar presente e clicar
        select2_container = wait.until(
            EC.element_to_be_clickable((By.XPATH, "//span[@aria-labelledby='select2-responsavelMacroscopiaId-container']"))
        )
        select2_container.click()
        time.sleep(0.2)
        
        # Aguardar e clicar na opção "Nathalia Fernanda da Silva Lopes"
        opcao_nathalia = wait.until(
            EC.element_to_be_clickable((By.XPATH, "//li[contains(text(), 'Nathalia Fernanda da Silva Lopes')]"))
        )
        opcao_nathalia.click()
        log_message("✅ Nathalia Fernanda da Silva Lopes selecionada como responsável", "SUCCESS")
        time.sleep(0.2)

    def selecionar_auxiliar_macroscopia(self, driver, wait):
        """Seleciona 'Renata Silva Sevidanis' como auxiliar da macroscopia"""
        # Aguardar o componente Select2 estar presente e clicar
        select2_container = wait.until(
            EC.element_to_be_clickable((By.XPATH, "//span[@aria-labelledby='select2-auxiliarMacroscopiaId-container']"))
        )
        select2_container.click()
        time.sleep(0.2)
        
        # Aguardar e clicar na opção "Renata Silva Sevidanis"
        opcao_renata = wait.until(
            EC.element_to_be_clickable((By.XPATH, "//li[contains(text(), 'Renata Silva Sevidanis')]"))
        )
        opcao_renata.click()
        log_message("✅ Renata Silva Sevidanis selecionada como auxiliar", "SUCCESS")
        time.sleep(0.2)

    def definir_data_fixacao(self, driver, wait, data_fixacao=None):
        """Define a data de fixação no campo de data de fixação"""
        try:
            if not data_fixacao:
                data_fixacao = '21082025'  # fallback para data padrão se não vier da planilha
            # Converter 21082025 para 2025-08-21
            if len(data_fixacao) == 8 and data_fixacao.isdigit():
                data_formatada = f"{data_fixacao[4:8]}-{data_fixacao[2:4]}-{data_fixacao[0:2]}"
            else:
                data_formatada = '2025-08-21'
            campo_data = wait.until(
                EC.presence_of_element_located((By.XPATH, "//input[@type='date' and @name='dataFixacao']"))
            )
            driver.execute_script("""
                var campo = arguments[0];
                campo.value = arguments[1];
                campo.dispatchEvent(new Event('change', { bubbles: true }));
            """, campo_data, data_formatada)
            log_message(f"📅 Data de fixação definida para: {data_formatada}", "SUCCESS")
            time.sleep(0.1)
        except Exception as e:
            log_message(f"⚠️ Erro ao definir data de fixação: {e}", "WARNING")

    def definir_hora_fixacao(self, driver, wait):
        """Define 18:00 no campo de hora de fixação"""
        # Aguardar o campo de hora estar presente
        campo_hora = wait.until(
            EC.presence_of_element_located((By.XPATH, "//input[@type='time' and @name='dataFixacao']"))
        )
        
        # Limpar e definir a hora
        campo_hora.clear()
        campo_hora.send_keys("18:00")
        log_message("🕕 Hora de fixação definida para: 18:00", "SUCCESS")
        time.sleep(0.1)

    def fechar_exame(self, driver, wait):
        """Clica no botão de fechar exame"""
        try:
            botao_fechar = wait.until(
                EC.element_to_be_clickable((By.ID, "fecharExameBarraFerramenta"))
            )
            botao_fechar.click()
            log_message("📁 Exame fechado", "INFO")
            
            # Aguardar retornar à tela principal
            try:
                # Verificar se voltou à tela principal aguardando o campo de código aparecer
                wait.until(EC.presence_of_element_located((By.ID, "inputSearchCodBarra")))
                log_message("✅ Retornou à tela principal após fechar exame", "INFO")
            except:
                log_message("⚠️ Pode não ter retornado à tela principal", "WARNING")
                # Tentar navegar de volta ao módulo se necessário
                try:
                    current_url = driver.current_url
                    if "modulo=1" not in current_url:
                        modulo_link = driver.find_element(By.CSS_SELECTOR, "a[href='/site/trocarModulo?modulo=1']")
                        modulo_link.click()
                        time.sleep(1.5)
                        log_message("🔄 Navegou de volta ao módulo de exames", "INFO")
                except:
                    pass
                    
        except Exception as e:
            log_message(f"Erro ao fechar exame: {e}", "ERROR")

    def digitar_mascara_e_buscar(self, driver, wait, mascara):
        """Digita a máscara no campo buscaArvore e pressiona Enter"""
        # Aguardar o campo estar presente e clicável
        campo_busca = wait.until(EC.element_to_be_clickable((By.ID, "buscaArvore")))
        
        # Digitar a máscara e pressionar Enter
        campo_busca.send_keys(mascara)
        campo_busca.send_keys(Keys.ENTER)
        log_message(f"✍️ Máscara '{mascara}' digitada no campo buscaArvore", "SUCCESS")
        time.sleep(0.5)

    def abrir_modal_variaveis_e_preencher(self, driver, wait, campo_d, campo_e, campo_f, campo_g):
        """Abre o modal de variáveis e preenche os campos"""
        try:
            # Clicar no botão "Pesquisar variáveis (F7)"
            botao_variaveis = wait.until(
                EC.element_to_be_clickable((By.ID, "cke_70"))
            )
            botao_variaveis.click()
            log_message("🔍 Clicou no botão de variáveis", "INFO")
            
            # Aguardar um pouco para o sistema processar
            time.sleep(0.8)
            
            # Verificar se apareceu um alerta
            try:
                alert = driver.switch_to.alert
                alert_text = alert.text
                if "não há variáveis" in alert_text.lower():
                    log_message(f"⚠️ Alerta detectado: {alert_text}", "WARNING")
                    alert.accept()  # Aceitar o alerta
                    log_message("⚠️ Pulando preenchimento de variáveis - não há variáveis no texto", "WARNING")
                    return
                else:
                    alert.accept()  # Aceitar qualquer outro alerta
            except:
                # Não há alerta, continuar normalmente
                pass
            
            # Aguardar o modal aparecer
            wait.until(EC.presence_of_element_located((By.CLASS_NAME, "swal2-popup")))
            log_message("🔍 Modal de variáveis aberto", "SUCCESS")
            time.sleep(0.3)
        
            # Preencher os campos usando classe genérica (IDs podem mudar)
            campos_input = driver.find_elements(By.CSS_SELECTOR, "input[style*='width: 100px'][style*='color: red']")
            log_message(f"🔍 Encontrados {len(campos_input)} campos de input no modal", "INFO")
            
            # Mapear valores para os campos na ordem que aparecem
            valores = [campo_d, campo_e, campo_f, campo_g, campo_d]  # Último é campo 334 (mesmo valor de D)
            
            for i, campo in enumerate(campos_input[:5]):  # Limitar aos 5 primeiros campos
                if i < len(valores) and valores[i]:
                    try:
                        campo.clear()
                        campo.send_keys(valores[i])
                        log_message(f"✍️ Campo {i+1} preenchido com: {valores[i]}", "SUCCESS")
                    except Exception as e:
                        log_message(f"⚠️ Erro ao preencher campo {i+1}: {e}", "WARNING")
            
            time.sleep(0.2)
            
            # Clicar no botão "Inserir"
            botao_inserir = wait.until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, ".swal2-confirm"))
            )
            botao_inserir.click()
            log_message("✅ Campos inseridos no modal", "SUCCESS")
            
            # Aguardar o modal fechar completamente
            try:
                wait.until(EC.invisibility_of_element_located((By.CLASS_NAME, "swal2-popup")))
                log_message("✅ Modal fechado completamente", "SUCCESS")
            except:
                # Se não conseguir detectar fechamento, aguardar um tempo fixo
                time.sleep(1)
                log_message("⏳ Aguardou fechamento do modal", "INFO")
            
        except Exception as e:
            log_message(f"⚠️ Erro ao preencher modal de variáveis: {e}", "WARNING")
            log_message("⚠️ Continuando sem preencher as variáveis", "WARNING")

    def salvar_macroscopia(self, driver, wait):
        """Clica no botão Salvar da macroscopia"""
        # Verificar se ainda há modal aberto e fechar se necessário
        try:
            modal = driver.find_element(By.CLASS_NAME, "swal2-popup")
            if modal.is_displayed():
                log_message("⚠️ Modal ainda aberto, tentando fechar...", "WARNING")
                # Tentar fechar o modal
                try:
                    botao_cancelar = driver.find_element(By.CSS_SELECTOR, ".swal2-cancel")
                    botao_cancelar.click()
                    time.sleep(0.5)
                except:
                    # Se não conseguir fechar, pressionar ESC
                    driver.find_element(By.TAG_NAME, "body").send_keys(Keys.ESCAPE)
                    time.sleep(0.5)
        except:
            # Não há modal, continuar normalmente
            pass
        
        # Aguardar o botão estar presente e clicável
        botao_salvar = wait.until(
            EC.element_to_be_clickable((By.ID, "salvarMacro"))
        )
        botao_salvar.click()
        log_message("💾 Macroscopia salva", "SUCCESS")
        time.sleep(0.3)

    def definir_grupo_baseado_mascara(self, driver, wait, mascara):
        """Define o grupo baseado na máscara (Estômago ou Intestino) - versão simplificada e confiável."""
        if not mascara:
            log_message("⚠️ Nenhuma máscara fornecida para definir grupo", "WARNING")
            return

        mascaras_estomago = ['A/C', 'A/I', 'AIC', 'AIF', 'ANTRO', 'COTO', 'DUO ', 'ESOFF', 'GASTRICA', 'POLIPO', 'G/POLIPO', 'ULCERA']
        mascaras_intestino = ['B/COLON', 'ICR', 'P/COLON']

        grupo_selecionado = None
        if mascara.upper() in mascaras_estomago:
            grupo_selecionado = "Estomago"
        elif mascara.upper() in mascaras_intestino:
            grupo_selecionado = "Intestino"
        else:
            log_message(f"⚠️ Máscara '{mascara}' não encontrada nas regras definidas", "WARNING")
            return

        try:
            # Tentar clicar na âncora de grupo (apenas a primeira encontrada com 'Vazio')
            try:
                campo_grupo = wait.until(
                    EC.element_to_be_clickable((By.XPATH, "//td[contains(text(), 'Grupo')]/following-sibling::td//a[contains(@class, 'autocomplete') and contains(text(), 'Vazio')]"))
                )
            except:
                # Fallback: procurar qualquer âncora de autocomplete vazia
                campo_grupo = wait.until(
                    EC.element_to_be_clickable((By.XPATH, "//a[contains(@class, 'autocomplete') and contains(text(), 'Vazio')]"))
                )
            campo_grupo.click()
            log_message(f"🔍 Clicou no campo de grupo", "INFO")
            time.sleep(0.3)

            # Aguardar o campo de input aparecer
            input_grupo = wait.until(
                EC.presence_of_element_located((By.ID, "idRegiao"))
            )
            input_grupo.clear()
            input_grupo.send_keys(grupo_selecionado)
            time.sleep(0.5)
            input_grupo.send_keys(Keys.TAB)
            log_message(f"✍️ Digitou '{grupo_selecionado}' no campo grupo", "SUCCESS")
            time.sleep(0.3)
        except Exception as e:
            log_message(f"⚠️ Erro ao definir grupo: {e}", "WARNING")

    def definir_representacao_secao(self, driver, wait):
        """Define a representação como 'Seção'"""
        try:
            # Procurar especificamente o campo de representação na linha correta
            try:
                campo_representacao = wait.until(
                    EC.element_to_be_clickable((By.XPATH, "//td[contains(text(), 'Representação')]/following-sibling::td//a[contains(@class, 'table-editable-ancora')]"))
                )
            except:
                # Fallback: procurar qualquer âncora de representação
                campo_representacao = wait.until(
                    EC.element_to_be_clickable((By.XPATH, "//a[contains(@class, 'table-editable-ancora')]"))
                )
            
            # Verificar se já está definida como "Seção"
            if "Seção" in campo_representacao.text:
                log_message("✅ Representação já está definida como 'Seção'", "SUCCESS")
                return
            
            # Se não estiver, clicar para alterar
            campo_representacao.click()
            log_message("🔍 Clicou no campo de representação", "INFO")
            time.sleep(0.3)
            
            # Aguardar o select aparecer
            select_representacao = wait.until(
                EC.presence_of_element_located((By.ID, "representacao"))
            )
            
            # Aguardar o select ficar visível
            wait.until(EC.element_to_be_clickable(select_representacao))
            
            # Selecionar "Seção" (valor "S")
            select = Select(select_representacao)
            select.select_by_value("S")
            log_message("✅ Representação definida como 'Seção'", "SUCCESS")
            time.sleep(0.3)
            
        except Exception as e:
            log_message(f"⚠️ Erro ao definir representação: {e}", "WARNING")

    def definir_regiao_gastrica(self, driver, wait, mascara=None):
        """Define a região de acordo com a máscara, conforme regras fornecidas"""
        try:
            if not mascara:
                log_message("⚠️ Nenhuma máscara fornecida para definir região", "WARNING")
                return

            # Regras de máscara para região
            mascara_regiao = {
                'A/C': 'AC: Antro/Corpo',
                'A/I': 'AI: Antro/Incisura',
                'AIC': 'AIC: Antro/Incisura/Corpo',
                'AIF': 'AIF: Antro/Incisura/Fundo',
                'ANTRO': 'AN: Antro',
                'COTO': 'COTO: Coto',
                'DUO': 'DUO: Duodeno',
                'ESOFF': 'ESÔF: Esôfago',
                'GASTRICA': 'GA: Gastrica',
                'G/POLIPO': 'POL/GASTRICA: Pólipo e Biópsia Gástrica',
                'POLIPO': 'POLG: Pólipo Gástrico',
                'ICR': 'ICR: Íleo/Cólon/Reto',
            }
            mascaras_sem_regiao = ['B/COLON', 'P/COLON', 'ULCERA']

            mascara_upper = mascara.upper().replace('Ó', 'O').replace('Ô', 'O')
            mascara_map = {k.upper().replace('Ó', 'O').replace('Ô', 'O'): v for k, v in mascara_regiao.items()}
            mascaras_sem_regiao_norm = [m.upper().replace('Ó', 'O').replace('Ô', 'O') for m in mascaras_sem_regiao]

            if mascara_upper in mascaras_sem_regiao_norm:
                log_message(f"⚠️ Máscara '{mascara}' não exige preenchimento de região (manual)", "WARNING")
                # Forçar foco no campo de quantidade de fragmentos para evitar erro de interatividade
                try:
                    input_quantidade = wait.until(
                        EC.presence_of_element_located((By.XPATH, "//input[contains(@name, 'quantidade_')]"))
                    )
                    driver.execute_script("arguments[0].focus();", input_quantidade)
                    log_message("🔍 Foco forçado no campo de quantidade de fragmentos (manual)", "INFO")
                except Exception as e:
                    log_message(f"⚠️ Não foi possível forçar foco no campo de quantidade: {e}", "WARNING")
                return

            regiao_valor = mascara_map.get(mascara_upper)
            if not regiao_valor:
                log_message(f"⚠️ Máscara '{mascara}' não encontrada nas regras de região", "WARNING")
                return

            # Procurar a âncora correspondente à região
            try:
                campo_regiao = wait.until(
                    EC.element_to_be_clickable((By.XPATH,
                                                "//input[contains(@name, 'regiao_')]/following-sibling::a[contains(@class, 'table-editable-ancora')]"))
                )
            except:
                campos_vazios = driver.find_elements(By.XPATH,
                                                     "//a[contains(@class, 'table-editable-ancora') and contains(text(), 'Vazio')]")
                if len(campos_vazios) >= 3:
                    campo_regiao = campos_vazios[2]
                else:
                    campo_regiao = campos_vazios[-1]

            # Clicar na âncora apenas se for visível e habilitada
            if campo_regiao.is_displayed() and campo_regiao.is_enabled():
                campo_regiao.click()
                log_message("🔍 Clicou no campo de região", "INFO")
                time.sleep(0.3)
            else:
                log_message("⚠️ Campo de região não está interativo, pulando clique", "WARNING")
                return

            # Esperar o input ficar clicável e visível
            input_regiao = wait.until(EC.element_to_be_clickable((By.XPATH, "//input[contains(@name, 'regiao_')]")))

            # Usar JavaScript para garantir que o campo está visível e interativo
            driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", input_regiao)
            driver.execute_script("arguments[0].focus();", input_regiao)

            # Definir o valor conforme a regra
            input_regiao.clear()
            input_regiao.send_keys(regiao_valor)
            log_message(f"✍️ Definiu região como '{regiao_valor}'", "SUCCESS")
            time.sleep(0.3)

            # Pressionar Tab para confirmar o valor
            input_regiao.send_keys(Keys.TAB)

        except Exception as e:
            log_message(f"⚠️ Erro ao definir região: {e}", "WARNING")

    def definir_quantidade_fragmentos(self, driver, wait, campo_d):
        """Define a quantidade de fragmentos baseado no campo D da planilha, sempre via JavaScript, sem scrollIntoView."""
        try:
            if not campo_d or campo_d.strip() == "":
                log_message("⚠️ Campo D está vazio, não definindo quantidade", "WARNING")
                return

            input_quantidade = wait.until(
                EC.presence_of_element_located((By.XPATH, "//input[contains(@name, 'quantidade_')]"))
            )

            # Preencher o campo diretamente via JavaScript, sem scroll nem focus
            driver.execute_script("arguments[0].value = arguments[1]; arguments[0].dispatchEvent(new Event('input', { bubbles: true }));", input_quantidade, campo_d.strip())
            log_message(f"✍️ Definiu quantidade como '{campo_d.strip()}' via JS", "SUCCESS")

            # Pressionar Tab para confirmar
            input_quantidade.send_keys(Keys.TAB)

        except Exception as e:
            log_message(f"⚠️ Erro ao definir quantidade: {e}", "WARNING")

    def definir_quantidade_blocos(self, driver, wait):
        """Define a quantidade de blocos como '1', sempre via JavaScript, sem scrollIntoView."""
        try:
            input_blocos = wait.until(
                EC.presence_of_element_located((By.XPATH, "//input[contains(@name, 'quantidadeBlocos_')]")))

            # Preencher o campo diretamente via JavaScript, sem scroll nem focus
            driver.execute_script("arguments[0].value = '1'; arguments[0].dispatchEvent(new Event('input', { bubbles: true }));", input_blocos)
            log_message("✍️ Definiu quantidade de blocos como '1' via JS", "SUCCESS")

            # Pressionar Tab para confirmar
            input_blocos.send_keys(Keys.TAB)

        except Exception as e:
            log_message(f"⚠️ Erro ao definir quantidade de blocos: {e}", "WARNING")

    def salvar_fragmentos(self, driver, wait):
        """Clica no botão Salvar dos fragmentos"""
        try:
            # Aguardar o botão estar presente e clicável
            botao_salvar_fragmentos = wait.until(
                EC.element_to_be_clickable((By.XPATH, "//a[contains(@class, 'btn-primary') and contains(@data-url, '/macroscopia/saveMacroscopiaFragAjax')]"))
            )
            
            # Verificar se o botão está visível
            if not botao_salvar_fragmentos.is_displayed():
                log_message("⚠️ Botão salvar fragmentos não está visível", "WARNING")
                return
            
            # Rolar até o botão para garantir visibilidade
            driver.execute_script("arguments[0].scrollIntoView({behavior: 'smooth', block: 'center'});", botao_salvar_fragmentos)
            time.sleep(0.5)
            
            # Clicar no botão
            botao_salvar_fragmentos.click()
            log_message("💾 Clicou em Salvar fragmentos", "SUCCESS")
            time.sleep(1.5)  # Aguardar o processamento
            
        except Exception as e:
            log_message(f"⚠️ Erro ao salvar fragmentos: {e}", "WARNING")
            # Tentar encontrar o botão por outras formas
            try:
                # Tentar por título
                botao_titulo = wait.until(
                    EC.element_to_be_clickable((By.XPATH, "//a[@title='Salvar' and contains(@class, 'btn-primary')]"))
                )
                driver.execute_script("arguments[0].scrollIntoView({behavior: 'smooth', block: 'center'});", botao_titulo)
                time.sleep(0.5)
                botao_titulo.click()
                log_message("💾 Clicou em Salvar fragmentos (por título)", "SUCCESS")
                time.sleep(1.5)
                return
            except:
                pass
            
            try:
                # Tentar por texto do botão
                botao_texto = wait.until(
                    EC.element_to_be_clickable((By.XPATH, "//a[contains(@class, 'btn-primary') and contains(text(), 'Salvar')]"))
                )
                driver.execute_script("arguments[0].scrollIntoView({behavior: 'smooth', block: 'center'});", botao_texto)
                time.sleep(0.5)
                botao_texto.click()
                log_message("💾 Clicou em Salvar fragmentos (por texto)", "SUCCESS")
                time.sleep(1.5)
                return
            except:
                pass
            
            log_message(f"❌ Não foi possível encontrar o botão Salvar fragmentos: {e}", "ERROR")
            raise

    def preencher_campos_pre_envio(self, driver, wait, mascara, campo_d):
        """Preenche todos os campos necessários antes de enviar para próxima etapa"""
        try:
            log_message("📝 Iniciando preenchimento dos campos pré-envio...", "INFO")
            
            # 1. Definir grupo baseado na máscara
            self.definir_grupo_baseado_mascara(driver, wait, mascara)
            
            # 2. Definir representação como "Seção"
            self.definir_representacao_secao(driver, wait)
            
            # 3. Definir região como "GA: Gastrica"
            self.definir_regiao_gastrica(driver, wait, mascara)

            # 4. Definir quantidade de fragmentos (campo D)
            self.definir_quantidade_fragmentos(driver, wait, campo_d)
            
            # 5. Definir quantidade de blocos como "1"
            self.definir_quantidade_blocos(driver, wait)
            
            log_message("✅ Campos pré-envio preenchidos com sucesso!", "SUCCESS")
            
        except Exception as e:
            log_message(f"⚠️ Erro no preenchimento dos campos pré-envio: {e}", "WARNING")
            log_message("⚠️ Continuando com o envio para próxima etapa", "WARNING")

    def enviar_proxima_etapa(self, driver, wait):
        """Clica no botão de enviar para próxima etapa"""
        try:
            botao_enviar = wait.until(
                EC.element_to_be_clickable((By.ID, "btn-enviar-proxima-etapa"))
            )
            botao_enviar.click()
            log_message("➡️ Clicou em Enviar para próxima etapa", "INFO")
            time.sleep(1.5)
        except Exception as e:
            log_message(f"Erro ao enviar para próxima etapa: {e}", "ERROR")
            raise

    def assinar_com_george(self, driver, wait):
        """Faz o processo de assinatura com Dr. George"""
        try:
            # Aguardar o modal de assinatura aparecer
            wait.until(EC.presence_of_element_located((By.ID, "assinatura")))
            log_message("📋 Modal de assinatura aberto", "INFO")
            
            # Encontrar e clicar no checkbox do Dr. George (value="2173")
            checkbox_george = wait.until(
                EC.element_to_be_clickable((By.XPATH, "//input[@type='checkbox' and @value='2173']"))
            )
            checkbox_george.click()
            log_message("✅ Checkbox do Dr. George marcado", "INFO")
            time.sleep(1)
            
            # Aguardar o campo de senha aparecer e digitar a senha
            campo_senha = wait.until(
                EC.presence_of_element_located((By.NAME, "senha_2173"))
            )
            campo_senha.send_keys("1323")
            log_message("🔐 Senha digitada", "INFO")
            time.sleep(1)
            
            # Clicar no botão Assinar
            botao_assinar = wait.until(
                EC.element_to_be_clickable((By.ID, "salvarAss"))
            )
            botao_assinar.click()
            log_message("✍️ Clicou em Assinar", "INFO")
            time.sleep(1.5)
            
        except Exception as e:
            log_message(f"Erro no processo de assinatura: {e}", "ERROR")
            raise

    def run(self, params: dict):
        username = params.get("username")
        password = params.get("password")
        excel_file = params.get("excel_file")
        cancel_flag = params.get("cancel_flag")
        
        try:
            # Lê os dados dos exames da planilha (código e máscara)
            dados_exames = self.get_dados_exames(excel_file)
            if not dados_exames:
                messagebox.showerror("Erro", "Nenhum dado de exame encontrado na planilha.")
                return
            
            log_message(f"Encontrados {len(dados_exames)} exames para processar", "INFO")
            
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao ler o Excel: {e}")
            return

        url = os.getenv("SYSTEM_URL", "https://pathoweb.com.br/login/auth")
        driver = None
        resultados = []
        
        try:
            driver = BrowserFactory.create_chrome()
            wait = WebDriverWait(driver, 10)
            
            log_message("Iniciando automação de macroscopia gástrica...", "INFO")
            
            # Login
            log_message("Fazendo login...", "INFO")
            driver.get(url)
            
            # Aguardar página carregar completamente
            wait.until(EC.presence_of_element_located((By.ID, "username")))
            
            username_field = driver.find_element(By.ID, "username")
            username_field.clear()
            username_field.send_keys(username)
            
            password_field = driver.find_element(By.ID, "password")
            password_field.clear()
            password_field.send_keys(password)
            
            submit_button = driver.find_element(By.CSS_SELECTOR, "button[type='submit']")
            submit_button.click()
            
            log_message("Navegando para módulo de exames...", "INFO")
            
            # Navegar para o módulo de exames (módulo 1)
            modulo_link = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "a[href='/site/trocarModulo?modulo=1']")))
            modulo_link.click()
            time.sleep(2)
            
            # Fechar modal se aparecer
            try:
                modal_close_button = driver.find_element(By.CSS_SELECTOR, "#mensagemParaClienteModal .modal-footer button")
                if modal_close_button.is_displayed():
                    modal_close_button.click()
            except Exception:
                pass

            log_message("✅ Login realizado com sucesso. Iniciando processamento dos exames.", "SUCCESS")
            
            # Processar cada exame da planilha
            for i, exame_data in enumerate(dados_exames, 1):
                if cancel_flag and cancel_flag.is_set():
                    log_message("Execução cancelada pelo usuário.", "WARNING")
                    break
                
                codigo = exame_data['codigo']
                mascara = exame_data['mascara']
                campo_d = exame_data['campo_d']
                campo_e = exame_data['campo_e']
                campo_f = exame_data['campo_f']
                campo_g = exame_data['campo_g']
                
                log_message(f"\n➡️ Processando exame {i}/{len(dados_exames)}: {codigo} (máscara: {mascara})", "INFO")
                
                try:
                    # Verificar se o browser ainda está ativo
                    if not self.verificar_sessao_browser(driver):
                        log_message("🔄 Recriando browser devido à sessão perdida...", "WARNING")
                        try:
                            driver.quit()
                        except:
                            pass
                        
                        # Recriar browser e fazer login novamente
                        driver = BrowserFactory.create_chrome()
                        wait = WebDriverWait(driver, 10)
                        
                        # Fazer login novamente
                        log_message("🔄 Fazendo login novamente...", "INFO")
                        driver.get(url)
                        
                        # Aguardar página carregar completamente
                        wait.until(EC.presence_of_element_located((By.ID, "username")))
                        time.sleep(2)
                        
                        username_field = driver.find_element(By.ID, "username")
                        username_field.clear()
                        username_field.send_keys(username)
                        
                        password_field = driver.find_element(By.ID, "password")
                        password_field.clear()
                        password_field.send_keys(password)
                        
                        submit_button = driver.find_element(By.CSS_SELECTOR, "button[type='submit']")
                        submit_button.click()
                        
                        log_message("🔄 Navegando para módulo de exames novamente...", "INFO")
                        
                        # Navegar para o módulo de exames (módulo 1)
                        modulo_link = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "a[href='/site/trocarModulo?modulo=1']")))
                        modulo_link.click()
                        time.sleep(2.5)
                        
                        # Fechar modal se aparecer
                        try:
                            modal_close_button = driver.find_element(By.CSS_SELECTOR, "#mensagemParaClienteModal .modal-footer button")
                            if modal_close_button.is_displayed():
                                modal_close_button.click()
                                time.sleep(1)
                        except Exception:
                            pass
                        
                        log_message("✅ Browser recriado e login realizado novamente", "SUCCESS")
                    
                    # Processar este exame específico
                    resultado = self.processar_exame(driver, wait, codigo, mascara, campo_d, campo_e, campo_f, campo_g)
                    resultados.append({
                        'codigo': codigo,
                        'mascara': mascara,
                        'campo_d': campo_d,
                        'campo_e': campo_e,
                        'campo_f': campo_f,
                        'campo_g': campo_g,
                        'status': resultado['status'],
                        'detalhes': resultado.get('detalhes', '')
                    })
                    
                except Exception as e:
                    log_message(f"❌ Erro ao processar exame {codigo}: {e}", "ERROR")
                    resultados.append({
                        'codigo': codigo,
                        'mascara': mascara,
                        'campo_d': campo_d,
                        'campo_e': campo_e,
                        'campo_f': campo_f,
                        'campo_g': campo_g,
                        'status': 'erro',
                        'detalhes': str(e)
                    })
            
            # Mostrar resumo final
            self.mostrar_resumo_final(resultados)
            
        except Exception as e:
            log_message(f"❌ Erro durante a automação: {e}", "ERROR")
            messagebox.showerror("Erro", f"❌ Erro durante a automação:\n{str(e)[:200]}...")
        finally:
            if driver:
                try:
                    driver.quit()
                    log_message("Browser fechado", "INFO")
                except Exception as quit_error:
                    log_message(f"Erro ao fechar browser: {quit_error}", "WARNING")

    def processar_exame(self, driver, wait, codigo, mascara, campo_d, campo_e, campo_f, campo_g):
        """Processa um exame individual"""
        try:
            # Verificar se a sessão do browser ainda está ativa
            if not self.verificar_sessao_browser(driver):
                raise Exception("Sessão do browser perdida - necessário reiniciar")
            
            # Aguardar e encontrar o campo de código de barras diretamente pelo placeholder (mais confiável)
            try:
                campo_codigo = wait.until(EC.presence_of_element_located((By.XPATH, "//input[@placeholder='Leitor de código de barras']")))
                log_message("✅ Campo de código encontrado", "INFO")
            except:
                # Fallback para ID se placeholder não funcionar
                campo_codigo = wait.until(EC.presence_of_element_located((By.ID, "inputSearchCodBarra")))
                log_message("✅ Campo de código encontrado pelo ID", "INFO")

            # Preencher o campo de código e clicar no botão de pesquisar
            campo_codigo.clear()
            campo_codigo.send_keys(codigo)
            log_message(f"✍️ Código '{codigo}' digitado no campo", "SUCCESS")

            # Clicar no botão de pesquisar (consultarExameBarraAbrirPorBarCode)
            try:
                botao_pesquisar = wait.until(EC.element_to_be_clickable((By.ID, "consultarExameBarraAbrirPorBarCode")))
                botao_pesquisar.click()
                log_message("🔍 Clicou no botão de pesquisar exame", "SUCCESS")
            except Exception as e:
                log_message(f"⚠️ Não foi possível clicar no botão de pesquisar: {e}", "WARNING")
                raise

            # Aguardar div de andamento aparecer
            return self.aguardar_e_processar_andamento(driver, wait, mascara, campo_d, campo_e, campo_f, campo_g)

        except Exception as e:
            error_message = str(e)
            log_message(f"Erro ao processar exame {codigo}: {error_message}", "ERROR")
            
            # Verificar se é erro de sessão inválida
            if "invalid session id" in error_message.lower():
                log_message("❌ Erro de sessão inválida detectado", "ERROR")
                return {'status': 'erro_sessao', 'detalhes': 'Sessão do browser perdida'}
            
            # Screenshot do erro para outros tipos de erro
            try:
                screenshot_path = f"erro_exame_{codigo}_{int(time.time())}.png"
                driver.save_screenshot(screenshot_path)
                log_message(f"Screenshot do erro salvo em: {screenshot_path}", "INFO")
            except:
                pass
            return {'status': 'erro', 'detalhes': error_message}

    def aguardar_e_processar_andamento(self, driver, wait, mascara, campo_d, campo_e, campo_f, campo_g):
        """Aguarda a div de andamento e processa o exame"""
        # Aguardar div de andamento aparecer (otimizado)
        try:
            wait.until(EC.presence_of_element_located((By.ID, "divAndamentoExame")))
            log_message("📋 Div de andamento do exame encontrada!", "SUCCESS")
            time.sleep(0.5)  # Reduzido de 2 para 0.5
        except:
            log_message("⚠️ Div de andamento não apareceu no tempo esperado", "WARNING")
            return {'status': 'sem_andamento', 'detalhes': 'Exame não encontrado ou não carregou'}
        
        # Processar conclusão diretamente sem verificar SVG
        log_message("✅ Exame carregado - iniciando processo de conclusão", "SUCCESS")
        return self.processar_conclusao_completa(driver, wait, mascara, campo_d, campo_e, campo_f, campo_g)

    def processar_conclusao_completa(self, driver, wait, mascara, campo_d, campo_e, campo_f, campo_g):
        """Processa a conclusão completa do exame"""
        try:
            # 1. Selecionar Nathalia como Macroscopista Responsável
            self.selecionar_responsavel_macroscopia(driver, wait)
            
            # 2. Selecionar Renata como Auxiliar da Macroscopia
            self.selecionar_auxiliar_macroscopia(driver, wait)
            
            # 3. Definir data de hoje
            self.definir_data_fixacao(driver, wait)

            # 4. Definir hora 18:00
            self.definir_hora_fixacao(driver, wait)
            
            # 5. Digitar a máscara e buscar (se houver)
            if mascara:
                self.digitar_mascara_e_buscar(driver, wait, mascara)
            else:
                log_message("⚠️ Nenhuma máscara encontrada, pulando busca", "WARNING")
            
            # 6. Abrir modal de variáveis e preencher campos (opcional)
            try:
                self.abrir_modal_variaveis_e_preencher(driver, wait, campo_d, campo_e, campo_f, campo_g)
            except Exception as var_error:
                log_message(f"⚠️ Erro no modal de variáveis: {var_error}", "WARNING")
                log_message("⚠️ Continuando o processo sem as variáveis", "WARNING")
            
            # 7. Salvar macroscopia
            self.salvar_macroscopia(driver, wait)
            
            # 8. Preencher campos necessários antes de enviar para próxima etapa
            self.preencher_campos_pre_envio(driver, wait, mascara, campo_d)
            
            # 9. Salvar fragmentos
            self.salvar_fragmentos(driver, wait)
            
            # 10. Enviar para próxima etapa
            self.enviar_proxima_etapa(driver, wait)
            
            log_message("🎉 Processo de macroscopia finalizado com sucesso!", "SUCCESS")
            return {'status': 'sucesso', 'detalhes': 'Macroscopia processada com sucesso'}
            
        except Exception as e:
            log_message(f"Erro durante processo de macroscopia: {e}", "ERROR")
            return {'status': 'erro_macroscopia', 'detalhes': str(e)}

    def mostrar_resumo_final(self, resultados):
        """Mostra o resumo final do processamento"""
        total = len(resultados)
        sucesso = len([r for r in resultados if r['status'] == 'sucesso'])
        sem_andamento = len([r for r in resultados if r['status'] == 'sem_andamento'])
        erro_sessao = len([r for r in resultados if r['status'] == 'erro_sessao'])
        erros = len([r for r in resultados if 'erro' in r['status'] and r['status'] != 'erro_sessao'])
        
        log_message("\n" + "="*50, "INFO")
        log_message("RESUMO FINAL DO PROCESSAMENTO", "INFO")
        log_message("="*50, "INFO")
        log_message(f"Total de exames: {total}", "INFO")
        log_message(f"✅ Processados com sucesso: {sucesso}", "SUCCESS")
        log_message(f"⚠️ Exames não encontrados: {sem_andamento}", "WARNING")
        log_message(f"🔄 Erros de sessão (browser perdido): {erro_sessao}", "WARNING")
        log_message(f"❌ Outros erros de processamento: {erros}", "ERROR")
        
        # Mostrar detalhes dos erros se houver
        erros_totais = erro_sessao + erros
        if erros_totais > 0:
            log_message("\nDetalhes dos erros:", "ERROR")
            for r in resultados:
                if 'erro' in r['status']:
                    log_message(f"- {r['codigo']}: {r['detalhes']}", "ERROR")
        
        messagebox.showinfo("Processamento Concluído", 
            f"✅ Processamento finalizado!\n\n"
            f"Total: {total}\n"
            f"Sucesso: {sucesso}\n"
            f"Não encontrados: {sem_andamento}\n"
            f"Erros de sessão: {erro_sessao}\n"
            f"Outros erros: {erros}")

def run(params: dict):
    module = MacroGastricaModule()
    module.run(params)
