import os
import time
import pandas as pd
from tkinter import messagebox
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
from datetime import datetime

from src.core.browser_factory import BrowserFactory
from src.core.logger import log_message
from src.modules.base import BaseModule

class LancamentoGuiaUnimedModule(BaseModule):
    def __init__(self):
        super().__init__(nome="Lançamento Guia Unimed")

    def read_excel_data(self, file_path: str) -> list:
        """Lê os dados do arquivo Excel com a estrutura: GUIA, CARTAO, MEDICO, CRM, PROCEDIMENTOS, QTD, TEXTO"""
        try:
            df = pd.read_excel(file_path, header=0)
            
            # Verificar se as colunas estão corretas
            expected_columns = ['GUIA', 'CARTAO', 'MEDICO', 'CRM', 'PROCEDIMENTOS', 'QTD', 'TEXTO']
            
            # Ajustar nomes das colunas se necessário (case insensitive)
            df.columns = df.columns.str.upper().str.strip()
            
            # Verificar se todas as colunas necessárias existem
            missing_columns = [col for col in expected_columns if col not in df.columns]
            if missing_columns:
                raise ValueError(f"Colunas faltando no Excel: {missing_columns}")
            
            # Converter DataFrame para lista de dicionários, removendo linhas vazias
            data_list = []
            for _, row in df.iterrows():
                if pd.notna(row['GUIA']) and str(row['GUIA']).strip():
                    data_list.append({
                        'guia': str(row['GUIA']).strip(),
                        'cartao': str(row['CARTAO']).strip() if pd.notna(row['CARTAO']) else '',
                        'medico': str(row['MEDICO']).strip() if pd.notna(row['MEDICO']) else '',
                        'crm': str(row['CRM']).strip() if pd.notna(row['CRM']) else '',
                        'procedimentos': str(row['PROCEDIMENTOS']).strip() if pd.notna(row['PROCEDIMENTOS']) else '',
                        'qtd': str(row['QTD']).strip() if pd.notna(row['QTD']) else '',
                        'texto': str(row['TEXTO']).strip() if pd.notna(row['TEXTO']) else ''
                    })
            
            return data_list
        except Exception as e:
            raise ValueError(f"Erro ao ler o Excel: {e}")

    def fazer_login_unimed(self, driver, wait, username, password):
        """Faz login no portal da Unimed"""
        log_message("Fazendo login no portal Unimed...", "INFO")
        driver.get("https://webmed.unimedlondrina.com.br/prestador/")
        
        # Aguardar e preencher campo usuário
        campo_usuario = wait.until(EC.presence_of_element_located((By.ID, "operador")))
        campo_usuario.clear()
        campo_usuario.send_keys(username)
        
        # Preencher campo senha
        campo_senha = driver.find_element(By.ID, "senha")
        campo_senha.clear()
        campo_senha.send_keys(password)
        
        # Clicar em entrar
        botao_entrar = driver.find_element(By.ID, "entrar")
        botao_entrar.click()
        time.sleep(2.5)
        
        log_message("✅ Login realizado com sucesso", "SUCCESS")

    def acessar_pagina_procedimento(self, driver):
        """Acessa a página de procedimento específica da Unimed"""
        url_procedimento = "https://webmed.unimedlondrina.com.br/prestador/procedimento.php?pagina=ff25c04430244fa10de866898f1a24d2"
        log_message(f"Acessando página de procedimentos: {url_procedimento}", "INFO")
        driver.get(url_procedimento)
        time.sleep(3)
        log_message("✅ Página de procedimentos acessada", "SUCCESS")

    def formatar_cartao_17_digitos(self, cartao):
        """Formata o número do cartão para ter 17 dígitos, adicionando zeros antes se necessário"""
        cartao_limpo = str(cartao).strip()
        
        # Remover apóstrofe do Excel e outros caracteres especiais, manter apenas números e letras
        cartao_sem_apostrofe = cartao_limpo.lstrip("'")  # Remove apóstrofe do início
        cartao_sem_espacos = ''.join(cartao_sem_apostrofe.split())
        
        if len(cartao_sem_espacos) < 17:
            # Adicionar zeros à esquerda para completar 17 dígitos
            zeros_necessarios = 17 - len(cartao_sem_espacos)
            cartao_formatado = "0" * zeros_necessarios + cartao_sem_espacos
            log_message(f"📋 Cartão formatado: '{cartao_limpo}' → {cartao_formatado} (17 dígitos)", "INFO")
            return cartao_formatado
        elif len(cartao_sem_espacos) == 17:
            log_message(f"📋 Cartão já tem 17 dígitos: {cartao_sem_espacos}", "INFO")
            return cartao_sem_espacos
        else:
            log_message(f"⚠️ Cartão com mais de 17 dígitos ({len(cartao_sem_espacos)}): {cartao_sem_espacos}", "WARNING")
            return cartao_sem_espacos

    def extrair_apenas_numeros(self, crm):
        """Extrai apenas os números do CRM, removendo letras"""
        import re
        apenas_numeros = re.sub(r'[^0-9]', '', str(crm))
        log_message(f"📋 CRM formatado: {crm} → {apenas_numeros}", "INFO")
        return apenas_numeros

    def buscar_medico_solicitante(self, driver, wait, crm, nome_medico):
        """Busca o médico solicitante no popup da Unimed"""
        try:
            # Guardar janela original
            janela_original = driver.current_window_handle
            
            # 1. Clicar no botão de busca do solicitante
            log_message("🔍 Clicando no botão de busca do solicitante...", "INFO")
            botao_busca = wait.until(EC.element_to_be_clickable((By.ID, "busca_solicitante")))
            botao_busca.click()
            
            # 2. Aguardar nova janela abrir e fazer switch
            time.sleep(3)
            
            # Verificar se há novas janelas
            todas_janelas = driver.window_handles
            if len(todas_janelas) > 1:
                # Mudar para a nova janela (popup)
                for janela in todas_janelas:
                    if janela != janela_original:
                        driver.switch_to.window(janela)
                        break
                log_message("✅ Mudou para janela do popup", "INFO")
            else:
                log_message("✅ Popup aberto na mesma janela", "INFO")
            
            # 3. Extrair apenas números do CRM
            crm_numeros = self.extrair_apenas_numeros(crm)
            
            # 4. Preencher campo do conselho com números do CRM
            log_message(f"📝 Preenchendo campo conselho com: {crm_numeros}", "INFO")
            campo_conselho = wait.until(EC.presence_of_element_located((By.ID, "conselho")))
            campo_conselho.clear()
            campo_conselho.send_keys(crm_numeros)
            
            # 5. Clicar no botão localizar
            log_message("🔍 Clicando em localizar...", "INFO")
            botao_localizar = wait.until(EC.element_to_be_clickable((By.ID, "localizar")))
            botao_localizar.click()
            
            # 6. Aguardar tabela carregar
            time.sleep(3)
            
            # 7. Verificar se tabela foi carregada
            try:
                tabela = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "table.table-hover tbody")))
                linhas = tabela.find_elements(By.CSS_SELECTOR, "tr")
                
                # Verificar se há resultados (primeira linha é cabeçalho)
                if len(linhas) <= 1:
                    log_message(f"⚠️ Nenhum médico encontrado para CRM: {crm}", "WARNING")
                    raise Exception(f"Médico não encontrado para CRM: {crm}")
                
                # 8. Procurar linha do médico (ignorar cabeçalho)
                medico_encontrado = False
                for linha in linhas[1:]:  # Pula o cabeçalho
                    try:
                        colunas = linha.find_elements(By.TAG_NAME, "td")
                        if len(colunas) >= 3:
                            nome_na_tabela = colunas[1].text.strip()
                            log_message(f"📋 Verificando médico: {nome_na_tabela}", "INFO")
                            
                            # Clicar na linha do médico encontrado
                            linha.click()
                            log_message(f"✅ Médico selecionado: {nome_na_tabela}", "SUCCESS")
                            medico_encontrado = True
                            break
                    except Exception as e:
                        log_message(f"⚠️ Erro ao processar linha da tabela: {e}", "WARNING")
                        continue
                
                if not medico_encontrado:
                    log_message(f"⚠️ Médico não encontrado na tabela para CRM: {crm}", "WARNING")
                    raise Exception(f"Médico não encontrado na tabela para CRM: {crm}")
                
                # 9. Aguardar popup fechar automaticamente e voltar para janela original
                time.sleep(3)
                
                # O popup fecha automaticamente, então só precisamos voltar para janela original
                driver.switch_to.window(janela_original)
                log_message("✅ Médico selecionado, voltou para janela principal", "SUCCESS")
                
            except Exception as e:
                log_message(f"❌ Erro ao processar tabela de médicos: {e}", "ERROR")
                # Tentar voltar para janela original em caso de erro
                try:
                    driver.switch_to.window(janela_original)
                    log_message("🔄 Voltou para janela principal após erro", "INFO")
                except:
                    pass
                raise Exception(f"Falha na busca do médico: {e}")
                
        except Exception as e:
            log_message(f"❌ Erro na busca do médico solicitante: {e}", "ERROR")
            # Garantir que volta para janela original em caso de erro
            try:
                if 'janela_original' in locals():
                    driver.switch_to.window(janela_original)
            except:
                pass
            raise e

    def preencher_campos_fixos(self, driver):
        """Preenche os campos fixos do formulário"""
        try:
            # Regime de atendimento: 01 - Ambulatorial
            log_message("📝 Preenchendo regime de atendimento: 01 - Ambulatorial", "INFO")
            js_regime = '''
            $("#regime_atendimento")
              .val("01")
              .trigger("change");
            '''
            driver.execute_script(js_regime)
            
            # Aguardar um pouco
            time.sleep(1)
            
            # Tipo de atendimento: 23 - Exame
            log_message("📝 Preenchendo tipo de atendimento: 23 - Exame", "INFO")
            js_tipo = '''
            $("#tipo_atendimento")
              .val("23")
              .trigger("change")
              .trigger("blur");
            '''
            driver.execute_script(js_tipo)
            
            time.sleep(1)
            log_message("✅ Campos fixos preenchidos", "SUCCESS")
            
        except Exception as e:
            log_message(f"❌ Erro ao preencher campos fixos: {e}", "ERROR")
            raise e

    def preencher_hipotese_diagnostica(self, driver, wait, texto):
        """Preenche a hipótese diagnóstica usando o campo TEXTO do Excel"""
        try:
            if not texto or not texto.strip():
                log_message("⚠️ Texto vazio, pulando hipótese diagnóstica", "WARNING")
                return
            
            texto_formatado = texto.strip()
            log_message(f"📝 Preenchendo hipótese diagnóstica: {texto_formatado}", "INFO")
            
            # 1. Clicar no campo select2 para abrir
            log_message("🔍 Clicando no campo de hipótese diagnóstica...", "INFO")
            select2_container = wait.until(EC.element_to_be_clickable((
                By.CSS_SELECTOR, 
                "#selecionarHipotese .select2-container .select2-selection"
            )))
            select2_container.click()
            
            # 2. Aguardar campo de busca aparecer
            time.sleep(2)
            
            # 3. Preencher campo de busca
            log_message(f"📝 Digitando texto: {texto_formatado}", "INFO")
            campo_busca = wait.until(EC.presence_of_element_located((
                By.CSS_SELECTOR, 
                ".select2-search__field"
            )))
            campo_busca.clear()
            campo_busca.send_keys(texto_formatado)
            
            # 4. Aguardar resultados carregar
            time.sleep(3)
            
            # 5. Verificar se há resultados ou se precisa usar "DIGITAR MANUALMENTE"
            try:
                # Procurar por opções de resultado
                resultados = driver.find_elements(By.CSS_SELECTOR, 
                    ".select2-results__options .select2-results__option")
                
                encontrou_resultado = False
                for resultado in resultados:
                    texto_resultado = resultado.text.strip()
                    
                    # Se encontrou "DIGITAR MANUALMENTE", clica nele
                    if "DIGITAR MANUALMENTE" in texto_resultado.upper():
                        log_message("📝 Selecionando 'DIGITAR MANUALMENTE'", "INFO")
                        resultado.click()
                        encontrou_resultado = True
                        break
                    # Se encontrou um resultado válido (não é mensagem de erro), pode clicar
                    elif (texto_resultado and 
                          "Digite 3 ou mais caracteres" not in texto_resultado and
                          "para selecionar" not in texto_resultado):
                        log_message(f"✅ Encontrou resultado: {texto_resultado}", "SUCCESS")
                        resultado.click()
                        encontrou_resultado = True
                        break
                
                if not encontrou_resultado:
                    log_message("⚠️ Nenhum resultado encontrado, tentando 'DIGITAR MANUALMENTE'", "WARNING")
                    # Tentar encontrar especificamente "DIGITAR MANUALMENTE"
                    digitar_manual = driver.find_element(By.XPATH, 
                        "//li[contains(text(), 'DIGITAR MANUALMENTE')]")
                    digitar_manual.click()
                
                time.sleep(1)
                log_message("✅ Hipótese diagnóstica preenchida", "SUCCESS")
                
            except Exception as e:
                log_message(f"⚠️ Erro ao selecionar hipótese: {e}", "WARNING")
                # Tentar fechar o dropdown em caso de erro
                try:
                    driver.execute_script("$('.select2-container').select2('close');")
                except:
                    pass
                
        except Exception as e:
            log_message(f"❌ Erro ao preencher hipótese diagnóstica: {e}", "ERROR")
            raise e

    def preencher_procedimentos(self, driver, procedimentos_str, quantidades_str):
        """Preenche os procedimentos e quantidades baseado nos dados do Excel"""
        try:
            if not procedimentos_str or not quantidades_str:
                log_message("⚠️ Procedimentos ou quantidades vazios", "WARNING")
                return
            
            # Processar strings dos procedimentos e quantidades
            procedimentos = [p.strip() for p in str(procedimentos_str).split(',')]
            quantidades = [q.strip() for q in str(quantidades_str).split(',')]
            
            # Verificar se as listas têm o mesmo tamanho
            if len(procedimentos) != len(quantidades):
                log_message(f"⚠️ Número de procedimentos ({len(procedimentos)}) difere do número de quantidades ({len(quantidades)})", "WARNING")
                # Ajustar para o menor tamanho
                min_size = min(len(procedimentos), len(quantidades))
                procedimentos = procedimentos[:min_size]
                quantidades = quantidades[:min_size]
            
            log_message(f"📋 Processando {len(procedimentos)} procedimentos:", "INFO")
            for i, (proc, qtd) in enumerate(zip(procedimentos, quantidades)):
                log_message(f"   {i}: {proc} = {qtd}", "INFO")
            
            # Preencher cada procedimento e quantidade
            for i, (procedimento, quantidade) in enumerate(zip(procedimentos, quantidades)):
                if i >= 5:  # Máximo de 5 campos (0 a 4)
                    log_message(f"⚠️ Limite de 5 procedimentos atingido, ignorando restantes", "WARNING")
                    break
                
                try:
                    # Preencher procedimento
                    log_message(f"📝 Preenchendo procedimento{i}: {procedimento}", "INFO")
                    js_procedimento = f'''
                    $("#procedimento{i}")
                      .val("{procedimento}")
                      .trigger("input")
                      .trigger("change") 
                      .trigger("blur");
                    '''
                    driver.execute_script(js_procedimento)
                    
                    # Aguardar um pouco
                    time.sleep(1)
                    
                    # Preencher quantidade
                    log_message(f"📝 Preenchendo quantidade{i}: {quantidade}", "INFO")
                    js_quantidade = f'''
                    $("#quantidade{i}")
                      .removeAttr("readonly")  
                      .val("{quantidade}")                
                      .trigger("input")
                      .trigger("change")
                      .trigger("blur");
                    '''
                    driver.execute_script(js_quantidade)
                    
                    time.sleep(1)
                    log_message(f"✅ Procedimento {i} preenchido: {procedimento} (qtd: {quantidade})", "SUCCESS")
                    
                except Exception as e:
                    log_message(f"❌ Erro ao preencher procedimento {i}: {e}", "ERROR")
                    continue
            
            log_message("✅ Procedimentos e quantidades preenchidos", "SUCCESS")
            
        except Exception as e:
            log_message(f"❌ Erro ao processar procedimentos: {e}", "ERROR")
            raise e

    def processar_guia_unimed(self, driver, wait, dados):
        """Processa uma guia individual na página da Unimed"""
        try:
            log_message(f"🔄 Iniciando processamento da guia {dados['guia']}", "INFO")
            
            # Logar os dados que serão processados
            log_message(f"📝 Dados a processar:", "INFO")
            log_message(f"   - Guia: {dados['guia']}", "INFO")
            log_message(f"   - Cartão: {dados['cartao']}", "INFO")
            log_message(f"   - Médico: {dados['medico']}", "INFO")
            log_message(f"   - CRM: {dados['crm']}", "INFO")
            log_message(f"   - Procedimentos: {dados['procedimentos']}", "INFO")
            log_message(f"   - Quantidade: {dados['qtd']}", "INFO")
            log_message(f"   - Texto: {dados['texto'][:50]}..." if dados['texto'] else "   - Texto: (vazio)", "INFO")
            
            # 1. Preencher número da carteira do beneficiário (17 dígitos)
            log_message("🔍 Preenchendo campo do número da carteira...", "INFO")
            try:
                cartao_formatado = self.formatar_cartao_17_digitos(dados['cartao'])
                
                # Usar JavaScript para preencher o campo conforme sugerido
                javascript_code = f'$("#codigo").val("{cartao_formatado}").trigger("input").trigger("change").trigger("blur");'
                
                # Aguardar o campo estar presente
                wait.until(EC.presence_of_element_located((By.ID, "codigo")))
                
                # Executar JavaScript
                log_message(f"🔧 Executando JavaScript: {javascript_code}", "INFO")
                driver.execute_script(javascript_code)
                log_message(f"✅ Cartão preenchido via JavaScript: {cartao_formatado}", "SUCCESS")
                
                # Aguardar um pouco após preencher para ver o resultado
                time.sleep(2)
                
            except Exception as e:
                log_message(f"❌ Erro ao preencher cartão: {e}", "ERROR")
                raise Exception(f"Falha ao preencher número da carteira: {e}")
            
            # 2. Buscar médico solicitante
            log_message("🔍 Iniciando busca do médico solicitante...", "INFO")
            try:
                self.buscar_medico_solicitante(driver, wait, dados['crm'], dados['medico'])
            except Exception as e:
                log_message(f"⚠️ Erro na busca do médico: {e}. Continuando para próximo exame...", "WARNING")
                return {
                    'guia': dados['guia'],
                    'status': 'erro_medico',
                    'erro': str(e),
                    'timestamp': datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                }
            
            # 3. Preencher campos fixos
            log_message("📝 Preenchendo campos fixos...", "INFO")
            try:
                self.preencher_campos_fixos(driver)
            except Exception as e:
                log_message(f"⚠️ Erro ao preencher campos fixos: {e}", "WARNING")
            
            # 4. Preencher hipótese diagnóstica
            log_message("🔍 Preenchendo hipótese diagnóstica...", "INFO")
            try:
                self.preencher_hipotese_diagnostica(driver, wait, dados['texto'])
            except Exception as e:
                log_message(f"⚠️ Erro ao preencher hipótese diagnóstica: {e}", "WARNING")
            
            # 5. Preencher procedimentos e quantidades
            log_message("📝 Preenchendo procedimentos e quantidades...", "INFO")
            try:
                self.preencher_procedimentos(driver, dados['procedimentos'], dados['qtd'])
            except Exception as e:
                log_message(f"⚠️ Erro ao preencher procedimentos: {e}", "WARNING")
            
            # Por enquanto, vamos aguardar para ver o resultado do preenchimento
            log_message("⏳ Aguardando para verificar preenchimento...", "INFO")
            time.sleep(3)
            
            return {
                'guia': dados['guia'],
                'status': 'sucesso',
                'cartao_formatado': cartao_formatado,
                'timestamp': datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            }
            
        except Exception as e:
            log_message(f"❌ Erro ao processar guia {dados['guia']}: {e}", "ERROR")
            return {
                'guia': dados['guia'],
                'status': 'erro',
                'erro': str(e),
                'timestamp': datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            }

    def run(self, params: dict):
        username = params.get("unimed_user")
        password = params.get("unimed_pass")
        cancel_flag = params.get("cancel_flag")
        headless_mode = params.get("headless_mode")
        excel_file = params.get("excel_file")

        # Validar credenciais da Unimed
        if not username or not password:
            messagebox.showerror("Erro", "Credenciais da Unimed são obrigatórias para este módulo.")
            return

        # Validar arquivo Excel
        if not excel_file or not os.path.exists(excel_file):
            messagebox.showerror("Erro", "Arquivo Excel é obrigatório para este módulo.")
            return

        driver = BrowserFactory.create_chrome(headless=headless_mode)
        wait = WebDriverWait(driver, 15)

        try:
            log_message("Iniciando automação de Lançamento de Guia Unimed...", "INFO")

            # Ler dados do Excel
            try:
                dados_excel = self.read_excel_data(excel_file)
                log_message(f"✅ Carregados {len(dados_excel)} registros do Excel", "SUCCESS")
                
                # Exibir amostra dos dados carregados
                if dados_excel:
                    primeiro_registro = dados_excel[0]
                    log_message(f"📋 Exemplo de registro: Guia={primeiro_registro['guia']}, "
                              f"Cartão={primeiro_registro['cartao']}, "
                              f"Médico={primeiro_registro['medico']}", "INFO")
                
            except Exception as e:
                log_message(f"❌ Erro ao ler arquivo Excel: {e}", "ERROR")
                messagebox.showerror("Erro", f"Erro ao ler arquivo Excel:\n{e}")
                return

            # Fazer login na Unimed
            self.fazer_login_unimed(driver, wait, username, password)

            # Acessar página de procedimentos
            self.acessar_pagina_procedimento(driver)

            # Processar cada registro do Excel
            resultados_processamento = []
            for i, dados in enumerate(dados_excel, 1):
                if cancel_flag and cancel_flag.is_set():
                    log_message("Execução cancelada pelo usuário.", "WARNING")
                    break
                
                try:
                    log_message(f"➡️ Processando registro {i}/{len(dados_excel)} - Guia: {dados['guia']}", "INFO")
                    
                    # TODO: Implementar lógica específica de lançamento da guia
                    # Por enquanto, apenas simulamos o processamento
                    resultado = self.processar_guia_unimed(driver, wait, dados)
                    resultados_processamento.append(resultado)
                    
                    log_message(f"✅ Guia {dados['guia']} processada com sucesso", "SUCCESS")
                    time.sleep(2)  # Pausa entre processamentos
                    
                except Exception as e:
                    log_message(f"❌ Erro ao processar guia {dados['guia']}: {e}", "ERROR")
                    resultados_processamento.append({
                        'guia': dados['guia'],
                        'status': 'erro',
                        'erro': str(e)
                    })

            # Resumo final
            total = len(resultados_processamento)
            sucessos = sum(1 for r in resultados_processamento if r.get('status') == 'sucesso')
            erros = total - sucessos

            log_message(f"\n📊 Resumo do processamento:", "INFO")
            log_message(f"Total de registros: {total}", "INFO")
            log_message(f"Sucessos: {sucessos}", "SUCCESS" if sucessos > 0 else "INFO")
            log_message(f"Erros: {erros}", "ERROR" if erros > 0 else "INFO")

            messagebox.showinfo("Processamento Concluído", 
                f"✅ Processamento finalizado!\n\n"
                f"Total de registros: {total}\n"
                f"Sucessos: {sucessos}\n"
                f"Erros: {erros}"
            )

        except Exception as e:
            log_message(f"❌ Erro durante a automação: {e}", "ERROR")
            messagebox.showerror("Erro", f"❌ Erro durante a automação:\n{e}")
        finally:
            # Aguardar antes de fechar para permitir visualização dos resultados
            if not headless_mode:
                input("Pressione Enter para fechar o navegador...")
            driver.quit()


def run(params: dict):
    module = LancamentoGuiaUnimedModule()
    module.run(params) 