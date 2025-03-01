# emissao_nf_2.3.ipynb

# Emissão de Notas Fiscais Unificado

## Importações Iniciais
### Dependências - Anaconda Prompt (adm)
#### conda install selenium pandas numpy openpyxl requests
#### pip install webdriver-manager

##########################################################################################################

# Cell 1
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException, ElementNotInteractableException
import pandas as pd
import math
import re
from time import sleep
import logging
import os
from openpyxl import Workbook
import csv
from webdriver_manager.chrome import ChromeDriverManager

# Configurações opcionais do Chrome
chrome_options = webdriver.ChromeOptions()
chrome_options.add_argument('--start-maximized')  # Iniciar maximizado
chrome_options.add_argument('--disable-notifications')  # Desabilitar notificações
# chrome_options.add_argument('--headless')  # Modo headless (sem interface gráfica)

# Inicialização do navegador usando webdriver-manager
service = Service(ChromeDriverManager().install())
navegador = webdriver.Chrome(service=service, options=chrome_options)

# Configuração de timeout global
navegador.set_page_load_timeout(30)  # timeout de 30 segundos para carregamento de página

# Acesso à URL
try:
    navegador.get("")
except TimeoutException:
    print("Timeout ao carregar a página. Verifique sua conexão com a internet.")
    navegador.quit()

##########################################################################################################

# Cell 2
# Preencher usuario
navegador.find_element("xpath", '//*[@id="usuario"]').send_keys("ocultado")

# Preencher e-mail
navegador.find_element("xpath", '//*[@id="senha"]').send_keys("ocultado")

# Clicar no botão (ARRUMAR)
navegador.find_element("xpath", '/html/body/div/main/div/div/div/div[2]/form/div[6]/span/button').click()

##########################################################################################################

#Cell 3
# Importação da planilha
arquivo_planilha = "ocultado.xlsx"
df = pd.read_excel(arquivo_planilha)
# Visualizar a planilha
display(df)

##########################################################################################################

#Cell 4
# Configuração do log
log_directory = "logs"
if not os.path.exists(log_directory):
    os.makedirs(log_directory)
log_file = os.path.join(log_directory, 'fluxo_nf.log')

logging.basicConfig(
    level=logging.DEBUG,
    format='%(asctime)s - %(levelname)s - %(message)s',
    filename=log_file,
    filemode='a'
)

print(f"O arquivo de log está sendo salvo em: {os.path.abspath(log_file)}")

# Inicializar arquivo de erros
erro_file_path = 'erros.xlsx'
if not os.path.exists(erro_file_path):
    wb = Workbook()
    ws = wb.active
    ws.append(["CPF", "Nome completo", "CEP", "Bairro", "Tipo Logradouro", "Logradouro", "Número", "Complemento", "E-mail", "Telefone", "Valor", "Procedimento"])
    wb.save(erro_file_path)

# Arquivo para armazenar registros processados
registros_processados_file = 'registros_processados.csv'

def carregar_registros_processados():
    processados = set()
    if os.path.exists(registros_processados_file):
        with open(registros_processados_file, 'r') as f:
            reader = csv.reader(f)
            for row in reader:
                if len(row) == 3:  # Garantir que temos CPF, valor e procedimento
                    processados.add((row[0], row[1], row[2]))
    logging.info(f"Carregados {len(processados)} registros processados.")
    return processados

def salvar_registro_processado(cpf, valor, procedimento):
    with open(registros_processados_file, 'a', newline='') as f:
        writer = csv.writer(f)
        writer.writerow([cpf, valor, procedimento])

def save_error_data(row_data):
    try:
        df = pd.read_excel(erro_file_path)
        df = df.append(row_data, ignore_index=True)
        df.to_excel(erro_file_path, index=False)
        logging.info("Dados do erro salvos no arquivo de erros.")
    except Exception as e:
        logging.error(f"Erro ao salvar dados no arquivo de erros: {e}")

def normalize_cep(cep):
    return re.sub(r'\D', '', cep)

def validar_cpf(cpf):
    cpf_str = re.sub(r'\D', '', str(cpf))
    return cpf_str.isdigit() and len(cpf_str) == 11

def validar_nome(nome):
    return isinstance(nome, str) and nome.strip() != ""

def validar_cep(cep):
    cep_str = normalize_cep(str(cep))
    return cep_str.isdigit() and len(cep_str) == 8

def preencher_campo(elemento, valor, descricao, max_tentativas=4):
    tentativas = 0
    while tentativas < max_tentativas:
        try:
            # Verifica se já existe algo no campo
            valor_atual = elemento.get_attribute("value")
            if valor_atual:
                logging.info(f"Campo {descricao} já contém valor: {valor_atual}")
                if valor_atual == str(valor):
                    logging.info(f"Valor existente em {descricao} é igual ao desejado. Mantendo.")
                    return True
                else:
                    logging.info(f"Valor existente em {descricao} é diferente. Atualizando.")
            
            elemento.click()
            sleep(0.5)
            elemento.send_keys(Keys.CONTROL + "a")
            elemento.send_keys(Keys.DELETE)
            if isinstance(valor, (float, int)):
                if math.isnan(valor):
                    logging.info(f"{descricao} está em branco. Continuando sem preencher.")
                    return True
                valor = str(int(valor))
            elif isinstance(valor, str):
                valor = valor.strip()
            elemento.send_keys(valor)
            sleep(0.5)
            elemento.send_keys(Keys.TAB)
            valor_preenchido = re.sub(r'\D', '', elemento.get_attribute("value"))
            if valor_preenchido == re.sub(r'\D', '', str(valor)):
                logging.info(f"{descricao} preenchido com sucesso.")
                return True
            else:
                logging.warning(f"Tentativa {tentativas + 1} falhou para {descricao}. Retentando...")
        except ElementNotInteractableException:
            logging.error(f"Erro: Não foi possível interagir com o elemento {descricao}. Tentando novamente...")
        tentativas += 1
    logging.error(f"Erro: Não foi possível preencher {descricao} após {max_tentativas} tentativas.")
    return False

def clicar_elemento_com_verificacao(xpath, descricao, verif_xpath, max_tentativas=4):
    tentativas = 0
    while tentativas < max_tentativas:
        try:
            elemento = WebDriverWait(navegador, 3).until(
                EC.element_to_be_clickable((By.XPATH, xpath))
            )
            elemento.click()
            sleep(0.5)

            try:
                if WebDriverWait(navegador, 3).until(
                    EC.visibility_of_element_located((By.XPATH, verif_xpath))
                ):
                    logging.info(f"{descricao} clicado com sucesso e ação verificada.")
                    return True
            except TimeoutException:
                logging.warning(f"Primeira verificação falhou para {descricao}, tentando novamente.")

            sleep(0.5)
            if WebDriverWait(navegador, 3).until(
                EC.visibility_of_element_located((By.XPATH, verif_xpath))
            ):
                logging.info(f"{descricao} clicado com sucesso e verificação na segunda tentativa.")
                return True

            logging.warning(f"Tentativa {tentativas + 1} falhou para verificar {descricao}. Retentando...")

        except (TimeoutException, NoSuchElementException, ElementNotInteractableException) as e:
            logging.error(f"Tentativa {tentativas + 1} falhou para clicar em {descricao}. Erro: {e}. Retentando...")
        tentativas += 1

    logging.error(f"Erro: Não foi possível clicar em {descricao} após {max_tentativas} tentativas. Salvando dados do erro e continuando.")
    save_error_data(row_data)
    navegador.find_element(By.TAG_NAME, 'body').send_keys(Keys.ESCAPE)
    return False

def tentar_fechar_tela_cadastro(navegador, max_tentativas=3):
    for tentativa in range(max_tentativas):
        try:
            navegador.find_element(By.TAG_NAME, 'body').send_keys(Keys.ESCAPE)
            sleep(1)
            cpf_campo = WebDriverWait(navegador, 3).until(
                EC.visibility_of_element_located((By.XPATH, '//*[@id="tomador"]'))
            )
            logging.info("Tela de cadastro fechada com sucesso.")
            return True
        except (TimeoutException, NoSuchElementException):
            logging.warning(f"Tentativa {tentativa + 1} de fechar tela de cadastro falhou.")
    logging.error("Não foi possível fechar a tela de cadastro após 3 tentativas.")
    return False

def preencher_nome_tomador(nome, max_tentativas=3):
    for tentativa in range(max_tentativas):
        try:
            logging.debug(f"Tentativa {tentativa + 1} de preencher o nome: {nome}")
            nome_elemento = WebDriverWait(navegador, 10).until(
                EC.visibility_of_element_located((By.XPATH, '//*[@id="topo"]/div[2]/div[3]/span/input'))
            )
            logging.debug("Elemento do nome encontrado")
            
            if not preencher_campo(nome_elemento, nome, "Nome do tomador"):
                logging.warning(f"Tentativa {tentativa + 1}: Falha ao preencher o nome do tomador")
                continue
            
            logging.info(f"Nome do tomador preenchido com sucesso: {nome}")
            return True
        except Exception as e:
            logging.error(f"Erro ao preencher nome do tomador (tentativa {tentativa + 1}): {str(e)}", exc_info=True)
        sleep(1)
    logging.error(f"Falha ao preencher nome do tomador após {max_tentativas} tentativas")
    return False

def fluxo_comum(cpf, valor_x100, procedimento, nome, cep, bairro, tipo, endereco, num, complemento, email, tel):
    logging.info(f"Iniciando o fluxo para CPF: {cpf}")

    for tentativa in range(4):
        try:
            cpf_elemento = WebDriverWait(navegador, 6).until(
                EC.presence_of_element_located((By.XPATH, '//*[@id="tomador"]'))
            )
            if not preencher_campo(cpf_elemento, str(cpf), "CPF"):
                raise Exception("Falha ao preencher CPF no início do fluxo.")

            # Verificar se o CPF já está cadastrado
            try:
                WebDriverWait(navegador, 3).until(
                    EC.presence_of_element_located((By.XPATH, '//*[contains(@id, "_list")]//li'))
                )
                logging.info(f"CPF {cpf} já cadastrado. Prosseguindo para emissão de NF.")
                # Seleciona o CPF cadastrado
                elemento = WebDriverWait(navegador, 3).until(
                    EC.element_to_be_clickable((By.XPATH, '//*[@id[contains(.,"_list")]]//li/div'))
                )
                elemento.click()
                logging.info("CPF cadastrado selecionado.")
            except TimeoutException:
                logging.info(f"CPF {cpf} não encontrado. Iniciando cadastro.")
                # Cadastro do CPF
                if not clicar_elemento_com_verificacao(
                    '//*[@id="btnNovoTomador"]',
                    "Botão Novo Tomador",
                    '//*[@id="topo"]/div[4]/div[1]/span/div/input'
                ):
                    raise Exception("Falha ao clicar no botão Novo Tomador.")

                cpf_form_elemento = WebDriverWait(navegador, 4).until(
                    EC.visibility_of_element_located((By.XPATH, '/html/body/div/main/div/div/div/div/div[2]/div[2]/div/div/div[2]/div[1]/div[2]/div[2]/span/div/input'))
                )
                cpf_form_elemento.click()
                sleep(0.5)
                if not preencher_campo(cpf_form_elemento, str(cpf), "CPF no formulário"):
                    raise Exception("Falha ao preencher CPF no formulário de cadastro.")

                if not clicar_elemento_com_verificacao(
                    '//*[@id="topo"]/div[2]/div[2]/span/div/button',
                    "Lupa do CPF do formulário",
                    '//*[@id="topo"]/div[4]/div[1]/span/div/input'
                ):
                    raise Exception("Falha ao clicar na lupa do CPF do formulário.")

                cep_elemento = WebDriverWait(navegador, 4).until(
                    EC.visibility_of_element_located((By.XPATH, '//*[@id="topo"]/div[4]/div[1]/span/div/input'))
                )
                if not preencher_campo(cep_elemento, str(cep), "CEP"):
                    raise Exception("Falha ao preencher o CEP.")

                # Tenta clicar no botão de confirmação do CEP, mas continua mesmo se falhar
                try:
                    if clicar_elemento_com_verificacao(
                        '//*[@id="app"]/main/div/div/div/div/div[2]/div[2]/div[2]/div/div[2]/div[1]/div/div/div[1]/table/tbody/tr[1]/td[2]/button',
                        "Botão de seleção do CEP",
                        '//*[@id="topo"]/div[4]/div[5]/span/input',
                        max_tentativas=2
                    ):
                        logging.info("CEP confirmado com sucesso.")
                    else:
                        logging.warning("Não foi possível confirmar o CEP, mas continuando o fluxo.")
                except Exception as e:
                    logging.warning(f"Erro ao tentar confirmar o CEP: {str(e)}. Continuando o fluxo.")

                if not preencher_nome_tomador(nome):
                    raise Exception("Falha ao preencher o nome do tomador")
                
                if bairro and isinstance(bairro, str) and bairro.strip():
                    preencher_campo(navegador.find_element(By.XPATH, '//*[@id="topo"]/div[4]/div[5]/span/input'), bairro, "Bairro")
                
                if tipo and isinstance(tipo, str) and tipo.strip():
                    preencher_campo(navegador.find_element(By.XPATH, '//*[@id="topo"]/div[4]/div[6]/span/input'), tipo, "Tipo Logradouro")
                
                if endereco and isinstance(endereco, str) and endereco.strip():
                    preencher_campo(navegador.find_element(By.XPATH, '//*[@id="topo"]/div[4]/div[7]/span/input'), endereco, "Logradouro")
                
                numero_elemento = navegador.find_element(By.XPATH, '//*[@id="topo"]/div[4]/div[8]/span/input')
                if num and isinstance(num, str) and num.strip():
                    preencher_campo(numero_elemento, num, "Número")
                else:
                    preencher_campo(numero_elemento, "s/n", "Número (s/n)")
                
                if complemento and isinstance(complemento, str) and complemento.strip():
                    preencher_campo(navegador.find_element(By.XPATH, '//*[@id="topo"]/div[4]/div[9]/span/input'), complemento, "Complemento")
                
                if email and isinstance(email, str) and email.strip():
                    preencher_campo(navegador.find_element(By.XPATH, '//*[@id="topo"]/div[5]/div[1]/span/input'), email, "Email")
                
                if tel and isinstance(tel, (int, float, str)):
                    if isinstance(tel, float) and math.isnan(tel):
                        logging.info("Telefone está em branco. Continuando sem preencher.")
                    else:
                        tel_str = re.sub(r'\D', '', str(tel))
                        preencher_campo(navegador.find_element(By.XPATH, '//*[@id="topo"]/div[5]/div[4]/span/input'), tel_str, "Celular")

                gravar_botao_cadastro = WebDriverWait(navegador, 6).until(
                    EC.element_to_be_clickable((By.XPATH, '/html/body/div/main/div/div/div/div/div[2]/div[2]/div/div/div[2]/div[2]/button[1]/span[2]'))
                )
                gravar_botao_cadastro.click()
                logging.info("Botão Gravar cadastro clicado.")

            botao_servico = WebDriverWait(navegador, 6).until(
                EC.element_to_be_clickable((By.XPATH, '//*[@id="app"]/main/div/div/div/div/div[2]/div[5]/div[2]/span/span/button'))
            )
            botao_servico.click()
            logging.info("Botão Serviço clicado.")

            sleep(0.5)

            # Selecionar o oitavo elemento da lista, cuja descrição começa com 'ocultado'
            servico_elemento = WebDriverWait(navegador, 6).until(
                EC.element_to_be_clickable((By.XPATH, '//*[@id[contains(.,"_list")]]//li[8]/div'))
            )
            servico_elemento.click()
            logging.info("")

            sleep(0.5)

            # Preenchimento do valor do serviço
            valor_input = WebDriverWait(navegador, 3).until(
                EC.element_to_be_clickable((By.XPATH, '//*[@id="valorServico"]'))
            )
            valor_input.click()
            valor_input.send_keys(Keys.CONTROL + "a")
            valor_input.send_keys(Keys.DELETE)
            valor_input.send_keys(str(valor_x100))
            logging.info(f"Valor do Serviço preenchido com: {valor_x100}")

            sleep(0.5)
            
            # Preenchimento da descrição do serviço
            desc_input = WebDriverWait(navegador, 3).until(
                EC.element_to_be_clickable((By.XPATH, '//*[@id="descServico"]'))
            )
            desc_input.click()
            desc_input.send_keys(Keys.CONTROL + "a")  # Seleciona todo o conteúdo
            desc_input.send_keys(Keys.DELETE)  # Apaga o conteúdo selecionado
            desc_input.send_keys(procedimento)
            logging.info(f"Discriminação do Serviço preenchida com: {procedimento}")
            
            sleep(1)
            
            def clicar_botao_gravar(max_tentativas=3):

                for tentativa in range(max_tentativas):
                    try:
                        # Tenta clicar no botão Gravar
                        gravar_botao = WebDriverWait(navegador, 6).until(
                            EC.element_to_be_clickable((By.XPATH, '/html/body/div/main/div/div/div/div/div[2]/div[14]/button[3]'))
                        )
                        gravar_botao.click()
                        logging.info(f"Botão Gravar clicado na tentativa {tentativa + 1}")
            
                        # Verifica se a ação foi bem-sucedida esperando o botão Novo aparecer
                        WebDriverWait(navegador, 6).until(
                            EC.visibility_of_element_located((By.XPATH, '/html/body/div/main/div/div/div/div/div[2]/div[14]/button[5]'))
                        )
                        logging.info("Gravação confirmada, botão Novo visível")
                        return True
            
                    except Exception as e:
                        logging.warning(f"Tentativa {tentativa + 1} falhou: {str(e)}")
                        sleep(2)
            
                logging.error("Todas as tentativas de clicar no botão Gravar falharam")
                return False
            
            # Uso da função no fluxo principal
            if clicar_botao_gravar():
                logging.info("Registro gravado com sucesso")
            else:
                logging.error("Falha ao gravar o registro")
                # Aqui você pode adicionar a lógica para lidar com a falha, se necessário
            sleep(3)
            
            novo_botao = WebDriverWait(navegador, 6).until(
                EC.element_to_be_clickable((By.XPATH, '/html/body/div/main/div/div/div/div/div[2]/div[14]/button[5]'))
            )
            novo_botao.click()
            logging.info("Botão Novo clicado.")

            sleep(3)

            WebDriverWait(navegador, 10).until(
                EC.visibility_of_element_located((By.XPATH, '//*[@id="tomador"]'))
            )
            logging.info("Campo CPF apareceu novamente, pronto para o próximo lançamento.")
            
            # Registro processado com sucesso
            salvar_registro_processado(cpf, valor_x100, procedimento)
            logging.info(f"Registro processado com sucesso e salvo: CPF {cpf}, Valor {valor_x100}, Procedimento {procedimento}")
            
            break  # Se chegou até aqui sem erros, sai do loop de tentativas

        except Exception as e:
            logging.error(f"Erro no fluxo (tentativa {tentativa + 1}): {e}")
            if tentativa == 3:  # Se for a última tentativa
                raise Exception(f"Falha após 4 tentativas: {e}")

    else:
        logging.error("Falha ao tentar cadastrar ou emitir NF após várias tentativas. Encerrando o processo para este CPF.")
        raise Exception("Interrupção do processo após várias tentativas falhas.")

# Carregar a tabela de dados
tabela = pd.read_excel(".xlsx")

# Carregar registros já processados
registros_processados = carregar_registros_processados()

# Lista para armazenar registros com erro
registros_com_erro = []

# Loop para percorrer a tabela e preencher o formulário
for i in range(len(tabela)):
    cpf = str(tabela.loc[i, "CPF"])
    valor = tabela.loc[i, "Valor"]
    procedimento = str(tabela.loc[i, "Procedimento"])
    valor_x100 = int(valor) * 100
    
    # Verifica se o registro já foi processado
    registro_atual = (cpf, str(valor_x100), procedimento)
    if registro_atual in registros_processados:
        logging.info(f"Registro já processado anteriormente: CPF {cpf}, Valor {valor_x100}, Procedimento {procedimento}. Pulando.")
        continue
    
    logging.info(f"Processando novo registro: CPF {cpf}, Valor {valor_x100}, Procedimento {procedimento}")
    
    try:
        nome = tabela.loc[i, "Nome completo"]
        cep = tabela.loc[i, "CEP"]
        bairro = tabela.loc[i, "Bairro"]
        tipo = tabela.loc[i, "Tipo Logradouro"]
        endereco = tabela.loc[i, "Logradouro"]
        num = tabela.loc[i, "Número"]
        complemento = tabela.loc[i, "Complemento"]
        email = tabela.loc[i, "E-mail"]
        tel = tabela.loc[i, "Telefone"]

        logging.debug(f"Processando linha {i+1}: CPF={cpf}, Nome={nome}, Valor={valor_x100}, Procedimento={procedimento}")
        
        fluxo_comum(cpf, valor_x100, procedimento, nome, cep, bairro, tipo, endereco, num, complemento, email, tel)
        
        # Adiciona o registro processado ao conjunto
        registros_processados.add(registro_atual)
        
        logging.debug(f"Processamento bem-sucedido para CPF={cpf}, Nome={nome}, Valor={valor_x100}, Procedimento={procedimento}")
    except Exception as e:
        logging.error(f"Erro ao processar registro: CPF {cpf}, Valor {valor_x100}, Procedimento {procedimento}: {str(e)}", exc_info=True)
        registros_com_erro.append({
            "CPF": cpf,
            "Nome": nome,
            "Valor": valor_x100,
            "Procedimento": procedimento,
            "Erro": str(e)
        })
        continue  # Continua com o próximo registro

# Após o loop, salvar o relatório de erros
if registros_com_erro:
    df_erros = pd.DataFrame(registros_com_erro)
    nome_arquivo_erros = "relatorio_registros_com_erro.xlsx"
    df_erros.to_excel(nome_arquivo_erros, index=False)
    logging.info(f"Relatório de registros com erro salvo em {nome_arquivo_erros}")

logging.info("Processamento concluído.")