from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import pandas as pd
import pyautogui as gui
from urllib.parse import urlparse, parse_qs
from time import sleep
from classes import ConexaoBancoDados, Acessos

def Processo_SEI(cpf, senha):
    # INSTANCIA A CLASSE DA CONEXÃO COM O BANCO
    bd = ConexaoBancoDados()
    ac = Acessos()
    bd.Update_certifi()
    
    # FILTRA AS ESCALAS PARA MONTAR
    sei = bd.SEI
    lista_escalas_sem_processo = bd.Escalas_Para_SEI()
    lista_escalas_com_processo = bd.Escalas_Para_SEI_com_processo()

    df_escalas_sem_processo = pd.DataFrame(lista_escalas_sem_processo)
    df_escalas_com_processo = pd.DataFrame(lista_escalas_com_processo)

    navegador = ac.inicia_navegador()
   
    def insere_documentos():        
        #SELECIONA MESA SEI CONFORME A PRIMEIRA LINHA DO DATAFRAME
        navegador.maximize_window()

        #VARIÁVEL PARA CAPTURAR A MESA SEI DA PRIMEIRA LINHA
        mesa_sei_atual = lista_escalas_sem_processo.loc[0, "MESA_SEI"]

        #EXECUTA FUNÇÃO PARA SELECIONAR A MESA SEI
        ac.seleciona_mesa_SEI(mesa_sei=mesa_sei_atual)
        
        # VARIÁVEL VAZIA PARA COMPARAR A DATA DA ESCALA ANTERIOR COM A ATUAL
        data_linha_anterior = None

        wait = WebDriverWait(navegador, 100)

        # EXECUTA O CADASTRO DA ESCALA PARA CADA LINHA DO DATAFRAME
        for index, linha in df_escalas_sem_processo.iterrows(): 
            # BUSCA DADOS DAS LINHAS
            id_escala = (linha['ID_ESCALA'])
            nome_escala = (linha['NOME_ESCALAS'])
            data_escala = (linha['DATA_ESCALA'])
            mesa_sei_linha = (linha['MESA_SEI'])
            data = (linha['DDMM'])
            hora = (linha['hhHmm'])
            tipo_processo = f'Atestado de Frequência'
            dia = data[:2]
            meses = {'01': 'JAN', '02': 'FEV', '03': 'MAR', '04': 'ABR', '05': 'MAI', '06': 'JUN', '07': 'JUL', '08': 'AGO', '09': 'SET', '10': 'OUT', '11': 'NOV', '12': 'DEZ'}
            n_mes = data[-2:]
            mes = meses.get(n_mes, 'Mês Invalido')
            ano = data_escala[8:10]
            especificacao = f'Relatório de Presença de Escala DEJEM | {dia}{mes}{ano}'
            tipo_documento = f'Escala de Serviço'
            descricao_doc = f'Relatório de Presença de Escala DEJEM ref. ID Nº {id_escala}'
            caminho_pasta_escala = 'C:\\AUTOMACAO_DEJEM\\ESCALAS_PNG'
            sleep(1)

            # SE MESA SEI FOR OUTRA, TROCAR DE MESA
            if mesa_sei_linha != mesa_sei_atual:
                ac.seleciona_mesa_SEI(mesa_sei=mesa_sei_atual)
            else:
                pass

            filtro_processos_cadastrados = (df_escalas_com_processo['DDMM'] == data)&(df_escalas_com_processo["MESA_SEI"] == mesa_sei_linha)

            # SE EXISTIR UM PROCESSO COM A MESMA DATA E MESA, CAPTURA O LINK, ABRE O PROCESSO E INSERE O DOCUMENTO
            if filtro_processos_cadastrados.any():
                # BUSCA LINHA DO PROCESSO EXISTENTE
                linha_escala_existente = df_escalas_com_processo.loc[filtro_processos_cadastrados]

                if linha_escala_existente.empty:
                    print("⚠ Nenhuma escala existente encontrada com o filtro aplicado.")
                    continue

                link_escala_existente = linha_escala_existente["LINK_SEI"].values[0]

                # EXTRAI O ID DO PROCESSO
                parse_link = urlparse(link_escala_existente)
                param = parse_qs(parse_link.query)

                id_processo = param.get("id_procedimento", [None])[0]
                if id_processo is None:
                    print("⚠ ID do procedimento não encontrado na URL!")
                    continue

                # ABRE O PROCESSO EXISTENTE
                link_processo = f'https://sei.sp.gov.br/sei/controlador.php?acao=procedimento_trabalhar&id_procedimento={id_processo}'
                navegador.get(link_processo)

                wait.until(EC.presence_of_element_located((By.XPATH, "//*[@id='ifrVisualizacao']")))
                navegador.switch_to.frame("ifrVisualizacao")

                # CLIQUE NO NOVO DOCUMENTO
                incluir_documento = wait.until(
                    EC.presence_of_element_located((By.XPATH, "//img[@alt='Incluir Documento']"))
                )
                incluir_documento.click()
                sleep(3)

                # ESCOLHE O TIPO DE DOCUMENTO
                navegador.find_element(By.XPATH, '//*[@id="txtFiltro"]').click()
                sleep(1)
                navegador.find_element(By.XPATH, '//*[@id="txtFiltro"]').send_keys(tipo_documento)
                sleep(0.5)
                gui.press('tab')
                sleep(0.5)
                gui.press('enter')
                sleep(2)

                # PREENCHIMENTO DOS CAMPOS
                navegador.find_element(By.XPATH, '//*[@id="txtDescricao"]').send_keys(descricao_doc)
                sleep(1)
                navegador.find_element(By.XPATH, '//*[@id="txtNomeArvore"]').send_keys(hora)
                sleep(1)
                navegador.find_element(By.XPATH, '//*[@id="divOptPublico"]/div/label').click()
                sleep(1)
                navegador.find_element(By.XPATH, '//*[@id="btnSalvar"]').click()
                sleep(2)

                # VERIFICA SE A ABA DE EDIÇÃO FOI ABERTA
                if len(navegador.window_handles) < 2:
                    print("⚠ Aba de edição do documento não foi aberta!")
                    continue

                edita_doc = navegador.window_handles[1]
                navegador.switch_to.window(edita_doc)

                wait.until(EC.presence_of_all_elements_located((By.XPATH, "/html/body")))

                # APAGA INFORMAÇÕES DO MODELO
                for _ in range(3): gui.press('tab'); sleep(0.2)
                for _ in range(3): gui.press('delete'); sleep(0.2)
                gui.press('down'); sleep(0.2)
                gui.press('end'); sleep(0.2)
                for _ in range(43): gui.press('backspace'); sleep(0.02)
                gui.press('down'); sleep(0.2)
                gui.press('end'); sleep(0.2)
                gui.press('delete'); sleep(0.2)
                gui.press('tab'); sleep(0.2)
                gui.hotkey('ctrl', 'a'); sleep(0.2)
                gui.press('delete'); sleep(0.2)

                # INSERE A ESCALA
                navegador.find_element(By.XPATH, '//*[@id="cke_235"]/span[1]').click()
                sleep(2)
                gui.press('enter')
                sleep(3)
                gui.hotkey('ctrl','l')
                sleep(0.5)
                gui.write(caminho_pasta_escala)
                sleep(1)
                gui.press('enter')
                sleep(1)
                gui.hotkey('ctrl','f')
                sleep(1)
                gui.write(nome_escala)
                sleep(1)
                for _ in range(3): gui.press('tab'); sleep(0.3)
                gui.press('space')
                sleep(1)
                gui.press('enter')
                sleep(3)
                gui.press('tab'); sleep(0.5)
                gui.press('enter')
                sleep(2)
                gui.hotkey('ctrl','alt','s')
                sleep(5)

                # EXTRAI A NOVA URL
                url_doc = navegador.current_url
                parsed_url = urlparse(url_doc)
                query = parse_qs(parsed_url.query)

                id_procedimento = query.get("id_procedimento", [None])[0]
                id_documento = query.get("id_documento", [None])[0]

                if not id_procedimento or not id_documento:
                    print("⚠ ID do documento ou do procedimento não foram encontrados!")
                    continue

                link_sei = f'https://sei.sp.gov.br/sei/controlador.php?acao=procedimento_trabalhar&id_procedimento={id_procedimento}&id_documento={id_documento}'
                lista_sei = [id_escala, link_sei]
                sei.append_row(lista_sei)

                # FECHAMENTO E VOLTA PARA ABA PRINCIPAL
                data_linha_anterior = data
                mesa_sei_atual = mesa_sei_linha
                navegador.close()
                navegador.switch_to.window(navegador.window_handles[0])


            # SE A ESCALA É PARA MESMA DATA E MESA DA ESCALA ANTERIOR
            elif data == data_linha_anterior and mesa_sei_linha == mesa_sei_atual:
                navegador.refresh()
                sleep(5)
                # ENTRA NO IFRAME QUE CONTÉM OS MENUS DE DOCUMENTOS
                navegador.switch_to.frame("ifrVisualizacao")

                # ESPERA O ICONE "NOVO DOCUMENTO APARECER E CLICA NELE"
                incluir_documento = wait.until(
                    EC.presence_of_element_located((By.XPATH, "//img[@alt='Incluir Documento']"))
                )
                incluir_documento.click()
                sleep(3)

                # SELECIONA TIPO DE PROCESSO COMO "Escala de Serviço"
                navegador.find_element('xpath','//*[@id="txtFiltro"]').click()
                sleep(1)
                navegador.find_element('xpath','//*[@id="txtFiltro"]').send_keys(tipo_documento)
                sleep(0.5)                    
                gui.press('tab')
                sleep(0.5)
                gui.press('enter')
                sleep(2)

                # PREENCHE OS DADOS DO DOCUMENTO
                navegador.find_element('xpath','//*[@id="txtDescricao"]').send_keys(descricao_doc)
                sleep(1)
                navegador.find_element('xpath','//*[@id="txtNomeArvore"]').send_keys(hora)
                sleep(1)
                navegador.find_element('xpath','//*[@id="divOptPublico"]/div/label').click()
                sleep(1)
                navegador.find_element('xpath','//*[@id="btnSalvar"]').click()
                sleep(10)

                # ABRE NOVA JANELA PARA EDIÇÃO DO DOCUMENTO SEI
                edita_doc = navegador.window_handles[1]
                navegador.switch_to.window(edita_doc)

                # APAGA AS INFORMAÇÕES DO DOCUMENTO MODELO
                for _ in range(3):
                    gui.press('tab')
                sleep(0.5)
                for _ in range(3):
                    gui.press('delete')
                sleep(0.5)
                gui.press('down')
                sleep(0.5)
                gui.press('end')
                sleep(0.5)
                for _ in range(43):
                    gui.press('backspace')
                sleep(0.5)
                gui.press('down')
                sleep(0.5)
                sleep(0.5)
                gui.press('end')
                sleep(0.5)
                gui.press('delete')

                for _ in range(3):
                    gui.press('delete')
                sleep(0.5)
                gui.press('tab')
                sleep(0.5)
                gui.hotkey('ctrl', 'a')
                sleep(0.5)
                gui.press('delete')
                sleep(0.5)
                gui.press('delete')
                sleep(0.5)

                # INSERE A IMAGEM DA ESCALA
                navegador.find_element('xpath','//*[@id="cke_235"]/span[1]').click()
                sleep(2)
                gui.press('enter')
                sleep(3)
                gui.hotkey('ctrl','l')
                sleep(0.5)
                gui.write(caminho_pasta_escala)
                sleep(1)
                gui.press('enter')
                sleep(1)
                gui.hotkey('ctrl','f')
                sleep(1)
                gui.write(nome_escala)
                sleep(1)
                for _ in range(3):
                    gui.press('tab')
                sleep(1)
                gui.press('space')
                sleep(1)
                gui.press('enter')
                sleep(3)
                gui.press('tab')
                sleep(0.5)
                gui.press('enter')
                sleep(2)
                gui.hotkey('ctrl','alt','s')
                sleep(5)
                url_doc = navegador.current_url
                parsed_url = urlparse(url_doc)
                
                # EXTRAI APENAS A QUERY DA URL
                query = parse_qs(parsed_url.query)

                # OS VALORES DAS QUERY SÃO OS NÚMEROS DE PROCESSO E DOCUMENTO
                id_procedimento = query.get("id_procedimento", [None])[0]
                id_documento = query.get("id_documento", [None])[0]

                if not id_procedimento or not id_documento:
                    print("⚠ ID do documento ou do procedimento não foram encontrados!")
                    continue

                # SALVA LINK DO DOCUMENTO NA PLANILHA SHEETS
                link_sei = f'https://sei.sp.gov.br/sei/controlador.php?acao=procedimento_trabalhar&id_procedimento={id_procedimento}&id_documento={id_documento}'
                lista_sei = [id_escala,link_sei]
                sei.append_row(lista_sei)
                data_linha_anterior = data
                mesa_sei_atual = mesa_sei_linha

                # FECHA A PÁGINA DE EDIÇÃO
                navegador.close()

                # RETORNA A PÁGINA DO PROCESSO
                processo = navegador.window_handles[0]
                navegador.switch_to.window(processo)

            # SE NÃO EXISTE A ESCALA E NÃO ESTÁ NO PROCESSO, CADASTRAR NOVO PROCESSO.
            else: 
                navegador.refresh()
                sleep(5)
                if navegador.find_element('xpath', '//*[@id="divInfraSidebarMenu"]').is_displayed():
                    sleep(1)
                    navegador.find_element('xpath', '/html/body/div[1]/div/div[1]/div/ul/li[10]').click()
                else:
                    navegador.find_element('xpath', '//*[@id="divInfraBarraSistemaPadraoD"]/div[1]').click()
                    sleep(1)
                    navegador.find_element('xpath', '/html/body/div[1]/div/div[1]/div/ul/li[10]').click()
                sleep(5)
                # SELECIONA O TIPO COMO "Atestado de Frequência"
                tipo_do_processo = wait.until(
                    EC.presence_of_element_located((By.XPATH, '//*[@id="txtFiltro"]'))
                )
                tipo_do_processo.click()                    
                navegador.find_element('xpath','//*[@id="txtFiltro"]').send_keys(tipo_processo)
                sleep(1)
                gui.press('down')
                gui.press('enter')
                #navegador.find_element('xpath','//*[@id="tblTipoProcedimento"]/tbody/tr[5]').click()
                sleep(2)
                navegador.find_element('xpath','//*[@id="txtDescricao"]').click()
                sleep(0.5)
                navegador.find_element('xpath','//*[@id="txtDescricao"]').send_keys(especificacao)
                sleep(0.5)
                navegador.find_element('xpath','//*[@id="divOptPublico"]/div/label').click()
                sleep(0.5)
                navegador.find_element('xpath','//*[@id="btnSalvar"]').click()
                sleep(2)
                
                # ENTRA NO IFRAME QUE CONTÉM OS MENUS DE DOCUMENTOS
                navegador.switch_to.frame("ifrVisualizacao")

                # DEFINIÇÃO DO MARCADOR
                tag_dejem = wait.until(
                    EC.presence_of_element_located((By.XPATH, "//img[@alt='Gerenciar Marcador']"))
                )
                tag_dejem.click()
                sleep(1)

                def encontrar_posicao_label(navegador, menu_xpath, lista_xpath, texto_procurado):
                    # Clique no menu suspenso
                    menu_element = WebDriverWait(navegador, 10).until(
                        EC.element_to_be_clickable((By.XPATH, menu_xpath))
                    )
                    menu_element.click()

                    # Pausa para garantir o carregamento dos itens (ajuste conforme necessário)
                    sleep(2)

                    # Liste todos os itens do menu suspenso
                    elementos = navegador.find_elements(By.XPATH, lista_xpath)
                    print(f"Total de elementos encontrados: {len(elementos)}")

                    for index, elemento in enumerate(elementos, start=1):
                        try:
                            # Localize o <a> dentro do li e então o label
                            a_element = elemento.find_element(By.TAG_NAME, "a")
                            label_element = a_element.find_element(By.TAG_NAME, "label")
                            label_text = label_element.text.strip().lower()
                            print(f"Elemento {index}: {label_text}")  # Exibe o texto do label
                            
                            if label_text == texto_procurado:
                                return index  # Retorna a posição encontrada
                        except Exception as e:
                            print(f"Erro ao verificar elemento na posição {index}: {e}")
                    
                    return None  # Retorna None se o texto não for encontrado
                
                # XPath para o menu suspenso e itens
                menu_xpath = '//*[@id="selMarcador"]'
                lista_xpath = '//*[@id="selMarcador"]/ul/li'  # Alterado para capturar <li> corretamente

                # Texto a ser encontrado
                texto_procurado = "dejem | relatório de presença de escala"
                posicao = encontrar_posicao_label(navegador, menu_xpath, lista_xpath, texto_procurado)

                if posicao:
                    print(f"O item '{texto_procurado}' está na posição li[{posicao}].")

                    xpath_para_click = f'//*[@id="selMarcador"]/ul/li[{posicao}]/a/label'
                    navegador.find_element('xpath', xpath_para_click).click()

                else:
                    print(f"O item '{texto_procurado}' não foi encontrado.")
                sleep(2)
                navegador.find_element('xpath','//*[@id="sbmSalvar"]').click()
                
                sleep(1)
                navegador.refresh()
                sleep(8)


                # ENTRA NOVAMENTE NO IFRAME QUE CONTÉM OS MENUS DE DOCUMENTOS
                navegador.switch_to.frame("ifrVisualizacao")

                # ESPERA O ICONE "NOVO DOCUMENTO APARECER E CLICA NELE"
                incluir_documento = wait.until(
                    EC.presence_of_element_located((By.XPATH, "//img[@alt='Incluir Documento']"))
                )
                incluir_documento.click()
                sleep(3)

                # SELECIONA TIPO DE PROCESSO COMO "Escala de Serviço"
                navegador.find_element('xpath','//*[@id="txtFiltro"]').click()
                sleep(1)
                navegador.find_element('xpath','//*[@id="txtFiltro"]').send_keys(tipo_documento)
                sleep(0.5)                    
                gui.press('tab')
                sleep(0.5)
                gui.press('enter')
                sleep(2)

                # PREENCHE OS DADOS DO DOCUMENTO
                navegador.find_element('xpath','//*[@id="txtDescricao"]').send_keys(descricao_doc)
                sleep(1)
                navegador.find_element('xpath','//*[@id="txtNomeArvore"]').send_keys(hora)
                sleep(1)
                navegador.find_element('xpath','//*[@id="divOptPublico"]/div/label').click()
                sleep(1)
                navegador.find_element('xpath','//*[@id="btnSalvar"]').click()
                sleep(10)

                # ABRE NOVA JANELA PARA EDIÇÃO DO DOCUMENTO SEI
                edita_doc = navegador.window_handles[1]
                navegador.switch_to.window(edita_doc)

                # APAGA AS INFORMAÇÕES DO DOCUMENTO MODELO
                for _ in range(3):
                    gui.press('tab')
                sleep(0.5)
                for _ in range(3):
                    gui.press('delete')
                sleep(0.5)
                gui.press('down')
                sleep(0.5)
                gui.press('end')
                sleep(0.5)
                for _ in range(43):
                    gui.press('backspace')
                sleep(0.5)
                gui.press('down')
                sleep(0.5)
                sleep(0.5)
                gui.press('end')
                sleep(0.5)
                gui.press('delete')
                sleep(0.5)
                gui.press('tab')
                sleep(0.5)
                gui.hotkey('ctrl', 'a')
                sleep(0.5)
                gui.press('delete')
                sleep(0.5)
                gui.press('delete')
                sleep(0.5)

                # INSERE A IMAGEM DA ESCALA
                navegador.find_element('xpath','//*[@id="cke_235"]/span[1]').click()
                sleep(2)
                gui.press('enter')
                sleep(3)
                gui.hotkey('ctrl','l')
                sleep(0.5)
                gui.write(caminho_pasta_escala)
                sleep(1)
                gui.press('enter')
                sleep(1)
                gui.hotkey('ctrl','f')
                sleep(1)
                gui.write(nome_escala)
                sleep(1)
                for _ in range(3):
                    gui.press('tab')
                sleep(1)
                gui.press('space')
                sleep(1)
                gui.press('enter')
                sleep(3)
                gui.press('tab')
                sleep(0.5)
                gui.press('enter')
                sleep(2)
                gui.hotkey('ctrl','alt','s')
                sleep(5)
                url_doc = navegador.current_url
                parsed_url = urlparse(url_doc)

                # EXTRAI APENAS A QUERY DA URL
                param_if = parse_qs(parsed_url.query)

                # OS VALORES DAS QUERY SÃO OS NÚMEROS DE PROCESSO E DOCUMENTO
                id_procedimento = param_if.get("id_procedimento", [None])[0]
                id_documento = param_if.get("id_documento", [None])[0]

                if not id_procedimento or not id_documento:
                    print("⚠ ID do documento ou do procedimento não foram encontrados!")
                    continue


                # SALVA LINK NO BANCO NA PLANILHA SHEETS
                link_sei = f'https://sei.sp.gov.br/sei/controlador.php?acao=procedimento_trabalhar&id_procedimento={id_procedimento}&id_documento={id_documento}'
                lista_sei = [id_escala,link_sei]
                sei.append_row(lista_sei)
                data_linha_anterior = data
                mesa_sei_atual = mesa_sei_linha

                # FECHA A PÁGINA DE EDIÇÃO
                navegador.close()

                # RETORNA A PÁGINA DO PROCESSO
                processo = navegador.window_handles[0]
                navegador.switch_to.window(processo)


    ac.acessa_SEI(lg=cpf,sn=senha)
    insere_documentos()

    navegador.quit()