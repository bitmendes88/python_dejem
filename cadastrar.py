from time import sleep
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.by import By
import pandas as pd
from classes import ConexaoBancoDados, Acessos
from openpyxl import load_workbook

def CadastrarEscalas(cpf, senha, aba, conexao):
    # BUSCA OS DADOS DA PLANILHA GOOGLE E SALVA AS INFORMAÇÕES NAS VARIÁVEIS
    bd = ConexaoBancoDados()
    ac = Acessos()
    bd.Update_certifi()

    cadastro = bd.cadastro
    df_tabela_plano = bd.Escalas_Para_Cadastrar()
    filtro_df_tabela_plano = pd.DataFrame(df_tabela_plano)
    df_info_OPM = bd.df_tabela_info_opm

    if conexao == "VPN":
        # SE FOR SALVAR NUMA PLANILHA OFLINE, COMENTAR geração E USAR OPENPYXL
        pasta_excel = load_workbook(bd.caminho_relativo("BD_GESTAO_DEJEM.xlsx"))
        planilha_excel = pasta_excel.worksheets[0]
        print("Aguardando vpn")
        sleep(40)

    else:
        print("Conexão via Intranet.")

    navegador = ac.inicia_navegador()

    def executa_cadastro():
        for index, row in filtro_df_tabela_plano.iterrows():
            data_inicio = row[f'DATA_ESCALA']
            data_termino = row[f'STR_DATA_TERMINO']
            convenio = row[f'CONVENIO']
            tipo_escala = row[f'TIPO_ESCALA']
            of_sup = row[f'OF_SUP']
            of_int = row[f'OF_INT']
            of_sub = row[f'OF_SUB']
            sgt = row[f'SGT']
            cbsd = row[f'CBSD']
            aisp = row[f'AISP']
            local = row[f'LOCAL']
            opm = row[f'OPM']
            data_limite_incricao = row[f'DATA_LIMITE_INSCRICAO']
            nome_atividade = row[f'NOME_ATIVIDADE']
            tipo_atividade = row[f'TIPO_ATIVIDADE']
            hora_revista = row[f'HORA_REVISTA']
            min_revista = row[f'MINUTO_REVISTA']
            local_revista = row[f'LOCAL_REVISTA']
            epi = row[f'EPI']
            obs = row[f'OBS']

            of_b1 = df_info_OPM[f'ch_b1'].iloc[0]
            of_ch_op = df_info_OPM[f'ch_sec_op'].iloc[0]
            of_scmt = df_info_OPM[f'ch_em'].iloc[0]

            # RECARREGA A PÁGINA
            navegador.refresh()
            sleep(5)

            # INSERE A DATA INÍCIO
            navegador.find_element('xpath','//*[@id="vESCOPRDATINIESC"]').send_keys(data_inicio)
            #sleep(1)

            # INSERE A DATA TÉRMINO
            navegador.find_element('xpath','//*[@id="vESCOPRHORFIMESC"]').click()
            sleep(0.3)
            navegador.find_element('xpath','//*[@id="vESCOPRHORFIMESC"]').send_keys(data_termino)
            sleep(0.3)

            # SELECIONA O TURNO COM BASE NA HORA DA REVISTA
            madrugada = "Madrugada"
            matutino = "Matutino"
            diurno = "Diurno"
            verpertino = "Vespertino"
            noturno = "Noturno"
            if hora_revista < 4:
                navegador.find_element('xpath',f"//*[@id='vESCOPRIDCTRN']/option[contains(text(),'{madrugada}')]").click()
            elif 4 <= hora_revista < 8:
                navegador.find_element('xpath',f"//*[@id='vESCOPRIDCTRN']/option[contains(text(),'{matutino}')]").click()
            elif 8 <= hora_revista < 12:
                navegador.find_element('xpath',f"//*[@id='vESCOPRIDCTRN']/option[contains(text(),'{diurno}')]").click()
            elif 12 <= hora_revista < 18:
                navegador.find_element('xpath',f"//*[@id='vESCOPRIDCTRN']/option[contains(text(),'{verpertino}')]").click()
            elif hora_revista >= 18:
                navegador.find_element('xpath',f"//*[@id='vESCOPRIDCTRN']/option[contains(text(),'{noturno}')]").click()
            sleep(0.3)

            # INSERE O CONVÊNIO
            navegador.find_element('xpath','//*[@id="vCODAJU"]').click()
            sleep(0.3)
            navegador.find_element('xpath','//*[@id="vCODAJU"]').send_keys(convenio)
            sleep(0.3)

            # INSERE O TIPO DA ESCALA
            navegador.find_element('xpath','//*[@id="vTIPESCIDF"]').click()
            sleep(0.3)
            navegador.find_element('xpath','//*[@id="vTIPESCIDF"]').send_keys(tipo_escala)            
            sleep(0.3)

            # DIGITA OF_SUP
            navegador.find_element('xpath','//*[@id="vQTDOFSUP"]').click()
            sleep(0.3)
            navegador.find_element('xpath','//*[@id="vQTDOFSUP"]').send_keys(of_sup)
            sleep(0.3)

            # DIGITA OF_INT
            navegador.find_element('xpath','//*[@id="vQTDOFINT"]').click()
            sleep(0.3)
            navegador.find_element('xpath','//*[@id="vQTDOFINT"]').send_keys(of_int)
            sleep(0.3)

            # DIGITA OF_SUB
            navegador.find_element('xpath','//*[@id="vQTDOFSUB"]').click()
            sleep(0.3)
            navegador.find_element('xpath','//*[@id="vQTDOFSUB"]').send_keys(of_sub)
            sleep(0.3)

            # DIGITA SGT
            navegador.find_element('xpath','//*[@id="vQTDPRASUBTEN"]').click()
            sleep(0.3)
            navegador.find_element('xpath','//*[@id="vQTDPRASUBTEN"]').send_keys(sgt)
            sleep(0.3)

            # DIGITA CB/SD
            navegador.find_element('xpath','//*[@id="vQTDPRACBSD"]').click()
            sleep(0.3)
            navegador.find_element('xpath','//*[@id="vQTDPRACBSD"]').send_keys(cbsd)
            sleep(0.3)

            # DIGITA AISP
            navegador.find_element('xpath','//*[@id="vIDF_AGP_GEO_SST"]').click()
            sleep(0.3)
            navegador.find_element('xpath','//*[@id="vIDF_AGP_GEO_SST"]').send_keys(str(aisp))
            sleep(0.3)

            # DIGITA LOCAL
            navegador.find_element('xpath','//*[@id="vLOCDJMCOD"]').click()
            sleep(0.3)
            navegador.find_element('xpath','//*[@id="vLOCDJMCOD"]').send_keys(local)
            sleep(0.3)

            # DIGITA OPM
            navegador.find_element('xpath','//*[@id="vESCOPROPMRSP"]').click()
            sleep(0.3)
            navegador.find_element('xpath','//*[@id="vESCOPROPMRSP"]').send_keys(opm)
            sleep(0.3)

            # DIGITA DATA LIMITE DE INSCRIÇÃO
            element = WebDriverWait(navegador, 6000).until(
                EC.presence_of_element_located((By.ID,"vESCOPRDATFIMISC"))
            )
            navegador.find_element('xpath','//*[@id="vESCOPRDATFIMISC"]').click()
            sleep(0.3)
            navegador.find_element('xpath','//*[@id="vESCOPRDATFIMISC"]').send_keys(data_limite_incricao)
            sleep(0.3)

            # DIGITA O NOME DA ATIVIDADE
            element = WebDriverWait(navegador, 6000).until(
                EC.presence_of_element_located((By.ID,"vESCOPRDLGATVNOM"))
            )
            atividade = navegador.find_element('xpath','//*[@id="vESCOPRDLGATVNOM"]')
            atividade.click()
            navegador.find_element('xpath','//*[@id="vESCOPRDLGATVNOM"]').send_keys(nome_atividade)
            sleep(0.3)

            # DIGITA O TIPO DE ATIVIDADE
            element = WebDriverWait(navegador, 6000).until(
                EC.presence_of_element_located((By.ID,"vESCATVTIPCOD"))
            )
            tipo = navegador.find_element('xpath','//*[@id="vESCATVTIPCOD"]')
            tipo.click()
            navegador.find_element('xpath','//*[@id="vESCATVTIPCOD"]').send_keys(tipo_atividade)
            sleep(1)

            # DIGITA A HORA DA REVISTA
            horas_map = {
                "00": "00h", "01": "01h", "02": "02h", "03": "03h", "04": "04h", "05": "05h", "06": "06h",
                "07": "07h", "08": "08h", "09": "09h", "10": "10h", "11": "11h", "12": "12h", "13": "13h",
                "14": "14h", "15": "15h", "16": "16h", "17": "17h", "18": "18h", "19": "19h", "20": "20h",
                "21": "21h", "22": "22h", "23": "23h"
            }
            hora_revista_str = str(hora_revista).zfill(2)
            navegador.find_element('xpath', f"//*[@id='vHOR']/option[contains(text(),'{horas_map[hora_revista_str]}')]").click()
            sleep(1)

            # DIGITA O MINUTO DA REVISTA
            minutos_map = {
                0: "00min", 5: "05min", 10: "10min", 15: "15min", 20: "20min",
                25: "25min", 30: "30min", 35: "35min", 40: "40min", 45: "45min",
                50: "50min", 55: "55min"
            }
            minuto_opcao = max(k for k in minutos_map if k <= min_revista)
            navegador.find_element('xpath', f"//*[@id='vMIN']/option[contains(text(),'{minutos_map[minuto_opcao]}')]").click()
            sleep(1)

            # DIGITA O LOCAL DA REVISTA
            navegador.find_element('xpath','//*[@id="vESCOPRLOCAPS"]').click()
            navegador.find_element('xpath','//*[@id="vESCOPRLOCAPS"]').send_keys(local_revista)
            sleep(1)

            # CLICA NO CAMPO EPI
            navegador.find_element('xpath','//*[@id="vEPIOUTROS"]').click()
            sleep(1)

            # DIGITA O EPI
            navegador.find_element('xpath','//*[@id="vEPIOUTROSTEXTO"]').click()
            sleep(0.3)
            navegador.find_element('xpath','//*[@id="vEPIOUTROSTEXTO"]').send_keys(epi)

            # INSERE A OBSERVAÇÃO
            navegador.find_element('xpath','//*[@id="vESCOPRDLGOBS"]').click()
            sleep(0.3)
            navegador.find_element('xpath','//*[@id="vESCOPRDLGOBS"]').send_keys(obs)
            sleep(1)

            # CASO A ESCALA SEJA CADASTRADA COMO EMERGENCIAL, PRECISA SELECIONAR OS OFICIAIS B1, CH OP E SCMT
            if navegador.find_element('xpath', '//*[@id="TEXTBLOCK78"]').is_displayed():
                navegador.find_element('xpath', '//*[@id="vRADIOJUSTDTESCALAREMANEJ"]').click()
                sleep(1)

                # DROPDOWN PARA OFICIAL B1
                navegador.find_element('xpath', '//*[@id="vCOMBOP1"]').click()
                sleep(0.3)                
                dropdown1 = Select(navegador.find_element('xpath', '//*[@id="vCOMBOP1"]'))
                for option in dropdown1.options:
                    if of_b1 in option.text.lower():  # Comparando em minúsculo para evitar problemas de case-sensitive
                        option.click()
                        break
                sleep(0.3)

                # DROPDOWN PARA CH OPERAÇÕES
                navegador.find_element('xpath', '//*[@id="vCOMBOCOORD"]').click()
                sleep(0.3)
                dropdown2 = Select(navegador.find_element('xpath', '//*[@id="vCOMBOCOORD"]'))
                for option in dropdown2.options:
                    if of_ch_op in option.text.lower():
                        option.click()
                        break
                sleep(0.3)

                # DROPDOWN CH EM
                navegador.find_element('xpath', '//*[@id="vCOMBOSUCMD"]').click()
                sleep(0.3)
                dropdown3 = Select(navegador.find_element('xpath', '//*[@id="vCOMBOSUCMD"]'))
                for option in dropdown3.options:
                    if of_scmt in option.text.lower():
                        option.click()
                        break
                sleep(0.3)

                # BOTÃO PARA INCLUIR A ESCALA
                navegador.find_element('xpath','//*[@id="TABLE4"]/tbody/tr[7]/td/input[3]').click()
                sleep(5)
            else:        
                navegador.find_element('xpath','//*[@id="TABLE4"]/tbody/tr[7]/td/input[3]').click()
                sleep(5)

    def busca_escala_cadastradas():
        # PESQUISA AS ESCALAS CADASTRADAS E CARREGA NA PLANILHA SHEETS
        aisp = filtro_df_tabela_plano['AISP'].iloc[0]
        navegador.find_element(By.XPATH,'//*[@id="vIDFAGPGEOSST"]').send_keys(str(aisp))
        sleep(0.5)
        navegador.find_element(By.XPATH,'//*[@id="IMAGE4"]').click()

        element = WebDriverWait(navegador, 6000).until(
            EC.presence_of_element_located((By.XPATH,'//*[@id="Grid1ContainerRow_0001"]'))
        )

        linhas = navegador.find_elements(By.XPATH, '//*[@id="Grid1ContainerTbl"]/tbody/tr[position() >= 2]')

        dados_tabela = []

        for linha in linhas:
            colunas = linha.find_elements(By.TAG_NAME, "td")
            dados_linha = []
            
            for i, coluna in enumerate(colunas[:9]):
                try:
                    span = coluna.find_element(By.TAG_NAME, "span")
                    dados_linha.append(span.text)
                except:
                    dados_linha.append("")
            
            dados_tabela.append(dados_linha)

        # ATUALIZA O SHEETS COM AS ESCALAS CADASTRADAS
        df_escalas_cadastradas = pd.DataFrame(dados_tabela)
        df_escalas_cadastradas.rename(columns={df_escalas_cadastradas.columns[2]: 'DATA_ESCALA'}, inplace=True)
        df_escalas_cadastradas_filtro = pd.merge(filtro_df_tabela_plano, df_escalas_cadastradas, on='DATA_ESCALA', how='inner')
        df_escalas_cadastradas_filtro.reset_index(inplace=True)
        colunas_id_dataini_datalim = [0, 'DATA_ESCALA', 4]
        df_colunas_interesse = df_escalas_cadastradas_filtro[colunas_id_dataini_datalim]
        lista_escaldas_cadastradas = df_colunas_interesse.values.tolist()
        if conexao == "VPN":
            for linha in lista_escaldas_cadastradas:
                planilha_excel.append(linha)
            pasta_excel.save(bd.caminho_relativo("BD_GESTAO_DEJEM.xlsx"))
        else:
            cadastro.append_rows(lista_escaldas_cadastradas)

    ac.acessa_aba_procedimentos(lg=cpf, sn=senha, opcao=aba)
    executa_cadastro()
    busca_escala_cadastradas()
    
    navegador.quit()

#bd = ConexaoBancoDados()
#bd.Update_certifi()
#CadastrarEscalas()