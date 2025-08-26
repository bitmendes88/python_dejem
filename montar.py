from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
import pandas as pd
from classes import ConexaoBancoDados, Acessos
from time import sleep
from openpyxl import load_workbook

def MontarEscalas(cpf, senha, aba, conexao, gb):
    # ATUALIZA O CERTIFICADO SSL
    bd = ConexaoBancoDados(gb)
    ac = Acessos(gb)
    bd.Update_certifi()
    # BUSCA FUNÇÃO QUE FILTRA AS ESCALAS PARA MONTAR
    geracao = bd.geracao
    df_lista_escalas = bd.Escalas_para_Montar()

    if conexao == "VPN":
        pasta_excel = load_workbook(bd.caminho_relativo("BD_GESTAO_DEJEM.xlsx"))
        planilha_excel = pasta_excel.worksheets[0]
        print("Aguardando VPN")
        sleep(40)

    else:
        print("Conexão via Intranet.")

    navegador = ac.inicia_navegador()

    def executa_montagem():
        # CRIA NOVO DATAFRAME VAZIO PARA GUARDAR OS NOMES DOS ESCALADOS
        lista_colunas = ['ID_ESCALA', 'MILITAR01', 'MILITAR02', 'MILITAR03', 'MILITAR04', 'MILITAR05', 'MILITAR06', 'MILITAR07', 'MILITAR08', 'MILITAR09']
        df_escalados = pd.DataFrame(columns=lista_colunas)
        wait = WebDriverWait(navegador, timeout=1000)

        # PECORRE CADA LINHA DO DATAFRAME RETORNADO DE "Escalas_para_Montar"
        for index, row in df_lista_escalas.iterrows():
            id_escala = row[f'ID_ESCALA']
            navegador.refresh()
            navegador.find_element('xpath','//*[@id="vESCOPRIDF"]').send_keys(id_escala)
            sleep(0.5)
            navegador.find_element('xpath','//*[@id="IMAGE2"]').click()
            sleep(0.5)
            botao_gerar_automaticamente = wait.until(EC.element_to_be_clickable((By.NAME,"MONESCAUT1")))
            botao_gerar_automaticamente.click()
            sleep(0.5)
            wait.until(EC.element_to_be_clickable(('xpath','//*[@id="gxErrorViewer"]/div')))
            efetivo_total = navegador.find_element(By.ID,"span_vESCOPRQTDTOT").text
            efetivo_escalado = navegador.find_element(By.ID,"span_vTOTESC").text
            sleep(0.5)
            navegador.find_element('xpath','//*[@id="PROCESSOS"]/tbody/tr[3]/td/span/span/span/span/input').click()
            sleep(0.5)
            if efetivo_escalado == "0":
                pass
            else:                
                btn_retornar = wait.until(EC.element_to_be_clickable(('xpath','//*[@id="TABLE1"]/tbody/tr/td[3]/span/span/span/span/input')))
            if efetivo_escalado == "0":
                if efetivo_total == "1":
                    lista01 = [id_escala, 'SEM INSCRITOS']
                    linha01 = lista01 + [''] * (len(lista_colunas) - len(lista01))
                    df_escalados = pd.DataFrame([linha01])
                elif efetivo_total == "2":
                    lista02 = [id_escala, 'SEM INSCRITOS', 'SEM INSCRITOS']
                    linha02 = lista02 + [''] * (len(lista_colunas) - len(lista02))
                    df_escalados = pd.DataFrame([linha02])
                elif efetivo_total == "3":
                    lista03 = [id_escala, 'SEM INSCRITOS', 'SEM INSCRITOS', 'SEM INSCRITOS']
                    linha03 = lista03 + [''] * (len(lista_colunas) - len(lista03))
                    df_escalados = pd.DataFrame([linha03])
                elif efetivo_total == "4":
                    lista04 = [id_escala, 'SEM INSCRITOS', 'SEM INSCRITOS', 'SEM INSCRITOS', 'SEM INSCRITOS']
                    linha04 = lista04 + [''] * (len(lista_colunas) - len(lista04))
                    df_escalados = pd.DataFrame([linha04])
                elif efetivo_total == "5":
                    lista05 = [id_escala, 'SEM INSCRITOS', 'SEM INSCRITOS', 'SEM INSCRITOS', 'SEM INSCRITOS', 'SEM INSCRITOS']
                    linha05 = lista05 + [''] * (len(lista_colunas) - len(lista05))
                    df_escalados = pd.DataFrame([linha05])
                elif efetivo_total == "6":
                    lista06 = [id_escala, 'SEM INSCRITOS', 'SEM INSCRITOS', 'SEM INSCRITOS', 'SEM INSCRITOS', 'SEM INSCRITOS', 'SEM INSCRITOS']
                    linha06 = lista06 + [''] * (len(lista_colunas) - len(lista06))
                    df_escalados = pd.DataFrame([linha06])
                elif efetivo_total == "7":
                    lista07 = [id_escala, 'SEM INSCRITOS', 'SEM INSCRITOS', 'SEM INSCRITOS', 'SEM INSCRITOS', 'SEM INSCRITOS', 'SEM INSCRITOS', 'SEM INSCRITOS']
                    linha07 = lista07 + [''] * (len(lista_colunas) - len(lista07))
                    df_escalados = pd.DataFrame([linha07])
                elif efetivo_total == "8":
                    lista08 = [id_escala, 'SEM INSCRITOS', 'SEM INSCRITOS', 'SEM INSCRITOS', 'SEM INSCRITOS', 'SEM INSCRITOS', 'SEM INSCRITOS', 'SEM INSCRITOS', 'SEM INSCRITOS']
                    linha08 = lista08 + [''] * (len(lista_colunas) - len(lista08))
                    df_escalados = pd.DataFrame([linha08])
                elif efetivo_total == "9":
                    lista09 = [id_escala, 'SEM INSCRITOS', 'SEM INSCRITOS', 'SEM INSCRITOS', 'SEM INSCRITOS', 'SEM INSCRITOS', 'SEM INSCRITOS', 'SEM INSCRITOS', 'SEM INSCRITOS', 'SEM INSCRITOS']
                    linha09 = lista09 + [''] * (len(lista_colunas) - len(lista09))
                    df_escalados = pd.DataFrame([linha09])
                lista_sem_inscritos = df_escalados.values.tolist()
                if conexao == "VPN":
                    for dados_linha in lista_sem_inscritos:
                        planilha_excel.append(dados_linha)
                    pasta_excel.save(bd.caminho_relativo("DB_GESTÃO_DEJEM.xlsx"))
                else:
                    geracao.append_rows(lista_sem_inscritos, value_input_option='USER_ENTERED')
                navegador.refresh()
                sleep(5)
            elif efetivo_escalado == "1":           
                re1 = navegador.find_element(By.ID,"span_vVW_MOT_ESCEFE_ESCEFERE_0001").text
                posto1 = navegador.find_element(By.ID,"span_vGRIDPOSDES2_0001").text
                militar1 = navegador.find_element(By.ID,"span_vGRIDPMNOM2_0001").text
                opm1 = navegador.find_element(By.ID,"span_vGRIDPMOPMDES2_0001").text
                escalado1 = f"{posto1} {re1} {militar1} {opm1}"
                if efetivo_total == "1":
                    lista11 = [id_escala, escalado1]
                    linha11 = lista11 + [''] * (len(lista_colunas) - len(lista11))
                    df_escalados = pd.DataFrame([linha11])
                elif efetivo_total == "2":
                    lista12 = [id_escala, escalado1, 'SEM INSCRITOS']
                    linha12 = lista12 + [''] * (len(lista_colunas) - len(lista12))
                    df_escalados = pd.DataFrame([linha12])
                elif efetivo_total == "3":
                    lista13 = [id_escala, escalado1, 'SEM INSCRITOS', 'SEM INSCRITOS']
                    linha13 = lista13 + [''] * (len(lista_colunas) - len(lista13))
                    df_escalados = pd.DataFrame([linha13])
                elif efetivo_total == "4":
                    lista14 = [id_escala, escalado1, 'SEM INSCRITOS', 'SEM INSCRITOS', 'SEM INSCRITOS']
                    linha14 = lista14 + [''] * (len(lista_colunas) - len(lista14))
                    df_escalados = pd.DataFrame([linha14])
                elif efetivo_total == "5":
                    lista15 = [id_escala, escalado1, 'SEM INSCRITOS', 'SEM INSCRITOS', 'SEM INSCRITOS', 'SEM INSCRITOS']
                    linha15 = lista15 + [''] * (len(lista_colunas) - len(lista15))
                    df_escalados = pd.DataFrame([linha15])
                elif efetivo_total == "6":
                    lista16 = [id_escala, escalado1, 'SEM INSCRITOS', 'SEM INSCRITOS', 'SEM INSCRITOS', 'SEM INSCRITOS', 'SEM INSCRITOS']
                    linha16 = lista16 + [''] * (len(lista_colunas) - len(lista16))
                    df_escalados = pd.DataFrame([linha16])
                elif efetivo_total == "7":
                    lista17 = [id_escala, escalado1, 'SEM INSCRITOS', 'SEM INSCRITOS', 'SEM INSCRITOS', 'SEM INSCRITOS', 'SEM INSCRITOS', 'SEM INSCRITOS']
                    linha17 = lista17 + [''] * (len(lista_colunas) - len(lista17))
                    df_escalados = pd.DataFrame([linha17])
                elif efetivo_total == "8":
                    lista18 = [id_escala, escalado1, 'SEM INSCRITOS', 'SEM INSCRITOS', 'SEM INSCRITOS', 'SEM INSCRITOS', 'SEM INSCRITOS', 'SEM INSCRITOS', 'SEM INSCRITOS']
                    linha18 = lista18 + [''] * (len(lista_colunas) - len(lista18))
                    df_escalados = pd.DataFrame([linha18])
                elif efetivo_total == "9":
                    lista19 = [id_escala, escalado1, 'SEM INSCRITOS', 'SEM INSCRITOS', 'SEM INSCRITOS', 'SEM INSCRITOS', 'SEM INSCRITOS', 'SEM INSCRITOS', 'SEM INSCRITOS', 'SEM INSCRITOS']
                    linha19 = lista19 + [''] * (len(lista_colunas) - len(lista19))
                    df_escalados = pd.DataFrame([linha19])
                lista_1escalado = df_escalados.values.tolist()
                if conexao == "VPN":
                    for dados_linha1 in lista_1escalado:
                        planilha_excel.append(dados_linha1)
                    pasta_excel.save(bd.caminho_relativo("DB_GESTÃO_DEJEM.xlsx"))
                else:
                    geracao.append_rows(lista_1escalado, value_input_option='USER_ENTERED')
                btn_retornar.click()
                sleep(5)
            elif efetivo_escalado == "2":
                re1 = navegador.find_element(By.ID,"span_vVW_MOT_ESCEFE_ESCEFERE_0001").text
                posto1 = navegador.find_element(By.ID,"span_vGRIDPOSDES2_0001").text
                militar1 = navegador.find_element(By.ID,"span_vGRIDPMNOM2_0001").text
                opm1 = navegador.find_element(By.ID,"span_vGRIDPMOPMDES2_0001").text
                escalado1 = f"{posto1} {re1} {militar1} {opm1}"
                re2 = navegador.find_element(By.ID,"span_vVW_MOT_ESCEFE_ESCEFERE_0002").text
                posto2 = navegador.find_element(By.ID,"span_vGRIDPOSDES2_0002").text
                militar2 = navegador.find_element(By.ID,"span_vGRIDPMNOM2_0002").text
                opm2 = navegador.find_element(By.ID,"span_vGRIDPMOPMDES2_0002").text
                escalado2 = f"{posto2} {re2} {militar2} {opm2}"
                if efetivo_total == "2":
                    lista22 = [id_escala, escalado1, escalado2]
                    linha22 = lista22 + [''] * (len(lista_colunas) - len(lista22))
                    df_escalados = pd.DataFrame([linha22])
                elif efetivo_total == "3":
                    lista23 = [id_escala, escalado1, escalado2, 'SEM INSCRITOS']
                    linha23 = lista23 + [''] * (len(lista_colunas) - len(lista23))
                    df_escalados = pd.DataFrame([linha23])
                elif efetivo_total == "4":
                    lista24 = [id_escala, escalado1, escalado2, 'SEM INSCRITOS', 'SEM INSCRITOS']
                    linha24 = lista24 + [''] * (len(lista_colunas) - len(lista24))
                    df_escalados = pd.DataFrame([linha24])
                elif efetivo_total == "5":
                    lista25 = [id_escala, escalado1, escalado2, 'SEM INSCRITOS', 'SEM INSCRITOS', 'SEM INSCRITOS']
                    linha25 = lista25 + [''] * (len(lista_colunas) - len(lista25))
                    df_escalados = pd.DataFrame([linha25])
                elif efetivo_total == "6":
                    lista26 = [id_escala, escalado1, escalado2, 'SEM INSCRITOS', 'SEM INSCRITOS', 'SEM INSCRITOS', 'SEM INSCRITOS']
                    linha26 = lista16 + [''] * (len(lista_colunas) - len(lista26))
                    df_escalados = pd.DataFrame([linha26])
                elif efetivo_total == "7":
                    lista27 = [id_escala, escalado1, escalado2, 'SEM INSCRITOS', 'SEM INSCRITOS', 'SEM INSCRITOS', 'SEM INSCRITOS', 'SEM INSCRITOS']
                    linha27 = lista17 + [''] * (len(lista_colunas) - len(lista27))
                    df_escalados = pd.DataFrame([linha27])
                elif efetivo_total == "8":
                    lista28 = [id_escala, escalado1, escalado2, 'SEM INSCRITOS', 'SEM INSCRITOS', 'SEM INSCRITOS', 'SEM INSCRITOS', 'SEM INSCRITOS', 'SEM INSCRITOS']
                    linha28 = lista28 + [''] * (len(lista_colunas) - len(lista28))
                    df_escalados = pd.DataFrame([linha28])
                elif efetivo_total == "9":
                    lista29 = [id_escala, escalado1, escalado2, 'SEM INSCRITOS', 'SEM INSCRITOS', 'SEM INSCRITOS', 'SEM INSCRITOS', 'SEM INSCRITOS', 'SEM INSCRITOS', 'SEM INSCRITOS']
                    linha29 = lista29 + [''] * (len(lista_colunas) - len(lista29))
                    df_escalados = pd.DataFrame([linha29])
                lista_2escalados = df_escalados.values.tolist()
                if conexao == "VPN":
                    for dados_linha2 in lista_2escalados:
                        planilha_excel.append(dados_linha2)
                    pasta_excel.save(bd.caminho_relativo("DB_GESTÃO_DEJEM.xlsx"))
                else:
                    geracao.append_rows(lista_2escalados, value_input_option='USER_ENTERED')
                btn_retornar.click()
                sleep(5)   
            elif efetivo_escalado == "3":
                re1 = navegador.find_element(By.ID,"span_vVW_MOT_ESCEFE_ESCEFERE_0001").text
                posto1 = navegador.find_element(By.ID,"span_vGRIDPOSDES2_0001").text
                militar1 = navegador.find_element(By.ID,"span_vGRIDPMNOM2_0001").text
                opm1 = navegador.find_element(By.ID,"span_vGRIDPMOPMDES2_0001").text
                escalado1 = f"{posto1} {re1} {militar1} {opm1}"
                re2 = navegador.find_element(By.ID,"span_vVW_MOT_ESCEFE_ESCEFERE_0002").text
                posto2 = navegador.find_element(By.ID,"span_vGRIDPOSDES2_0002").text
                militar2 = navegador.find_element(By.ID,"span_vGRIDPMNOM2_0002").text
                opm2 = navegador.find_element(By.ID,"span_vGRIDPMOPMDES2_0002").text
                escalado2 = f"{posto2} {re2} {militar2} {opm2}"
                re3 = navegador.find_element(By.ID,"span_vVW_MOT_ESCEFE_ESCEFERE_0003").text
                posto3 = navegador.find_element(By.ID,"span_vGRIDPOSDES2_0003").text
                militar3 = navegador.find_element(By.ID,"span_vGRIDPMNOM2_0003").text
                opm3 = navegador.find_element(By.ID,"span_vGRIDPMOPMDES2_0003").text
                escalado3 = f"{posto3} {re3} {militar3} {opm3}"
                if efetivo_total == "3":
                    lista33 = [id_escala, escalado1, escalado2, escalado3]
                    linha33 = lista33 + [''] * (len(lista_colunas) - len(lista33))
                    df_escalados = pd.DataFrame([linha33                ])
                elif efetivo_total == "4":
                    lista34 = [id_escala, escalado1, escalado2, escalado3, 'SEM INSCRITOS']
                    linha34 = lista34 + [''] * (len(lista_colunas) - len(lista34))
                    df_escalados = pd.DataFrame([linha34])
                elif efetivo_total == "5":
                    lista35 = [id_escala, escalado1, escalado2, escalado3, 'SEM INSCRITOS', 'SEM INSCRITOS']
                    linha35 = lista35 + [''] * (len(lista_colunas) - len(lista35))
                    df_escalados = pd.DataFrame([linha35])
                elif efetivo_total == "6":
                    lista36 = [id_escala, escalado1, escalado2, escalado3, 'SEM INSCRITOS', 'SEM INSCRITOS', 'SEM INSCRITOS']
                    linha36 = lista36 + [''] * (len(lista_colunas) - len(lista36))
                    df_escalados = pd.DataFrame([linha36])
                elif efetivo_total == "7":
                    lista37 = [id_escala, escalado1, escalado2, escalado3, 'SEM INSCRITOS', 'SEM INSCRITOS', 'SEM INSCRITOS', 'SEM INSCRITOS']
                    linha37 = lista37 + [''] * (len(lista_colunas) - len(lista37))
                    df_escalados = pd.DataFrame([linha37])
                elif efetivo_total == "8":
                    lista38 = [id_escala, escalado1, escalado2, escalado3, 'SEM INSCRITOS', 'SEM INSCRITOS', 'SEM INSCRITOS', 'SEM INSCRITOS', 'SEM INSCRITOS']
                    linha38 = lista38 + [''] * (len(lista_colunas) - len(lista38))
                    df_escalados = pd.DataFrame([linha38])
                elif efetivo_total == "9":
                    lista39 = [id_escala, escalado1, escalado2, escalado3, 'SEM INSCRITOS', 'SEM INSCRITOS', 'SEM INSCRITOS', 'SEM INSCRITOS', 'SEM INSCRITOS', 'SEM INSCRITOS']
                    linha39 = lista39 + [''] * (len(lista_colunas) - len(lista39))
                    df_escalados = pd.DataFrame([linha39])
                lista_3escalados = df_escalados.values.tolist()
                if conexao == "VPN":
                    for dados_linha3 in lista_3escalados:
                        planilha_excel.append(dados_linha3)
                    pasta_excel.save(bd.caminho_relativo("DB_GESTÃO_DEJEM.xlsx"))
                else:
                    geracao.append_rows(lista_3escalados, value_input_option='USER_ENTERED')
                btn_retornar.click()
                sleep(5)
            elif efetivo_escalado == "4":
                re1 = navegador.find_element(By.ID,"span_vVW_MOT_ESCEFE_ESCEFERE_0001").text
                posto1 = navegador.find_element(By.ID,"span_vGRIDPOSDES2_0001").text
                militar1 = navegador.find_element(By.ID,"span_vGRIDPMNOM2_0001").text
                opm1 = navegador.find_element(By.ID,"span_vGRIDPMOPMDES2_0001").text
                escalado1 = f"{posto1} {re1} {militar1} {opm1}"
                re2 = navegador.find_element(By.ID,"span_vVW_MOT_ESCEFE_ESCEFERE_0002").text
                posto2 = navegador.find_element(By.ID,"span_vGRIDPOSDES2_0002").text
                militar2 = navegador.find_element(By.ID,"span_vGRIDPMNOM2_0002").text
                opm2 = navegador.find_element(By.ID,"span_vGRIDPMOPMDES2_0002").text
                escalado2 = f"{posto2} {re2} {militar2} {opm2}"
                re3 = navegador.find_element(By.ID,"span_vVW_MOT_ESCEFE_ESCEFERE_0002").text
                posto3 = navegador.find_element(By.ID,"span_vGRIDPOSDES2_0003").text
                militar3 = navegador.find_element(By.ID,"span_vGRIDPMNOM2_0003").text
                opm3 = navegador.find_element(By.ID,"span_vGRIDPMOPMDES2_0003").text
                escalado3 = f"{posto3} {re3} {militar3} {opm3}"
                re4 = navegador.find_element(By.ID,"span_vVW_MOT_ESCEFE_ESCEFERE_0004").text
                posto4 = navegador.find_element(By.ID,"span_vGRIDPOSDES2_0004").text
                militar4 = navegador.find_element(By.ID,"span_vGRIDPMNOM2_0004").text
                opm4 = navegador.find_element(By.ID,"span_vGRIDPMOPMDES2_0004").text
                escalado4 = f"{posto4} {re4} {militar4} {opm4}"
                if efetivo_total == "4":
                    lista44 = [id_escala, escalado1, escalado2, escalado3, escalado4]
                    linha44 = lista44 + [''] * (len(lista_colunas) - len(lista44))
                    df_escalados = pd.DataFrame([linha44])
                elif efetivo_total == "5":
                    lista45 = [id_escala, escalado1, escalado2, escalado3, escalado4, 'SEM INSCRITOS']
                    linha45 = lista35 + [''] * (len(lista_colunas) - len(lista45))
                    df_escalados = pd.DataFrame([linha45])
                elif efetivo_total == "6":
                    lista46 = [id_escala, escalado1, escalado2, escalado3, escalado4, 'SEM INSCRITOS', 'SEM INSCRITOS']
                    linha46 = lista46 + [''] * (len(lista_colunas) - len(lista46))
                    df_escalados = pd.DataFrame([linha46])
                elif efetivo_total == "7":
                    lista47 = [id_escala, escalado1, escalado2, escalado3, escalado4, 'SEM INSCRITOS', 'SEM INSCRITOS', 'SEM INSCRITOS']
                    linha47 = lista37 + [''] * (len(lista_colunas) - len(lista47))
                    df_escalados = pd.DataFrame([linha47])
                elif efetivo_total == "8":
                    lista48 = [id_escala, escalado1, escalado2, escalado3, escalado4, 'SEM INSCRITOS', 'SEM INSCRITOS', 'SEM INSCRITOS', 'SEM INSCRITOS']
                    linha48 = lista38 + [''] * (len(lista_colunas) - len(lista48))
                    df_escalados = pd.DataFrame([linha48])
                elif efetivo_total == "9":
                    lista49 = [id_escala, escalado1, escalado2, escalado3, escalado4, 'SEM INSCRITOS', 'SEM INSCRITOS', 'SEM INSCRITOS', 'SEM INSCRITOS', 'SEM INSCRITOS']
                    linha49 = lista49 + [''] * (len(lista_colunas) - len(lista49))
                    df_escalados = pd.DataFrame([linha49                    ])
                lista_4escalados = df_escalados.values.tolist()
                if conexao == "VPN":
                    for dados_linha4 in lista_4escalados:
                        planilha_excel.append(dados_linha4)
                    pasta_excel.save(bd.caminho_relativo("DB_GESTÃO_DEJEM.xlsx"))
                else:
                    geracao.append_rows(lista_4escalados, value_input_option='USER_ENTERED')
                btn_retornar.click()
                sleep(5)
            elif efetivo_escalado == "5":
                re1 = navegador.find_element(By.ID,"span_vVW_MOT_ESCEFE_ESCEFERE_0001").text
                posto1 = navegador.find_element(By.ID,"span_vGRIDPOSDES2_0001").text
                militar1 = navegador.find_element(By.ID,"span_vGRIDPMNOM2_0001").text
                opm1 = navegador.find_element(By.ID,"span_vGRIDPMOPMDES2_0001").text
                escalado1 = f"{posto1} {re1} {militar1} {opm1}"
                re2 = navegador.find_element(By.ID,"span_vVW_MOT_ESCEFE_ESCEFERE_0002").text
                posto2 = navegador.find_element(By.ID,"span_vGRIDPOSDES2_0002").text
                militar2 = navegador.find_element(By.ID,"span_vGRIDPMNOM2_0002").text
                opm2 = navegador.find_element(By.ID,"span_vGRIDPMOPMDES2_0002").text
                escalado2 = f"{posto2} {re2} {militar2} {opm2}"
                re3 = navegador.find_element(By.ID,"span_vVW_MOT_ESCEFE_ESCEFERE_0002").text
                posto3 = navegador.find_element(By.ID,"span_vGRIDPOSDES2_0003").text
                militar3 = navegador.find_element(By.ID,"span_vGRIDPMNOM2_0003").text
                opm3 = navegador.find_element(By.ID,"span_vGRIDPMOPMDES2_0003").text
                escalado3 = f"{posto3} {re3} {militar3} {opm3}"
                re4 = navegador.find_element(By.ID,"span_vVW_MOT_ESCEFE_ESCEFERE_0004").text
                posto4 = navegador.find_element(By.ID,"span_vGRIDPOSDES2_0004").text
                militar4 = navegador.find_element(By.ID,"span_vGRIDPMNOM2_0004").text
                opm4 = navegador.find_element(By.ID,"span_vGRIDPMOPMDES2_0004").text
                escalado4 = f"{posto4} {re4} {militar4} {opm4}"
                re5 = navegador.find_element(By.ID,"span_vVW_MOT_ESCEFE_ESCEFERE_0005").text
                posto5 = navegador.find_element(By.ID,"span_vGRIDPOSDES2_0005").text
                militar5 = navegador.find_element(By.ID,"span_vGRIDPMNOM2_0005").text
                opm5 = navegador.find_element(By.ID,"span_vGRIDPMOPMDES2_0005").text
                escalado5 = f"{posto5} {re5} {militar5} {opm5}"            
                if efetivo_total == "5":
                    lista55 = [id_escala, escalado1, escalado2, escalado3, escalado4, escalado5]
                    linha55 = lista55 + [''] * (len(lista_colunas) - len(lista55))
                    df_escalados = pd.DataFrame([linha55])
                elif efetivo_total == "6":
                    lista56 = [id_escala, escalado1, escalado2, escalado3, escalado4, escalado5, 'SEM INSCRITOS']
                    linha56 = lista56 + [''] * (len(lista_colunas) - len(lista56))
                    df_escalados = pd.DataFrame([linha56])
                elif efetivo_total == "7":
                    lista57 = [id_escala, escalado1, escalado2, escalado3, escalado4, escalado5, 'SEM INSCRITOS', 'SEM INSCRITOS']
                    linha57 = lista57 + [''] * (len(lista_colunas) - len(lista57))
                    df_escalados = pd.DataFrame([linha57])
                elif efetivo_total == "8":
                    lista58 = [id_escala, escalado1, escalado2, escalado3, escalado4,escalado5, 'SEM INSCRITOS', 'SEM INSCRITOS', 'SEM INSCRITOS']
                    linha58 = lista58 + [''] * (len(lista_colunas) - len(lista58))
                    df_escalados = pd.DataFrame([linha58])
                elif efetivo_total == "9":
                    lista59 = [id_escala, escalado1, escalado2, escalado3, escalado4, escalado5, 'SEM INSCRITOS', 'SEM INSCRITOS', 'SEM INSCRITOS', 'SEM INSCRITOS']
                    linha59 = lista59 + [''] * (len(lista_colunas) - len(lista59))
                    df_escalados = pd.DataFrame([linha59])
                lista_5escalados = df_escalados.values.tolist()
                if conexao == "VPN":
                    for dados_linha5 in lista_5escalados:
                        planilha_excel.append(dados_linha5)
                    pasta_excel.save(bd.caminho_relativo("DB_GESTÃO_DEJEM.xlsx"))
                else:
                    geracao.append_rows(lista_5escalados, value_input_option='USER_ENTERED')
                btn_retornar.click()
                sleep(5)
            elif efetivo_escalado == "6":
                re1 = navegador.find_element(By.ID,"span_vVW_MOT_ESCEFE_ESCEFERE_0001").text
                posto1 = navegador.find_element(By.ID,"span_vGRIDPOSDES2_0001").text
                militar1 = navegador.find_element(By.ID,"span_vGRIDPMNOM2_0001").text
                opm1 = navegador.find_element(By.ID,"span_vGRIDPMOPMDES2_0001").text
                escalado1 = f"{posto1} {re1} {militar1} {opm1}"
                re2 = navegador.find_element(By.ID,"span_vVW_MOT_ESCEFE_ESCEFERE_0002").text
                posto2 = navegador.find_element(By.ID,"span_vGRIDPOSDES2_0002").text
                militar2 = navegador.find_element(By.ID,"span_vGRIDPMNOM2_0002").text
                opm2 = navegador.find_element(By.ID,"span_vGRIDPMOPMDES2_0002").text
                escalado2 = f"{posto2} {re2} {militar2} {opm2}"
                re3 = navegador.find_element(By.ID,"span_vVW_MOT_ESCEFE_ESCEFERE_0002").text
                posto3 = navegador.find_element(By.ID,"span_vGRIDPOSDES2_0003").text
                militar3 = navegador.find_element(By.ID,"span_vGRIDPMNOM2_0003").text
                opm3 = navegador.find_element(By.ID,"span_vGRIDPMOPMDES2_0003").text
                escalado3 = f"{posto3} {re3} {militar3} {opm3}"
                re4 = navegador.find_element(By.ID,"span_vVW_MOT_ESCEFE_ESCEFERE_0004").text
                posto4 = navegador.find_element(By.ID,"span_vGRIDPOSDES2_0004").text
                militar4 = navegador.find_element(By.ID,"span_vGRIDPMNOM2_0004").text
                opm4 = navegador.find_element(By.ID,"span_vGRIDPMOPMDES2_0004").text
                escalado4 = f"{posto4} {re4} {militar4} {opm4}"
                re5 = navegador.find_element(By.ID,"span_vVW_MOT_ESCEFE_ESCEFERE_0005").text
                posto5 = navegador.find_element(By.ID,"span_vGRIDPOSDES2_0005").text
                militar5 = navegador.find_element(By.ID,"span_vGRIDPMNOM2_0005").text
                opm5 = navegador.find_element(By.ID,"span_vGRIDPMOPMDES2_0005").text
                escalado5 = f"{posto5} {re5} {militar5} {opm5}"            
                re6 = navegador.find_element(By.ID,"span_vVW_MOT_ESCEFE_ESCEFERE_0006").text
                posto6 = navegador.find_element(By.ID,"span_vGRIDPOSDES2_0006").text
                militar6 = navegador.find_element(By.ID,"span_vGRIDPMNOM2_0006").text
                opm6 = navegador.find_element(By.ID,"span_vGRIDPMOPMDES2_0006").text
                escalado6 = f"{posto6} {re6} {militar6} {opm6}"
                if efetivo_total == "6":
                    lista66 = [id_escala, escalado1, escalado2, escalado3, escalado4, escalado5, escalado6]
                    linha66 = lista66 + [''] * (len(lista_colunas) - len(lista66))
                    df_escalados = pd.DataFrame([linha66])
                elif efetivo_total == "7":
                    lista67 = [id_escala, escalado1, escalado2, escalado3, escalado4, escalado5, escalado6, 'SEM INSCRITOS']
                    linha67 = lista67 + [''] * (len(lista_colunas) - len(lista67))
                    df_escalados = pd.DataFrame([linha67])
                elif efetivo_total == "8":
                    lista68 = [id_escala, escalado1, escalado2, escalado3, escalado4,escalado5, escalado6, 'SEM INSCRITOS', 'SEM INSCRITOS']
                    linha68 = lista68 + [''] * (len(lista_colunas) - len(lista68))
                    df_escalados = pd.DataFrame([linha68])
                elif efetivo_total == "9":
                    lista69 = [id_escala, escalado1, escalado2, escalado3, escalado4, escalado5, escalado6, 'SEM INSCRITOS', 'SEM INSCRITOS', 'SEM INSCRITOS']
                    linha69 = lista69 + [''] * (len(lista_colunas) - len(lista69))
                    df_escalados = pd.DataFrame([linha69])
                lista_6escalados = df_escalados.values.tolist()
                if conexao == "VPN":
                    for dados_linha6 in lista_6escalados:
                        planilha_excel.append(dados_linha6)
                    pasta_excel.save(bd.caminho_relativo("DB_GESTÃO_DEJEM.xlsx"))
                else:
                    geracao.append_rows(lista_6escalados, value_input_option='USER_ENTERED')
                btn_retornar.click()
                sleep(5)
            elif efetivo_escalado == "7":
                re1 = navegador.find_element(By.ID,"span_vVW_MOT_ESCEFE_ESCEFERE_0001").text
                posto1 = navegador.find_element(By.ID,"span_vGRIDPOSDES2_0001").text
                militar1 = navegador.find_element(By.ID,"span_vGRIDPMNOM2_0001").text
                opm1 = navegador.find_element(By.ID,"span_vGRIDPMOPMDES2_0001").text
                escalado1 = f"{posto1} {re1} {militar1} {opm1}"
                re2 = navegador.find_element(By.ID,"span_vVW_MOT_ESCEFE_ESCEFERE_0002").text
                posto2 = navegador.find_element(By.ID,"span_vGRIDPOSDES2_0002").text
                militar2 = navegador.find_element(By.ID,"span_vGRIDPMNOM2_0002").text
                opm2 = navegador.find_element(By.ID,"span_vGRIDPMOPMDES2_0002").text
                escalado2 = f"{posto2} {re2} {militar2} {opm2}"
                re3 = navegador.find_element(By.ID,"span_vVW_MOT_ESCEFE_ESCEFERE_0002").text
                posto3 = navegador.find_element(By.ID,"span_vGRIDPOSDES2_0003").text
                militar3 = navegador.find_element(By.ID,"span_vGRIDPMNOM2_0003").text
                opm3 = navegador.find_element(By.ID,"span_vGRIDPMOPMDES2_0003").text
                escalado3 = f"{posto3} {re3} {militar3} {opm3}"
                re4 = navegador.find_element(By.ID,"span_vVW_MOT_ESCEFE_ESCEFERE_0004").text
                posto4 = navegador.find_element(By.ID,"span_vGRIDPOSDES2_0004").text
                militar4 = navegador.find_element(By.ID,"span_vGRIDPMNOM2_0004").text
                opm4 = navegador.find_element(By.ID,"span_vGRIDPMOPMDES2_0004").text
                escalado4 = f"{posto4} {re4} {militar4} {opm4}"
                re5 = navegador.find_element(By.ID,"span_vVW_MOT_ESCEFE_ESCEFERE_0005").text
                posto5 = navegador.find_element(By.ID,"span_vGRIDPOSDES2_0005").text
                militar5 = navegador.find_element(By.ID,"span_vGRIDPMNOM2_0005").text
                opm5 = navegador.find_element(By.ID,"span_vGRIDPMOPMDES2_0005").text
                escalado5 = f"{posto5} {re5} {militar5} {opm5}"            
                re6 = navegador.find_element(By.ID,"span_vVW_MOT_ESCEFE_ESCEFERE_0006").text
                posto6 = navegador.find_element(By.ID,"span_vGRIDPOSDES2_0006").text
                militar6 = navegador.find_element(By.ID,"span_vGRIDPMNOM2_0006").text
                opm6 = navegador.find_element(By.ID,"span_vGRIDPMOPMDES2_0006").text
                escalado6 = f"{posto6} {re6} {militar6} {opm6}"
                re7 = navegador.find_element(By.ID,"span_vVW_MOT_ESCEFE_ESCEFERE_0007").text
                posto7 = navegador.find_element(By.ID,"span_vGRIDPOSDES2_0007").text
                militar7 = navegador.find_element(By.ID,"span_vGRIDPMNOM2_0007").text
                opm7 = navegador.find_element(By.ID,"span_vGRIDPMOPMDES2_0007").text
                escalado7 = f"{posto7} {re7} {militar7} {opm7}"            
                if efetivo_total == "7":
                    lista77 = [id_escala, escalado1, escalado2, escalado3, escalado4, escalado5, escalado6, escalado7]
                    linha77 = lista77 + [''] * (len(lista_colunas) - len(lista77))
                    df_escalados = pd.DataFrame([linha77])
                elif efetivo_total == "8":
                    lista78 = [id_escala, escalado1, escalado2, escalado3, escalado4,escalado5, escalado6, escalado7, 'SEM INSCRITOS']
                    linha78 = lista78 + [''] * (len(lista_colunas) - len(lista78))
                    df_escalados = pd.DataFrame([linha78])
                elif efetivo_total == "9":
                    lista79 = [id_escala, escalado1, escalado2, escalado3, escalado4, escalado5, escalado6, escalado7, 'SEM INSCRITOS', 'SEM INSCRITOS']
                    linha79 = lista79 + [''] * (len(lista_colunas) - len(lista79))
                    df_escalados = pd.DataFrame([linha79])
                lista_7escalados = df_escalados.values.tolist()
                if conexao == "VPN":
                    for dados_linha7 in lista_7escalados:
                        planilha_excel.append(dados_linha7)
                    pasta_excel.save(bd.caminho_relativo("DB_GESTÃO_DEJEM.xlsx"))
                else:
                    geracao.append_rows(lista_7escalados, value_input_option='USER_ENTERED')
                btn_retornar.click()
                sleep(5)
            elif efetivo_escalado == "8":
                re1 = navegador.find_element(By.ID,"span_vVW_MOT_ESCEFE_ESCEFERE_0001").text
                posto1 = navegador.find_element(By.ID,"span_vGRIDPOSDES2_0001").text
                militar1 = navegador.find_element(By.ID,"span_vGRIDPMNOM2_0001").text
                opm1 = navegador.find_element(By.ID,"span_vGRIDPMOPMDES2_0001").text
                escalado1 = f"{posto1} {re1} {militar1} {opm1}"
                re2 = navegador.find_element(By.ID,"span_vVW_MOT_ESCEFE_ESCEFERE_0002").text
                posto2 = navegador.find_element(By.ID,"span_vGRIDPOSDES2_0002").text
                militar2 = navegador.find_element(By.ID,"span_vGRIDPMNOM2_0002").text
                opm2 = navegador.find_element(By.ID,"span_vGRIDPMOPMDES2_0002").text
                escalado2 = f"{posto2} {re2} {militar2} {opm2}"
                re3 = navegador.find_element(By.ID,"span_vVW_MOT_ESCEFE_ESCEFERE_0002").text
                posto3 = navegador.find_element(By.ID,"span_vGRIDPOSDES2_0003").text
                militar3 = navegador.find_element(By.ID,"span_vGRIDPMNOM2_0003").text
                opm3 = navegador.find_element(By.ID,"span_vGRIDPMOPMDES2_0003").text
                escalado3 = f"{posto3} {re3} {militar3} {opm3}"
                re4 = navegador.find_element(By.ID,"span_vVW_MOT_ESCEFE_ESCEFERE_0004").text
                posto4 = navegador.find_element(By.ID,"span_vGRIDPOSDES2_0004").text
                militar4 = navegador.find_element(By.ID,"span_vGRIDPMNOM2_0004").text
                opm4 = navegador.find_element(By.ID,"span_vGRIDPMOPMDES2_0004").text
                escalado4 = f"{posto4} {re4} {militar4} {opm4}"
                re5 = navegador.find_element(By.ID,"span_vVW_MOT_ESCEFE_ESCEFERE_0005").text
                posto5 = navegador.find_element(By.ID,"span_vGRIDPOSDES2_0005").text
                militar5 = navegador.find_element(By.ID,"span_vGRIDPMNOM2_0005").text
                opm5 = navegador.find_element(By.ID,"span_vGRIDPMOPMDES2_0005").text
                escalado5 = f"{posto5} {re5} {militar5} {opm5}"            
                re6 = navegador.find_element(By.ID,"span_vVW_MOT_ESCEFE_ESCEFERE_0006").text
                posto6 = navegador.find_element(By.ID,"span_vGRIDPOSDES2_0006").text
                militar6 = navegador.find_element(By.ID,"span_vGRIDPMNOM2_0006").text
                opm6 = navegador.find_element(By.ID,"span_vGRIDPMOPMDES2_0006").text
                escalado6 = f"{posto6} {re6} {militar6} {opm6}"
                re7 = navegador.find_element(By.ID,"span_vVW_MOT_ESCEFE_ESCEFERE_0007").text
                posto7 = navegador.find_element(By.ID,"span_vGRIDPOSDES2_0007").text
                militar7 = navegador.find_element(By.ID,"span_vGRIDPMNOM2_0007").text
                opm7 = navegador.find_element(By.ID,"span_vGRIDPMOPMDES2_0007").text
                escalado7 = f"{posto7} {re7} {militar7} {opm7}"
                re8 = navegador.find_element(By.ID,"span_vVW_MOT_ESCEFE_ESCEFERE_0008").text
                posto8 = navegador.find_element(By.ID,"span_vGRIDPOSDES2_0008").text
                militar8 = navegador.find_element(By.ID,"span_vGRIDPMNOM2_0008").text
                opm8 = navegador.find_element(By.ID,"span_vGRIDPMOPMDES2_0008").text
                escalado8 = f"{posto8} {re8} {militar8} {opm8}"
                if efetivo_total == "8":
                    lista88 = [id_escala, escalado1, escalado2, escalado3, escalado4,escalado5, escalado6, escalado7, escalado8]
                    linha88 = lista88 + [''] * (len(lista_colunas) - len(lista88))
                    df_escalados = pd.DataFrame([linha88])
                elif efetivo_total == "9":
                    lista89 = [id_escala, escalado1, escalado2, escalado3, escalado4, escalado5, escalado6, escalado7, escalado8, 'SEM INSCRITOS']
                    linha89 = lista89 + [''] * (len(lista_colunas) - len(lista89))
                    df_escalados = pd.DataFrame([linha89])
                lista_8escalados = df_escalados.values.tolist()
                if conexao == "VPN":
                    for dados_linha8 in lista_8escalados:
                        planilha_excel.append(dados_linha8)
                    pasta_excel.save(bd.caminho_relativo("DB_GESTÃO_DEJEM.xlsx"))
                else:
                    geracao.append_rows(lista_8escalados, value_input_option='USER_ENTERED')
                btn_retornar.click()
                sleep(5)
            elif efetivo_escalado == "9":
                re1 = navegador.find_element(By.ID,"span_vVW_MOT_ESCEFE_ESCEFERE_0001").text
                posto1 = navegador.find_element(By.ID,"span_vGRIDPOSDES2_0001").text
                militar1 = navegador.find_element(By.ID,"span_vGRIDPMNOM2_0001").text
                opm1 = navegador.find_element(By.ID,"span_vGRIDPMOPMDES2_0001").text
                escalado1 = f"{posto1} {re1} {militar1} {opm1}"
                re2 = navegador.find_element(By.ID,"span_vVW_MOT_ESCEFE_ESCEFERE_0002").text
                posto2 = navegador.find_element(By.ID,"span_vGRIDPOSDES2_0002").text
                militar2 = navegador.find_element(By.ID,"span_vGRIDPMNOM2_0002").text
                opm2 = navegador.find_element(By.ID,"span_vGRIDPMOPMDES2_0002").text
                escalado2 = f"{posto2} {re2} {militar2} {opm2}"
                re3 = navegador.find_element(By.ID,"span_vVW_MOT_ESCEFE_ESCEFERE_0002").text
                posto3 = navegador.find_element(By.ID,"span_vGRIDPOSDES2_0003").text
                militar3 = navegador.find_element(By.ID,"span_vGRIDPMNOM2_0003").text
                opm3 = navegador.find_element(By.ID,"span_vGRIDPMOPMDES2_0003").text
                escalado3 = f"{posto3} {re3} {militar3} {opm3}"
                re4 = navegador.find_element(By.ID,"span_vVW_MOT_ESCEFE_ESCEFERE_0004").text
                posto4 = navegador.find_element(By.ID,"span_vGRIDPOSDES2_0004").text
                militar4 = navegador.find_element(By.ID,"span_vGRIDPMNOM2_0004").text
                opm4 = navegador.find_element(By.ID,"span_vGRIDPMOPMDES2_0004").text
                escalado4 = f"{posto4} {re4} {militar4} {opm4}"
                re5 = navegador.find_element(By.ID,"span_vVW_MOT_ESCEFE_ESCEFERE_0005").text
                posto5 = navegador.find_element(By.ID,"span_vGRIDPOSDES2_0005").text
                militar5 = navegador.find_element(By.ID,"span_vGRIDPMNOM2_0005").text
                opm5 = navegador.find_element(By.ID,"span_vGRIDPMOPMDES2_0005").text
                escalado5 = f"{posto5} {re5} {militar5} {opm5}"            
                re6 = navegador.find_element(By.ID,"span_vVW_MOT_ESCEFE_ESCEFERE_0006").text
                posto6 = navegador.find_element(By.ID,"span_vGRIDPOSDES2_0006").text
                militar6 = navegador.find_element(By.ID,"span_vGRIDPMNOM2_0006").text
                opm6 = navegador.find_element(By.ID,"span_vGRIDPMOPMDES2_0006").text
                escalado6 = f"{posto6} {re6} {militar6} {opm6}"
                re7 = navegador.find_element(By.ID,"span_vVW_MOT_ESCEFE_ESCEFERE_0007").text
                posto7 = navegador.find_element(By.ID,"span_vGRIDPOSDES2_0007").text
                militar7 = navegador.find_element(By.ID,"span_vGRIDPMNOM2_0007").text
                opm7 = navegador.find_element(By.ID,"span_vGRIDPMOPMDES2_0007").text
                escalado7 = f"{posto7} {re7} {militar7} {opm7}"
                re8 = navegador.find_element(By.ID,"span_vVW_MOT_ESCEFE_ESCEFERE_0008").text
                posto8 = navegador.find_element(By.ID,"span_vGRIDPOSDES2_0008").text
                militar8 = navegador.find_element(By.ID,"span_vGRIDPMNOM2_0008").text
                opm8 = navegador.find_element(By.ID,"span_vGRIDPMOPMDES2_0008").text
                escalado8 = f"{posto8} {re8} {militar8} {opm8}"
                re9 = navegador.find_element(By.ID,"span_vVW_MOT_ESCEFE_ESCEFERE_0009").text
                posto9 = navegador.find_element(By.ID,"span_vGRIDPOSDES2_0009").text
                militar9 = navegador.find_element(By.ID,"span_vGRIDPMNOM2_0009").text
                opm9 = navegador.find_element(By.ID,"span_vGRIDPMOPMDES2_0009").text
                escalado9 = f"{posto9} {re9} {militar9} {opm9}"            
                if efetivo_total == "9":
                    lista99 = [id_escala, escalado1, escalado2, escalado3, escalado4, escalado5, escalado6, escalado7, escalado8, escalado9]
                    linha99 = lista99 + [''] * (len(lista_colunas) - len(lista99))
                    df_escalados = pd.DataFrame([linha99])
                lista_9escalados = df_escalados.values.tolist()
                if conexao == "VPN":
                    for dados_linha9 in lista_9escalados:
                        planilha_excel.append(dados_linha9)
                    pasta_excel.save(bd.caminho_relativo("DB_GESTÃO_DEJEM.xlsx"))
                else:
                    geracao.append_rows(lista_9escalados, value_input_option='USER_ENTERED')
                btn_retornar.click()
                sleep(5)
    
    ac.acessa_aba_procedimentos(lg=cpf,sn=senha, opcao=aba)

    executa_montagem()
    
    navegador.quit()

#bd = ConexaoBancoDados()
#bd.Update_certifi()
#MontarEscalas()