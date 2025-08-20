from selenium.webdriver import ActionChains
from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.alert import Alert
import pandas as pd
import time
from classes import ConexaoBancoDados, Acessos
from openpyxl import load_workbook

def ConfirmarEscalas(cpf, senha, aba, conexao):
    bd = ConexaoBancoDados()
    ac = Acessos()
    bd.Update_certifi()

    confirmacao = bd.confirmacao

    lista_id_confirmar = bd.Escalas_Para_Confirmar()
    df_id_confirmar = pd.DataFrame(lista_id_confirmar)

    if conexao == "VPN":
        pasta_excel = load_workbook(bd.caminho_relativo("BD_GESTAO_DEJEM.xlsx"))
        planilha_excel = pasta_excel.worksheets[0]
        print("Aguardando VPN")
        time.sleep(40)

    else:
        print("Conexão via Intranet.")

    navegador = ac.inicia_navegador()

    wait =  WebDriverWait(navegador, 100)
    alerta = Alert(navegador)

    def confirma_escalas():
        # PECORRE CADA LINHA DO DATAFRAME RETORNADO DE "Escalas_para_Confirmar" E EXECUTA AS CONFIRMAÇÕES
        for index, row in df_id_confirmar.iterrows():
            id_escala = row['ID_ESCALA']
            data_escala = row['DATA_ESCALA']
            navegador.refresh()
            
            # AGUARDA APARECER A LUPA PARA CONTINUAR
            wait.until(
                EC.presence_of_element_located((By.ID,"IMAGE3"))
            )
            # DIGITA O ID
            navegador.find_element('xpath','//*[@id="vESCOPRIDF"]').send_keys(id_escala)
            time.sleep(0.5)
            # CLICA NA LUPA
            lupa = navegador.find_element('xpath','//*[@id="IMAGE3"]')
            lupa.click()
            time.sleep(3)

            # VERIFICA SE A ESCALA FOI CONFIRMADA E EXECUTA CONFORME RESULTADO
            if navegador.find_element('xpath','//*[@id="TABLE4"]/tbody/tr[1]/td/div').is_displayed():
                mensagem = navegador.find_element('xpath','//*[@id="TABLE4"]/tbody/tr[1]/td/div').text
                data_hora_atual = time.localtime()
                data_hora = "{:04d}-{:02d}-{:02d} {:02d}:{:02d}:{:02d}".format(
                data_hora_atual[0], data_hora_atual[1], data_hora_atual[2],
                data_hora_atual[3], data_hora_atual[4], data_hora_atual[5]
                )
                lista_confirmado = [id_escala, data_escala, data_hora, mensagem]
                confirmacao.append_row(lista_confirmado)
                navegador.refresh()
                time.sleep(5)
            else:
                time.sleep(5)
                navegador.find_element('xpath','//*[@id="TABLE4"]/tbody/tr[5]/td/input[1]').click()
                time.sleep(2)
                alerta.accept()
                time.sleep(3)
                mensagem = navegador.find_element('xpath','//*[@id="TABLE4"]/tbody/tr[1]/td/div').text
                data_hora_atual = time.localtime()
                data_hora = "{:04d}-{:02d}-{:02d} {:02d}:{:02d}:{:02d}".format(
                data_hora_atual[0], data_hora_atual[1], data_hora_atual[2],
                data_hora_atual[3], data_hora_atual[4], data_hora_atual[5]
                )
                lista_confirmado = [id_escala, data_escala, data_hora, mensagem]

                # BUSCA A REFERÊNCIA DA PLANILHA NA CLASSE DE CONEXÃO COM BANCO DE DADOS                
                if conexao == "VPN":
                    for linha in lista_confirmado:
                        planilha_excel.append(linha)
                    pasta_excel.save(bd.caminho_relativo("BD_GESTAO_DEJEM.xlsx"))
                else:
                    confirmacao.append_row(lista_confirmado)

                navegador.refresh()
                time.sleep(5)
        
    ac.acessa_aba_procedimentos(lg=cpf,sn=senha,opcao=aba)
    confirma_escalas()
    navegador.quit()