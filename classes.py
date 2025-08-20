import gspread
import os
import pandas as pd
import subprocess
import sys
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver import ActionChains
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from time import sleep
import pyautogui as gui


class ConexaoBancoDados:
    def __init__(self):
        self.credentials = self.caminho_relativo("credentials.json")
        self.client = gspread.service_account(filename=self.credentials)
        
        self.BD_GESTAO_DEJEM = self.client.open('BD_GESTAO_DEJEM')
        self.bd_app = self.BD_GESTAO_DEJEM.get_worksheet(0)
        self.cadastro = self.BD_GESTAO_DEJEM.get_worksheet(1)
        self.geracao = self.BD_GESTAO_DEJEM.get_worksheet(2)
        self.SEI = self.BD_GESTAO_DEJEM.get_worksheet(3)
        self.confirmacao = self.BD_GESTAO_DEJEM.get_worksheet(4)
        self.info_opm = self.BD_GESTAO_DEJEM.get_worksheet(5)

        self.tabela_info_opm = self.info_opm.get_all_records()
        self.df_tabela_info_opm = pd.DataFrame(self.tabela_info_opm)

        self.tabela_db_app = self.bd_app.get_all_records()
        self.df_tabela_db_app = pd.DataFrame(self.tabela_db_app)

        #=================================================================

        self.BD_DEJEM_PLANO = self.client.open('BD_DEJEM_PLANO')
        self.plano = self.BD_DEJEM_PLANO.get_worksheet(1)

        self.tabela_plano = self.plano.get_all_records()
        self.df_tabela_plano = pd.DataFrame(self.tabela_plano)

    def Update_certifi(self):
        try:
            # Comando para atualizar o pacote certifi
            subprocess.check_call([sys.executable, "-m", "pip", "install", "--upgrade", "certifi"])
            print("certifi atualizado com sucesso!")
        except Exception as e:
            print(f"Erro ao atualizar o certifi: {e}")
    
        # Função para pegar o caminho relativo
    
    def caminho_relativo(self, relative_path):
        try:
            base_path = r"C:\\AUTOMACAO_DEJEM"
        except Exception:
            base_path = os.path.abspath(".")

        return os.path.join(base_path, relative_path)   
    
    def Escalas_Para_Cadastrar(self):
        # VERIFICA SE EXIXTE A COLUNA CADASTRAR NA PLANILHA
        if 'CADASTRAR' not in self.df_tabela_plano.columns:
            print("A coluna 'CADASTRAR' não foi encontrada na tabela.")
            return None

        # FILTRA PELO CRITÉRIO 'CADASTRAR' == "SIM"
        self.filtro_df_tabela_plano = self.df_tabela_plano.query('CADASTRAR == "SIM"').copy()

        # DATAHORA TÉRMINO DA ESCALA
        self.filtro_df_tabela_plano['DATA_ESCALA_CONVERTIDA'] = pd.to_datetime(self.filtro_df_tabela_plano['DATA_ESCALA'], dayfirst=True, errors='coerce')
        self.filtro_df_tabela_plano['DATA_TERMINO'] = self.filtro_df_tabela_plano['DATA_ESCALA_CONVERTIDA'] + pd.Timedelta(hours=8)
        self.filtro_df_tabela_plano['STR_DATA_TERMINO'] = self.filtro_df_tabela_plano['DATA_TERMINO'].dt.strftime('%d/%m/%Y %H:%M')

        # HORA E MINUTO DA REVISTA
        self.filtro_df_tabela_plano['ANTECEDENCIA_REVISTA'] = self.filtro_df_tabela_plano.apply(lambda row: row['DATA_ESCALA_CONVERTIDA'] - pd.Timedelta(minutes=row['ANTECEDENCIA']), axis=1)
        self.filtro_df_tabela_plano['HORA_REVISTA'] = self.filtro_df_tabela_plano['ANTECEDENCIA_REVISTA'].dt.hour
        self.filtro_df_tabela_plano['MINUTO_REVISTA'] = self.filtro_df_tabela_plano['ANTECEDENCIA_REVISTA'].dt.minute

        print(self.filtro_df_tabela_plano[['DATA_ESCALA','NOME_ATIVIDADE']])
        
        # RETORNA UMA LISTA COM OS DADOS PARA CADASTRO
        return self.filtro_df_tabela_plano[
            ['DATA_ESCALA', 'STR_DATA_TERMINO','NOME_ATIVIDADE', 'CONVENIO', 'TIPO_ESCALA', 'OF_SUP', 'OF_INT', 'OF_SUB', 
            'SGT', 'CBSD', 'AISP', 'LOCAL', 'OPM', 'DATA_LIMITE_INSCRICAO', 'TIPO_ATIVIDADE', 
            'HORA_REVISTA', 'MINUTO_REVISTA', 'LOCAL_REVISTA', 'EPI', 'OBS']
        ].reset_index(drop=True)

    def Escalas_Para_Baixar(self):
        self.filtro_df_tabela_db_app = self.df_tabela_db_app.query('STATUS == "3-CONFIRMAR" and MILITAR_01 != "SEM INSCRITOS" and LINK_SEI == ""')
        self.filtro_df_tabela_db_app = self.filtro_df_tabela_db_app.copy()
        self.filtro_df_tabela_db_app.loc[:, 'DATA_ESCALA_CONVERTIDA'] = pd.to_datetime(self.filtro_df_tabela_db_app['DATA_ESCALA'], dayfirst=True)
        self.filtro_df_tabela_db_app['NOME_PDF'] = (self.filtro_df_tabela_db_app['DATA_ESCALA_CONVERTIDA'].dt.strftime('%d%m') + self.filtro_df_tabela_db_app['DATA_ESCALA_CONVERTIDA'].dt.strftime('%H%M'))
        return self.filtro_df_tabela_db_app[['ID_ESCALA', 'NOME_PDF']].reset_index(drop=True)

    def Escalas_para_Montar(self):
        self.df_tabela_db_app['DATA_LIMITE_INSCRICAO_CONVERTIDA'] = pd.to_datetime(self.df_tabela_db_app['DATA_LIMITE_INSCRICAO'], dayfirst=True)
        self.filtro_df_tabela_db_app = self.df_tabela_db_app.query('STATUS == "2-MONTAR" and DATA_LIMITE_INSCRICAO_CONVERTIDA < @pd.Timestamp.now()')
        return self.filtro_df_tabela_db_app[['ID_ESCALA','DATA_ESCALA']].reset_index(drop=True)

    def Escalas_Para_SEI(self):
        self.pasta = r'C:\\AUTOMACAO_DEJEM\\ESCALAS_PNG'
        self.lista_arquivos = os.listdir(self.pasta)
        self.lista_png = [arquivo for arquivo in self.lista_arquivos if arquivo.endswith('.png')]
        self.df_png = pd.DataFrame(self.lista_png, columns=['LISTA_PNG'])
        self.df_png['NOME_ESCALAS'] = self.df_png['LISTA_PNG'].str[:8]
        self.df_tabela_db_app = pd.DataFrame(self.tabela_db_app)
        self.filtro_df_tabela_db_app = self.df_tabela_db_app.query('STATUS == "3-CONFIRMAR" and MILITAR_01 != "SEM INSCRITOS" and LINK_SEI == ""')
        self.filtro_df_tabela_db_app = self.filtro_df_tabela_db_app.copy()
        self.filtro_df_tabela_db_app.loc[:, 'DATA_ESCALA_CONVERTIDA'] = pd.to_datetime(self.filtro_df_tabela_db_app['DATA_ESCALA'], dayfirst=True)
        self.filtro_df_tabela_db_app['NOME_ESCALAS'] = (self.filtro_df_tabela_db_app['DATA_ESCALA_CONVERTIDA'].dt.strftime('%d%m') + self.filtro_df_tabela_db_app['DATA_ESCALA_CONVERTIDA'].dt.strftime('%H%M'))
        self.df_para_SEI = pd.merge(self.df_png,self.filtro_df_tabela_db_app, on='NOME_ESCALAS')
        self.df_para_SEI['DDMM'] = self.df_para_SEI['NOME_ESCALAS'].str[:4]
        self.df_para_SEI['hhHmm'] = self.df_para_SEI['NOME_ESCALAS'].str[4:6] + 'H' + self.df_para_SEI['NOME_ESCALAS'].str[6:8]
        print(self.df_para_SEI[['ID_ESCALA', 'DATA_ESCALA']])
        return self.df_para_SEI[['LISTA_PNG','NOME_ESCALAS', 'ID_ESCALA', 'DATA_ESCALA','ATIVIDADE',"MESA_SEI",'DDMM', 'hhHmm']].reset_index(drop=True)

    def Escalas_Para_SEI_com_processo(self):
        self.filtro_df_tabela_db_app = self.df_tabela_db_app.query('STATUS == "3-CONFIRMAR" and MILITAR_01 != "SEM INSCRITOS" and LINK_SEI != ""')
        self.filtro_df_tabela_db_app = self.filtro_df_tabela_db_app.copy()
        self.filtro_df_tabela_db_app.loc[:, 'DATA_ESCALA_CONVERTIDA'] = pd.to_datetime(self.filtro_df_tabela_db_app['DATA_ESCALA'], dayfirst=True)
        self.filtro_df_tabela_db_app['NOME_ESCALAS'] = (self.filtro_df_tabela_db_app['DATA_ESCALA_CONVERTIDA'].dt.strftime('%d%m') + self.filtro_df_tabela_db_app['DATA_ESCALA_CONVERTIDA'].dt.strftime('%H%M'))
        self.filtro_df_tabela_db_app['DDMM'] = self.filtro_df_tabela_db_app['NOME_ESCALAS'].str[:4]
        return self.filtro_df_tabela_db_app[['ID_ESCALA', 'DATA_ESCALA','ATIVIDADE',"LINK_SEI","MESA_SEI",'DDMM']].reset_index(drop=True)

    def Escalas_Para_Confirmar(self):
        self.df_tabela_db_app['DATA_ESCALA_CONVERTIDA'] = pd.to_datetime(self.df_tabela_db_app['DATA_ESCALA'], dayfirst=True)
        filtro_df_tabela_db_app = self.df_tabela_db_app.query('COMPARECEU == "SIM" and CONFIRMADO != "SIM"')
        print(filtro_df_tabela_db_app[['ID_ESCALA','DATA_ESCALA']])
        return filtro_df_tabela_db_app[['ID_ESCALA','DATA_ESCALA']].reset_index(drop=True)

class Acessos():
    def __init__(self):
        self.bd = ConexaoBancoDados()

    def inicia_navegador(self):
        options = Options()
        options.add_argument("--incognito")
        self.navegador = webdriver.Chrome(options=options)
        return self.navegador


    def acessa_aba_procedimentos(self, lg, sn, opcao):    
        # FAZ O LOGIN NA INTRANET
        self.navegador.get("http://ms.policiamilitar.sp.gov.br/login.aspx")
        sleep(2)
        self.navegador.find_element('xpath','//*[@id="vUSRNUMCPFAUX"]').send_keys(lg)
        self.navegador.find_element('xpath','//*[@id="vSENHA"]').send_keys(sn)
        self.navegador.find_element('xpath','//*[@id="TABLE2"]/tbody/tr[3]/td/input').click()
        sleep(2)
        sirh = self.navegador.window_handles[1]
        self.navegador.switch_to.window(sirh)
        sleep(2)

        if opcao == "baixar":
            pass
        else:
            #NAVEGAR NO MENU E REALIZAR OS CLIQUES
            def clicar_menu(driver, class_name, texto, hover=False):
                try:
                    # AGUARDA A PRESENÇA DO ELEMENTO NA PÁGINA
                    elementos = WebDriverWait(driver, 15).until(
                        EC.presence_of_all_elements_located((By.XPATH, f"//td[contains(@class, '{class_name}') and contains(text(), '{texto}')]"))
                    )

                    for elemento in elementos:
                        if elemento.is_displayed() and elemento.is_enabled():
                            if hover:
                                ActionChains(driver).move_to_element(elemento).perform()
                                print(f'Hover realizado: {texto}')
                            else:
                                WebDriverWait(driver, 5).until(EC.element_to_be_clickable(elemento)).click()
                                print(f'Clique realizado: {texto}')
                            return
                except Exception as e:
                    print(f"Erro ao interagir com {texto}: {e}")

            # SELECIONA MENU SIRH
            clicar_menu(self.navegador, "ThemeClassicMainFolderText", "SIRH", hover=True)
            
            # SELECIONA SUBMENU Escala
            clicar_menu(self.navegador, "ThemeClassicMenuFolderText", "Escala", hover=True)

            # CLICA NO ÍTEM ADEQUADO CONFORME OPÇÃO DE SERVIÇO
            if opcao == "cadastrar":
                clicar_menu(self.navegador, "ThemeClassicMenuItemText", "Cadastro de Escala")
            elif opcao == "montar":
                clicar_menu(self.navegador, "ThemeClassicMenuItemText", "Montar Escala")
            elif opcao == "confirmar":
                clicar_menu(self.navegador, "ThemeClassicMenuItemText", "Confirmação de Presença")
           
            sleep(2)
            # TROCA A ABA DO DE CADASTRO
            aba2_navegador = self.navegador.window_handles[2]
            self.navegador.switch_to.window(aba2_navegador)
            sleep(1)

    def acessa_SEI(self, lg, sn):
        # LOGIN NO SISTEMA SEI
        self.navegador.get("https://sei.sp.gov.br/sip/login.php?sigla_orgao_sistema=GESP&sigla_sistema=SEI")
        sleep(2)
        self.navegador.find_element('xpath','//*[@id="txtUsuario"]').send_keys(lg)
        self.navegador.find_element('xpath','//*[@id="pwdSenha"]').send_keys(sn)
        orgao = "PMESP"
        lista_itens = WebDriverWait(self.navegador, 5).until(
            EC.presence_of_all_elements_located((By.XPATH, f"//option[contains(text(), '{orgao}')]")))        
        item_PMESP = lista_itens[0]
        WebDriverWait(self.navegador,5).until(EC.element_to_be_clickable(item_PMESP)).click()
        self.navegador.find_element('xpath','//*[@id="sbmAcessar"]').click()

    def seleciona_mesa_SEI(self, mesa_sei):
        #SELECIONA MESA SEI CONFORME A PRIMEIRA LINHA DO DATAFRAME
        self.navegador.maximize_window()

        # SELECIONA A MESA SEI CONFORME PRIMEIRA LINHA DO DATAFRAME
        self.navegador.find_element('xpath','//*[@id="divInfraBarraSistemaPadraoD"]/div[3]').click()
        sleep(1)
        self.navegador.find_element('xpath','//*[@id="txtInfraSiglaUnidade"]').click()
        sleep(1)
        self.navegador.find_element('xpath','//*[@id="txtInfraSiglaUnidade"]').send_keys(mesa_sei)
        gui.press('enter')
        sleep(1)
        for _ in range(3):
            gui.press('tab')
        gui.press('space')
        sleep(1)