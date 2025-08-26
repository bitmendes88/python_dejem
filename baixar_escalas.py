import os
import requests
from pdf2image import convert_from_path
from PIL import Image
from classes import ConexaoBancoDados, Acessos
from time import sleep

def Baixar_Escalas(cpf, senha, aba, conexao, gb):
    # INSTANCIA A CLASSE PARA CONEXÃO COM O BANCO
    bd = ConexaoBancoDados(gb)
    ac = Acessos(gb)
    bd.Update_certifi()

    # BUSCA NO BANCO DE DADOS DO APPSHEET AS ESCALAS QUE SERÃO BAIXADAS
    df_lista_escalas = bd.Escalas_Para_Baixar()

    if conexao == "VPN":
        print("Aguardando VPN")
        sleep(40)
    else:
        print("Conexão via Intranet.")

    navegador = ac.inicia_navegador()

    # PASTA PARA SALVAR E LISTAR AS ESCALAS
    pasta = r'C:\\AUTOMACAO_DEJEM\\ESCALAS_PNG'

    # LINK DAS ESCALAS NA INTRANET
    link_intranet = "http://sistemasadmin.intranet.policiamilitar.sp.gov.br/Escala/arrelpreesc.aspx?"

    def baixa_pdf_escalas():

        os.makedirs(pasta, exist_ok=True)

        # FAZ A LEITURA DA LISTA DE ESCALAS E BAIXA TODAS
        for _, row in df_lista_escalas.iterrows():
            id_escala = row[f'ID_ESCALA']
            nome_pdf = row[f'NOME_PDF']

            link_pdf = f"{link_intranet}{id_escala}.pdf"
            file_name = f"{nome_pdf}.pdf"
            file_path = os.path.join(pasta, file_name)
            response = requests.get(link_pdf)
            if response.status_code == 200:
                with open(file_path, 'wb') as file:
                    file.write(response.content)
                print(f"PDF {file_name} salvo com sucesso.")
            else:
                print(f"Falha ao baixar o PDF do link {link_pdf}")

    def converte_para_png():

        caminho_poppler = bd.caminho_relativo("poppler/Library/bin")
        
        # Verifica se o caminho da pasta poppler está correto
        if not os.path.exists(caminho_poppler):
            raise FileNotFoundError(f"Pasta poppler não encontrada: {caminho_poppler}")
        
        # BUSCA OS ARQUIVOS EM PDF NA PASTA
        arquivos = os.listdir(pasta)
        pdfs = [arquivo for arquivo in arquivos if arquivo.endswith('.pdf')]

        # PERCORRE CADA DOCUMENTO
        for pdf in pdfs:
            escala_pdf = os.path.join(pasta, pdf)
            
            # Usa o caminho poppler ao chamar convert_from_path
            escala = convert_from_path(escala_pdf, poppler_path=caminho_poppler)
            
            for i, pagina in enumerate(escala):
                nome_png = pdf.replace('.pdf', f'_{i+1}.png') if len(escala) > 1 else pdf.replace('.pdf', '.png')
                escala_png = os.path.join(pasta, nome_png)

                # REDIMENSIONA A IMAGEM PROPORCIONALMENTE PARA 1500 PIXELS DE LARGURA
                largura_desejada = 1500
                largura_original, altura_original = pagina.size
                altura_desejada = int((largura_desejada / largura_original) * altura_original)
                pagina_redimensionada = pagina.resize((largura_desejada, altura_desejada), Image.Resampling.LANCZOS)

                # SALVA A IMAGEM REDIMENSIONADA EM PNG
                pagina_redimensionada.save(escala_png, 'PNG')
                
                # EXCLUI O PDF
                os.remove(escala_pdf)

        print("Escalas convertidas para PNG.")


    ac.acessa_aba_procedimentos(lg=cpf, sn=senha, opcao=aba)
    baixa_pdf_escalas()    
    converte_para_png()
    navegador.quit()