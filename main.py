import flet as ft
from cadastrar import CadastrarEscalas
from montar import MontarEscalas
from baixar_escalas import Baixar_Escalas
from processoSEI import Processo_SEI
from confirmar import ConfirmarEscalas

class PaginaInicial:
    pass

def main(page: ft.Page):
    # Configurações principais da página
    page.title = "Gestão de Escalas DEJEM"
    page.vertical_alignment = ft.MainAxisAlignment.CENTER
    page.horizontal_alignment = ft.CrossAxisAlignment.CENTER
    page.window.width = 600
    page.window.height = 930
    page.window.resizable = False
    background = ft.Container(
        #image_src="background_image.jpg",
        #image_fit=ft.ImageFit.COVER,
        expand=True  # Para cobrir toda a tela
    )

    cabecalho = ft.Text(
        "COMANDO DE BOMBEIROS DO INTERIOR 1",
        size=25,
        color="black",
        weight=ft.FontWeight.W_500,
        text_align=ft.TextAlign.CENTER,
    )

    # TOOLBAR
    toolbar = ft.Container(
        content=ft.Text(
            "GESTÃO DE ESCALAS DEJEM",
            size=20,
            color="white",
            weight=ft.FontWeight.BOLD,
        ),
        bgcolor="red",
        alignment=ft.alignment.center,
        height=50,
        padding=ft.padding.all(10),
    )

    # SUBTÍTULO
    subtitle = ft.Text(
        "Sistema para automação das rotinas DEJEM",
        size=16,
        color="red",
        weight=ft.FontWeight.W_500,
        text_align=ft.TextAlign.CENTER,
    )

    # CRIAÇÃO DINÂMICA DOS BOTÕES
    def create_button(text, on_click_action):
        return ft.Container(
            content=ft.Text(text, color="white", size=14, weight=ft.FontWeight.BOLD),
            bgcolor="blue",
            alignment=ft.alignment.center,
            width=200,
            height=50,
            border_radius=ft.border_radius.all(10),
            on_hover=lambda e: (setattr(e.control, "bgcolor", "red" if e.data == "true" else "blue"), e.control.update()),
            on_click=on_click_action
        )

    # DEFINIÇÃO DOS BOTÕES
    buttons = ft.Column(
        [
            create_button("CADASTRAR ESCALAS", lambda e: abrir_popup("cadastrar", selected_option)),
            create_button("MONTAR ESCALAS", lambda e: abrir_popup("montar", selected_option)),
            create_button("BAIXAR ESCALAS", lambda e: abrir_popup("baixar", selected_option)),
            create_button("GERAR PROCESSO SEI", lambda e: abrir_popup("sei", selected_option)),
            create_button("CONFIRMAR ESCALAS", lambda e: abrir_popup("confirmar", selected_option)),
        ],
        spacing=15,
        alignment=ft.MainAxisAlignment.CENTER,
        horizontal_alignment=ft.CrossAxisAlignment.CENTER,
    )

    # LINHA DIVISÓRIA
    divider = ft.Divider(color="red", thickness=2)

    # TÍTULO "LINKS ÚTEIS"
    links_title = ft.Text(
        "LINKS ÚTEIS",
        size=18,
        color="black",
        weight=ft.FontWeight.BOLD,
        text_align=ft.TextAlign.LEFT,
    )

    # LINKS ÚTEIS
    links = ft.Column(
        [
            ft.TextButton("APPSHEET - GESTÃO DAS ESCALAS DEJEM", on_click=lambda e: page.launch_url("https://www.appsheet.com/start/3149b7c4-fe3b-484c-9b65-149a511ec3fb")),
            ft.TextButton("PLANEJAR ESCALAS NA PLANILHA SHEETS", on_click=lambda e: page.launch_url("https://docs.google.com/spreadsheets/d/1x-Uj6ClJlQBqu87OQ_SB0kAiHU8BV62kqGdDcyHqT1c/edit?usp=sharing")),
            ft.TextButton("LOOKER - RELATÓRIO DE VAGAS CADASTRADAS", on_click=lambda e: page.launch_url("https://lookerstudio.google.com/reporting/0d55dff5-5f16-4f6d-8cb3-2b5473bb5f7e")),
            ft.TextButton("ABA PROCEDIMENTOS", on_click=lambda e: page.launch_url("http://ms.policiamilitar.sp.gov.br/login.aspx")),
            ft.TextButton("SEI - SISTEMA ELETRÔNICO DE INFORMAÇÕES", on_click=lambda e: page.launch_url("https://sei.sp.gov.br/sip/login.php?sigla_orgao_sistema=GESP&sigla_sistema=SEI")),
        ],
        spacing=10,
        alignment=ft.MainAxisAlignment.CENTER,
    )

    # Modal que será exibido como popup
    popup_modal = ft.AlertDialog(modal=True)

    # Função para abrir o popup
    def abrir_popup(btn_pop, conexao_btn_rotina):

        def continua_rotina(btn_cont, conexao_btn_pop):
            cpf = campo_cpf.value.strip()
            senha = campo_senha.value.strip()
            if not cpf or not senha:
                msg_erro.value = "Por favor, preencha todos os campos!"
                page.update()
                return
            try:
                if btn_cont == "cadastrar":
                    CadastrarEscalas(cpf, senha, btn_cont, conexao_btn_pop)
                elif btn_cont == "montar":
                    MontarEscalas(cpf, senha, btn_cont, conexao_btn_pop)
                elif btn_cont == "baixar":
                    Baixar_Escalas(cpf, senha, btn_cont, conexao_btn_pop)
                elif btn_cont == "sei":
                    Processo_SEI(cpf, senha)
                elif btn_cont == "confirmar":
                    ConfirmarEscalas(cpf, senha, btn_cont, conexao_btn_pop)
                popup_modal.open = False
            except Exception as ex:
                msg_erro.value = f"Erro: {str(ex)}"
            finally:
                page.update()

        # Componentes do popup
        titulo_popup = ""
        if btn_pop == "cadastrar":
            titulo_popup = "CADASTRAR ESCALAS | Digite os dados de login da ABA PROCEDIMENTOS."
        elif btn_pop == "montar":
            titulo_popup = "MONTAR ESCALAS | Digite os dados de login da ABA PROCEDIMENTOS."
        elif btn_pop == "baixar":
            titulo_popup = "BAIXAR ESCALAS | Digite os dados de login da ABA PROCEDIMENTOS."
        elif btn_pop == "sei":
            titulo_popup = "INSERIR ESCALAS NO SEI | Digite os dados de login do sistema SEI."
        elif btn_pop == "confirmar":
            titulo_popup = "CONFIRMAR ESCALAS | Digite os dados de login da ABA PROCEDIMENTOS."
            
        text_popup = ft.Text(titulo_popup,size=18, color="black", weight=ft.FontWeight.W_500, text_align=ft.TextAlign.CENTER)
        campo_cpf = ft.TextField(label="CPF", width=250)
        campo_senha = ft.TextField(label="Senha", password=True, width=250)
        msg_erro = ft.Text(value="", color="red")
        btn_continuar = ft.ElevatedButton("CONTINUAR", on_click=lambda e: continua_rotina(btn_pop, conexao_btn_rotina))
        btn_fechar = ft.IconButton(icon=ft.Icons.CLOSE, icon_color="red", on_click=fechar_popup, style=ft.ButtonStyle( padding=ft.padding.all(5), alignment=ft.alignment.top_right))

        # Criar o modal
        popup_modal.content = ft.Column(
            [
                ft.Row([btn_fechar],alignment=ft.MainAxisAlignment.END),
                text_popup,
                campo_cpf,
                campo_senha,
                msg_erro,
                btn_continuar,
            ],
            tight=True,
            horizontal_alignment=ft.CrossAxisAlignment.CENTER,
        )
        popup_modal.open = True
        page.update()

    def fechar_popup(e):
        popup_modal.open = False
        page.update()

    selected_option = "Intranet"
    def switch_changed(e):
        nonlocal selected_option
        selected_option = "VPN" if toggle_switch.value else "Intranet"

    toggle_switch = ft.Switch(value=False, on_change=switch_changed)
    page.update()

    # Adicionando elementos à página
    page.add(
        background,
        cabecalho,
        toolbar,
        ft.Container(
            content=subtitle,
            margin=ft.margin.symmetric(vertical=20),
            alignment=ft.alignment.center,
        ),
        ft.Row(
            [
                ft.Text("Intranet"),
                toggle_switch,
                ft.Text("VPN"),
            ],
            alignment=ft.MainAxisAlignment.CENTER,
            
        ),
        buttons,
        ft.Container(content=divider, margin=ft.margin.only(top=20)),
        ft.Container(
            content=links_title,
            margin=ft.margin.only(top=20),
            alignment=ft.alignment.center,
        ),

        links
    )
    page.overlay.append(popup_modal)
# Executa o app
ft.app(target=main)
