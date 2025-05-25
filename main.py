import flet as ft
from pdf2docx import Converter
from docx2pdf import convert
import os
import datetime
import traceback
import subprocess
import platform

def main(page: ft.Page):
    # Configurações fixas da janela
    page.title = "Conversor de Documentos - HSO Solutions"
    page.window_width = 800
    page.window_height = 700
    page.window_resizable = False
    page.vertical_alignment = ft.MainAxisAlignment.CENTER
    page.horizontal_alignment = ft.CrossAxisAlignment.CENTER
    page.theme_mode = ft.ThemeMode.DARK
    page.padding = 20
    page.bgcolor = ft.Colors.BLUE_GREY_900

    # Variáveis de estado
    selected_file_path = {"path": ""}
    output_path = {"path": ""}
    current_conversion = {"type": ""}

    # Componentes UI
    progress_ring = ft.ProgressRing(visible=False, width=20, height=20)
    file_picker = ft.FilePicker()

    def add_message(text, style="normal"):
        colors = {
            "normal": ft.Colors.BLUE_GREY_800,
            "user": ft.Colors.INDIGO_900,
            "success": ft.Colors.TEAL_800,
            "error": ft.Colors.RED_900,
            "warning": ft.Colors.AMBER_800
        }
        icon = {
            "normal": "📄",
            "user": "💾",
            "success": "✅",
            "error": "❌",
            "warning": "⚠️"
        }.get(style, "📄")
        
        message_log.controls.append(
            ft.ListTile(
                leading=ft.Text(icon, size=14),
                title=ft.Text(text, size=12, color=ft.Colors.WHITE),
                bgcolor=colors.get(style, ft.Colors.BLUE_GREY_800),
                height=60
            )
        )
        message_log.scroll_to(offset=-1, duration=100)
        page.update()

    def open_file(e=None):
        try:
            if output_path["path"] and os.path.exists(output_path["path"]):
                system = platform.system()
                if system == "Windows":
                    os.startfile(output_path["path"])
                elif system == "Darwin":
                    subprocess.run(["open", output_path["path"]])
                else:
                    subprocess.run(["xdg-open", output_path["path"]])
            else:
                add_message("Arquivo não encontrado", "warning")
        except Exception as err:
            add_message(f"Erro ao abrir: {str(err)}", "error")

    def ao_escolher(e: ft.FilePickerResultEvent):
        if e.files:
            selected_file_path["path"] = e.files[0].path
            file_ext = os.path.splitext(selected_file_path["path"])[1].lower()
            
            if file_ext == '.pdf':
                btn_pdf_to_word.disabled = False
                btn_word_to_pdf.disabled = True
                current_conversion["type"] = "pdf_to_word"
                add_message(f"PDF selecionado: {os.path.basename(selected_file_path['path'])}", "user")
            elif file_ext == '.docx':
                btn_pdf_to_word.disabled = True
                btn_word_to_pdf.disabled = False
                current_conversion["type"] = "word_to_pdf"
                add_message(f"Word selecionado: {os.path.basename(selected_file_path['path'])}", "user")
            else:
                add_message("Formato não suportado", "error")
            
            btn_abrir.disabled = True
        else:
            add_message("Nenhum arquivo selecionado", "warning")
        
        page.update()

    def escolher_arquivo(e):
        file_picker.pick_files(
            allowed_extensions=["pdf", "docx"],
            allow_multiple=False,
            dialog_title="Selecione um arquivo"
        )

    def converter_documento(e):
        if not selected_file_path["path"]:
            add_message("Selecione um arquivo primeiro", "error")
            return

        try:
            with open(selected_file_path["path"], "rb"):
                pass
        except IOError:
            add_message("Feche o arquivo antes de converter", "error")
            return

        # Desativa controles durante conversão
        btn_escolher.disabled = True
        btn_pdf_to_word.disabled = True
        btn_word_to_pdf.disabled = True
        btn_abrir.disabled = True
        progress_ring.visible = True
        page.update()

        try:
            timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
            if current_conversion["type"] == "pdf_to_word":
                add_message("Convertendo PDF para Word...", "user")
                output_file = os.path.splitext(selected_file_path["path"])[0] + f"{timestamp}.docx"
                cv = Converter(selected_file_path["path"])
                cv.convert(output_file, start=0, end=None)
                cv.close()
                output_path["path"] = output_file
                add_message(f"Conversão concluída: {os.path.basename(output_file)}", "success")
            elif current_conversion["type"] == "word_to_pdf":
                add_message("Convertendo Word para PDF...", "user")
                output_file = os.path.splitext(selected_file_path["path"])[0] + f"{timestamp}.pdf"
                convert(selected_file_path["path"], output_file)
                output_path["path"] = output_file
                add_message(f"Conversão concluída: {os.path.basename(output_file)}", "success")

            btn_abrir.disabled = False

        except Exception as err:
            add_message(f"Erro na conversão: {str(err)}", "error")
        finally:
            btn_escolher.disabled = False
            progress_ring.visible = False
            page.update()

    # Configuração do file picker
    file_picker.on_result = ao_escolher
    page.overlay.append(file_picker)

    # Área de mensagens
    message_log = ft.ListView(
        spacing=5,
        height=350,
        width=550,
        auto_scroll=False
    )

    # Botões
    btn_escolher = ft.ElevatedButton(
        "Selecionar Arquivo",
        icon=ft.Icons.UPLOAD_FILE,
        on_click=escolher_arquivo,
        width=200,
        height=40
    )

    btn_pdf_to_word = ft.ElevatedButton(
        "PDF para Word",
        icon=ft.Icons.PICTURE_AS_PDF,
        on_click=converter_documento,
        disabled=True,
        width=150,
        height=40
    )

    btn_word_to_pdf = ft.ElevatedButton(
        "Word para PDF",
        icon=ft.Icons.DOCUMENT_SCANNER,
        on_click=converter_documento,
        disabled=True,
        width=150,
        height=40
    )

    btn_abrir = ft.ElevatedButton(
        "Abrir Arquivo",
        icon=ft.Icons.OPEN_IN_NEW,
        on_click=open_file,
        disabled=True,
        width=200,
        height=40
    )

    # Layout principal
    page.add(
        ft.Column(
            [
                ft.Text(
                    "Conversor de Documentos",
                    size=20,
                    weight=ft.FontWeight.BOLD,
                    text_align=ft.TextAlign.CENTER
                ),
                
                ft.Container(
                    content=message_log,
                    padding=10,
                    border_radius=10,
                    bgcolor=ft.Colors.BLUE_GREY_800,
                    width=750,
                    height=350
                ),
                
                progress_ring,
                
                btn_escolher,
                
                ft.Row(
                    [btn_pdf_to_word, btn_word_to_pdf],
                    alignment=ft.MainAxisAlignment.CENTER,
                    spacing=20
                ),
                
                btn_abrir,
                
                ft.Text(
                    "HSO SOLUTIONS © 2025 - Todos os direitos reservados.",
                    text_align=ft.TextAlign.CENTER,
                    size=10,
                    color=ft.Colors.BLUE_GREY_500
                )
            ],
            spacing=15,
            horizontal_alignment=ft.CrossAxisAlignment.CENTER,
            scroll=ft.ScrollMode.AUTO
        )
    )

ft.app(target=main)