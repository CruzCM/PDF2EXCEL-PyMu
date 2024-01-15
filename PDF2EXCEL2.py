import PySimpleGUI as sg
from tkinter import filedialog
import os
import fitz  # PyMuPDF
import pandas as pd


# Função para converter PDF em Excel
def convert_pdf_to_excel(pdf_file, excel_file):
    try:
        # Abrir o arquivo PDF
        doc = fitz.open(pdf_file)
        text_cells = []
        # Extrair o texto das páginas
        for page in doc:
            for block in page.get_text("blocks"):
                text_cells.append(block[4])
        # Criar um DataFrame a partir do texto
        data = {'Texto do PDF': text_cells}
        df = pd.DataFrame(data)
        # Salvar o DataFrame em um arquivo Excel
        df.to_excel(excel_file, index=False)

        doc.close()

    except Exception as e:
        sg.popup_error(f"Ocorreu um erro: {str(e)}", title="Erro")


# Função para escolher PDFs
def browse_button_click():
    file_paths = filedialog.askopenfilenames(
        title="Escolha os arquivos PDF",
        filetypes=[("Arquivos PDF", "*.pdf")],
    )
    if file_paths:
        for pdf_path in file_paths:
            # Converter cada PDF para Excel e salvar na mesma pasta
            pdf_filename = os.path.splitext(os.path.basename(pdf_path))[0]
            excel_filename = os.path.join(os.path.dirname(pdf_path), f"{pdf_filename}.xlsx")
            convert_pdf_to_excel(pdf_path, excel_filename)


# Layout
sg.theme('LightGrey3')
layout = [
    [sg.Text("", size=(1, 1))],
    [sg.Text("Escolha os arquivos PDF para converter em Excel:"), sg.Button("Browse", key="-BROWSE-BUTTON-")]
]

window = sg.Window("Conversor PDF para Excel", layout, size=(400, 100))

# EVENTOS
while True:
    event, values = window.read()

    if event == sg.WINDOW_CLOSED:
        break
    elif event == "-BROWSE-BUTTON-":
        browse_button_click()
        sg.popup("Conversão concluída com sucesso!", title="Conversão de PDF para Excel")
        window.close()
