import os
from pdf2docx import Converter
import tkinter as tk
from tkinter import filedialog


def selecionar_e_converter():
    # Oculta a janela principal do Tkinter (para mostrar apenas o explorador de arquivos)
    root = tk.Tk()
    root.withdraw()

    print("Abrindo janela para selecionar o PDF...")

    # Abre a caixa de diálogo para escolher o PDF
    pdf_path = filedialog.askopenfilename(
        title="Selecione o arquivo PDF para converter",
        filetypes=[("Arquivos PDF", "*.pdf")],
    )

    # Verifica se o usuário cancelou a seleção
    if not pdf_path:
        print("Nenhum arquivo foi selecionado.")
        return

    # Define automaticamente o nome do arquivo de saída (mesmo nome, mas com .docx)
    docx_path = os.path.splitext(pdf_path)[0] + ".docx"

    print(f"Convertendo o arquivo: {os.path.basename(pdf_path)}...")

    try:
        # Cria o objeto conversor e converte
        cv = Converter(pdf_path)
        cv.convert(docx_path, start=0, end=None)
        cv.close()

        print(f"Sucesso! Arquivo salvo em: {docx_path}")

    except Exception as e:
        print(f"Ocorreu um erro durante a conversão: {e}")


if __name__ == "__main__":
    selecionar_e_converter()
