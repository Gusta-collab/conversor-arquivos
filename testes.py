import tkinter as tk
from tkinter import filedialog
from pdf2docx import Converter
import pandas as pd
import pdfplumber
import win32com.client
import os


root = tk.Tk()

# variaveis
arquivos_selecionados = []


# função que faz selelcionar apenas arquivos pdf
def selecionar_arquivos():
    arquivos = filedialog.askopenfilenames(
        title="selecione os arquivos para converter",
        filetypes =[("Arquivos PDF", "*.*")],
    )

    for arquivo in arquivos:
        if arquivo not in arquivos_selecionados:
            arquivos_selecionados.append(arquivo)
            lista.insert(tk.END, arquivo)


    print(arquivos)

def remover_arquivo():
    selecionado = lista.curselection()
    if selecionado:
        indice = selecionado[0]
        lista.delete(indice)

        arquivos_selecionados.pop(indice)

        print(arquivos_selecionados)

def conversao_word():
    if not arquivos_selecionados:
        print("nenhum arquivo selecionado")
        return

    for arquivo_pdf in arquivos_selecionados:
        arquivo_docx = arquivo_pdf.rsplit(".",1)[0] + ".docx"

        print(f"convertendo: {arquivo_pdf} ...")

        try:
            cv = Converter(arquivo_pdf)

            cv.convert(arquivo_docx, start=0, end=None)

            cv.close()

            print(f"Sucesso! Salvo como: {arquivo_docx}")

        except Exception as e:
            print(f"erro: {e}")


def conversao_excel():
    if not arquivos_selecionados:
        print("nenhum arquivo selecionado")
        return

    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = False

    for arquivo_excel in arquivos_selecionados:
        try:
            print("convertendo")

            caminho_completo = os.path.abspath(arquivo_excel)
            arquivo_pdf = caminho_completo.rsplit(".",1)[0] + ".pdf"

            pasta_trabalho = excel.Workbooks.Open(caminho_completo)
            pasta_trabalho.ExportAsFixedFormat(0,arquivo_pdf)
            pasta_trabalho.Close(False)

            print(f"sucesso! salvo:{arquivo_pdf}")

        except Exception as e:
            print(f"erro:{e}")

    excel.Quit()


    
def conversao_powerpoint():  
    print("opção selecionada: powerpoint")
    # código para conversão para powerpoint



def selecao_arquivo():
    if selecionado == "word":
        print("opção selecionada: word")
        modulo = conversao_word


def executar_comando_selecionado():
    arquivo_escolhido = selecionado.get()

    funcao_para_executar = comando_arquivo.get(arquivo_escolhido)

    if funcao_para_executar:
        funcao_para_executar()
    


# tela para visualização
root.title("conversor PDF -> word")
root.geometry("600x600")

comando_arquivo = {"word": conversao_word,"excel": conversao_excel,"powerpoint": conversao_powerpoint}



lista = tk.Listbox(root, width=50, height=10)
lista.pack(pady=20)

titulo = tk.Label(root, text="conversor pdf -> word", font=("arial", 20))
titulo.pack(pady=20)

# botão para o usuario

botao = tk.Button(root, text="adicionar arquivos", command=selecionar_arquivos)
botao.pack(pady=20)

botao_remover = tk.Button(root, text="remover arquivo", command = remover_arquivo)
botao_remover.pack(pady=20)

botao_converter = tk.Button(root, text="converter", command=executar_comando_selecionado)
botao_converter.pack(pady=20)

#selecionar arquivo
opcoes = list(comando_arquivo.keys())
selecionado = tk.StringVar(root)
selecionado.set(opcoes[0]) #valor padrão inicial

menu_suspenso = tk.OptionMenu(root, selecionado, *opcoes)
menu_suspenso.pack(pady=20)


root.mainloop()
