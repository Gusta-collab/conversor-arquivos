import tkinter as tk
from tkinter import filedialog
from pdf2docx import Converter


root = tk.Tk()

# variaveis
arquivos_selecionados = []

# função que faz selelcionar apenas arquivos pdf
def selecionar_arquivos():
    arquivos = filedialog.askopenfilenames(
        title="selecione os arquivos para converter",
        filetypes =[("Arquivos PDF", "*.pdf")],
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

def converter_arquivos():
    if not arquivos_selecionados:
        print("nenhum arquivo selecionado")
        return

    for arquivo_pdf in arquivos_selecionados:
        arquivo_docx = arquivo_pdf.rsplit(".",1)[0] + ".docx"

        print("convertendo: {arquivo_pdf} ...")

        try:
            cv = Converter(arquivo_pdf)

            cv.convert(arquivo_docx, start=0, end=None)

            cv.close()

            print(f"Sucesso! Salvo como: {arquivo_docx}")

        except Exception as e:
            print("erro: {e}")




# tela para visualização

root.title("conversor PDF -> word")

root.geometry("600x500")

lista = tk.Listbox(root, width=50, height=10)
lista.pack(pady=20)

titulo = tk.Label(root, text="conversor pdf -> word", font=("arial", 20))
titulo.pack(pady=20)

# botão para o usuario

botao = tk.Button(root, text="adicionar arquivos", command=selecionar_arquivos)
botao.pack(pady=20)

botao_remover = tk.Button(root, text="remover arquivo", command = remover_arquivo)
botao_remover.pack(pady=20)

botao_converter = tk.Button(root, text="converter", command=converter_arquivos)
botao_converter.pack(pady=20)


root.mainloop()
