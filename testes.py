import tkinter as tk
from tkinter import filedialog

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
        lista.insert(tk.END, arquivo)

    print(arquivos)

def remover_arquivo():
    selecionado = lista.curselection()
    if selecionado:
        indice = selecionado[0]
        lista.delete(indice)

        arquivos_selecionados.pop(indice)



# tela para visualização

root.title("conversor PDF -> word")

root.geometry("600x400")

lista = tk.Listbox(root, width=50, height=10)
lista.pack(pady=20)

titulo = tk.Label(root, text="conversor pdf -> word", font=("arial", 20))
titulo.pack(pady=20)

# botão para o usuario

botao = tk.Button(root, text="adicionar arquivos", command=selecionar_arquivos)
botao.pack(pady=20)

botao_remover = tk.Button(root, text="remover arquivo", command = remover_arquivo)
botao_remover.pack(pady=20)


root.mainloop()
