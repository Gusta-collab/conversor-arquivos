import os
import customtkinter as ctk
from tkinter import filedialog
from tkinterdnd2 import TkinterDnD, DND_FILES
import win32com.client

# unindo customtkinter com tkinterdnd2 para manter o arrastar e soltar funcionando no tema moderno
class TkinterDnD_CTk(ctk.CTk, TkinterDnD.DnDWrapper):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.TkdndVersion = TkinterDnD._require(self)

ctk.set_appearance_mode("Dark")

# cores base do nosso tema escuro
BG_ROOT = "#121212"
BG_ELEVATED = "#1e1e1e"
COR_PRINCIPAL = "#0A84FF"
TXT_TITULO = "#FFFFFF"
TXT_SECUNDARIO = "#8E8E93"

# deixando a janela bem menor e compacta
root = TkinterDnD_CTk(fg_color=BG_ROOT)
root.title("Conversor PDF")
root.geometry("400x500") 
root.resizable(False, False)

# variaveis
arquivos_selecionados = []

# criando a pasta padrão na área de trabalho onde os pdfs vão ser salvos
caminho_desktop = os.path.join(os.path.expanduser("~"), "Desktop", "PDFs_Salvos")
if not os.path.exists(caminho_desktop):
    os.makedirs(caminho_desktop)

# atualizando a lista na tela e colocando as cores corretas para cada tipo de arquivo
def atualizar_lista():
    # limpando a lista atual
    for widget in frame_lista.winfo_children():
        widget.destroy()

    # texto de fundo caso não tenha nada selecionado
    if not arquivos_selecionados:
        lbl = ctk.CTkLabel(frame_lista, text="Arraste os arquivos aqui\nou use o botão abaixo.", text_color=TXT_SECUNDARIO)
        lbl.pack(pady=80)
        return

    for arquivo in arquivos_selecionados:
        nome_arquivo = os.path.basename(arquivo)
        ext = os.path.splitext(nome_arquivo)[1].lower()

        # pintando de acordo com a extensão do arquivo
        if ext in ['.doc', '.docx']:
            cor_bg = '#2b579a' # azul do word
        elif ext in ['.xls', '.xlsx']:
            cor_bg = '#217346' # verde do excel
        elif ext in ['.ppt', '.pptx']:
            cor_bg = '#d24726' # laranja do powerpoint
        else:
            cor_bg = BG_ELEVATED

        # caixinha de cada arquivo na lista
        row_frame = ctk.CTkFrame(frame_lista, fg_color=cor_bg, corner_radius=6, height=30)
        row_frame.pack(fill="x", pady=3, padx=3)

        lbl_nome = ctk.CTkLabel(row_frame, text=nome_arquivo, text_color="#ffffff", font=("Arial", 11, "bold"))
        lbl_nome.pack(side="left", padx=10, pady=2)

        # botao para excluir individualmente
        btn_remover = ctk.CTkButton(
            row_frame, text="X", width=20, height=20, fg_color="transparent",
            hover_color="#ff4444", text_color="#ffffff", font=("Arial", 10, "bold"),
            command=lambda arq=arquivo: remover_arquivo(arq)
        )
        btn_remover.pack(side="right", padx=5, pady=2)

# função para lidar quando eu soltar os arquivos na tela
def ao_soltar_arquivos(evento):
    arquivos = root.tk.splitlist(evento.data)
    for arquivo in arquivos:
        if arquivo not in arquivos_selecionados:
            arquivos_selecionados.append(arquivo)
    atualizar_lista()

# função para buscar os arquivos clicando caso eu não queira arrastar
def selecionar_arquivos():
    arquivos = filedialog.askopenfilenames(
        title="Selecione os arquivos",
        filetypes=[("Arquivos Office", "*.docx *.doc *.xlsx *.xls *.pptx *.ppt")]
    )
    for arquivo in arquivos:
        if arquivo not in arquivos_selecionados:
            arquivos_selecionados.append(arquivo)
    atualizar_lista()

# tirando o arquivo da lista caso eu tenha selecionado errado
def remover_arquivo(arquivo):
    if arquivo in arquivos_selecionados:
        arquivos_selecionados.remove(arquivo)
        atualizar_lista()

# onde a mágica acontece, convertendo com o código nativo do office
def converter_arquivos():
    if not arquivos_selecionados:
        label_status.configure(text="Nenhum arquivo adicionado!", text_color="#FF453A")
        return

    label_status.configure(text="Convertendo arquivos...", text_color=COR_PRINCIPAL)
    root.update()

    word = None
    excel = None
    ppt = None

    for caminho_arquivo in arquivos_selecionados:
        caminho_completo = os.path.abspath(caminho_arquivo)
        nome_arquivo = os.path.basename(caminho_completo)
        nome_sem_ext, ext = os.path.splitext(nome_arquivo)
        ext = ext.lower()

        arquivo_pdf = os.path.join(caminho_desktop, f"{nome_sem_ext}.pdf")

        try:
            if ext in ['.doc', '.docx']:
                if word is None:
                    word = win32com.client.Dispatch("Word.Application")
                    word.Visible = False
                doc = word.Documents.Open(caminho_completo)
                doc.SaveAs(arquivo_pdf, FileFormat=17) 
                doc.Close()

            elif ext in ['.xls', '.xlsx']:
                if excel is None:
                    excel = win32com.client.Dispatch("Excel.Application")
                    excel.Visible = False
                wb = excel.Workbooks.Open(caminho_completo)
                wb.ExportAsFixedFormat(0, arquivo_pdf)
                wb.Close(False)

            elif ext in ['.ppt', '.pptx']:
                if ppt is None:
                    ppt = win32com.client.Dispatch("PowerPoint.Application")
                presentation = ppt.Presentations.Open(caminho_completo, WithWindow=False)
                presentation.SaveAs(arquivo_pdf, 32)
                presentation.Close()

        except Exception as e:
            print(f"Erro no arquivo {nome_arquivo}: {e}")

    # fechando os processos para não consumir memoria
    if word: word.Quit()
    if excel: excel.Quit()
    if ppt: ppt.Quit()

    label_status.configure(text="Concluído! PDFs gerados.", text_color="#30D158")


# montando a interface enxuta
ctk.CTkLabel(root, text="Conversor Office para PDF", font=("Arial", 16, "bold"), text_color=TXT_TITULO).pack(pady=(15, 5))

# a própria lista agora serve como área de arrastar e soltar para economizar espaço
frame_lista = ctk.CTkScrollableFrame(root, fg_color=BG_ELEVATED, border_width=1, border_color="#38383A", width=340, height=300)
frame_lista.pack(pady=10, padx=20, fill="both", expand=True)

# habilitando o drag and drop na tela toda para facilitar
root.drop_target_register(DND_FILES)
root.dnd_bind('<<Drop>>', ao_soltar_arquivos)
frame_lista.drop_target_register(DND_FILES)
frame_lista.dnd_bind('<<Drop>>', ao_soltar_arquivos)

# botões na parte de baixo
frame_botoes = ctk.CTkFrame(root, fg_color="transparent")
frame_botoes.pack(pady=10)

btn_add = ctk.CTkButton(frame_botoes, text="+ Adicionar", width=130, height=35, command=selecionar_arquivos, fg_color=BG_ELEVATED, text_color=TXT_TITULO, hover_color="#38383A", font=("Arial", 12, "bold"))
btn_add.grid(row=0, column=0, padx=10)

btn_converter = ctk.CTkButton(frame_botoes, text="GERAR PDF", width=130, height=35, command=converter_arquivos, fg_color=COR_PRINCIPAL, hover_color="#007AFF", text_color="white", font=("Arial", 12, "bold"))
btn_converter.grid(row=0, column=1, padx=10)

# label onde aviso que terminou
label_status = ctk.CTkLabel(root, text="", font=("Arial", 12, "bold"))
label_status.pack(pady=(0, 10))

# chamo a primeira vez para mostrar o texto de arrastar arquivos
atualizar_lista()

root.mainloop()