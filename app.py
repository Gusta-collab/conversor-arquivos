import os
import customtkinter as ctk
from tkinter import filedialog
from tkinterdnd2 import TkinterDnD, DND_FILES
import win32com.client
from PIL import Image  # Nova biblioteca para converter as imagens

# unindo customtkinter com tkinterdnd2 para manter o arrastar e soltar funcionando no tema moderno
class TkinterDnD_CTk(ctk.CTk, TkinterDnD.DnDWrapper):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.TkdndVersion = TkinterDnD._require(self)

ctk.set_appearance_mode("Dark")

# cores base do nosso tema escuro
BG_ROOT = "#121212"
BG_ELEVATED = "#1e1e1e"
COR_PRINCIPAL = "#ff0000"
TXT_TITULO = "#FFFFFF"
TXT_SECUNDARIO = "#8E8E93"

root = TkinterDnD_CTk(fg_color=BG_ROOT)
root.title("Conversor PDF")

# configurando tamanho e centralizando a janela principal na tela
largura_janela = 400
altura_janela = 500

largura_tela = root.winfo_screenwidth()
altura_tela = root.winfo_screenheight()

pos_x = int((largura_tela / 2) - (largura_janela / 2))
pos_y = int((altura_tela / 2) - (altura_janela / 2))

root.geometry(f"{largura_janela}x{altura_janela}+{pos_x}+{pos_y}")
root.resizable(False, False)

# variaveis
arquivos_selecionados = []

# variavel e pasta padrão na área de trabalho onde os pdfs vão ser salvos inicialmente
caminho_destino = ctk.StringVar()
pasta_padrao = os.path.join(os.path.expanduser("~"), "Desktop", "PDFs_Salvos")
if not os.path.exists(pasta_padrao):
    os.makedirs(pasta_padrao)
caminho_destino.set(pasta_padrao)

# janela suspensa que mostra todos os formatos suportados
def abrir_janela_ajuda():
    janela_ajuda = ctk.CTkToplevel(root)
    janela_ajuda.title("Arquivos Suportados")
    
    largura_ajuda = 340
    altura_ajuda = 380
    pos_x_ajuda = int((largura_tela / 2) - (largura_ajuda / 2))
    pos_y_ajuda = int((altura_tela / 2) - (altura_ajuda / 2))
    janela_ajuda.geometry(f"{largura_ajuda}x{altura_ajuda}+{pos_x_ajuda}+{pos_y_ajuda}")
    
    janela_ajuda.resizable(False, False)
    janela_ajuda.configure(fg_color="#181818")
    janela_ajuda.transient(root)
    janela_ajuda.grab_set()

    ctk.CTkLabel(
        janela_ajuda,
        text="Formatos Suportados",
        font=("Arial", 15, "bold"),
        text_color=TXT_TITULO
    ).pack(pady=(15, 10))

    # caixa com as categorias e extensões permitidas atualizadas
    card_info = ctk.CTkFrame(janela_ajuda, fg_color=BG_ELEVATED, corner_radius=8)
    card_info.pack(fill="both", expand=True, padx=15, pady=(0, 15))

    texto_formatos = (
        "• Textos e Web (Azul):\n"
        "   .docx, .doc, .txt, .html, .htm, .xml\n\n"
        "• Excel (Verde):\n"
        "   .xlsx, .xls, .xlsb\n\n"
        "• Apresentações (Laranja):\n"
        "   .pptx, .ppt, .pps, .ppsx, .odp\n\n"
        "• Imagens (Roxo):\n"
        "   .png, .jpg, .jpeg, .bmp, .webp, .tiff"
    )

    ctk.CTkLabel(
        card_info,
        text=texto_formatos,
        font=("Arial", 12),
        text_color=TXT_SECUNDARIO,
        justify="left"
    ).pack(padx=15, pady=15, anchor="w")

    btn_fechar = ctk.CTkButton(
        janela_ajuda,
        text="Entendi",
        width=100,
        height=30,
        command=janela_ajuda.destroy,
        fg_color=COR_PRINCIPAL,
        hover_color="#BB0000",
        text_color="white",
        font=("Arial", 11, "bold"),
        corner_radius=6
    )
    btn_fechar.pack(pady=(0, 15))

# atualizando a lista na tela e colocando as cores corretas para cada tipo de arquivo
def atualizar_lista():
    for widget in frame_lista.winfo_children():
        widget.destroy()

    if not arquivos_selecionados:
        lbl = ctk.CTkLabel(frame_lista, text="Arraste os arquivos aqui\nou use o botão abaixo.", text_color=TXT_SECUNDARIO)
        lbl.pack(pady=80)
        return

    for arquivo in arquivos_selecionados:
        nome_arquivo = os.path.basename(arquivo)
        ext = os.path.splitext(nome_arquivo)[1].lower()

        if ext in ['.doc', '.docx', '.txt', '.html', '.htm', '.xml']:
            cor_bg = '#2b579a' # Azul Word / Textos
        elif ext in ['.xls', '.xlsx', '.xlsb']:
            cor_bg = '#217346' # Verde Excel
        elif ext in ['.ppt', '.pptx', '.pps', '.ppsx', '.odp']:
            cor_bg = '#d24726' # Laranja PowerPoint
        elif ext in ['.png', '.jpg', '.jpeg', '.bmp', '.webp', '.tiff']:
            cor_bg = '#8e44ad' # Roxo para Imagens
        else:
            cor_bg = BG_ELEVATED

        row_frame = ctk.CTkFrame(frame_lista, fg_color=cor_bg, corner_radius=6, height=30)
        row_frame.pack(fill="x", pady=3, padx=3)

        lbl_nome = ctk.CTkLabel(row_frame, text=nome_arquivo, text_color="#ffffff", font=("Arial", 11, "bold"))
        lbl_nome.pack(side="left", padx=10, pady=2)

        btn_remover = ctk.CTkButton(
            row_frame, text="X", width=20, height=20, fg_color="transparent",
            hover_color="#ff4444", text_color="#ffffff", font=("Arial", 10, "bold"),
            command=lambda arq=arquivo: remover_arquivo(arq)
        )
        btn_remover.pack(side="right", padx=5, pady=2)

# lidando quando eu soltar os arquivos na tela
def ao_soltar_arquivos(evento):
    arquivos = root.tk.splitlist(evento.data)
    for arquivo in arquivos:
        if arquivo not in arquivos_selecionados:
            arquivos_selecionados.append(arquivo)
    atualizar_lista()

# buscando arquivos pelo botão
def selecionar_arquivos():
    arquivos = filedialog.askopenfilenames(
        title="Selecione os arquivos",
        filetypes=[
            ("Todos os Suportados", "*.docx *.doc *.txt *.html *.htm *.xml *.xlsx *.xls *.xlsb *.pptx *.ppt *.pps *.ppsx *.odp *.png *.jpg *.jpeg *.bmp *.webp *.tiff"),
            ("Documentos e Textos", "*.docx *.doc *.txt *.html *.htm *.xml"),
            ("Planilhas", "*.xlsx *.xls *.xlsb"),
            ("Apresentações", "*.pptx *.ppt *.pps *.ppsx *.odp"),
            ("Imagens", "*.png *.jpg *.jpeg *.bmp *.webp *.tiff")
        ]
    )
    for arquivo in arquivos:
        if arquivo not in arquivos_selecionados:
            arquivos_selecionados.append(arquivo)
    atualizar_lista()

def remover_arquivo(arquivo):
    if arquivo in arquivos_selecionados:
        arquivos_selecionados.remove(arquivo)
        atualizar_lista()

# função pra escolher onde vamos salvar os pdfs
def selecionar_pasta_destino():
    pasta_escolhida = filedialog.askdirectory(title="Onde quer salvar os PDFs?")
    if pasta_escolhida:
        caminho_destino.set(pasta_escolhida)

def converter_arquivos():
    if not arquivos_selecionados:
        label_status.configure(text="Nenhum arquivo adicionado!", text_color="#FF453A")
        return

    pasta_salvar = caminho_destino.get()
    if not os.path.exists(pasta_salvar):
        os.makedirs(pasta_salvar)

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

        arquivo_pdf = os.path.join(pasta_salvar, f"{nome_sem_ext}.pdf")

        try:
            # Documentos de texto, web e xml roteados para o Word
            if ext in ['.doc', '.docx', '.txt', '.html', '.htm', '.xml']:
                if word is None:
                    word = win32com.client.Dispatch("Word.Application")
                    word.Visible = False
                    word.DisplayAlerts = 0 # Evita caixas de diálogo ao abrir TXT/HTML/XML
                doc = word.Documents.Open(caminho_completo)
                doc.SaveAs(arquivo_pdf, FileFormat=17) 
                doc.Close()

            # Planilhas roteadas para o Excel
            elif ext in ['.xls', '.xlsx', '.xlsb']:
                if excel is None:
                    excel = win32com.client.Dispatch("Excel.Application")
                    excel.Visible = False
                    excel.DisplayAlerts = False
                wb = excel.Workbooks.Open(caminho_completo)
                wb.ExportAsFixedFormat(0, arquivo_pdf)
                wb.Close(False)

            # Slides roteados para o PowerPoint
            elif ext in ['.ppt', '.pptx', '.pps', '.ppsx', '.odp']:
                if ppt is None:
                    ppt = win32com.client.Dispatch("PowerPoint.Application")
                presentation = ppt.Presentations.Open(caminho_completo, WithWindow=False)
                presentation.SaveAs(arquivo_pdf, 32)
                presentation.Close()
                
            # Imagens tratadas pela biblioteca Pillow
            elif ext in ['.png', '.jpg', '.jpeg', '.bmp', '.webp', '.tiff']:
                imagem = Image.open(caminho_completo)
                # Converte para RGB (necessário para salvar como PDF pois formatos como PNG podem ter fundo transparente/RGBA)
                imagem_convertida = imagem.convert('RGB')
                imagem_convertida.save(arquivo_pdf)

        except Exception as e:
            print(f"Erro no arquivo {nome_arquivo}: {e}")

    # Encerra os processos
    if word: word.Quit()
    if excel: excel.Quit()
    if ppt: ppt.Quit()

    label_status.configure(text="Concluído! PDFs gerados.", text_color="#30D158")


# cabecalho com titulo e botao circular de ajuda
frame_topo = ctk.CTkFrame(root, fg_color="transparent")
frame_topo.pack(fill="x", padx=20, pady=(15, 5))

ctk.CTkLabel(
    frame_topo,
    text="Conversor Office para PDF",
    font=("Arial", 16, "bold"),
    text_color=TXT_TITULO
).pack(side="left")

# botao redondo com ponto de interrogacao
btn_ajuda = ctk.CTkButton(
    frame_topo,
    text="?",
    width=26,
    height=26,
    corner_radius=13,
    fg_color=BG_ELEVATED,
    hover_color="#38383A",
    text_color=TXT_TITULO,
    font=("Arial", 13, "bold"),
    command=abrir_janela_ajuda
)
btn_ajuda.pack(side="right")

# a própria lista serve como área de arrastar e soltar
frame_lista = ctk.CTkScrollableFrame(root, fg_color=BG_ELEVATED, border_width=1, border_color="#38383A", width=340, height=240)
frame_lista.pack(pady=10, padx=20, fill="both", expand=True)

root.drop_target_register(DND_FILES)
root.dnd_bind('<<Drop>>', ao_soltar_arquivos)
frame_lista.drop_target_register(DND_FILES)
frame_lista.dnd_bind('<<Drop>>', ao_soltar_arquivos)

# painel de escolha do destino
frame_destino = ctk.CTkFrame(root, fg_color="transparent")
frame_destino.pack(fill="x", padx=25, pady=5)

ctk.CTkLabel(frame_destino, text="Salvar em:", font=("Arial", 11, "bold"), text_color=TXT_SECUNDARIO).pack(anchor="w")

frame_destino_input = ctk.CTkFrame(frame_destino, fg_color="transparent")
frame_destino_input.pack(fill="x")

lbl_caminho = ctk.CTkLabel(frame_destino_input, textvariable=caminho_destino, font=("Arial", 10), text_color=TXT_TITULO, fg_color=BG_ELEVATED, corner_radius=6, height=28, anchor="w", padx=10)
lbl_caminho.pack(side="left", fill="x", expand=True, padx=(0, 5))

btn_mudar_pasta = ctk.CTkButton(frame_destino_input, text="Mudar", width=60, height=28, command=selecionar_pasta_destino, fg_color=BG_ELEVATED, text_color=TXT_TITULO, hover_color="#38383A", font=("Arial", 11, "bold"))
btn_mudar_pasta.pack(side="right")

# botões na parte de baixo
frame_botoes = ctk.CTkFrame(root, fg_color="transparent")
frame_botoes.pack(pady=15)

btn_add = ctk.CTkButton(frame_botoes, text="+ Adicionar", width=130, height=35, command=selecionar_arquivos, fg_color=BG_ELEVATED, text_color=TXT_TITULO, hover_color="#38383A", font=("Arial", 12, "bold"))
btn_add.grid(row=0, column=0, padx=10)

btn_converter = ctk.CTkButton(frame_botoes, text="GERAR PDF", width=130, height=35, command=converter_arquivos, fg_color=COR_PRINCIPAL, hover_color="#BB0000", text_color="white", font=("Arial", 12, "bold"))
btn_converter.grid(row=0, column=1, padx=10)

label_status = ctk.CTkLabel(root, text="", font=("Arial", 12, "bold"))
label_status.pack(pady=(0, 10))

atualizar_lista()

root.mainloop()