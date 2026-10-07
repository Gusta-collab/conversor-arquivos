import os
import customtkinter as ctk
from tkinter import filedialog
from tkinterdnd2 import TkinterDnD, DND_FILES
import win32com.client
from PIL import Image  
from pdf2docx import Converter

# unindo customtkinter com tkinterdnd2 para manter o arrastar e soltar funcionando no tema moderno
class TkinterDnD_CTk(ctk.CTk, TkinterDnD.DnDWrapper):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.TkdndVersion = TkinterDnD._require(self)

ctk.set_appearance_mode("Dark")

# cores base do tema escuro
BG_ROOT = "#121212"
BG_ELEVATED = "#1e1e1e"
COR_PRINCIPAL = "#ff0000"
TXT_TITULO = "#FFFFFF"
TXT_SECUNDARIO = "#8E8E93"

root = TkinterDnD_CTk(fg_color=BG_ROOT)
root.title("Conversor de Arquivos")

# configurando tamanho e centralizando a janela principal na tela
largura_janela = 400
altura_janela = 500

largura_tela = root.winfo_screenwidth()
altura_tela = root.winfo_screenheight()

pos_x = int((largura_tela / 2) - (largura_janela / 2))
pos_y = int((altura_tela / 2) - (altura_janela / 2))

root.geometry(f"{largura_janela}x{altura_janela}+{pos_x}+{pos_y}")
root.resizable(False, False)

# variaveis globais e listas de extensoes permitidas
arquivos_para_pdf = []
arquivos_de_pdf = []

EXTENSOES_PARA_PDF = ['.doc', '.docx', '.txt', '.html', '.htm', '.xml', '.xls', '.xlsx', '.xlsb', '.ppt', '.pptx', '.pps', '.ppsx', '.odp', '.png', '.jpg', '.jpeg', '.bmp', '.webp', '.tiff']

# variavel e pasta padrão na área de trabalho onde os arquivos vão ser salvos inicialmente
caminho_destino = ctk.StringVar()
pasta_padrao = os.path.join(os.path.expanduser("~"), "Desktop", "Arquivos_Convertidos")
if not os.path.exists(pasta_padrao):
    os.makedirs(pasta_padrao)
caminho_destino.set(pasta_padrao)

# ==========================================
# 1. INTERFACE GRÁFICA
# ==========================================

frame_topo = ctk.CTkFrame(root, fg_color="transparent")
frame_topo.pack(fill="x", padx=20, pady=(15, 5))

ctk.CTkLabel(frame_topo, text="Conversor de Arquivos", font=("Arial", 16, "bold"), text_color=TXT_TITULO).pack(side="left")

def abrir_janela_ajuda():
    janela_ajuda = ctk.CTkToplevel(root)
    janela_ajuda.title("Arquivos Suportados")
    
    largura_ajuda = 340
    altura_ajuda = 440
    pos_x_ajuda = int((largura_tela / 2) - (largura_ajuda / 2))
    pos_y_ajuda = int((altura_tela / 2) - (altura_ajuda / 2))
    janela_ajuda.geometry(f"{largura_ajuda}x{altura_ajuda}+{pos_x_ajuda}+{pos_y_ajuda}")
    
    janela_ajuda.resizable(False, False)
    janela_ajuda.configure(fg_color="#181818")
    janela_ajuda.transient(root)
    janela_ajuda.grab_set()

    ctk.CTkLabel(janela_ajuda, text="Formatos Suportados", font=("Arial", 15, "bold"), text_color=TXT_TITULO).pack(pady=(15, 10))

    card_info = ctk.CTkFrame(janela_ajuda, fg_color=BG_ELEVATED, corner_radius=8)
    card_info.pack(fill="both", expand=True, padx=15, pady=(0, 15))

    texto_formatos = (
        "• Textos e Web para PDF (Azul):\n   .docx, .doc, .txt, .html, .htm, .xml\n\n"
        "• Excel para PDF (Verde):\n   .xlsx, .xls, .xlsb\n\n"
        "• Apresentações para PDF (Laranja):\n   .pptx, .ppt, .pps, .ppsx, .odp\n\n"
        "• Imagens para PDF (Roxo):\n   .png, .jpg, .jpeg, .bmp, .webp, .tiff\n\n"
        "• PDF para Word (Vermelho):\n   .pdf"
    )

    ctk.CTkLabel(card_info, text=texto_formatos, font=("Arial", 12), text_color=TXT_SECUNDARIO, justify="left").pack(padx=15, pady=15, anchor="w")
    ctk.CTkButton(janela_ajuda, text="Entendi", width=100, height=30, command=janela_ajuda.destroy, fg_color=COR_PRINCIPAL, hover_color="#BB0000", text_color="white", font=("Arial", 11, "bold"), corner_radius=6).pack(pady=(0, 15))

btn_ajuda = ctk.CTkButton(frame_topo, text="?", width=26, height=26, corner_radius=13, fg_color=BG_ELEVATED, hover_color="#38383A", text_color=TXT_TITULO, font=("Arial", 13, "bold"), command=abrir_janela_ajuda)
btn_ajuda.pack(side="right")

# Sistema de Abas
tabview = ctk.CTkTabview(root, width=340, fg_color=BG_ELEVATED, segmented_button_selected_color=COR_PRINCIPAL, segmented_button_selected_hover_color="#BB0000", text_color=TXT_TITULO)
tabview.pack(pady=5, padx=20, fill="both", expand=True)

aba_para_pdf = tabview.add("Para PDF")
aba_de_pdf = tabview.add("PDF para Word")

# ----- CONTEÚDO DA ABA 1 (PARA PDF) -----
frame_lista_para_pdf = ctk.CTkScrollableFrame(aba_para_pdf, fg_color=BG_ROOT, border_width=1, border_color="#38383A")
frame_lista_para_pdf.pack(fill="both", expand=True, padx=5, pady=5)

frame_botoes_1 = ctk.CTkFrame(aba_para_pdf, fg_color="transparent")
frame_botoes_1.pack(pady=10)

btn_add_1 = ctk.CTkButton(frame_botoes_1, text="+ Adicionar", width=120, height=35, fg_color=BG_ELEVATED, text_color=TXT_TITULO, hover_color="#38383A", font=("Arial", 12, "bold"))
btn_add_1.pack(side="left", padx=10)

btn_converter_1 = ctk.CTkButton(frame_botoes_1, text="GERAR PDF", width=120, height=35, fg_color=COR_PRINCIPAL, hover_color="#BB0000", text_color="white", font=("Arial", 12, "bold"))
btn_converter_1.pack(side="right", padx=10)

# ----- CONTEÚDO DA ABA 2 (DE PDF PARA WORD) -----
frame_lista_de_pdf = ctk.CTkScrollableFrame(aba_de_pdf, fg_color=BG_ROOT, border_width=1, border_color="#38383A")
frame_lista_de_pdf.pack(fill="both", expand=True, padx=5, pady=5)

frame_botoes_2 = ctk.CTkFrame(aba_de_pdf, fg_color="transparent")
frame_botoes_2.pack(pady=10)

btn_add_2 = ctk.CTkButton(frame_botoes_2, text="+ Adicionar", width=120, height=35, fg_color=BG_ELEVATED, text_color=TXT_TITULO, hover_color="#38383A", font=("Arial", 12, "bold"))
btn_add_2.pack(side="left", padx=10)

btn_converter_2 = ctk.CTkButton(frame_botoes_2, text="GERAR WORD", width=120, height=35, fg_color=COR_PRINCIPAL, hover_color="#BB0000", text_color="white", font=("Arial", 12, "bold"))
btn_converter_2.pack(side="right", padx=10)

# ----- RODAPÉ GLOBAL (Destino e Status) -----
frame_destino = ctk.CTkFrame(root, fg_color="transparent")
frame_destino.pack(fill="x", padx=25, pady=5)

ctk.CTkLabel(frame_destino, text="Salvar em:", font=("Arial", 11, "bold"), text_color=TXT_SECUNDARIO).pack(anchor="w")

frame_destino_input = ctk.CTkFrame(frame_destino, fg_color="transparent")
frame_destino_input.pack(fill="x")

lbl_caminho = ctk.CTkLabel(frame_destino_input, textvariable=caminho_destino, font=("Arial", 10), text_color=TXT_TITULO, fg_color=BG_ELEVATED, corner_radius=6, height=28, anchor="w", padx=10)
lbl_caminho.pack(side="left", fill="x", expand=True, padx=(0, 5))

def selecionar_pasta_destino():
    pasta_escolhida = filedialog.askdirectory(title="Onde quer salvar os arquivos?")
    if pasta_escolhida:
        caminho_destino.set(pasta_escolhida)

btn_mudar_pasta = ctk.CTkButton(frame_destino_input, text="Mudar", width=60, height=28, command=selecionar_pasta_destino, fg_color=BG_ELEVATED, text_color=TXT_TITULO, hover_color="#38383A", font=("Arial", 11, "bold"))
btn_mudar_pasta.pack(side="right")

label_status = ctk.CTkLabel(root, text="", font=("Arial", 12, "bold"))
label_status.pack(pady=(0, 10))


# ==========================================
# 2. FUNÇÕES GERAIS E ARRASTAR/SOLTAR INTELIGENTE
# ==========================================

# A função deteta a aba ativa e insere os ficheiros no sítio certo!
def ao_soltar_arquivos(evento):
    aba_ativa = tabview.get()
    arquivos_arrastados = root.tk.splitlist(evento.data)
    
    if aba_ativa == "Para PDF":
        for arquivo in arquivos_arrastados:
            ext = os.path.splitext(arquivo)[1].lower()
            if ext in EXTENSOES_PARA_PDF and arquivo not in arquivos_para_pdf:
                arquivos_para_pdf.append(arquivo)
        atualizar_lista_para_pdf()
        
    elif aba_ativa == "PDF para Word":
        for arquivo in arquivos_arrastados:
            if arquivo.lower().endswith('.pdf') and arquivo not in arquivos_de_pdf:
                arquivos_de_pdf.append(arquivo)
        atualizar_lista_de_pdf()


# ==========================================
# 3. FUNÇÕES PARA A ABA 1 (PARA PDF)
# ==========================================
def atualizar_lista_para_pdf():
    for widget in frame_lista_para_pdf.winfo_children():
        widget.destroy()

    if not arquivos_para_pdf:
        lbl = ctk.CTkLabel(frame_lista_para_pdf, text="Arraste os arquivos originais aqui\nou use o botão abaixo.", text_color=TXT_SECUNDARIO)
        lbl.pack(pady=80)
        return

    for arquivo in arquivos_para_pdf:
        nome_arquivo = os.path.basename(arquivo)
        ext = os.path.splitext(nome_arquivo)[1].lower()

        if ext in ['.doc', '.docx', '.txt', '.html', '.htm', '.xml']:
            cor_bg = '#2b579a' 
        elif ext in ['.xls', '.xlsx', '.xlsb']:
            cor_bg = '#217346' 
        elif ext in ['.ppt', '.pptx', '.pps', '.ppsx', '.odp']:
            cor_bg = '#d24726' 
        elif ext in ['.png', '.jpg', '.jpeg', '.bmp', '.webp', '.tiff']:
            cor_bg = '#8e44ad' 
        else:
            cor_bg = BG_ELEVATED

        row_frame = ctk.CTkFrame(frame_lista_para_pdf, fg_color=cor_bg, corner_radius=6, height=30)
        row_frame.pack(fill="x", pady=3, padx=3)

        lbl_nome = ctk.CTkLabel(row_frame, text=nome_arquivo, text_color="#ffffff", font=("Arial", 11, "bold"))
        lbl_nome.pack(side="left", padx=10, pady=2)

        btn_remover = ctk.CTkButton(row_frame, text="X", width=20, height=20, fg_color="transparent", hover_color="#ff4444", text_color="#ffffff", font=("Arial", 10, "bold"), command=lambda arq=arquivo: remover_para_pdf(arq))
        btn_remover.pack(side="right", padx=5, pady=2)

def adicionar_para_pdf():
    arquivos = filedialog.askopenfilenames(
        title="Selecione os arquivos originais",
        filetypes=[
            ("Todos os Suportados", "*.docx *.doc *.txt *.html *.htm *.xml *.xlsx *.xls *.xlsb *.pptx *.ppt *.pps *.ppsx *.odp *.png *.jpg *.jpeg *.bmp *.webp *.tiff"),
            ("Documentos", "*.docx *.doc *.txt *.html *.htm *.xml"),
            ("Planilhas", "*.xlsx *.xls *.xlsb"),
            ("Apresentações", "*.pptx *.ppt *.pps *.ppsx *.odp"),
            ("Imagens", "*.png *.jpg *.jpeg *.bmp *.webp *.tiff")
        ]
    )
    for arquivo in arquivos:
        if arquivo not in arquivos_para_pdf:
            arquivos_para_pdf.append(arquivo)
    atualizar_lista_para_pdf()

def remover_para_pdf(arquivo):
    if arquivo in arquivos_para_pdf:
        arquivos_para_pdf.remove(arquivo)
        atualizar_lista_para_pdf()

def converter_para_pdf():
    if not arquivos_para_pdf:
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

    for caminho_arquivo in arquivos_para_pdf:
        caminho_completo = os.path.abspath(caminho_arquivo)
        nome_arquivo = os.path.basename(caminho_completo)
        nome_sem_ext, ext = os.path.splitext(nome_arquivo)
        ext = ext.lower()

        arquivo_pdf = os.path.join(pasta_salvar, f"{nome_sem_ext}.pdf")

        try:
            if ext in ['.doc', '.docx', '.txt', '.html', '.htm', '.xml']:
                if word is None:
                    word = win32com.client.Dispatch("Word.Application")
                    word.Visible = False
                    word.DisplayAlerts = 0 
                
                caminho_abrir = caminho_completo
                txt_temporario = None
                
                if ext == '.xml':
                    import shutil
                    txt_temporario = os.path.join(pasta_salvar, f"{nome_sem_ext}_temporario.txt")
                    shutil.copy2(caminho_completo, txt_temporario)
                    caminho_abrir = txt_temporario

                doc = word.Documents.Open(caminho_abrir)
                doc.SaveAs(arquivo_pdf, FileFormat=17) 
                doc.Close()
                
                if txt_temporario and os.path.exists(txt_temporario):
                    os.remove(txt_temporario)

            elif ext in ['.xls', '.xlsx', '.xlsb']:
                if excel is None:
                    excel = win32com.client.Dispatch("Excel.Application")
                    excel.Visible = False
                    excel.DisplayAlerts = False
                wb = excel.Workbooks.Open(caminho_completo)
                wb.ExportAsFixedFormat(0, arquivo_pdf)
                wb.Close(False)

            elif ext in ['.ppt', '.pptx', '.pps', '.ppsx', '.odp']:
                if ppt is None:
                    ppt = win32com.client.Dispatch("PowerPoint.Application")
                presentation = ppt.Presentations.Open(caminho_completo, WithWindow=False)
                presentation.SaveAs(arquivo_pdf, 32)
                presentation.Close()
                
            elif ext in ['.png', '.jpg', '.jpeg', '.bmp', '.webp', '.tiff']:
                imagem = Image.open(caminho_completo)
                imagem_convertida = imagem.convert('RGB')
                imagem_convertida.save(arquivo_pdf)

        except Exception as e:
            print(f"Erro no arquivo {nome_arquivo}: {e}")

    if word: word.Quit()
    if excel: excel.Quit()
    if ppt: ppt.Quit()

    label_status.configure(text="Concluído! PDFs gerados.", text_color="#30D158")


# ==========================================
# 4. FUNÇÕES PARA A ABA 2 (DE PDF PARA WORD)
# ==========================================
def atualizar_lista_de_pdf():
    for widget in frame_lista_de_pdf.winfo_children():
        widget.destroy()

    if not arquivos_de_pdf:
        lbl = ctk.CTkLabel(frame_lista_de_pdf, text="Arraste os arquivos PDF aqui\nou use o botão abaixo.", text_color=TXT_SECUNDARIO)
        lbl.pack(pady=80)
        return

    for arquivo in arquivos_de_pdf:
        nome_arquivo = os.path.basename(arquivo)
        row_frame = ctk.CTkFrame(frame_lista_de_pdf, fg_color="#b30b00", corner_radius=6, height=30)
        row_frame.pack(fill="x", pady=3, padx=3)

        lbl_nome = ctk.CTkLabel(row_frame, text=nome_arquivo, text_color="#ffffff", font=("Arial", 11, "bold"))
        lbl_nome.pack(side="left", padx=10, pady=2)

        btn_remover = ctk.CTkButton(row_frame, text="X", width=20, height=20, fg_color="transparent", hover_color="#ff4444", text_color="#ffffff", font=("Arial", 10, "bold"), command=lambda arq=arquivo: remover_de_pdf(arq))
        btn_remover.pack(side="right", padx=5, pady=2)

def adicionar_de_pdf():
    arquivos = filedialog.askopenfilenames(
        title="Selecione os arquivos PDF",
        filetypes=[("Arquivos PDF", "*.pdf")]
    )
    for arquivo in arquivos:
        if arquivo not in arquivos_de_pdf:
            arquivos_de_pdf.append(arquivo)
    atualizar_lista_de_pdf()

def remover_de_pdf(arquivo):
    if arquivo in arquivos_de_pdf:
        arquivos_de_pdf.remove(arquivo)
        atualizar_lista_de_pdf()

def converter_de_pdf():
    if not arquivos_de_pdf:
        label_status.configure(text="Nenhum arquivo adicionado!", text_color="#FF453A")
        return

    pasta_salvar = caminho_destino.get()
    if not os.path.exists(pasta_salvar):
        os.makedirs(pasta_salvar)

    label_status.configure(text="Convertendo para Word...", text_color=COR_PRINCIPAL)
    root.update()

    for pdf_path in arquivos_de_pdf:
        nome_sem_ext = os.path.splitext(os.path.basename(pdf_path))[0]
        docx_path = os.path.join(pasta_salvar, f"{nome_sem_ext}.docx")
        
        try:
            cv = Converter(pdf_path)
            cv.convert(docx_path, start=0, end=None)
            cv.close()
        except Exception as e:
            print(f"Erro na conversão de {nome_sem_ext}: {e}")

    label_status.configure(text="Concluído! Documentos Word gerados.", text_color="#30D158")


# ==========================================
# 5. LIGAÇÕES FINAIS E ARRANQUE DO SISTEMA
# ==========================================
btn_add_1.configure(command=adicionar_para_pdf)
btn_converter_1.configure(command=converter_para_pdf)

btn_add_2.configure(command=adicionar_de_pdf)
btn_converter_2.configure(command=converter_de_pdf)

# Regista a janela inteira como área de Drag and Drop!
root.drop_target_register(DND_FILES)
root.dnd_bind('<<Drop>>', ao_soltar_arquivos)

atualizar_lista_para_pdf()
atualizar_lista_de_pdf()

root.mainloop()