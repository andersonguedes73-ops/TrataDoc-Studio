import fitz  # PyMuPDF
import pytesseract
import re
import spacy
import os
import sys
import threading
import platform
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import customtkinter as ctk  # MOTOR VISUAL
from PIL import Image, ImageTk

# --- DETECÇÃO DE SISTEMA ---
SISTEMA = platform.system() # 'Windows' ou 'Darwin' (Mac)

# --- CONFIGURAÇÃO DE AMBIENTE ---
def obter_raiz():
    if getattr(sys, 'frozen', False): return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))

PASTA_RAIZ = obter_raiz()

# Configuração Tesseract (Apenas Windows)
if SISTEMA == "Windows":
    CAMINHO_TESS = os.path.join(PASTA_RAIZ, "Tesseract-OCR", "tesseract.exe")
    pytesseract.pytesseract.tesseract_cmd = CAMINHO_TESS
    os.environ['TESSDATA_PREFIX'] = os.path.join(PASTA_RAIZ, "Tesseract-OCR", "tessdata")

# --- CARREGAMENTO INTELIGENTE DA IA (Para garantir Status: ATIVO) ---
def carregar_ia():
    # 1. Tenta carregar do jeito padrão
    try:
        return spacy.load("pt_core_news_md")
    except:
        pass

    # 2. Busca em pastas locais (Essencial para o Executável .exe)
    caminhos_tentar = [
        os.path.join(PASTA_RAIZ, "pt_core_news_md"),
        os.path.join(PASTA_RAIZ, "_internal", "pt_core_news_md"),
        os.path.join(PASTA_RAIZ, "_internal", "pt_core_news_md", "pt_core_news_md")
    ]

    for caminho in caminhos_tentar:
        if os.path.exists(caminho):
            try:
                return spacy.load(caminho)
            except:
                continue
    return None

nlp = carregar_ia()

# Configuração Padrão do Tema Visual
ctk.set_appearance_mode("Light")  
ctk.set_default_color_theme("blue")  

class TrataDocApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        # Título dinâmico
        desc_sis = "MacOS Vision Engine" if SISTEMA == "Darwin" else "Windows Tesseract"
        self.title(f"TrataDoc Studio v8.2 - {desc_sis}")
        self.geometry("1450x950")
        
        # --- ESTRUTURA DE DADOS ---
        self.dados = {
            "ocr": {"entrada": [], "prontos": []},
            "merge": {"entrada": [], "prontos": []},
            "tarja": {"entrada": [], "prontos": []}
        }
        self.pagina_atual = 0
        self.zoom_level = 0.8 
        self.doc_aberto = None
        self.caminho_atual = None
        
        self.modo_tarja_manual = False
        self.rect_id = None
        self.start_x = 0
        self.start_y = 0

        # --- NOVA ARQUITETURA DE GRID ---
        self.grid_rowconfigure(0, weight=0) # Barra Superior (Fixa)
        self.grid_rowconfigure(1, weight=1) # Área de Trabalho
        self.grid_rowconfigure(2, weight=0) # Barra de Status
        
        self.grid_columnconfigure(0, weight=0) # Sidebar
        self.grid_columnconfigure(1, weight=1) # Área Ativa

        # 1. TOP BAR
        self.setup_topbar()

        # 2. ÁREA DE TRABALHO PRINCIPAL
        self.area_trabalho = ctk.CTkFrame(self, fg_color="transparent", corner_radius=0)
        self.area_trabalho.grid(row=1, column=1, sticky="nsew")
        self.area_trabalho.grid_rowconfigure(0, weight=1)
        self.area_trabalho.grid_columnconfigure(0, weight=1) # Painel Esquerdo
        self.area_trabalho.grid_columnconfigure(1, weight=2) # Visualizador de PDF
        
        # 3. MENU LATERAL
        self.setup_sidebar()

        # 4. ÁREAS DE CONTEÚDO 
        self.frame_ocr = ctk.CTkScrollableFrame(self.area_trabalho, corner_radius=0, fg_color="transparent")
        self.frame_merge = ctk.CTkScrollableFrame(self.area_trabalho, corner_radius=0, fg_color="transparent")
        self.frame_tarja = ctk.CTkScrollableFrame(self.area_trabalho, corner_radius=0, fg_color="transparent")

        self.setup_ocr_tab()
        self.setup_merge_tab()
        self.setup_tarja_tab()

        # 5. VISUALIZADOR
        self.frame_view = ctk.CTkFrame(self.area_trabalho, corner_radius=0, fg_color="#f2f2f2")
        self.frame_view.grid(row=0, column=1, sticky="nsew", padx=2)
        self.setup_viewer()

        # 6. BARRA DE STATUS
        self.f_status = tk.Frame(self, bd=1, relief=tk.SUNKEN, bg="#e9ecef")
        self.f_status.grid(row=2, column=0, columnspan=2, sticky="ew")
        
        status_ia = "🧠 Motor IA: ATIVO" if nlp else "⚠️ Motor IA: DESATIVADO"
        cor_ia = "#28a745" if nlp else "#dc3545"
        self.lbl_status_ia = tk.Label(self.f_status, text=status_ia, font=("Arial", 9, "bold"), fg=cor_ia, bg="#e9ecef")
        self.lbl_status_ia.pack(side="left", padx=10)
        
        self.progress = ttk.Progressbar(self.f_status, orient="horizontal", mode="determinate")
        self.progress.pack(side="right", fill="x", expand=True, padx=5)

        # Inicia na Aba Principal e esconde a sidebar
        self.select_frame_by_name("tarja")
        self.sidebar_frame.grid_remove() 
        self.sidebar_visible = False

        # Radar Global: Lê a posição do mouse na tela inteira
        self.bind_all("<Motion>", self.check_mouse_position)

    # ==========================================
    # BARRA SUPERIOR E MENU (RADAR DE FLUIDEZ)
    # ==========================================
    def setup_topbar(self):
        self.topbar = ctk.CTkFrame(self, height=50, corner_radius=0, fg_color="#102a43")
        self.topbar.grid(row=0, column=0, columnspan=2, sticky="ew")
        
        self.btn_menu = ctk.CTkButton(self.topbar, text="☰", width=50, font=("Segoe UI", 24), fg_color="transparent", text_color="white", hover_color="#243b55")
        self.btn_menu.pack(side="left", padx=5, pady=5)
        
        self.btn_menu.bind("<Enter>", self.show_sidebar)
        
        self.lbl_title = ctk.CTkLabel(self.topbar, text="TrataDoc Studio", font=("Segoe UI", 22, "bold"), text_color="#00d2ff")
        self.lbl_title.pack(side="left", padx=10)
        
        self.lbl_sub = ctk.CTkLabel(self.topbar, text="|  Corregedoria MPO", font=("Segoe UI", 16), text_color="#e0e0e0")
        self.lbl_sub.pack(side="left")

    def setup_sidebar(self):
        self.sidebar_frame = ctk.CTkFrame(self, width=250, corner_radius=0, fg_color="#f8f9fa", border_width=1, border_color="#dee2e6")
        self.sidebar_frame.grid(row=1, column=0, sticky="nsew")

        self.btn_tab_tarja = ctk.CTkButton(self.sidebar_frame, text="⬛ Ocultação de Dados", corner_radius=0, height=55, anchor="w", fg_color="transparent", text_color="black", font=("Segoe UI", 14), hover_color="#e2e6ea", command=lambda: self.select_frame_by_name("tarja"))
        self.btn_tab_tarja.pack(fill="x", pady=(10,0))
        
        self.btn_tab_merge = ctk.CTkButton(self.sidebar_frame, text="🔗 Unificar Arquivos", corner_radius=0, height=55, anchor="w", fg_color="transparent", text_color="black", font=("Segoe UI", 14), hover_color="#e2e6ea", command=lambda: self.select_frame_by_name("merge"))
        self.btn_tab_merge.pack(fill="x")
        
        self.btn_tab_ocr = ctk.CTkButton(self.sidebar_frame, text="📄 OCR Pesquisável", corner_radius=0, height=55, anchor="w", fg_color="transparent", text_color="black", font=("Segoe UI", 14), hover_color="#e2e6ea", command=lambda: self.select_frame_by_name("ocr"))
        self.btn_tab_ocr.pack(fill="x")

    def show_sidebar(self, event=None):
        if not self.sidebar_visible:
            self.sidebar_frame.grid()
            self.sidebar_frame.lift()
            self.sidebar_visible = True

    def check_mouse_position(self, event):
        if self.sidebar_visible:
            x = self.winfo_pointerx() - self.winfo_rootx()
            y = self.winfo_pointery() - self.winfo_rooty()
            if x > 250 or (y < 50 and x > 65):
                self.sidebar_frame.grid_remove()
                self.sidebar_visible = False

    def select_frame_by_name(self, name):
        self.btn_tab_tarja.configure(fg_color="#e2e6ea" if name == "tarja" else "transparent")
        self.btn_tab_merge.configure(fg_color="#e2e6ea" if name == "merge" else "transparent")
        self.btn_tab_ocr.configure(fg_color="#e2e6ea" if name == "ocr" else "transparent")

        if name == "tarja": self.frame_tarja.grid(row=0, column=0, sticky="nsew", padx=10, pady=10)
        else: self.frame_tarja.grid_forget()
        
        if name == "merge": self.frame_merge.grid(row=0, column=0, sticky="nsew", padx=10, pady=10)
        else: self.frame_merge.grid_forget()
        
        if name == "ocr": self.frame_ocr.grid(row=0, column=0, sticky="nsew", padx=10, pady=10)
        else: self.frame_ocr.grid_forget()

    # ==========================================
    # CONTEÚDO DAS ABAS 
    # ==========================================
    def criar_lista_ui(self, parent, h=5):
        f = tk.Frame(parent, bg="white", bd=1, relief="solid")
        listbox = tk.Listbox(f, height=h, font=("Segoe UI", 10), bg="white", fg="#333", borderwidth=0, highlightthickness=0, selectbackground="#3B8ED0", selectforeground="white", exportselection=False)
        scroll = ttk.Scrollbar(f, orient="vertical", command=listbox.yview)
        listbox.configure(yscrollcommand=scroll.set)
        scroll.pack(side="right", fill="y")
        listbox.pack(side="left", fill="both", expand=True, padx=5, pady=5)
        f.pack(fill="x", pady=5)
        return listbox

    def setup_tarja_tab(self):
        ctk.CTkLabel(self.frame_tarja, text="Ocultação de Dados", font=("Segoe UI", 22, "bold")).pack(anchor="w", pady=(10, 0))
        ctk.CTkLabel(self.frame_tarja, text="Mapeamento automático e auxílio na proteção de dados e informações sensíveis.", text_color="gray", wraplength=400, justify="left").pack(anchor="w", pady=(0, 20))

        ctk.CTkButton(self.frame_tarja, text="📂 Selecionar Documentos", font=("Segoe UI", 12, "bold"), command=lambda: self.importar("tarja", self.lst_tarja)).pack(anchor="w", pady=5)
        self.lst_tarja = self.criar_lista_ui(self.frame_tarja)
        self.lst_tarja.bind('<<ListboxSelect>>', lambda e: self.preview_selecao(e, "tarja"))

        f_btns = ctk.CTkFrame(self.frame_tarja, fg_color="transparent")
        f_btns.pack(anchor="w", pady=5)
        ctk.CTkButton(f_btns, text="Remover", fg_color="#dc3545", hover_color="#c82333", width=100, font=("Segoe UI", 12, "bold"), command=lambda: self.excluir_um(self.lst_tarja, "tarja")).pack(side="left", padx=(0, 5))
        ctk.CTkButton(f_btns, text="Limpar Lista", fg_color="gray", hover_color="darkgray", width=100, font=("Segoe UI", 12, "bold"), command=lambda: self.limpar_aba("tarja", self.lst_tarja)).pack(side="left")

        f_filtros = ctk.CTkFrame(self.frame_tarja, corner_radius=8, fg_color="#f8f9fa", border_width=1, border_color="#e2e6ea")
        f_filtros.pack(fill="x", pady=15)
        ctk.CTkLabel(f_filtros, text="Parâmetros de Auditoria e Proteção", font=("Segoe UI", 13, "bold")).pack(anchor="w", padx=15, pady=(10, 5))
        
        self.vars = {
            "ID (CPF / CNPJ / RG / Placas de Veículos)": ctk.BooleanVar(value=True), 
            "Bancário (Contas Corrente / Cartões / PIX)": ctk.BooleanVar(value=True),
            "Contatos (E-mail / Números de Telefone)": ctk.BooleanVar(value=True), 
            "Localização (Endereços Completos e CEP)": ctk.BooleanVar(value=True),
            "Pessoas (Nomes Próprios via IA e Gatilhos)": ctk.BooleanVar(value=True)
        }
        
        for l, v in self.vars.items():
            ctk.CTkCheckBox(f_filtros, text=l, variable=v, font=("Segoe UI", 13)).pack(anchor="w", padx=20, pady=6)

        ctk.CTkButton(self.frame_tarja, text="🔍 INICIAR MAPEAMENTO", command=self.thread_analise, font=("Segoe UI", 13, "bold"), height=35).pack(fill="x", pady=10)
        self.lbl_tarja_status = ctk.CTkLabel(self.frame_tarja, text="Aguardando...", text_color="gray", font=("Segoe UI", 12))
        self.lbl_tarja_status.pack()

        ctk.CTkLabel(self.frame_tarja, text="Caixa de Revisão (Edite ou apague dados que NÃO devem ser tarjados):", text_color="#102a43", font=("Segoe UI", 12, "bold")).pack(anchor="w", pady=(10,0))
        self.caixa_rev = ctk.CTkTextbox(self.frame_tarja, height=120, font=("Consolas", 13), border_width=1, border_color="gray")
        self.caixa_rev.pack(fill="x", pady=5)

        ctk.CTkButton(self.frame_tarja, text="🔒 EXECUTAR TARJAMENTO", fg_color="#198754", hover_color="#157347", font=("Segoe UI", 14, "bold"), height=45, command=self.thread_tarjar).pack(fill="x", pady=15)

        ctk.CTkLabel(self.frame_tarja, text="Arquivos Finalizados (Clique para visualizar):", font=("Segoe UI", 12, "bold")).pack(anchor="w")
        self.lst_prontos_tarja = self.criar_lista_ui(self.frame_tarja, h=3)
        self.lst_prontos_tarja.bind('<<ListboxSelect>>', lambda e: self.preview_pronto(e, "tarja"))
        ctk.CTkButton(self.frame_tarja, text="Limpar Resultados", fg_color="transparent", border_width=1, text_color="gray", font=("Segoe UI", 12, "bold"), command=lambda: self.limpar_resultados("tarja", self.lst_prontos_tarja)).pack(anchor="e", pady=5)

    def setup_merge_tab(self):
        ctk.CTkLabel(self.frame_merge, text="Unificar Arquivos", font=("Segoe UI", 22, "bold")).pack(anchor="w", pady=(10, 0))
        ctk.CTkLabel(self.frame_merge, text="Organize e junte seus PDFs na ordem correta do Processo.", text_color="gray", wraplength=400, justify="left").pack(anchor="w", pady=(0, 20))

        ctk.CTkButton(self.frame_merge, text="📂 Adicionar PDFs", font=("Segoe UI", 12, "bold"), command=lambda: self.importar("merge", self.lst_merge)).pack(anchor="w", pady=5)
        
        f = ctk.CTkFrame(self.frame_merge, fg_color="transparent")
        f.pack(fill="x")
        self.lst_merge = tk.Listbox(f, height=12, font=("Segoe UI", 10), relief="solid", bd=1)
        self.lst_merge.pack(side="left", fill="x", expand=True)
        self.lst_merge.bind('<<ListboxSelect>>', lambda e: self.preview_selecao(e, "merge"))
        
        f_s = ctk.CTkFrame(f, fg_color="transparent")
        f_s.pack(side="right", padx=5)
        ctk.CTkButton(f_s, text="▲ Subir", width=60, font=("Segoe UI", 12, "bold"), command=lambda: self.mover_item(-1)).pack(pady=2)
        ctk.CTkButton(f_s, text="▼ Descer", width=60, font=("Segoe UI", 12, "bold"), command=lambda: self.mover_item(1)).pack(pady=2)

        f_btns = ctk.CTkFrame(self.frame_merge, fg_color="transparent")
        f_btns.pack(anchor="w", pady=10)
        ctk.CTkButton(f_btns, text="Remover", fg_color="#dc3545", width=100, hover_color="#c82333", font=("Segoe UI", 12, "bold"), command=lambda: self.excluir_um(self.lst_merge, "merge")).pack(side="left", padx=(0, 5))
        ctk.CTkButton(f_btns, text="Limpar Lista", fg_color="gray", width=100, hover_color="darkgray", font=("Segoe UI", 12, "bold"), command=lambda: self.limpar_aba("merge", self.lst_merge)).pack(side="left")

        ctk.CTkButton(self.frame_merge, text="🔗 UNIFICAR DOCUMENTOS", font=("Segoe UI", 14, "bold"), height=45, command=self.thread_merge).pack(fill="x", pady=20)
        self.lbl_merge_status = ctk.CTkLabel(self.frame_merge, text="", text_color="gray", font=("Segoe UI", 12))
        self.lbl_merge_status.pack()

        ctk.CTkLabel(self.frame_merge, text="Arquivos Finalizados:", font=("Segoe UI", 12, "bold")).pack(anchor="w")
        self.lst_prontos_merge = self.criar_lista_ui(self.frame_merge)
        self.lst_prontos_merge.bind('<<ListboxSelect>>', lambda e: self.preview_pronto(e, "merge"))
        ctk.CTkButton(self.frame_merge, text="Limpar Resultados", fg_color="transparent", border_width=1, text_color="gray", font=("Segoe UI", 12, "bold"), command=lambda: self.limpar_resultados("merge", self.lst_prontos_merge)).pack(anchor="e", pady=5)

    def setup_ocr_tab(self):
        ctk.CTkLabel(self.frame_ocr, text="Conversão OCR", font=("Segoe UI", 22, "bold")).pack(anchor="w", pady=(10, 0))
        ctk.CTkLabel(self.frame_ocr, text="Torne imagens pesquisáveis mantendo a cor e a validade.", text_color="gray", wraplength=400, justify="left").pack(anchor="w", pady=(0, 20))

        ctk.CTkButton(self.frame_ocr, text="📂 Selecionar Imagens/PDFs", font=("Segoe UI", 12, "bold"), command=lambda: self.importar("ocr", self.lst_ocr)).pack(anchor="w", pady=5)
        self.lst_ocr = self.criar_lista_ui(self.frame_ocr)
        self.lst_ocr.bind('<<ListboxSelect>>', lambda e: self.preview_selecao(e, "ocr"))

        f_btns = ctk.CTkFrame(self.frame_ocr, fg_color="transparent")
        f_btns.pack(anchor="w", pady=5)
        ctk.CTkButton(f_btns, text="Remover", fg_color="#dc3545", width=100, hover_color="#c82333", font=("Segoe UI", 12, "bold"), command=lambda: self.excluir_um(self.lst_ocr, "ocr")).pack(side="left", padx=(0, 5))
        ctk.CTkButton(f_btns, text="Limpar Lista", fg_color="gray", width=100, hover_color="darkgray", font=("Segoe UI", 12, "bold"), command=lambda: self.limpar_aba("ocr", self.lst_ocr)).pack(side="left")

        ctk.CTkButton(self.frame_ocr, text="⚙️ INICIAR CONVERSÃO", fg_color="#0dcaf0", hover_color="#31d2f2", text_color="black", font=("Segoe UI", 14, "bold"), height=45, command=self.thread_ocr).pack(fill="x", pady=20)
        self.lbl_ocr_status = ctk.CTkLabel(self.frame_ocr, text="", text_color="gray", font=("Segoe UI", 12))
        self.lbl_ocr_status.pack()

        ctk.CTkLabel(self.frame_ocr, text="Arquivos Convertidos:", font=("Segoe UI", 12, "bold")).pack(anchor="w")
        self.lst_prontos_ocr = self.criar_lista_ui(self.frame_ocr)
        self.lst_prontos_ocr.bind('<<ListboxSelect>>', lambda e: self.preview_pronto(e, "ocr"))

    # ==========================================
    # VISUALIZADOR 
    # ==========================================
    def setup_viewer(self):
        self.view_toolbar = ctk.CTkFrame(self.frame_view, height=50, corner_radius=0, fg_color="#e9ecef")
        self.view_toolbar.pack(fill="x")
        
        ctk.CTkButton(self.view_toolbar, text="📂 Abrir Avulso", width=100, fg_color="white", text_color="black", hover_color="#e2e6ea", font=("Segoe UI", 12, "bold"), command=self.abrir_avulso).pack(side="left", padx=5, pady=10)
        ctk.CTkButton(self.view_toolbar, text="🖨️ Imprimir", width=100, fg_color="white", text_color="black", hover_color="#e2e6ea", font=("Segoe UI", 12, "bold"), command=self.imprimir_avulso).pack(side="left", padx=5, pady=10)
        ctk.CTkButton(self.view_toolbar, text="📠 Digitalizar", width=100, fg_color="white", text_color="black", hover_color="#e2e6ea", font=("Segoe UI", 12, "bold"), command=self.chamar_scanner).pack(side="left", padx=5, pady=10)
        
        ctk.CTkButton(self.view_toolbar, text="💾 Salvar Edição", width=120, fg_color="#0d6efd", hover_color="#0b5ed7", font=("Segoe UI", 12, "bold"), command=self.salvar_manual).pack(side="right", padx=10, pady=10)
        self.btn_manual = ctk.CTkButton(self.view_toolbar, text="⬛ Tarja Manual: OFF", width=150, fg_color="#6c757d", hover_color="#5c636a", font=("Segoe UI", 12, "bold"), command=self.toggle_tarja_manual)
        self.btn_manual.pack(side="right", padx=5, pady=10)

        c_nav = ctk.CTkFrame(self.frame_view, fg_color="transparent")
        c_nav.pack(pady=5, fill="x")
        ctk.CTkButton(c_nav, text="◀", width=40, font=("Segoe UI", 14, "bold"), command=self.pag_ant).pack(side="left", padx=10)
        self.lbl_pag = ctk.CTkLabel(c_nav, text="Pág: 0/0", font=("Segoe UI", 12, "bold"))
        self.lbl_pag.pack(side="left", padx=10)
        ctk.CTkButton(c_nav, text="▶", width=40, font=("Segoe UI", 14, "bold"), command=self.pag_prox).pack(side="left", padx=10)
        
        self.z_scale = ctk.CTkSlider(c_nav, from_=0.1, to=2.0, command=self.att_zoom, width=150)
        self.z_scale.set(0.8)
        self.z_scale.pack(side="left", padx=20)
        ctk.CTkLabel(c_nav, text="Zoom", text_color="gray", font=("Segoe UI", 12)).pack(side="left")

        f_canvas = tk.Frame(self.frame_view, bg="gray")
        f_canvas.pack(expand=True, fill="both", padx=10, pady=10)
        
        self.scroll_v_view = ttk.Scrollbar(f_canvas, orient="vertical")
        self.scroll_h_view = ttk.Scrollbar(f_canvas, orient="horizontal")
        
        self.can_view = tk.Canvas(f_canvas, bg="#cccccc", cursor="crosshair", highlightthickness=0,
                                  yscrollcommand=self.scroll_v_view.set,
                                  xscrollcommand=self.scroll_h_view.set)
                                  
        self.scroll_v_view.config(command=self.can_view.yview)
        self.scroll_h_view.config(command=self.can_view.xview)
        
        self.scroll_v_view.pack(side="right", fill="y")
        self.scroll_h_view.pack(side="bottom", fill="x")
        self.can_view.pack(side="left", expand=True, fill="both")
        
        self.can_view.bind("<ButtonPress-1>", self.on_press)
        self.can_view.bind("<B1-Motion>", self.on_drag)
        self.can_view.bind("<ButtonRelease-1>", self.on_release)
        
        def _on_mousewheel(event):
            self.can_view.yview_scroll(int(-1*(event.delta/120)), "units")
        self.can_view.bind("<Enter>", lambda e: self.can_view.bind_all("<MouseWheel>", _on_mousewheel))
        self.can_view.bind("<Leave>", lambda e: self.can_view.unbind_all("<MouseWheel>"))

    # --- MOTOR DE OCR (HÍBRIDO MAC/WINDOWS) ---
    def exec_ocr_engine(self, img):
        if SISTEMA == "Darwin": # MacOS: Usa Apple Vision nativo
            try:
                from ocrmac import ocrmac
                return ocrmac.OCR(img).pdf()
            except:
                return pytesseract.pytesseract.image_to_pdf_or_hocr(img, extension='pdf', lang='por')
        else: # Windows: Usa Tesseract local
            return pytesseract.pytesseract.image_to_pdf_or_hocr(img, extension='pdf', lang='por')

    # ==========================================
    # LÓGICA DE AUDITORIA E PDF
    # ==========================================
    def thread_analise(self): threading.Thread(target=self.analisar, daemon=True).start()
    def analisar(self):
        arquivos = self.dados["tarja"]["entrada"]
        if not arquivos: return
        self.after(0, lambda: self.lbl_tarja_status.configure(text="⏳ Mapeando documentos e tabelas...", text_color="#ffc107"))
        termos = set()
        
        reg_num_radical = r'\b(?:\d[\.\-\/\\]?){3,}\b'
        reg_placa = r'\b[A-Z]{3}-?\d{4}\b|\b[A-Z]{3}\d[A-Z]\d{2}\b'
        reg_contato = r'([a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,})'
        reg_end = r'(?i)(?:Rua|Av\.|Avenida|Travessa|Alameda|Rodovia|Quadra|SQN|SQS|SCLN|SCLS|Lote|Conjunto|Bloco|Fazenda|Vila|Praça|Bairro|Setor)\s+[A-ZÀ-Ú0-9].{1,100}?(?:\d+|S/N)'
        reg_gatilho_nome = r'(?i)(?:Servidor|Titular|civilmente pelo|Interessado|Sra?\.?|Dr?\.?|Senhor|Senhora|colaborador)\s+([A-ZÀ-Úa-zà-ú\s]{5,40})'

        key_id = "ID (CPF / CNPJ / RG / Placas de Veículos)"
        key_banco = "Bancário (Contas Corrente / Cartões / PIX)"
        key_contato = "Contatos (E-mail / Números de Telefone)"
        key_end = "Localização (Endereços Completos e CEP)"
        key_nome = "Pessoas (Nomes Próprios via IA e Gatilhos)"

        for cam in arquivos:
            try:
                with fitz.open(cam) as d:
                    for p in d:
                        txt = p.get_text("text")
                        if self.vars[key_id].get() or self.vars[key_banco].get() or self.vars[key_contato].get() or self.vars[key_end].get():
                            termos.update(re.findall(reg_num_radical, txt))
                        if self.vars[key_id].get(): termos.update(re.findall(reg_placa, txt, re.I))
                        if self.vars[key_contato].get(): termos.update(re.findall(reg_contato, txt))
                        if self.vars[key_end].get(): termos.update([m.group() for m in re.finditer(reg_end, txt, re.I)])
                        if self.vars[key_nome].get():
                            for m in re.finditer(reg_gatilho_nome, txt):
                                termos.add(m.group(1).strip().split('\n')[0])
                            if nlp:
                                res = nlp(txt)
                                for ent in res.ents:
                                    if ent.label_ == "PER":
                                        texto_ent = ent.text.strip()
                                        if not texto_ent.isupper() and len(texto_ent) > 3:
                                            termos.add(texto_ent)
            except: pass
        
        self.caixa_rev.delete("0.0", "end")
        lista_final = set()
        for t in termos:
            for p in str(t).split('\n'):
                p_limpo = p.strip()
                if len(p_limpo) >= 3 and not re.match(r'^20[1-3]\d$', p_limpo):
                    lista_final.add(p_limpo)
        for t in sorted(list(lista_final)): self.caixa_rev.insert("end", f"{t}\n")
        self.after(0, lambda: self.lbl_tarja_status.configure(text="✅ Mapeamento Concluído!", text_color="#28a745"))

    def thread_tarjar(self): threading.Thread(target=self.tarjar, daemon=True).start()
    def tarjar(self):
        arquivos = self.dados["tarja"]["entrada"]
        termos = [x.strip() for x in self.caixa_rev.get("0.0", "end").split("\n") if x.strip()]
        if not arquivos or not termos: return
        self.after(0, lambda: self.lbl_tarja_status.configure(text="🔒 Aplicando Tarjas...", text_color="#ffc107"))
        for cam in arquivos:
            try:
                doc = fitz.open(cam)
                for p in doc:
                    for t in termos:
                        for inst in p.search_for(t): p.add_redact_annot(inst, fill=(0,0,0))
                    p.apply_redactions()
                saida = os.path.splitext(cam)[0] + "_TARJADO.pdf"
                doc.save(saida); doc.close()
                if saida not in self.dados["tarja"]["prontos"]: self.dados["tarja"]["prontos"].append(saida)
                self.after(0, lambda: self.atualizar_lb(self.lst_prontos_tarja, self.dados["tarja"]["prontos"]))
            except: pass
        self.after(0, lambda: self.lbl_tarja_status.configure(text="✅ Tarjamento Finalizado!", text_color="#28a745"))

    def thread_ocr(self): threading.Thread(target=self.exec_ocr, daemon=True).start()
    def exec_ocr(self):
        self.after(0, lambda: self.lbl_ocr_status.configure(text="⏳ Convertendo..."))
        for cam in self.dados["ocr"]["entrada"]:
            doc = fitz.open(cam); res_pdf = fitz.open()
            for i in range(len(doc)):
                pix = doc[i].get_pixmap(dpi=300); img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
                res = self.exec_ocr_engine(img) # AQUI CHAMA O MOTOR HÍBRIDO!
                with fitz.open("pdf", res) as p: res_pdf.insert_pdf(p)
            saida = f"{os.path.splitext(cam)[0]}_PRONTO.pdf"; res_pdf.save(saida); res_pdf.close(); doc.close()
            if saida not in self.dados["ocr"]["prontos"]: self.dados["ocr"]["prontos"].append(saida)
            self.after(0, lambda: self.atualizar_lb(self.lst_prontos_ocr, self.dados["ocr"]["prontos"]))
        self.after(0, lambda: self.lbl_ocr_status.configure(text="✅ Concluído", text_color="#28a745"))

    def thread_merge(self): threading.Thread(target=self.exec_merge, daemon=True).start()
    def exec_merge(self):
        if not self.dados["merge"]["entrada"]: return
        self.after(0, lambda: self.lbl_merge_status.configure(text="⏳ Unificando..."))
        res = fitz.open()
        for f in self.dados["merge"]["entrada"]:
            try: res.insert_pdf(fitz.open(f))
            except: pass
        saida = os.path.join(os.path.dirname(self.dados["merge"]["entrada"][0]), "UNIFICADO.pdf")
        res.save(saida); res.close()
        if saida not in self.dados["merge"]["prontos"]: self.dados["merge"]["prontos"].append(saida)
        self.after(0, lambda: self.atualizar_lb(self.lst_prontos_merge, self.dados["merge"]["prontos"]))
        self.after(0, lambda: self.lbl_merge_status.configure(text="✅ Concluído", text_color="#28a745"))

    # Auxiliares do Viewer
    def abrir_avulso(self):
        f = filedialog.askopenfilename(filetypes=[("PDF", "*.pdf")])
        if f: self.carregar_pdf(f)

    def imprimir_avulso(self):
        if self.caminho_atual:
            try:
                messagebox.showinfo("Impressão", "O documento será aberto no seu leitor de PDF padrão.\nAperte 'Ctrl + P' no programa que abrir.")
                os.startfile(self.caminho_atual)
            except Exception as e: messagebox.showerror("Erro", f"Erro ao abrir: {e}")
        else: messagebox.showwarning("Aviso", "Abra um documento primeiro.")

    def chamar_scanner(self):
        try: os.system("start wiaacmgr")
        except: messagebox.showwarning("Aviso", "Scanner não encontrado.")

    def toggle_tarja_manual(self):
        self.modo_tarja_manual = not self.modo_tarja_manual
        if self.modo_tarja_manual:
            self.btn_manual.configure(text="⬛ Tarja Manual: ON (Arraste o mouse)", fg_color="#ffc107", text_color="black")
        else:
            self.btn_manual.configure(text="⬛ Tarja Manual: OFF", fg_color="#6c757d", text_color="white")

    def on_press(self, event):
        if not self.modo_tarja_manual or not self.doc_aberto: return
        self.start_x, self.start_y = self.can_view.canvasx(event.x), self.can_view.canvasy(event.y)
        self.rect_id = self.can_view.create_rectangle(self.start_x, self.start_y, self.start_x, self.start_y, outline="red", width=2)

    def on_drag(self, event):
        if not self.modo_tarja_manual or not self.doc_aberto or not self.rect_id: return
        self.can_view.coords(self.rect_id, self.start_x, self.start_y, self.can_view.canvasx(event.x), self.can_view.canvasy(event.y))

    def on_release(self, event):
        if not self.modo_tarja_manual or not self.doc_aberto or not self.rect_id: return
        end_x, end_y = self.can_view.canvasx(event.x), self.can_view.canvasy(event.y)
        x0, x1 = min(self.start_x, end_x) / self.zoom_level, max(self.start_x, end_x) / self.zoom_level
        y0, y1 = min(self.start_y, end_y) / self.zoom_level, max(self.start_y, end_y) / self.zoom_level
        
        page = self.doc_aberto.load_page(self.pagina_atual)
        page.add_redact_annot(fitz.Rect(x0, y0, x1, y1), fill=(0,0,0))
        page.apply_redactions()
        
        self.can_view.delete(self.rect_id)
        self.rect_id = None
        self.renderizar()

    def salvar_manual(self):
        if self.doc_aberto and self.caminho_atual:
            saida = filedialog.asksaveasfilename(defaultextension=".pdf", initialfile=os.path.basename(self.caminho_atual).replace(".pdf", "_EDITADO.pdf"))
            if saida:
                self.doc_aberto.save(saida)
                messagebox.showinfo("Salvo", "Edição manual salva!")
                self.carregar_pdf(saida)

    def carregar_pdf(self, c):
        if not os.path.exists(c): return
        if self.doc_aberto: self.doc_aberto.close()
        self.caminho_atual = c
        self.doc_aberto = fitz.open(c)
        self.pagina_atual = 0
        self.renderizar()

    def renderizar(self):
        if not self.doc_aberto: return
        pix = self.doc_aberto.load_page(self.pagina_atual).get_pixmap(matrix=fitz.Matrix(self.zoom_level, self.zoom_level))
        img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
        self.img_tk = ImageTk.PhotoImage(img)
        self.can_view.delete("all")
        self.can_view.create_image(0, 0, anchor="nw", image=self.img_tk)
        self.can_view.config(scrollregion=(0, 0, pix.width, pix.height))
        self.lbl_pag.configure(text=f"Pág: {self.pagina_atual+1}/{len(self.doc_aberto)}")

    def att_zoom(self, v): self.zoom_level = float(v); self.renderizar()
    def pag_ant(self):
        if self.pagina_atual > 0: self.pagina_atual -= 1; self.renderizar()
    def pag_prox(self):
        if self.doc_aberto and self.pagina_atual < len(self.doc_aberto)-1: self.pagina_atual += 1; self.renderizar()

    # Controles de UI Compartilhados
    def importar(self, secao, lb):
        f = filedialog.askopenfilenames(filetypes=[("Docs", "*.pdf *.jpg *.jpeg *.png")])
        if f: self.dados[secao]["entrada"].extend(list(f)); self.atualizar_lb(lb, self.dados[secao]["entrada"]); self.carregar_pdf(f[0])
    
    def limpar_aba(self, secao, lb):
        self.dados[secao]["entrada"] = []
        self.atualizar_lb(lb, self.dados[secao]["entrada"])
        self.can_view.delete("all")
        self.caminho_atual = None
        if secao == "tarja": self.caixa_rev.delete("0.0", "end")
        
    def limpar_resultados(self, secao, lb):
        self.dados[secao]["prontos"] = []
        self.atualizar_lb(lb, self.dados[secao]["prontos"])
        self.can_view.delete("all")
        self.caminho_atual = None
        
    def excluir_um(self, lb, secao):
        s = lb.curselection()
        if s: 
            self.dados[secao]["entrada"].pop(s[0])
            self.atualizar_lb(lb, self.dados[secao]["entrada"])
            self.can_view.delete("all")
            self.caminho_atual = None
            
    def atualizar_lb(self, lb, lista):
        lb.delete(0, tk.END)
        for f in lista: lb.insert(tk.END, os.path.basename(f))
        
    def preview_selecao(self, event, secao):
        s = event.widget.curselection()
        if s: self.carregar_pdf(self.dados[secao]["entrada"][s[0]])
        
    def preview_pronto(self, event, secao):
        s = event.widget.curselection()
        if s: self.carregar_pdf(self.dados[secao]["prontos"][s[0]])
        
    def mover_item(self, d):
        s = self.lst_merge.curselection()
        if s:
            i = s[0]; j = i + d
            if 0 <= j < len(self.dados["merge"]["entrada"]):
                self.dados["merge"]["entrada"][i], self.dados["merge"]["entrada"][j] = self.dados["merge"]["entrada"][j], self.dados["merge"]["entrada"][i]
                self.atualizar_lb(self.lst_merge, self.dados["merge"]["entrada"])
                self.lst_merge.select_set(j)
                self.carregar_pdf(self.dados["merge"]["entrada"][j])

if __name__ == "__main__":
    app = TrataDocApp()
    app.mainloop()