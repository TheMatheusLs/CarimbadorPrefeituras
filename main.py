import os
import sys
import io
import json
import math
import threading
import webbrowser
import tkinter as tk
from tkinter import filedialog, messagebox
import ttkbootstrap as ttk
from ttkbootstrap.constants import *
from PIL import Image, ImageDraw, ImageFont
from pypdf import PdfReader, PdfWriter
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4

# --- FUNÇÃO DE ROBUSTEZ DE CAMINHO ---
def obter_diretorio_base():
    """Retorna o diretório onde o Executável ou o Script está localizado."""
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    else:
        return os.path.dirname(os.path.abspath(__file__))

# --- CONSTANTES ---
BASE_DIR = obter_diretorio_base()
ARQUIVO_CONFIG = os.path.join(BASE_DIR, "config.json")
PASTA_CARIMBOS = os.path.join(BASE_DIR, "Carimbos Prefeituras")

class GeradorCarimbos:
    def __init__(self):
        self.tamanho_canvas = 1200
        self.centro = 600
        self.raio_circulo = 480
        self.raio_texto = 400
        try:
            ImageFont.truetype("arial.ttf", 10)
            self.font_path_bold = "arialbd.ttf"
            self.font_path_normal = "arial.ttf"
        except:
            self.font_path_bold = "arial.ttf"
            self.font_path_normal = "arial.ttf"

    def gerar(self, nome_cidade):
        nome_limpo = nome_cidade.strip().upper()
        img = Image.new("RGBA", (self.tamanho_canvas, self.tamanho_canvas), (255, 255, 255, 0))
        draw = ImageDraw.Draw(img)
        
        # Círculo Externo
        draw.ellipse([self.centro-480, self.centro-480, self.centro+480, self.centro+480], outline="black", width=12)
        
        # Texto Curvo (Prefeitura)
        texto = f"PREFEITURA MUNICIPAL DE {nome_limpo}"
        tamanho_fonte = 75
        try: fonte = ImageFont.truetype(self.font_path_bold, tamanho_fonte)
        except: fonte = ImageFont.load_default()
        
        self._desenhar_texto_curvo(img, texto, fonte)
        
        # Miolo sempre presente (para permitir preenchimento manual)
        try: font_miolo = ImageFont.truetype(self.font_path_normal, 40)
        except: font_miolo = ImageFont.load_default()

        y_linha = self.centro + 60 
        
        # Desenha Linha Centralizada
        draw.line([(350, y_linha), (850, y_linha)], fill="black", width=6)
        
        # Texto "Folha"
        draw.text((self.centro, y_linha + 15), "Folha", font=font_miolo, fill="black", anchor="mt")
        
        if not os.path.exists(PASTA_CARIMBOS): os.makedirs(PASTA_CARIMBOS)
        
        # Salva o arquivo normal único
        nome_arquivo = f"{nome_limpo.replace(' ', '_')}.png"
        path = os.path.join(PASTA_CARIMBOS, nome_arquivo)
        img.save(path)
        return nome_limpo

    def _desenhar_texto_curvo(self, img, texto, fonte):
        perimetro = 2 * math.pi * self.raio_texto
        largura_px = sum(fonte.getbbox(c)[2] - fonte.getbbox(c)[0] for c in texto) + (len(texto)*10)
        angulo_total = (largura_px / perimetro) * 360
        angulo_atual = -90 - (angulo_total / 2)
        
        for char in texto:
            w = fonte.getbbox(char)[2] - fonte.getbbox(char)[0]
            passo = (w / perimetro) * 360
            angulo_centro = angulo_atual + (passo / 2)
            rad = math.radians(angulo_centro)
            
            x = self.centro + self.raio_texto * math.cos(rad)
            y = self.centro + self.raio_texto * math.sin(rad)
            
            txt_img = Image.new('RGBA', (200, 200), (255, 255, 255, 0))
            ImageDraw.Draw(txt_img).text((100, 100), char, font=fonte, fill="black", anchor="mm")
            
            txt_rot = txt_img.rotate(-(90 + angulo_centro), resample=Image.BICUBIC)
            img.paste(txt_rot, (int(x - txt_rot.width/2), int(y - txt_rot.height/2)), txt_rot)
            angulo_atual += passo + ((10 / perimetro) * 360)

class AppMaster:
    def __init__(self, root):
        self.root = root
        self.root.title("Carimbador Automático Municipal")
        self.root.geometry("780x940")
        
        # Ícone se existir
        try: 
            if os.path.exists(os.path.join(BASE_DIR, 'icone.ico')):
                self.root.iconbitmap(os.path.join(BASE_DIR, 'icone.ico'))
        except: pass
        
        self.gerador = GeradorCarimbos()
        self.config = self.carregar_config()

        # Variáveis
        self.var_pdf = tk.StringVar()
        self.var_carimbo = tk.StringVar()
        self.var_inicio = tk.IntVar(value=1)
        self.var_pular_capa = tk.BooleanVar(value=False)
        self.var_qtd_paginas_branco = tk.IntVar(value=1)
        self.var_sem_numero = tk.BooleanVar(value=False) 
        
        self.var_pos_canto = tk.StringVar(value=self.config.get("canto", "sup_dir"))
        self.var_ajuste_manual = tk.BooleanVar(value=self.config.get("manual", False))
        self.var_pos_x = tk.IntVar(value=self.config.get("pos_x", 480))
        self.var_pos_y = tk.IntVar(value=self.config.get("pos_y", 730))
        self.var_tamanho = tk.IntVar(value=self.config.get("tamanho", 110))

        # Cache para guardar coordenadas separadas [X, Y, TAMANHO]
        self.cache_coords = {
            "retrato": [self.var_pos_x.get(), self.var_pos_y.get(), self.var_tamanho.get()],
            "paisagem": [700, 450, 110]
        }
        self.modo_atual = "retrato"

        self._montar_layout()

    def carregar_config(self):
        if os.path.exists(ARQUIVO_CONFIG):
            try:
                with open(ARQUIVO_CONFIG, 'r') as f: return json.load(f)
            except: pass
        return {"canto": "sup_dir", "manual": False, "pos_x": 480, "pos_y": 730, "tamanho": 110}

    def salvar_config(self):
        self.config.update({"canto": self.var_pos_canto.get(), "manual": self.var_ajuste_manual.get(),
                            "pos_x": self.var_pos_x.get(), "pos_y": self.var_pos_y.get(), "tamanho": self.var_tamanho.get()})
        try:
            with open(ARQUIVO_CONFIG, 'w') as f: json.dump(self.config, f)
        except Exception as e:
            print(f"Erro config: {e}")

    def _montar_layout(self):
        container = ttk.Frame(self.root, padding=20)
        container.pack(fill='both', expand=True)

        ttk.Label(container, text="Carimbador Automático", font=("Helvetica", 18, "bold"), bootstyle="primary").pack(pady=(0, 20))

        nb = ttk.Notebook(container, bootstyle="primary")
        nb.pack(fill='both', expand=True)
        
        f_exec = ttk.Frame(nb, padding=20)
        f_branco = ttk.Frame(nb, padding=20)
        f_conf = ttk.Frame(nb, padding=20)
        
        nb.add(f_exec, text="  Processar PDF  ")
        nb.add(f_branco, text="  Gerar Folhas  ")
        nb.add(f_conf, text="  Configurações  ")

        # === ABA 1 ===
        ttk.Label(f_exec, text="Selecione o arquivo PDF:", font=("Helvetica", 10)).pack(anchor="w")
        
        f_file = ttk.Frame(f_exec)
        f_file.pack(fill="x", pady=5)
        ttk.Entry(f_file, textvariable=self.var_pdf).pack(side="left", fill="x", expand=True, padx=(0, 10))
        ttk.Button(f_file, text="Buscar PDF", command=lambda: self.var_pdf.set(filedialog.askopenfilename(filetypes=[("PDF","*.pdf")])), bootstyle="secondary").pack(side="left")

        ttk.Separator(f_exec, orient="horizontal").pack(fill="x", pady=15)

        ttk.Label(f_exec, text="Selecione a Prefeitura:", font=("Helvetica", 10)).pack(anchor="w")
        self.cb_c = ttk.Combobox(f_exec, textvariable=self.var_carimbo, state="readonly", bootstyle="primary")
        self.cb_c.pack(fill="x", pady=5)
        
        f_opts = ttk.Labelframe(f_exec, text="Opções de Numeração", padding=15, bootstyle="info")
        f_opts.pack(fill="x", pady=20)
        
        ttk.Label(f_opts, text="Iniciar em:").pack(side="left")
        ttk.Spinbox(f_opts, textvariable=self.var_inicio, from_=1, to=999999, width=8).pack(side="left", padx=10)
        ttk.Checkbutton(f_opts, text="Pular 1ª Pág.", variable=self.var_pular_capa, bootstyle="round-toggle").pack(side="left", padx=10)
        # Nova opção de carimbo limpo
        ttk.Checkbutton(f_opts, text="Apenas Carimbo (Sem nº)", variable=self.var_sem_numero, bootstyle="round-toggle").pack(side="left", padx=10)

        self.btn_run = ttk.Button(f_exec, text="CARIMBAR AGORA", command=self.processar_pdf_existente, bootstyle="success", width=25)
        self.btn_run.pack(pady=30, ipady=5)
        
        self.progresso = ttk.Progressbar(f_exec, mode='indeterminate', bootstyle="success-striped")
        self.lbl_st = ttk.Label(f_exec, text="Aguardando...", foreground="gray", font=("Helvetica", 9))
        self.lbl_st.pack(pady=5)

        # === ABA 2 ===
        ttk.Label(f_branco, text="Gerar folhas em branco numeradas com carimbo.", font=("Helvetica", 10)).pack(anchor="w", pady=(0, 15))
        
        ttk.Label(f_branco, text="Prefeitura:").pack(anchor="w")
        self.cb_c2 = ttk.Combobox(f_branco, textvariable=self.var_carimbo, state="readonly", bootstyle="info")
        self.cb_c2.pack(fill="x", pady=5)
        
        f_grid = ttk.Frame(f_branco); f_grid.pack(fill="x", pady=15)
        
        # Coluna 1
        f_c1 = ttk.Frame(f_grid); f_c1.pack(side="left", fill="x", expand=True, padx=(0, 10))
        ttk.Label(f_c1, text="Quantidade:").pack(anchor="w")
        ttk.Spinbox(f_c1, textvariable=self.var_qtd_paginas_branco, from_=1, to=1000).pack(fill="x")
        
        # Coluna 2
        f_c2 = ttk.Frame(f_grid); f_c2.pack(side="left", fill="x", expand=True, padx=(0, 10))
        ttk.Label(f_c2, text="Início da Numeração:").pack(anchor="w")
        ttk.Spinbox(f_c2, textvariable=self.var_inicio, from_=1, to=999999).pack(fill="x")

        # Coluna 3 (Nova)
        f_c3 = ttk.Frame(f_grid); f_c3.pack(side="left", fill="x", expand=True)
        ttk.Checkbutton(f_c3, text="Apenas Carimbo", variable=self.var_sem_numero, bootstyle="round-toggle").pack(anchor="w", pady=(20,0))

        self.btn_run_blank = ttk.Button(f_branco, text="GERAR PDF EM BRANCO", command=self.processar_em_branco, bootstyle="primary", width=25)
        self.btn_run_blank.pack(pady=40, ipady=5)
        self.lbl_st_blank = ttk.Label(f_branco, text="Pronto.", foreground="gray", font=("Helvetica", 9)); self.lbl_st_blank.pack()

        # === ABA 3 ===
        f_conf.columnconfigure(0, weight=1)
        
        f_add = ttk.Frame(f_conf, padding=10)
        f_add.pack(side="bottom", fill="x", pady=10)
        
        ttk.Label(f_add, text="Nova Prefeitura:", font=("Helvetica", 10, "bold")).pack(anchor="w")
        f_in = ttk.Frame(f_add)
        f_in.pack(fill="x", pady=5)
        
        self.entry_add = ttk.Entry(f_in, font=("Helvetica", 11))
        self.entry_add.pack(side="left", fill="x", expand=True, padx=(0, 10))
        ttk.Button(f_in, text="CRIAR CARIMBO", command=self.add_prefeitura_inline, bootstyle="warning").pack(side="left")

        f_preview = ttk.Labelframe(f_conf, text="Visualização", padding=10, bootstyle="info")
        f_preview.pack(side="top", fill="x", pady=(0, 5))
        
        f_top_row = ttk.Frame(f_preview)
        f_top_row.pack(fill="x")
        
        f_tog = ttk.Frame(f_top_row)
        f_tog.pack(side="top", pady=(0, 5)) 
        self.var_orientacao = tk.StringVar(value="retrato")
        ttk.Radiobutton(f_tog, text="Retrato", variable=self.var_orientacao, value="retrato", command=self._mudanca_orientacao, bootstyle="toolbutton-outline-secondary").pack(side="left", padx=5)
        ttk.Radiobutton(f_tog, text="Paisagem", variable=self.var_orientacao, value="paisagem", command=self._mudanca_orientacao, bootstyle="toolbutton-outline-secondary").pack(side="left", padx=5)

        self.cv_width = 300
        self.cv_height = 220
        f_cv_container = ttk.Frame(f_preview)
        f_cv_container.pack()
        self.canvas = tk.Canvas(f_cv_container, width=self.cv_width, height=self.cv_height, bg="#f0f0f0", highlightthickness=0)
        self.canvas.pack()

        f_ctrl = ttk.Labelframe(f_conf, text="Ajuste Fino", padding=10)
        f_ctrl.pack(side="top", fill="x", pady=5)
        
        ttk.Label(f_ctrl, text="Posição Horizontal (X):").pack(anchor="w")
        self.scale_x = ttk.Scale(f_ctrl, variable=self.var_pos_x, from_=0, to=600, command=lambda v: self._atualizar_preview())
        self.scale_x.pack(fill="x", pady=(0, 5))
        
        ttk.Label(f_ctrl, text="Posição Vertical (Y):").pack(anchor="w")
        self.scale_y = ttk.Scale(f_ctrl, variable=self.var_pos_y, from_=0, to=850, command=lambda v: self._atualizar_preview())
        self.scale_y.pack(fill="x", pady=(0, 5))
        
        ttk.Label(f_ctrl, text="Tamanho do Carimbo:").pack(anchor="w")
        self.scale_tam = ttk.Scale(f_ctrl, variable=self.var_tamanho, from_=50, to=300, command=lambda v: self._atualizar_preview())
        self.scale_tam.pack(fill="x", pady=(0, 5))
        
        ttk.Button(f_ctrl, text="Restaurar Padrão (Sup. Direito)", command=self.restaurar_padrao, bootstyle="secondary-outline", width=30).pack(pady=10)

        self.atualizar_lista()
        self.root.after(100, self._atualizar_preview)
        
        lbl_ft = ttk.Label(self.root, text="© Matheus Lôbo  |  www.matheuslobo.com  |  versão 3.0.1", font=("Helvetica", 8), foreground="#999", cursor="hand2")
        lbl_ft.bind("<Button-1>", lambda e: webbrowser.open("https://matheuslobo.com"))
        lbl_ft.pack(side="bottom", pady=5)

    def restaurar_padrao(self):
        self.cache_coords["retrato"] = [480, 730, 110]
        self.cache_coords["paisagem"] = [730, 480, 110]
        self.var_tamanho.set(110)
        mode = self.modo_atual
        if mode in self.cache_coords:
            coords = self.cache_coords[mode]
            self.var_pos_x.set(coords[0])
            self.var_pos_y.set(coords[1])
            if len(coords) > 2: self.var_tamanho.set(coords[2])
        self._atualizar_preview()
        messagebox.showinfo("Configuração", "As coordenadas foram resetadas para o padrão.")

    def _mudanca_orientacao(self):
        novo_modo = self.var_orientacao.get()
        if novo_modo == self.modo_atual: return
        self.cache_coords[self.modo_atual] = [self.var_pos_x.get(), self.var_pos_y.get(), self.var_tamanho.get()]
        default_coords = [480, 730, 110] if novo_modo == "retrato" else [730, 480, 110]
        coords = self.cache_coords.get(novo_modo, default_coords)
        self.var_pos_x.set(coords[0])
        self.var_pos_y.set(coords[1])
        if len(coords) > 2: self.var_tamanho.set(coords[2])
        else: self.var_tamanho.set(110)
        self.modo_atual = novo_modo
        self._atualizar_preview()

    def _atualizar_preview(self, event=None):
        self.canvas.delete("all")
        a4_w, a4_h = 595, 842
        orientacao = self.modo_atual
        img_w, img_h = (a4_w, a4_h) if orientacao == "retrato" else (a4_h, a4_w)
        tam_pdf = self.var_tamanho.get()
        limit_x = max(0, img_w - tam_pdf)
        limit_y = max(0, img_h - tam_pdf)
        self.scale_x.config(to=limit_x)
        self.scale_y.config(to=limit_y)
        curr_x = self.var_pos_x.get()
        curr_y = self.var_pos_y.get()
        new_x = min(curr_x, limit_x)
        new_y = min(curr_y, limit_y)
        if new_x != curr_x: self.var_pos_x.set(new_x)
        if new_y != curr_y: self.var_pos_y.set(new_y)
        
        self.cache_coords[orientacao] = [new_x, new_y, tam_pdf]
        
        padding = 20
        scale = min((self.cv_width - padding*2) / img_w, (self.cv_height - padding*2) / img_h)
        draw_w = img_w * scale
        draw_h = img_h * scale
        off_x = (self.cv_width - draw_w) / 2
        off_y = (self.cv_height - draw_h) / 2
        
        self.canvas.create_rectangle(off_x+3, off_y+3, off_x+draw_w+3, off_y+draw_h+3, fill="#ccc", outline="")
        self.canvas.create_rectangle(off_x, off_y, off_x+draw_w, off_y+draw_h, fill="white", outline="#999")
        
        screen_h = tam_pdf * scale
        screen_base_y = (off_y + draw_h) - (new_y * scale)
        screen_top_y = screen_base_y - screen_h
        screen_left_x = off_x + (new_x * scale)
        
        self.canvas.create_oval(screen_left_x, screen_top_y, screen_left_x+screen_h, screen_base_y, outline="red", width=2, fill="#ffcccc")
        self.canvas.create_text(self.cv_width/2, 10, text=f"[{orientacao.upper()}] Pos: ({int(new_x)}, {int(new_y)}) | Tam: {int(tam_pdf)}", font=("Arial", 8), fill="#555")

    def atualizar_lista(self):
        if not os.path.exists(PASTA_CARIMBOS): os.makedirs(PASTA_CARIMBOS)
        # Mantém filtro apenas por precaução caso existam resquícios da versão anterior
        arquivos = sorted([f for f in os.listdir(PASTA_CARIMBOS) if f.lower().endswith('.png') and not f.endswith('_vazio.png')])
        nomes_limpos = [f.upper().replace(".PNG", "").replace("_", " ") for f in arquivos]
        self.cb_c['values'] = nomes_limpos
        self.cb_c2['values'] = nomes_limpos
        if nomes_limpos: 
            self.cb_c.current(0)
            self.cb_c2.current(0)

    def add_prefeitura_inline(self):
        nome = self.entry_add.get().strip()
        if nome:
            try:
                # Agora gera apenas um tipo de imagem
                nome_criado = self.gerador.gerar(nome)
                self.atualizar_lista()
                self.entry_add.delete(0, tk.END)
                self.var_carimbo.set(nome_criado) 
                messagebox.showinfo("Sucesso", f"Carimbo '{nome_criado}' criado!")
            except Exception as e:
                messagebox.showerror("Erro ao gerar", str(e))
        else:
            messagebox.showwarning("Aviso", "Digite o nome da cidade.")

    # --- PROCESSAMENTO PDF EXISTENTE ---
    def processar_pdf_existente(self):
        self.salvar_config()
        dados = {
            "nome_carimbo": self.var_carimbo.get(),
            "caminho_pdf": self.var_pdf.get(),
            "inicio": self.var_inicio.get(),
            "pular_capa": self.var_pular_capa.get(),
            "cache_coords": self.cache_coords.copy(),
            "sem_numero": self.var_sem_numero.get() 
        }
        
        if not dados["caminho_pdf"]:
            messagebox.showwarning("Aviso", "Selecione um arquivo PDF.")
            return

        threading.Thread(target=self._thread_existente, args=(dados,), daemon=True).start()

    def _thread_existente(self, dados):
        self.root.after(0, self._iniciar_gui_processamento)
        try:
            nome_selecionado = dados["nome_carimbo"]
            if not nome_selecionado: raise Exception("Selecione um carimbo.")
            
            com_num = not dados["sem_numero"]
            
            arquivo_img = nome_selecionado.replace(" ", "_") + ".png"
            path_img = os.path.join(PASTA_CARIMBOS, arquivo_img)
            
            if not os.path.exists(path_img): 
                self.gerador.gerar(nome_selecionado)

            reader = PdfReader(dados["caminho_pdf"])
            writer = PdfWriter()
            start_num = dados["inicio"]
            delta = 1 if dados["pular_capa"] else 0
            
            coords_retrato = dados["cache_coords"]["retrato"]
            coords_paisagem = dados["cache_coords"]["paisagem"]

            for i, page in enumerate(reader.pages):
                if dados["pular_capa"] and i == 0:
                    writer.add_page(page)
                    continue
                
                pg_width = float(page.mediabox.width)
                pg_height = float(page.mediabox.height)
                is_landscape = pg_width > pg_height
                
                if is_landscape: coords = coords_paisagem
                else: coords = coords_retrato
                
                x_f, y_f = coords[0], coords[1]
                tam = coords[2] if len(coords) > 2 else 110
                num_pag = start_num + i - delta
                
                packet = io.BytesIO()
                can = canvas.Canvas(packet, pagesize=(pg_width, pg_height))
                
                can.drawImage(path_img, x_f, y_f, width=tam, height=tam, mask='auto')
                
                # Se for carimbo limpo, ele NÃO vai executar o código abaixo, deixando em branco para escrita manual
                if com_num:
                    can.setFont("Helvetica-Bold", 11)
                    can.drawCentredString(x_f + (tam/2), y_f + (tam*0.45), str(num_pag))
                
                can.save()
                packet.seek(0)
                
                overlay_pdf = PdfReader(packet)
                page.merge_page(overlay_pdf.pages[0])
                writer.add_page(page)

            out = dados["caminho_pdf"].replace(".pdf", "_CARIMBADO.pdf")
            try:
                with open(out, "wb") as f: writer.write(f)
            except PermissionError:
                raise Exception(f"O arquivo de destino está aberto.\nFeche: {out}")
            
            proximo = start_num + len(reader.pages) - delta
            self.root.after(0, lambda: self._finalizar_gui(True, f"Salvo em: {out}", proximo))

        except Exception as e:
            self.root.after(0, lambda e=e: self._finalizar_gui(False, str(e)))

    # --- PROCESSAMENTO EM BRANCO ---
    def processar_em_branco(self):
        self.salvar_config()
        file_path = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF", "*.pdf")])
        if not file_path: return
        
        dados = {
            "nome_carimbo": self.var_carimbo.get(),
            "inicio": self.var_inicio.get(),
            "qtd": self.var_qtd_paginas_branco.get(),
            "coords": self.cache_coords["retrato"],
            "sem_numero": self.var_sem_numero.get() 
        }
        
        threading.Thread(target=self._thread_branco, args=(file_path, dados), daemon=True).start()

    def _thread_branco(self, output_path, dados):
        self.root.after(0, lambda: self.btn_run_blank.config(state="disabled"))
        self.root.after(0, lambda: self.lbl_st_blank.config(text="Gerando...", bootstyle="primary"))
        
        try:
            nome_selecionado = dados["nome_carimbo"]
            com_num = not dados["sem_numero"]
            
            path_img = os.path.join(PASTA_CARIMBOS, nome_selecionado.replace(" ", "_") + ".png")
            
            if not os.path.exists(path_img): 
                self.gerador.gerar(nome_selecionado)

            c = canvas.Canvas(output_path, pagesize=A4)
            
            coords = dados["coords"]
            x_f, y_f = coords[0], coords[1]
            tam = coords[2] if len(coords) > 2 else 110
            
            start_num = dados["inicio"]
            qtd = dados["qtd"]

            for i in range(qtd):
                num_pag = start_num + i
                c.drawImage(path_img, x_f, y_f, width=tam, height=tam, mask='auto')
                
                # Se for carimbo limpo, ele NÃO vai executar o código abaixo
                if com_num:
                    c.setFont("Helvetica-Bold", 11)
                    c.drawCentredString(x_f + (tam/2), y_f + (tam*0.45), str(num_pag))
                    
                c.showPage()
            
            try:
                c.save()
            except PermissionError:
                raise Exception(f"O arquivo de destino está aberto.\nFeche: {output_path}")

            proximo = start_num + qtd
            self.root.after(0, lambda: self._finalizar_blank(True, f"Gerado com sucesso!", proximo))
            
        except Exception as e:
            self.root.after(0, lambda e=e: self._finalizar_blank(False, str(e)))

    # --- MÉTODOS VISUAIS AUXILIARES ---
    def _iniciar_gui_processamento(self):
        self.btn_run.config(state="disabled")
        self.lbl_st.config(text="Processando...", bootstyle="primary")
        self.progresso.pack(fill='x', pady=(0, 10))
        self.progresso.start(10)

    def _finalizar_gui(self, sucesso, msg, proximo=None):
        self.progresso.stop()
        self.progresso.pack_forget()
        self.btn_run.config(state="normal")
        if sucesso:
            self.var_inicio.set(proximo)
            self.lbl_st.config(text=f"Concluído! Próx: {proximo}", bootstyle="success")
            messagebox.showinfo("Sucesso", msg)
        else:
            self.lbl_st.config(text="Erro.", bootstyle="danger")
            messagebox.showerror("Erro", msg)

    def _finalizar_blank(self, sucesso, msg, proximo=None):
        self.btn_run_blank.config(state="normal")
        if sucesso:
            self.var_inicio.set(proximo)
            self.lbl_st_blank.config(text=f"Concluído.", bootstyle="success")
            messagebox.showinfo("Sucesso", msg)
        else:
            self.lbl_st_blank.config(text="Erro.", bootstyle="danger")
            messagebox.showerror("Erro", msg)

if __name__ == "__main__":
    app = ttk.Window(themename="litera")
    AppMaster(app)
    app.mainloop()