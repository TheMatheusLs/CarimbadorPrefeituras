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
ARQUIVO_CONFIG = os.path.join(BASE_DIR, "config_v10.json")
PASTA_CARIMBOS = os.path.join(BASE_DIR, "Carimbos Prefeituras")
# Nota: LARGURA_A4 e ALTURA_A4 são usados apenas como padrão/fallback agora

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
        
        # Miolo
        try: font_miolo = ImageFont.truetype(self.font_path_normal, 40)
        except: font_miolo = ImageFont.load_default()

        # Ajuste vertical
        y_linha = self.centro + 60 
        
        # Desenha Linha Centralizada
        draw.line([(350, y_linha), (850, y_linha)], fill="black", width=6)
        
        # Texto "Folha"
        draw.text((self.centro, y_linha + 15), "Folha", font=font_miolo, fill="black", anchor="mt")
        
        if not os.path.exists(PASTA_CARIMBOS): os.makedirs(PASTA_CARIMBOS)
        
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
        self.root.title("Carimbador Pro - v10.0")
        self.root.geometry("680x800")
        
        # Ícone se existir
        # try: self.root.iconbitmap(os.path.join(BASE_DIR, 'icone.ico'))
        # except: pass
        
        self.gerador = GeradorCarimbos()
        self.config = self.carregar_config()

        # Variáveis
        self.var_pdf = tk.StringVar()
        self.var_carimbo = tk.StringVar()
        self.var_inicio = tk.IntVar(value=1)
        self.var_pular_capa = tk.BooleanVar(value=False)
        self.var_qtd_paginas_branco = tk.IntVar(value=1)
        
        self.var_pos_canto = tk.StringVar(value=self.config.get("canto", "sup_dir"))
        self.var_ajuste_manual = tk.BooleanVar(value=self.config.get("manual", False))
        self.var_pos_x = tk.IntVar(value=self.config.get("pos_x", 480))
        self.var_pos_y = tk.IntVar(value=self.config.get("pos_y", 730))
        self.var_tamanho = tk.IntVar(value=self.config.get("tamanho", 110))

        self._montar_layout()
        self.alternar_modo_posicao()

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

    def alternar_modo_posicao(self):
        estado = "normal" if self.var_ajuste_manual.get() else "disabled"
        self.spin_x.config(state=estado)
        self.spin_y.config(state=estado)

    def _montar_layout(self):
        # Container Principal
        container = ttk.Frame(self.root, padding=20)
        container.pack(fill='both', expand=True)

        # Cabeçalho Simples
        ttk.Label(container, text="Carimbador Automático", font=("Helvetica", 18, "bold"), bootstyle="primary").pack(pady=(0, 20))

        # Notebook (Abas Modernas)
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
        ttk.Spinbox(f_opts, textvariable=self.var_inicio, from_=1, to=999999, width=10).pack(side="left", padx=10)
        ttk.Checkbutton(f_opts, text="Pular 1ª Página (Capa)", variable=self.var_pular_capa, bootstyle="round-toggle").pack(side="left", padx=20)

        # Botão de Ação destaque
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
        f_c2 = ttk.Frame(f_grid); f_c2.pack(side="left", fill="x", expand=True)
        ttk.Label(f_c2, text="Início da Numeração:").pack(anchor="w")
        ttk.Spinbox(f_c2, textvariable=self.var_inicio, from_=1, to=999999).pack(fill="x")

        self.btn_run_blank = ttk.Button(f_branco, text="GERAR PDF EM BRANCO", command=self.processar_em_branco, bootstyle="primary", width=25)
        self.btn_run_blank.pack(pady=40, ipady=5)
        self.lbl_st_blank = ttk.Label(f_branco, text="Pronto.", foreground="gray", font=("Helvetica", 9)); self.lbl_st_blank.pack()

        # === ABA 3 (CONFIG) ===
        f_pos = ttk.Labelframe(f_conf, text="Posição do Carimbo", padding=15)
        f_pos.pack(fill="both", expand=True)

        ttk.Label(f_pos, text="Clique na posição desejada (folha A4):", font=("Helvetica", 9, "bold")).pack(pady=5)
        
        # Simulador Visual
        canvas_papel = tk.Canvas(f_pos, bg="white", highlightthickness=1, width=150, height=210) # Canvas puro para desenhar
        canvas_papel.pack(pady=10)
        
        # Pseudo-radiobuttons visuais
        # Para simplificar com ttkbootstrap, usaremos Radiobuttons normais posicionados
        # Truque: frame transparente ou place dentro do canvas
        
        fr_botoes = ttk.Frame(f_pos)
        fr_botoes.pack(pady=10)
        
        # 'toolbutton-outline' faz eles parecerem botões de alternância
        ttk.Radiobutton(fr_botoes, text="↖ Sup. Esq.", variable=self.var_pos_canto, value="sup_esq", bootstyle="toolbutton-outline").grid(row=0, column=0, padx=5, pady=5)
        ttk.Radiobutton(fr_botoes, text="Sup. Dir. ↗", variable=self.var_pos_canto, value="sup_dir", bootstyle="toolbutton-outline").grid(row=0, column=1, padx=5, pady=5)
        ttk.Radiobutton(fr_botoes, text="↙ Inf. Esq.", variable=self.var_pos_canto, value="inf_esq", bootstyle="toolbutton-outline").grid(row=1, column=0, padx=5, pady=5)
        ttk.Radiobutton(fr_botoes, text="Inf. Dir. ↘", variable=self.var_pos_canto, value="inf_dir", bootstyle="toolbutton-outline").grid(row=1, column=1, padx=5, pady=5)

        ttk.Separator(f_pos, orient="horizontal").pack(fill="x", pady=15)
        
        ttk.Checkbutton(f_pos, text="Modo Manual (Coordenadas Exatas)", variable=self.var_ajuste_manual, command=self.alternar_modo_posicao, bootstyle="round-toggle").pack(anchor="w")
        
        f_man = ttk.Frame(f_pos); f_man.pack(pady=10, fill="x")
        ttk.Label(f_man, text="X:").pack(side="left")
        self.spin_x = ttk.Spinbox(f_man, textvariable=self.var_pos_x, from_=0, to=600, width=8); self.spin_x.pack(side="left", padx=5)
        ttk.Label(f_man, text="Y:").pack(side="left", padx=(10,0))
        self.spin_y = ttk.Spinbox(f_man, textvariable=self.var_pos_y, from_=0, to=850, width=8); self.spin_y.pack(side="left", padx=5)
        
        f_tam = ttk.Frame(f_pos); f_tam.pack(pady=5, fill="x")
        ttk.Label(f_tam, text="Tamanho (pts):").pack(side="left")
        ttk.Spinbox(f_tam, textvariable=self.var_tamanho, from_=50, to=300, width=8).pack(side="left", padx=5)

        # Adicionar Prefeitura
        f_add = ttk.Labelframe(f_conf, text="Nova Prefeitura", padding=15, bootstyle="warning")
        f_add.pack(fill="x", pady=20)
        
        self.entry_add = ttk.Entry(f_add)
        self.entry_add.pack(side="left", fill="x", expand=True, padx=(0,10))
        ttk.Button(f_add, text="Criar", command=self.add_prefeitura_inline, bootstyle="warning").pack(side="left")

        self.atualizar_lista()
        
        # Footer
        lbl_ft = ttk.Label(self.root, text="Desenvolvido por Matheus Lôbo  |  v10.0", font=("Helvetica", 8), foreground="#999")
        lbl_ft.pack(side="bottom", pady=10)

    def atualizar_lista(self):
        if not os.path.exists(PASTA_CARIMBOS): os.makedirs(PASTA_CARIMBOS)
        arquivos = sorted([f for f in os.listdir(PASTA_CARIMBOS) if f.lower().endswith('.png')])
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
                nome_criado = self.gerador.gerar(nome)
                self.atualizar_lista()
                self.entry_add.delete(0, tk.END)
                self.var_carimbo.set(nome_criado) 
                messagebox.showinfo("Sucesso", f"Carimbo '{nome_criado}' criado!")
            except Exception as e:
                messagebox.showerror("Erro ao gerar", str(e))
        else:
            messagebox.showwarning("Aviso", "Digite o nome da cidade.")

    # --- MÉTODO AUXILIAR PARA CÁLCULO DINÂMICO ---
    def _calcular_posicao_dinamica(self, pg_width, pg_height, tam_carimbo, config_manual, config_canto, config_pos):
        """
        Calcula X e Y baseado no tamanho REAL da página atual.
        Isso resolve problemas de Landscape (Paisagem) e folhas fora do padrão A4.
        """
        # Desempacota configs para thread-safety
        manual_ativo, manual_x, manual_y = config_manual, config_pos[0], config_pos[1]
        
        if manual_ativo:
            return manual_x, manual_y
        
        canto = config_canto
        margem = 20
        
        if canto == "sup_esq":
            x = margem
            y = pg_height - tam_carimbo - margem
        elif canto == "sup_dir":
            x = pg_width - tam_carimbo - margem
            y = pg_height - tam_carimbo - margem
        elif canto == "inf_esq":
            x = margem
            y = margem
        elif canto == "inf_dir":
            x = pg_width - tam_carimbo - margem
            y = margem
        else:
            # Padrão
            x = pg_width - tam_carimbo - 20
            y = pg_height - tam_carimbo - 20
            
        return x, y

    # --- PROCESSAMENTO PDF EXISTENTE (CORRIGIDO PARA LANDSCAPE) ---
    def processar_pdf_existente(self):
        self.salvar_config()
        
        # CAPTURA DE DADOS NA MAIN THREAD (PREVINE ERROS TÉCNICOS)
        dados = {
            "nome_carimbo": self.var_carimbo.get(),
            "caminho_pdf": self.var_pdf.get(),
            "tamanho": self.var_tamanho.get(),
            "inicio": self.var_inicio.get(),
            "pular_capa": self.var_pular_capa.get(),
            "config_manual": self.var_ajuste_manual.get(),
            "config_canto": self.var_pos_canto.get(),
            "config_pos": (self.var_pos_x.get(), self.var_pos_y.get())
        }
        
        threading.Thread(target=self._thread_existente, args=(dados,), daemon=True).start()

    def _thread_existente(self, dados):
        self.root.after(0, self._iniciar_gui_processamento)
        try:
            nome_selecionado = dados["nome_carimbo"]
            if not nome_selecionado: raise Exception("Selecione um carimbo.")
            
            arquivo_img = nome_selecionado.replace(" ", "_") + ".png"
            path_img = os.path.join(PASTA_CARIMBOS, arquivo_img)
            
            # Gera imagem se não existir (thread-safe pois é file system)
            if not os.path.exists(path_img): self.gerador.gerar(nome_selecionado)

            reader = PdfReader(dados["caminho_pdf"])
            writer = PdfWriter()
            tam = dados["tamanho"]
            start_num = dados["inicio"]
            delta = 1 if dados["pular_capa"] else 0

            for i, page in enumerate(reader.pages):
                if dados["pular_capa"] and i == 0:
                    writer.add_page(page)
                    continue
                
                # Dimensões REAIS
                pg_width = float(page.mediabox.width)
                pg_height = float(page.mediabox.height)
                
                # Calcula X e Y com dados passados
                x_f, y_f = self._calcular_posicao_dinamica(
                    pg_width, pg_height, tam, 
                    dados["config_manual"], dados["config_canto"], dados["config_pos"]
                )
                
                num_pag = start_num + i - delta
                
                packet = io.BytesIO()
                can = canvas.Canvas(packet, pagesize=(pg_width, pg_height))
                
                can.drawImage(path_img, x_f, y_f, width=tam, height=tam, mask='auto')
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
            err_msg = str(e)
            self.root.after(0, lambda: self._finalizar_gui(False, err_msg))

    # --- PROCESSAMENTO EM BRANCO (MANTÉM A4 FIXO) ---
    def processar_em_branco(self):
        self.salvar_config()
        file_path = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF", "*.pdf")])
        if not file_path: return
        
        # CAPTURA DE DADOS NA MAIN THREAD
        dados = {
            "nome_carimbo": self.var_carimbo.get(),
            "tamanho": self.var_tamanho.get(),
            "inicio": self.var_inicio.get(),
            "qtd": self.var_qtd_paginas_branco.get(),
            "config_manual": self.var_ajuste_manual.get(),
            "config_canto": self.var_pos_canto.get(),
            "config_pos": (self.var_pos_x.get(), self.var_pos_y.get())
        }
        
        threading.Thread(target=self._thread_branco, args=(file_path, dados), daemon=True).start()

    def _thread_branco(self, output_path, dados):
        self.root.after(0, lambda: self.btn_run_blank.config(state="disabled"))
        self.root.after(0, lambda: self.lbl_st_blank.config(text="Gerando...", bootstyle="primary"))
        
        try:
            nome_selecionado = dados["nome_carimbo"]
            path_img = os.path.join(PASTA_CARIMBOS, nome_selecionado.replace(" ", "_") + ".png")
            if not os.path.exists(path_img): self.gerador.gerar(nome_selecionado)

            c = canvas.Canvas(output_path, pagesize=A4)
            tam = dados["tamanho"]
            
            # Cálculo de Posição (A4 = 595x842)
            x_f, y_f = self._calcular_posicao_dinamica(
                595, 842, tam,
                dados["config_manual"], dados["config_canto"], dados["config_pos"]
            )
            
            start_num = dados["inicio"]
            qtd = dados["qtd"]

            for i in range(qtd):
                num_pag = start_num + i
                c.drawImage(path_img, x_f, y_f, width=tam, height=tam, mask='auto')
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
            err_msg = str(e)
            self.root.after(0, lambda: self._finalizar_blank(False, err_msg))

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
    # Inicialização do Tema Moderno "Litera"
    app = ttk.Window(themename="litera")
    AppMaster(app)
    app.mainloop()