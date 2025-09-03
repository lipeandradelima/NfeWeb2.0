import os
import time
import json
import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

# Importações para imagens e controle de mouse/teclado
from PIL import Image, ImageTk, ImageDraw  # Adicionado ImageDraw
from pynput import keyboard
import pyautogui

import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager

# Constantes
NOME_PASTA_XML = "XMLs_Baixados"
NOME_PASTA_PDF = "PDFs_Baixados"
FSIST_URL = "https://www.fsist.com.br/"


class App:
    def __init__(self, root):
        self.root = root
        self.root.title("Automação NF-e FSist")
        self.root.geometry("800x750")

        # ... (variáveis de controle) ...
        self.chaves = []
        self.driver = None
        self.processing = False
        self.download_choice = tk.StringVar(value="XML")
        self.output_dir_var = tk.StringVar(value=os.path.dirname(os.path.abspath(__file__)))
        self.auto_clicking_active = False
        self.autoclick_interval_var = tk.StringVar(value="0.5")
        self.animation_job = None
        self.animation_frame = 0

        # --- MUDANÇA: Carregar e cortar os frames do GIF em formato redondo ---
        self.animation_images = []
        try:
            gif_filename = "animation.gif"
            canvas_display_size = 38  # O tamanho da imagem no canvas para deixar espaço para a borda

            with Image.open(gif_filename) as gif_image:
                for i in range(gif_image.n_frames):
                    gif_image.seek(i)

                    # 1. Converte para RGBA e redimensiona
                    frame_original = gif_image.convert("RGBA").resize((canvas_display_size, canvas_display_size),
                                                                      Image.Resampling.LANCZOS)

                    # 2. Cria uma imagem de máscara com o mesmo tamanho
                    mask = Image.new("L", frame_original.size, 0)  # "L" é para grayscale (preto e branco)
                    draw = ImageDraw.Draw(mask)

                    # 3. Desenha um círculo branco na máscara (o que for branco será visível)
                    # O círculo tem o tamanho total do frame para um corte perfeito
                    draw.ellipse((0, 0, frame_original.size[0], frame_original.size[1]), fill=255)  # 255 é branco

                    # 4. Aplica a máscara ao frame
                    # O 'putalpha' usa a imagem 'mask' como o canal alfa (transparência)
                    frame_original.putalpha(mask)

                    self.animation_images.append(ImageTk.PhotoImage(frame_original))
        except FileNotFoundError:
            self.animation_images = []
            print(f"Aviso: Arquivo '{gif_filename}' não encontrado. A animação não será exibida.")
        # ---------------------------------------------------------------------------------------

        self.setup_ui()

        listener_thread = threading.Thread(target=self.key_listener_thread, daemon=True)
        listener_thread.start()

    def setup_ui(self):
        # --- NOVO: Frame principal do topo para alinhar o botão e a animação ---
        top_bar_frame = ttk.Frame(self.root)
        top_bar_frame.pack(fill=tk.X, padx=10, pady=5, anchor='n')

        # O frame do botão de carregar arquivo agora é filho do top_bar_frame
        load_file_frame = ttk.Frame(top_bar_frame)
        # E é alinhado à esquerda
        load_file_frame.pack(side=tk.LEFT)
        self.btn_load = ttk.Button(load_file_frame, text="Selecionar arquivo Excel", command=self.load_excel)
        self.btn_load.pack()

        # --- MUDANÇA: O canvas agora também é filho do top_bar_frame ---
        canvas_size = 40
        # O canvas é alinhado à direita, o que o coloca na mesma "linha" do botão
        self.canvas = tk.Canvas(top_bar_frame, width=canvas_size, height=canvas_size, bg=self.root.cget('bg'),
                                highlightthickness=0)
        self.canvas.pack(side=tk.RIGHT)
        # A linha .place() foi removida

        # O resto da UI continua como antes, mas agora abaixo do top_bar_frame
        output_dir_frame = ttk.LabelFrame(self.root, text="Pasta de Destino")
        output_dir_frame.pack(fill=tk.X, padx=10, pady=5)
        dir_entry = ttk.Entry(output_dir_frame, textvariable=self.output_dir_var, state="readonly")
        dir_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5, pady=5)
        browse_btn = ttk.Button(output_dir_frame, text="Procurar...", command=self.select_output_directory)
        browse_btn.pack(side=tk.RIGHT, padx=5, pady=5)

        columns = ("chave", "status")
        self.tree = ttk.Treeview(self.root, columns=columns, show="headings")
        self.tree.heading("chave", text="Chave NF-e (44 caracteres)")
        self.tree.heading("status", text="Status")
        self.tree.column("chave", width=550)
        self.tree.column("status", width=150, anchor=tk.CENTER)
        self.tree.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        self.tree.tag_configure('baixado', foreground='green')
        self.tree.tag_configure('erro', foreground='red')

        choice_frame = ttk.Frame(self.root)
        choice_frame.pack(pady=5)
        ttk.Label(choice_frame, text="Selecione o tipo de download:").pack(side=tk.LEFT, padx=5)
        xml_radio = ttk.Radiobutton(choice_frame, text="XML", variable=self.download_choice, value="XML")
        xml_radio.pack(side=tk.LEFT, padx=5)
        pdf_radio = ttk.Radiobutton(choice_frame, text="PDF", variable=self.download_choice, value="PDF")
        pdf_radio.pack(side=tk.LEFT, padx=5)
        self.btn_start = ttk.Button(self.root, text="Iniciar Automação", command=self.start_automation,
                                    state=tk.DISABLED)
        self.btn_start.pack(pady=5)
        autoclick_frame = ttk.LabelFrame(self.root, text="Controle do Clique Automático")
        autoclick_frame.pack(fill=tk.X, padx=10, pady=10)
        ttk.Label(autoclick_frame, text="Iniciar: Pressione Ctrl + S ou clique no botão abaixo").pack(pady=2)
        ttk.Label(autoclick_frame, text="Parar a qualquer momento: Tecla Esc").pack(pady=2)
        interval_frame = ttk.Frame(autoclick_frame)
        interval_frame.pack(pady=5)
        ttk.Label(interval_frame, text="Intervalo (segundos):").pack(side=tk.LEFT, padx=(0, 5))
        interval_options = ('5', '4', '3', '2', '1', '0.5')
        self.interval_combo = ttk.Combobox(interval_frame, textvariable=self.autoclick_interval_var,
                                           values=interval_options, width=5, state='readonly')
        self.interval_combo.pack(side=tk.LEFT)
        self.btn_start_autoclick = ttk.Button(autoclick_frame, text="Iniciar Cliques", command=self.start_countdown,
                                              state=tk.NORMAL)
        self.btn_start_autoclick.pack(pady=5)
        self.btn_stop_autoclick = ttk.Button(autoclick_frame, text="Parar Cliques", command=self.stop_auto_clicking,
                                             state=tk.DISABLED)
        self.btn_stop_autoclick.pack(pady=5)
        self.autoclick_status_label = ttk.Label(autoclick_frame, text="Status: Inativo", font=('Helvetica', 10, 'bold'))
        self.autoclick_status_label.pack(pady=5)

    # --- A função de animação agora exibe as imagens do GIF já cortadas ---
    def update_animation(self):
        """Atualiza a animação mostrando a imagem já cortada e a moldura circular."""
        if not self.processing or not self.animation_images:
            return

        self.canvas.delete("all")

        current_image = self.animation_images[self.animation_frame]

        w, h = 40, 40

        # Etapa 1: Exibe a imagem do frame atual no centro.
        # Como a imagem já está cortada, ela se encaixará perfeitamente.
        # As coordenadas (w/2, h/2) centralizam a imagem.
        self.canvas.create_image(w / 2, h / 2, image=current_image)

        # Etapa 2: Desenha a moldura circular preta por cima da imagem.
        # Agora a borda se alinha perfeitamente com o corte da imagem.
        self.canvas.create_oval(1, 1, w - 1, h - 1, outline='black', width=2)

        # Avança para o próximo frame
        self.animation_frame = (self.animation_frame + 1) % len(self.animation_images)

        # Agenda a próxima atualização
        self.animation_job = self.root.after(100, self.update_animation)

    # ... (Resto do código sem mudanças) ...
    def key_listener_thread(self):
        hotkeys = {'<ctrl>+s': self.start_countdown, '<esc>': self.stop_auto_clicking}
        with keyboard.GlobalHotKeys(hotkeys) as listener:
            listener.join()

    def start_countdown(self):
        if self.auto_clicking_active: return
        self.btn_start_autoclick.config(state=tk.DISABLED)
        self.btn_stop_autoclick.config(state=tk.NORMAL)
        self.interval_combo.config(state=tk.DISABLED)

        def countdown(count):
            if not self.processing and not self.auto_clicking_active and count != 5:
                self.stop_auto_clicking()
                return
            if count > 0:
                self.autoclick_status_label.config(text=f"Iniciando cliques em... {count}")
                self.root.after(1000, countdown, count - 1)
            else:
                self.autoclick_status_label.config(text="Status: CLICANDO!")
                self.start_auto_clicking()

        countdown(5)

    def start_auto_clicking(self):
        if not self.auto_clicking_active:
            self.auto_clicking_active = True
            threading.Thread(target=self.autoclick_loop, daemon=True).start()

    def autoclick_loop(self):
        pyautogui.FAILSAFE = True
        try:
            interval = float(self.autoclick_interval_var.get())
        except ValueError:
            interval = 0.5
        try:
            while self.auto_clicking_active:
                pyautogui.click()
                time.sleep(interval)
        except pyautogui.FailSafeException:
            self.root.after(0, lambda: self.autoclick_status_label.config(text="Status: PARADA DE EMERGÊNCIA!"))
            self.root.after(0, self.stop_auto_clicking)

    def stop_auto_clicking(self):
        if self.auto_clicking_active:
            self.auto_clicking_active = False
        self.btn_start_autoclick.config(state=tk.NORMAL)
        self.btn_stop_autoclick.config(state=tk.DISABLED)
        self.autoclick_status_label.config(text="Status: Inativo")
        self.interval_combo.config(state=tk.NORMAL)

    def select_output_directory(self):
        directory = filedialog.askdirectory(title="Selecione a pasta para salvar os arquivos",
                                            initialdir=self.output_dir_var.get())
        if directory: self.output_dir_var.set(directory)

    def load_excel(self):
        filepath = filedialog.askopenfilename(title="Selecionar arquivo Excel",
                                              filetypes=[("Arquivos Excel", "*.xlsx *.xls")])
        if not filepath: return
        try:
            df = pd.read_excel(filepath, dtype=str)
        except Exception as e:
            messagebox.showerror("Erro", f"Falha ao ler o arquivo Excel:\n{e}")
            return
        self.chaves.clear()
        self.tree.delete(*self.tree.get_children())
        found_keys = False
        for col in df.columns:
            valid_chaves = [val.strip() for val in df[col].dropna() if isinstance(val, str) and len(val.strip()) == 44]
            if valid_chaves:
                self.chaves = valid_chaves
                found_keys = True
                break
        if not found_keys:
            messagebox.showwarning("Aviso", "Nenhuma coluna com chaves de 44 caracteres encontrada!")
            self.btn_start.config(state=tk.DISABLED)
            return
        base_output_dir = self.output_dir_var.get()
        PDF_DIR = os.path.join(base_output_dir, NOME_PASTA_PDF)
        XML_DIR = os.path.join(base_output_dir, NOME_PASTA_XML)
        for chave in self.chaves:
            if not chave: continue
            status = "Espera"
            tags = ()
            pdf_path = os.path.join(PDF_DIR, f"{chave}.pdf")
            xml_path = os.path.join(XML_DIR, f"{chave}.xml")
            if os.path.exists(pdf_path) or os.path.exists(xml_path):
                status = "✔ Já baixado"
                tags = ('baixado',)
            self.tree.insert("", tk.END, values=(chave, status), tags=tags)
        self.btn_start.config(state=tk.NORMAL)

    def start_automation(self):
        if self.processing: return
        if not self.output_dir_var.get():
            messagebox.showwarning("Aviso", "Por favor, selecione uma pasta de destino antes de iniciar.")
            return
        self.processing = True
        self.btn_start.config(state=tk.DISABLED)
        if self.auto_clicking_active:
            self.stop_auto_clicking()

        self.update_animation()
        threading.Thread(target=self.automate, daemon=True).start()

    def update_status(self, index, status, is_success=False, is_error=False):
        def update():
            all_items = self.tree.get_children()
            if index < len(all_items):
                item_id = all_items[index]
                current_values = self.tree.item(item_id, 'values')
                tags = ()
                if is_success:
                    tags = ('baixado',)
                elif is_error:
                    tags = ('erro',)
                if current_values:
                    self.tree.item(item_id, values=(current_values[0], status), tags=tags)

        self.root.after(0, update)

    def automate(self):
        download_type = self.download_choice.get()
        base_output_dir = self.output_dir_var.get()
        XML_DIR = os.path.join(base_output_dir, NOME_PASTA_XML)
        PDF_DIR = os.path.join(base_output_dir, NOME_PASTA_PDF)
        os.makedirs(XML_DIR, exist_ok=True)
        os.makedirs(PDF_DIR, exist_ok=True)
        options = webdriver.ChromeOptions()
        settings = {"recentDestinations": [{"id": "Save as PDF", "origin": "local", "account": ""}],
                    "selectedDestinationId": "Save as PDF", "version": 2}
        prefs = {"printing.print_preview_sticky_settings.appState": json.dumps(settings),
                 "savefile.default_directory": os.path.abspath(PDF_DIR), "download.prompt_for_download": False,
                 "download.directory_upgrade": True, "download.default_directory": os.path.abspath(XML_DIR)}
        options.add_experimental_option("prefs", prefs)
        options.add_experimental_option("excludeSwitches", ["enable-automation"])
        options.add_experimental_option("useAutomationExtension", False)
        options.add_argument("--disable-blink-features=AutomationControlled")
        options.add_argument('--kiosk-printing')
        service = Service(ChromeDriverManager().install())
        self.driver = webdriver.Chrome(service=service, options=options)
        wait = WebDriverWait(self.driver, 120)
        try:
            all_item_ids = self.tree.get_children()
            for idx, item_id in enumerate(all_item_ids):
                item_values = self.tree.item(item_id, 'values')
                if not item_values: continue
                chave = item_values[0]
                status = item_values[1]
                if "Já baixado" in status:
                    continue
                self.update_status(idx, "Processando...")
                self.driver.get(FSIST_URL)
                campo_chave = wait.until(EC.presence_of_element_located((By.ID, "chave")))
                campo_chave.clear()
                campo_chave.send_keys(chave)
                self.driver.find_element(By.ID, "butconsulta").click()
                if download_type == "XML":
                    try:
                        self.update_status(idx, "Baixando XML...")
                        btn_xml = wait.until(EC.element_to_be_clickable((By.ID, "butComCertificado")))
                        self.driver.execute_script("arguments[0].scrollIntoView(true);", btn_xml)
                        self.driver.execute_script("arguments[0].click();", btn_xml)
                        time.sleep(1)
                        self.update_status(idx, "✔ XML Baixado", is_success=True)
                    except Exception:
                        self.update_status(idx, "Erro: Botão XML não encontrado", is_error=True)
                        continue
                elif download_type == "PDF":
                    try:
                        self.update_status(idx, "Baixando PDF...")
                        original_window = self.driver.current_window_handle
                        btn_imprimir = wait.until(EC.element_to_be_clickable((By.ID, "butImprimir")))
                        btn_imprimir.click()
                        wait.until(EC.number_of_windows_to_be(2))
                        for window_handle in self.driver.window_handles:
                            if window_handle != original_window:
                                self.driver.switch_to.window(window_handle)
                                break
                        time.sleep(1.5)
                        self.driver.execute_script("window.print();")
                        time.sleep(0.8)
                        self.driver.close()
                        self.driver.switch_to.window(original_window)
                        time.sleep(0.2)
                        list_of_files = os.listdir(PDF_DIR)
                        full_path_files = [os.path.join(PDF_DIR, f) for f in list_of_files if f.endswith('.pdf')]
                        if full_path_files:
                            latest_file = max(full_path_files, key=os.path.getctime)
                            os.rename(latest_file, os.path.join(PDF_DIR, f"{chave}.pdf"))
                        self.update_status(idx, "✔ PDF Baixado", is_success=True)
                    except Exception as e:
                        self.update_status(idx, "Erro ao baixar PDF", is_error=True)
                time.sleep(0.5)
        except Exception as e:
            def show_error():
                messagebox.showerror("Erro Crítico", f"Ocorreu um erro durante a automação:\n{e}")

            self.root.after(0, show_error)
        finally:
            if self.driver: self.driver.quit()
            self.processing = False
            if self.animation_job:
                self.root.after_cancel(self.animation_job)
                self.animation_job = None

            self.root.after(100, lambda: self.canvas.delete("all"))

            self.root.after(0, lambda: self.btn_start.config(state=tk.NORMAL))
            self.root.after(0, lambda: messagebox.showinfo("Info", "Automação finalizada!"))


if __name__ == "__main__":
    root = tk.Tk()
    app = App(root)
    root.mainloop()