import customtkinter as ctk
from tkinter import filedialog, messagebox
import sys
import os
from logica_ccts import processar_arquivos, carregar_credenciais, salvar_credenciais  # credenciais agora vêm do config.ini (via logica_ccts)
import threading

def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

def get_app_path():
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    else:
        return os.path.dirname(os.path.abspath(__file__))

theme_path = resource_path("marsh.json")
ctk.set_appearance_mode("dark")
ctk.set_default_color_theme(theme_path)

class App:
    def __init__(self, root, auto_start=False):
        self.root = root
        self.auto_start = auto_start  # --auto-start (agendador): usa a pasta do mês atual, inicia sozinho e fecha ao terminar
        self.diretorio_ccts = None
        self.root.title("Robô CCTs para ComuniQ")
        self.root.geometry("1000x400")

        self.usuario = ""
        self.senha = ""
        self.carregar_credenciais()

        self.parar_event = threading.Event()
        self.thread = None

        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)
        root.grid_rowconfigure(0, weight=1)
        root.grid_columnconfigure(0, weight=1)

        self.mainframe = ctk.CTkFrame(root)
        self.mainframe.grid(row=0, column=0, padx=(5,5), pady=(12,12), sticky="nsew")

        for i in range(4):
            self.mainframe.grid_columnconfigure(i, weight=1)
        for j in range(4):
            self.mainframe.grid_rowconfigure(j, weight=1)

        self.label = ctk.CTkLabel(self.mainframe, text="Escolha a pasta com os documentos. Eles serão padronizados e salvos no ComuniQ.", font=("Inter", 16, "bold"), fg_color="#397F5A", corner_radius=15)
        self.label.grid(row=0, column=0, columnspan=4, sticky="n", padx=15, pady=(20,0))

        self.botao_caminho = ctk.CTkButton(self.mainframe, text="Selecionar o caminho", text_color="white", font=("Inter", 12, "bold"), width=130, height=30, command=self.seleciona_diretorio_ccts)
        self.botao_caminho.grid(row=1, column=1, sticky="w", padx=5, pady=(0,5))

        self.botao_inicia = ctk.CTkButton(self.mainframe, text="Iniciar o processo", text_color="white", font=("Inter", 12, "bold"), width=130, height=30, command=self.iniciar_processo)
        self.botao_inicia.grid(row=1, column=2, sticky="e", padx=5, pady=(0,5))
        
        self.botao_ajuda = ctk.CTkButton(self.mainframe, text="Como usar", text_color="white", font=("Inter", 12, "bold"), width=130, height=30, command=self.mostra_guia)
        self.botao_ajuda.grid(row=4, column=1, sticky="w", padx=20, pady=15)

        self.botao_editar_login = ctk.CTkButton(self.mainframe, text="Editar Login", text_color="white", font=("Inter", 12, "bold"), width=130, height=30, command=self.abrir_janela_login)
        self.botao_editar_login.grid(row=4, column=2, sticky="e", padx=20, pady=15)

        self.status_frame = ctk.CTkFrame(self.mainframe, fg_color="transparent")
        self.status_frame.grid(row=3, column=1, columnspan=2, sticky="ew", padx=15, pady=(10, 5))

        self.status_frame.grid_columnconfigure(0, weight=1) 
        self.status_frame.grid_columnconfigure(1, weight=0) 

        self.spinner_label = ctk.CTkLabel(self.status_frame, text="", font=("Courier", 20, "bold"))
        self.spinner_label.grid(row=0, column=1, padx=(10, 0))

        self.label_status = ctk.CTkLabel(self.status_frame, text="", font=("arial", 14, "bold"), fg_color="#333333", corner_radius=10, wraplength=700)
        self.label_status.grid(row=0, column=0, sticky="ew")
        
        self.spinner_running = False
        self.spinner_chars = ["|", "/", "-", "\\"]

        if self.auto_start:
            self.root.after(500, self.iniciar_processo)  # espera a janela abrir antes de iniciar

    def animate_spinner(self, index=0):
        if self.spinner_running:
            self.spinner_label.configure(text=self.spinner_chars[index])
            
            next_index = (index + 1) % len(self.spinner_chars)
            
            self.root.after(150, self.animate_spinner, next_index)

    def carregar_credenciais(self):
        self.usuario, self.senha = carregar_credenciais()  # lê [Login] do config.ini (o arquivo é criado ao importar logica_ccts)

    def gravar_credenciais(self):
        try:
            salvar_credenciais(self.usuario, self.senha)  # grava [Login] no config.ini preservando [Settings] e [Email]
        except Exception as e:
            messagebox.showerror("Erro", f"Não foi possível salvar as credenciais no arquivo de configuração.\nErro: {e}")

    def seleciona_diretorio_ccts(self, *args):
        arquivos = filedialog.askopenfilenames(title="Selecione um arquivo na pasta desejada")
        if arquivos:
            self.diretorio_ccts = os.path.dirname(arquivos[0])
            self.label_status.configure(text=f"PASTA SELECIONADA: {self.diretorio_ccts}")
        else:
            messagebox.showwarning("Aviso", "Nenhum arquivo selecionado.")

    def iniciar_processo(self):
        if not self.usuario or not self.senha:
             messagebox.showwarning("Aviso", "Usuário ou senha não configurados. Por favor, edite o login.")
             return
        if self.diretorio_ccts is None and not self.auto_start:  # no automático a pasta é encontrada pela lógica
            messagebox.showwarning("Aviso", "Por favor, selecione um arquivo da pasta dos CCTs antes de iniciar.")
            return

        BASE_DIR = os.path.dirname(os.path.abspath(sys.executable if getattr(sys, 'frozen', False) else __file__))
        pasta_temp = os.path.join(BASE_DIR, "temp_upload")
        log_path = os.path.join(BASE_DIR, "log_historico.txt")
        self.parar_event.clear()

        def executar_processo():
            try:
                self.spinner_running = True
                self.root.after(0, self.animate_spinner)

                self.label_status.configure(text="Por favor aguarde, isso pode levar alguns minutos.")
               
                processar_arquivos(
                    self.diretorio_ccts, pasta_temp, log_path,
                    self.usuario, self.senha, self.parar_event,
                    automatico=self.auto_start
                )
                if not self.parar_event.is_set():
                    self.label_status.configure(text=f"Processo concluído! Log gerado em: {log_path}")
                else:
                    self.label_status.configure(text="Processo interrompido pelo usuário.")
            except Exception as e:
                self.label_status.configure(text=f"Erro durante o processo: {e}")
                messagebox.showerror("Erro", f"Ocorreu um erro durante o processamento: {e}. Verifique o log para mais detalhes.")

            finally:
                self.spinner_running = False
                self.spinner_label.configure(text="")
                if self.auto_start:
                    self.root.after(4000, self.root.destroy)  # fecha sozinho (mesmo comportamento do mediador)


        self.thread = threading.Thread(target=executar_processo)
        self.thread.start()

    def on_closing(self):
        if messagebox.askokcancel("Sair", "Deseja realmente encerrar o robô?"):
            self.parar_event.set()
            if self.thread and self.thread.is_alive():
                self.thread.join(timeout=5)
            self.root.destroy()

    def mostra_guia(self):
        guia = (
            "Os documentos da pasta selecionada devem seguir o seguinte padrão de nomenclatura:\n\n"
            " - número das lojas_cidade_código do sindicato_vigência\n\n"
            "Exemplos:\n"
            " - 160,310_LAJEADO_SIND0088_2026\n"
            " - 642,285,187,259_VIAMAO_SIND0173_2024-2026\n\n"
            "Observações:\n"
            " - Se necessário, pode-se adicionar complementos após a vigência, como: _ADITIVOS, _FERIADOS, etc.\n\n"
            "Exemplos com complemento:\n"
            " - 127,139,140,155_LAJEADO_SIND0077_2025-2027_ADITIVOS\n"
            " - 503,631,659,671_UMUARAMA_SIND0180_2025-2027_NATAL\n\n"
            "Eles serão salvos no ComuniQ com o seguinte padrão:\n"
            "   - CCT_vigência\n"
            "   - CCT_vigência_complemento (caso exista complemento)\n"
        )
        messagebox.showinfo("Padrão necessário", guia)

    def abrir_janela_login(self):
        self.login_window = ctk.CTkToplevel(self.root)
        self.login_window.title("Editar Credenciais de Login")
        self.login_window.geometry("400x250")
        self.login_window.transient(self.root)
        self.login_window.grab_set()
        self.login_window.after(10, self.login_window.lift)

        ctk.CTkLabel(self.login_window, text="Usuário (e-mail):", font=("Inter", 12)).pack(padx=20, pady=(20, 5))
        self.usuario_entry = ctk.CTkEntry(self.login_window, width=250)
        self.usuario_entry.insert(0, self.usuario)
        self.usuario_entry.pack(padx=20, pady=5)

        ctk.CTkLabel(self.login_window, text="Senha:", font=("Inter", 12)).pack(padx=20, pady=5)
        self.senha_entry = ctk.CTkEntry(self.login_window, show="*", width=250)
        self.senha_entry.insert(0, self.senha)
        self.senha_entry.pack(padx=20, pady=5)

        self.show_password_checkbox = ctk.CTkCheckBox(
            self.login_window, 
            text="Mostrar Senha",
            command=self.toggle_password_visibility
        )
        self.show_password_checkbox.pack(padx=20, pady=10)

        ctk.CTkButton(self.login_window, text="Salvar", command=self.salvar_e_gravar_credenciais).pack(pady=10)

    def toggle_password_visibility(self):

        if self.show_password_checkbox.get() == 1:
            self.senha_entry.configure(show="")
        else:
            self.senha_entry.configure(show="*")

    def salvar_e_gravar_credenciais(self):
        novo_usuario = self.usuario_entry.get().strip()
        nova_senha = self.senha_entry.get()
        if not novo_usuario or not nova_senha:
            messagebox.showwarning("Aviso", "Usuário e senha não podem estar em branco.", parent=self.login_window)
            return

        self.usuario = novo_usuario
        self.senha = nova_senha
        self.gravar_credenciais()
        messagebox.showinfo("Sucesso", "Credenciais salvas com sucesso!", parent=self.login_window)
        self.login_window.destroy()

if __name__ == "__main__":
    root = ctk.CTk()
    app = App(root, auto_start='--auto-start' in sys.argv)
    root.mainloop()
