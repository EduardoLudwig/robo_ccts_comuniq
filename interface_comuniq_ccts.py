import os
import sys
import threading
import traceback
import tkinter as tk
from tkinter import filedialog, messagebox

import customtkinter as ctk

from logica_ccts import (
    APP_CONFIG_DIR,
    carregar_credenciais,
    log,
    processar_arquivos,
    salvar_credenciais,
)

AUTO_CLOSE_MS = 4000
APP_TITLE = "Robô CCTS ComuniQ"
LIMITE_CAMINHO = 78
LIMITE_DETALHE = 100

COR_VERDE = "#397F5A"
COR_VERDE_HOVER = "#4E8F69"
COR_FUNDO = "#070C09"
COR_CARD = "#0B1410"
COR_CARD_HOVER = "#11201A"
COR_BORDA = "#1C2B23"
COR_TEXTO = "#E9F5EE"
COR_TEXTO_SECUNDARIO = "#8AA396"
COR_SUCESSO = "#74C9A5"
COR_ERRO = "#E07E72"
COR_NEUTRO = "#F0F8ED"
COR_ESCURO = "#F0F8ED"


def resource_path(relative_path):
    if hasattr(sys, "_MEIPASS"):
        base_path = sys._MEIPASS
    else:
        base_path = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base_path, relative_path)


class App:
    def __init__(self, root, auto_start=False):
        self.root = root
        self.auto_start = auto_start
        self.diretorio_ccts = None
        self.processando = False
        self.thread = None
        self.log_path = os.path.join(APP_CONFIG_DIR, "log_historico.txt")
        self.pasta_temp = os.path.join(APP_CONFIG_DIR, "temp_upload")
        self.erro_fatal = None
        self.resumo = {
            "arquivos": 0,
            "enviados": 0,
            "ja_existiam": 0,
            "falhas": 0,
        }
        self._tooltip = None
        self._tooltip_after = None

        self.usuario = ""
        self.senha = ""
        self.carregar_credenciais()

        self.parar_event = threading.Event()

        self.root.title(APP_TITLE)
        self.root.geometry("960x660")
        self.root.minsize(860, 640)
        self.root.configure(fg_color=COR_FUNDO)
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)

        try:
            self.root.iconbitmap(resource_path("ccts_icon.ico"))
        except Exception:
            pass

        self._montar_layout()
        self._configurar_estado_inicial()
        self.root.after(50, self._centralizar)

    def _montar_layout(self):
        self.root.grid_rowconfigure(0, weight=1)
        self.root.grid_columnconfigure(0, weight=1)

        self.container = ctk.CTkFrame(self.root, fg_color="transparent")
        self.container.grid(row=0, column=0, sticky="nsew", padx=30, pady=24)
        self.container.grid_columnconfigure(0, weight=1)
        self.container.grid_rowconfigure(3, weight=1)

        self._montar_cabecalho()
        self._montar_card_origem()
        self._montar_card_status()
        self._montar_card_atividade()
        self._montar_rodape()

    def _montar_cabecalho(self):
        self.header = ctk.CTkFrame(self.container, fg_color="transparent")
        self.header.grid(row=0, column=0, sticky="ew")
        self.header.grid_columnconfigure(1, weight=1)

        self.header_marca = ctk.CTkFrame(
            self.header,
            width=18,
            height=18,
            fg_color=COR_VERDE,
            corner_radius=0,
        )
        self.header_marca.grid(row=0, column=0, sticky="w", padx=(0, 16))
        self.header_marca.grid_propagate(False)

        self.label_titulo = ctk.CTkLabel(
            self.header,
            text=APP_TITLE.upper(),
            font=("Segoe UI", 26, "bold"),
            text_color=COR_TEXTO,
            anchor="w",
        )
        self.label_titulo.grid(row=0, column=1, sticky="w")

        self.botao_ajuda = ctk.CTkButton(
            self.header,
            text="COMO USAR",
            width=126,
            height=36,
            fg_color="transparent",
            hover_color=COR_CARD_HOVER,
            border_width=1,
            border_color=COR_BORDA,
            text_color=COR_TEXTO_SECUNDARIO,
            font=("Segoe UI", 11, "bold"),
            corner_radius=0,
            command=self.mostra_guia,
        )
        self.botao_ajuda.grid(row=0, column=2, sticky="e")

        self.header_linha = ctk.CTkFrame(
            self.container, height=1, fg_color=COR_BORDA
        )
        self.header_linha.grid(row=0, column=0, sticky="sew", pady=(16, 0))
        self.header_linha.grid_propagate(False)

    def _montar_card_origem(self):
        self.card_origem = ctk.CTkFrame(self.container, fg_color="transparent")
        self.card_origem.grid(row=1, column=0, sticky="ew", pady=(24, 18))
        self.card_origem.grid_columnconfigure(0, weight=1)

        self.label_origem = ctk.CTkLabel(
            self.card_origem,
            text="ORIGEM DOS DOCUMENTOS",
            font=("Segoe UI", 10, "bold"),
            text_color=COR_TEXTO_SECUNDARIO,
            anchor="w",
        )
        self.label_origem.grid(row=0, column=0, columnspan=2, sticky="w", padx=0)

        self.label_origem_desc = ctk.CTkLabel(
            self.card_origem,
            text="Selecione a pasta que contém os documentos que serão enviados ao ComuniQ.",
            font=("Segoe UI", 13),
            text_color=COR_TEXTO,
            anchor="w",
        )
        self.label_origem_desc.grid(
            row=1, column=0, columnspan=2, sticky="w", padx=0, pady=(6, 0)
        )

        self.path_frame = ctk.CTkFrame(
            self.card_origem,
            fg_color=COR_CARD,
            border_width=1,
            border_color=COR_BORDA,
            corner_radius=0,
        )
        self.path_frame.grid(row=2, column=0, sticky="ew", padx=(0, 14), pady=(16, 0))
        self.path_frame.grid_columnconfigure(0, weight=1)

        self.label_caminho = ctk.CTkLabel(
            self.path_frame,
            text="Nenhuma pasta selecionada",
            width=180,
            height=22,
            font=("Segoe UI", 12),
            text_color=COR_TEXTO_SECUNDARIO,
            anchor="w",
            justify="left",
            wraplength=0,
        )
        self.label_caminho.grid(row=0, column=0, sticky="ew", padx=16, pady=15)
        self.label_caminho.bind("<Enter>", self._agendar_tooltip_caminho)
        self.label_caminho.bind("<Leave>", self._cancelar_tooltip_caminho)

        self.botao_caminho = ctk.CTkButton(
            self.card_origem,
            text="SELECIONAR PASTA  ↗",
            width=180,
            height=48,
            fg_color="transparent",
            hover_color=COR_CARD_HOVER,
            border_width=1,
            border_color=COR_VERDE,
            text_color=COR_VERDE,
            font=("Segoe UI", 11, "bold"),
            corner_radius=0,
            command=self.seleciona_diretorio_ccts,
        )
        self.botao_caminho.grid(row=2, column=1, sticky="e", pady=(16, 0))

    def _montar_card_status(self):
        self.card_status = ctk.CTkFrame(self.container, fg_color="transparent")
        self.card_status.grid(row=2, column=0, sticky="ew", pady=(0, 20))
        self.card_status.grid_columnconfigure(0, weight=1)

        self.label_status_titulo = ctk.CTkLabel(
            self.card_status,
            text="STATUS DA EXECUÇÃO",
            font=("Segoe UI", 10, "bold"),
            text_color=COR_TEXTO_SECUNDARIO,
            anchor="w",
        )
        self.label_status_titulo.grid(row=0, column=0, sticky="w", padx=0)

        self.status_line = ctk.CTkFrame(self.card_status, fg_color="transparent")
        self.status_line.grid(row=1, column=0, sticky="ew", padx=0, pady=(10, 0))
        self.status_line.grid_columnconfigure(1, weight=1)

        self.status_dot = ctk.CTkLabel(
            self.status_line,
            text="■",
            width=20,
            font=("Segoe UI", 14),
            text_color=COR_NEUTRO,
        )
        self.status_dot.grid(row=0, column=0, sticky="nw", padx=(0, 12), pady=(6, 0))

        self.label_status = ctk.CTkLabel(
            self.status_line,
            text="SELECIONE A PASTA DE ORIGEM PARA COMEÇAR",
            font=("Segoe UI", 30, "bold"),
            text_color=COR_VERDE,
            anchor="w",
            justify="left",
            wraplength=820,
        )
        self.label_status.grid(row=0, column=1, sticky="ew")

        self.label_detalhe = ctk.CTkLabel(
            self.card_status,
            text="",
            height=18,
            font=("Segoe UI", 12),
            text_color=COR_TEXTO_SECUNDARIO,
            anchor="w",
            justify="left",
            wraplength=760,
        )
        self.label_detalhe.grid(row=2, column=0, sticky="w", padx=0, pady=(10, 0))

        self.progress = ctk.CTkProgressBar(
            self.card_status,
            height=2,
            corner_radius=0,
            mode="indeterminate",
            fg_color=COR_BORDA,
            progress_color=COR_VERDE,
            border_width=0,
        )
        self.progress.grid(row=3, column=0, sticky="ew", padx=0, pady=(6, 0))
        self.progress.grid_remove()

        self.actions = ctk.CTkFrame(self.card_status, fg_color="transparent")
        self.actions.grid(row=4, column=0, sticky="ew", padx=0, pady=(12, 0))
        self.actions.grid_columnconfigure(0, weight=0)
        self.actions.grid_columnconfigure(1, weight=0)
        self.actions.grid_columnconfigure(2, weight=0)
        self.actions.grid_columnconfigure(3, weight=1)

        self.botao_inicia = ctk.CTkButton(
            self.actions,
            text="INICIAR PROCESSAMENTO",
            width=248,
            height=48,
            fg_color=COR_VERDE,
            hover_color=COR_VERDE_HOVER,
            text_color=COR_ESCURO,
            text_color_disabled="#5C6B7E",
            font=("Segoe UI", 11, "bold"),
            corner_radius=0,
            command=self.iniciar_processo,
        )
        self.botao_inicia.grid(row=0, column=0, sticky="w")

        self.botao_logs = ctk.CTkButton(
            self.actions,
            text="ABRIR LOGS",
            width=146,
            height=48,
            fg_color="transparent",
            hover_color=COR_CARD_HOVER,
            border_width=1,
            border_color=COR_BORDA,
            text_color=COR_TEXTO,
            font=("Segoe UI", 11, "bold"),
            corner_radius=0,
            command=self.abrir_logs,
        )
        self.botao_logs.grid(row=0, column=1, sticky="w", padx=(12, 0))

        self.botao_editar_login = ctk.CTkButton(
            self.actions,
            text="EDITAR LOGIN",
            width=146,
            height=48,
            fg_color="transparent",
            hover_color=COR_CARD_HOVER,
            border_width=1,
            border_color=COR_BORDA,
            text_color=COR_TEXTO,
            font=("Segoe UI", 11, "bold"),
            corner_radius=0,
            command=self.abrir_janela_login,
        )
        self.botao_editar_login.grid(row=0, column=2, sticky="w", padx=(12, 0))

    def _montar_card_atividade(self):
        self.card_atividade = ctk.CTkFrame(self.container, fg_color="transparent")
        self.card_atividade.grid(row=3, column=0, sticky="nsew", pady=(0, 16))
        self.card_atividade.grid_columnconfigure(0, weight=1)

        self.label_atividade_titulo = ctk.CTkLabel(
            self.card_atividade,
            text="ATIVIDADE",
            height=16,
            font=("Segoe UI", 10, "bold"),
            text_color=COR_TEXTO_SECUNDARIO,
            anchor="w",
        )
        self.label_atividade_titulo.grid(row=0, column=0, sticky="w", padx=0)

        self.label_loja_atual = ctk.CTkLabel(
            self.card_atividade,
            text="AGUARDANDO INÍCIO",
            font=("Segoe UI", 21, "bold"),
            text_color=COR_TEXTO,
            anchor="w",
            justify="left",
            wraplength=820,
        )
        self.label_loja_atual.grid(row=1, column=0, sticky="ew", padx=0, pady=(4, 0))

        self.label_detalhe_atividade = ctk.CTkLabel(
            self.card_atividade,
            text="",
            height=18,
            font=("Segoe UI", 12),
            text_color=COR_TEXTO_SECUNDARIO,
            anchor="w",
            justify="left",
            wraplength=820,
        )
        self.label_detalhe_atividade.grid(
            row=2,
            column=0,
            sticky="ew",
            padx=0,
            pady=(2, 8),
        )

        self.resumo_divisor = ctk.CTkFrame(
            self.card_atividade, height=1, fg_color=COR_BORDA
        )
        self.resumo_divisor.grid(row=3, column=0, sticky="ew", pady=(0, 6))
        self.resumo_divisor.grid_propagate(False)

        self.label_resumo_titulo = ctk.CTkLabel(
            self.card_atividade,
            text="RESUMO DA EXECUÇÃO",
            height=14,
            font=("Segoe UI", 10, "bold"),
            text_color=COR_TEXTO_SECUNDARIO,
            anchor="w",
        )
        self.label_resumo_titulo.grid(
            row=4,
            column=0,
            sticky="w",
            padx=0,
            pady=(0, 4),
        )

        self.resumo_frame = ctk.CTkFrame(self.card_atividade, fg_color="transparent")
        self.resumo_frame.grid(row=5, column=0, sticky="ew", padx=0)
        for coluna in range(4):
            self.resumo_frame.grid_columnconfigure(
                coluna, weight=1, uniform="resumo"
            )

        self.valor_arquivos = self._criar_metrica(self.resumo_frame, 0, "ARQUIVOS")
        self.valor_enviados = self._criar_metrica(self.resumo_frame, 1, "ENVIADOS")
        self.valor_ja_existiam = self._criar_metrica(
            self.resumo_frame, 2, "JÁ EXISTIAM"
        )
        self.valor_falhas = self._criar_metrica(self.resumo_frame, 3, "FALHAS")

    def _criar_metrica(self, pai, coluna, rotulo):
        bloco = ctk.CTkFrame(pai, fg_color="transparent")
        bloco.grid(row=0, column=coluna, sticky="w")

        valor = ctk.CTkLabel(
            bloco,
            text="0",
            height=30,
            font=("Segoe UI", 20, "bold"),
            text_color=COR_VERDE,
            anchor="w",
        )
        valor.grid(row=0, column=0, sticky="w")

        legenda = ctk.CTkLabel(
            bloco,
            text=rotulo,
            height=14,
            font=("Segoe UI", 9, "bold"),
            text_color=COR_TEXTO_SECUNDARIO,
            anchor="w",
        )
        legenda.grid(row=1, column=0, sticky="w", pady=(2, 0))
        return valor

    def _montar_rodape(self):
        self.footer = ctk.CTkFrame(self.container, fg_color="transparent")
        self.footer.grid(row=4, column=0, sticky="ew", pady=(14, 0))
        self.footer.grid_columnconfigure(0, weight=1)

        self.label_destino = ctk.CTkLabel(
            self.footer,
            text=(
                f"DESTINO  ·  {self.pasta_temp}\n"
                f"LOG  ·  {self.log_path}"
            ),
            font=("Segoe UI", 10),
            text_color=COR_TEXTO_SECUNDARIO,
            anchor="w",
            justify="left",
            wraplength=860,
        )
        self.label_destino.grid(row=0, column=0, sticky="ew")

    def _configurar_estado_inicial(self):
        self._resetar_resumo()
        if self.auto_start:
            self.label_caminho.configure(
                text="Meses atual e anterior",
                text_color=COR_TEXTO_SECUNDARIO,
            )
            self.botao_caminho.configure(state="disabled")
            self.botao_logs.configure(state="disabled")
            self._atualizar_botao_iniciar()
            self._definir_status(
                "Execução automática",
                COR_VERDE,
                "Preparando processamento dos meses atual e anterior.",
            )
            self.label_atividade_titulo.configure(text="EM EXECUÇÃO")
            self._definir_atividade(
                "Preparando execução automática",
                "Localizando os documentos dos meses atual e anterior.",
            )
            self.root.after(500, self.iniciar_processo)
        else:
            self._definir_status(
                "Selecione a pasta de origem para começar",
                COR_NEUTRO,
                "",
            )
            self.label_atividade_titulo.configure(text="ATIVIDADE")
            self._definir_atividade("Aguardando início", "")
            self._atualizar_botao_iniciar()

    def _centralizar(self):
        largura = self.root.winfo_width()
        altura = self.root.winfo_height()
        x = (self.root.winfo_screenwidth() // 2) - (largura // 3)
        y = (self.root.winfo_screenheight() // 2) - (altura // 2)
        self.root.geometry(f"+{x}+{y}")

    def _definir_status(self, titulo, cor, detalhe):
        self.status_dot.configure(text_color=cor)
        self.label_status.configure(text=(titulo or "").upper())
        self.label_detalhe.configure(
            text=self._resumir_texto(detalhe or "", LIMITE_DETALHE)
        )
        if detalhe:
            self.label_detalhe.grid()
        else:
            self.label_detalhe.grid_remove()

    def _definir_atividade(self, loja, detalhe):
        self.label_loja_atual.configure(text=(loja or "—").upper())
        self.label_detalhe_atividade.configure(
            text=self._resumir_texto(detalhe or "", LIMITE_DETALHE)
        )

    def _resumir_texto(self, texto, limite):
        if len(texto) <= limite:
            return texto
        return texto[: limite - 3].rstrip() + "..."

    def _resumir_caminho(self, caminho):
        return self._resumir_texto(caminho, LIMITE_CAMINHO)

    def _agendar_tooltip_caminho(self, _evento=None):
        self._cancelar_tooltip_caminho()
        if not self.diretorio_ccts or len(self.diretorio_ccts) <= LIMITE_CAMINHO:
            return
        self._tooltip_after = self.root.after(350, self._mostrar_tooltip_caminho)

    def _mostrar_tooltip_caminho(self):
        self._tooltip_after = None
        if not self.diretorio_ccts or len(self.diretorio_ccts) <= LIMITE_CAMINHO:
            return
        if self._tooltip is not None and self._tooltip.winfo_exists():
            return

        self._tooltip = tk.Toplevel(self.root)
        self._tooltip.wm_overrideredirect(True)
        self._tooltip.configure(bg=COR_BORDA)
        try:
            self._tooltip.attributes("-topmost", True)
        except Exception:
            pass

        quadro = tk.Frame(self._tooltip, bg=COR_CARD, padx=12, pady=9)
        quadro.pack(padx=1, pady=1)
        tk.Label(
            quadro,
            text=self.diretorio_ccts,
            bg=COR_CARD,
            fg=COR_TEXTO,
            font=("Segoe UI", 10),
            justify="left",
            wraplength=640,
        ).pack()

        x = self.label_caminho.winfo_rootx()
        y = self.label_caminho.winfo_rooty() + self.label_caminho.winfo_height() + 8
        self._tooltip.wm_geometry(f"+{x}+{y}")

    def _cancelar_tooltip_caminho(self, _evento=None):
        if self._tooltip_after is not None:
            try:
                self.root.after_cancel(self._tooltip_after)
            except Exception:
                pass
            self._tooltip_after = None
        if self._tooltip is not None:
            try:
                self._tooltip.destroy()
            except Exception:
                pass
            self._tooltip = None

    def _resetar_resumo(self):
        self.resumo = {
            "arquivos": 0,
            "enviados": 0,
            "ja_existiam": 0,
            "falhas": 0,
        }
        self._atualizar_resumo()

    def _atualizar_resumo(self):
        self.valor_arquivos.configure(text=str(self.resumo["arquivos"]))
        self.valor_enviados.configure(text=str(self.resumo["enviados"]))
        self.valor_ja_existiam.configure(text=str(self.resumo["ja_existiam"]))
        self.valor_falhas.configure(text=str(self.resumo["falhas"]))

    def _registrar_resultado_arquivo(self, resultado):
        self.resumo["arquivos"] += 1

        if resultado == "SUCESSO":
            self.resumo["enviados"] += 1
        elif resultado == "JA_EXISTE":
            self.resumo["ja_existiam"] += 1
        else:
            self.resumo["falhas"] += 1

        self._atualizar_resumo()

    def _atualizar_botao_iniciar(self):
        if self.processando:
            self.botao_inicia.configure(
                state="disabled",
                text="PROCESSANDO…",
                fg_color=COR_VERDE,
                text_color=COR_ESCURO,
                border_width=0,
            )
        elif self.auto_start or self.diretorio_ccts:
            self.botao_inicia.configure(
                state="normal",
                text="INICIAR PROCESSAMENTO",
                fg_color=COR_VERDE,
                text_color=COR_ESCURO,
                border_width=0,
            )
        else:
            self.botao_inicia.configure(
                state="disabled",
                text="INICIAR PROCESSAMENTO",
                fg_color=COR_CARD,
                text_color="#4E5C6E",
                border_width=1,
                border_color=COR_BORDA,
            )

    def _set_processing_state(self, ativo):
        self.processando = ativo
        estado = "disabled" if ativo else "normal"
        self.botao_caminho.configure(state=estado)
        self.botao_logs.configure(state=estado)
        self.botao_editar_login.configure(state=estado)
        self._atualizar_botao_iniciar()

        if ativo:
            self.progress.grid()
            self.progress.start()

    def carregar_credenciais(self):
        self.usuario, self.senha = carregar_credenciais()

    def seleciona_diretorio_ccts(self, *_args):
        self._cancelar_tooltip_caminho()
        pasta = filedialog.askdirectory(title="Selecione a pasta de origem")
        if not pasta:
            return

        self.diretorio_ccts = os.path.normpath(pasta)
        self.label_caminho.configure(
            text=self._resumir_caminho(self.diretorio_ccts),
            text_color=COR_TEXTO,
        )
        self._definir_status(
            "Pasta de origem selecionada",
            COR_NEUTRO,
            "",
        )
        self._atualizar_botao_iniciar()

    def iniciar_processo(self):
        if self.processando:
            return

        if not self.usuario or not self.senha:
            if self.auto_start:
                try:
                    with open(self.log_path, "a", encoding="utf-8") as log_file:
                        log(
                            "[ERRO] Usuário ou senha não configurados. Execução automática encerrada.",
                            log_file,
                        )
                except Exception:
                    pass
                self._definir_status(
                    "Erro ao processar",
                    COR_ERRO,
                    "Usuário ou senha não configurados.",
                )
                self.root.after(AUTO_CLOSE_MS, self.root.destroy)
                return
            messagebox.showwarning(
                "Aviso",
                "Usuário ou senha não configurados. Por favor, edite o login.",
            )
            return

        if self.diretorio_ccts is None and not self.auto_start:
            messagebox.showwarning(
                "Aviso",
                "Por favor, selecione a pasta dos CCTs antes de iniciar.",
            )
            return

        self._iniciar_processamento()

    def _iniciar_processamento(self):
        if self.processando:
            return

        self.parar_event.clear()
        self.erro_fatal = None
        self._set_processing_state(True)
        self._definir_status(
            "Processando...",
            COR_VERDE,
            "Enviando os documentos ao ComuniQ e preparando o relatório.",
        )
        self.label_atividade_titulo.configure(text="EM EXECUÇÃO")
        self._definir_atividade(
            "Preparando...",
            "Abrindo o navegador e validando as credenciais.",
        )
        self._resetar_resumo()

        def executar():
            log_path = None
            erro = None
            try:
                log_path = os.path.join(APP_CONFIG_DIR, "log_historico.txt")
                processar_arquivos(
                    self.diretorio_ccts,
                    self.pasta_temp,
                    log_path,
                    self.usuario,
                    self.senha,
                    self.parar_event,
                    automatico=self.auto_start,
                    progresso_callback=self._receber_progresso,
                )
            except Exception as exc:
                erro = str(exc)
                if self.auto_start:
                    try:
                        with open(log_path or self.log_path, "a", encoding="utf-8") as log_file:
                            log(
                                f"[ERRO FATAL] Erro inesperado durante a execução automática: {exc}",
                                log_file,
                                espacamento=True,
                            )
                            log(traceback.format_exc(), log_file)
                    except Exception:
                        pass

            self.root.after(0, self._finalizar_processamento, log_path, erro)

        self.thread = threading.Thread(target=executar)
        self.thread.start()

    def _receber_progresso(self, dados):
        try:
            self.root.after(0, self._atualizar_atividade, dados)
        except Exception:
            pass

    def _traduzir_resultado(self, resultado):
        return {
            "SUCESSO": "ENVIADO",
            "JA_EXISTE": "JÁ EXISTIA",
            "FALHA": "FALHA",
        }.get(resultado, resultado or "")

    def _atualizar_atividade(self, dados):
        etapa = dados.get("etapa")

        if etapa == "navegador":
            self._definir_atividade(
                "Abrindo navegador",
                dados.get("mensagem", "Entrando no ComuniQ."),
            )
        elif etapa == "inicio":
            total = dados.get("total_arquivos", 0)
            pasta = dados.get("pasta", "")
            self._definir_atividade(
                "Documentos localizados",
                f"{total} arquivo(s) serão avaliados em {pasta}",
            )
        elif etapa == "reprocessar":
            total = dados.get("total_arquivos", 0)
            pasta = dados.get("pasta", "")
            self._definir_atividade(
                "Reprocessando pendências",
                f"{total} arquivo(s) em {pasta}",
            )
        elif etapa == "arquivo":
            arquivo = dados.get("arquivo", "")
            total = dados.get("total_lojas", 0)
            self._definir_atividade(
                arquivo,
                f"{total} loja(s) serão avaliadas.",
            )
        elif etapa == "loja":
            loja = dados.get("loja", "")
            arquivo = dados.get("arquivo", "")
            indice = dados.get("indice", 0)
            total = dados.get("total_lojas", 0)

            partes = [arquivo] if arquivo else []
            if indice and total:
                partes.append(f"Loja {indice} de {total}")

            self._definir_atividade(
                f"Loja {loja}",
                " | ".join(partes) or "Loja em processamento.",
            )
        elif etapa == "loja_resultado":
            loja = dados.get("loja", "")
            arquivo = dados.get("arquivo", "")
            resultado = self._traduzir_resultado(dados.get("resultado", ""))
            self._definir_atividade(
                f"Loja {loja}",
                f"{arquivo} | Resultado: {resultado}",
            )
        elif etapa == "arquivo_finalizado":
            arquivo = dados.get("arquivo", "")
            resultado = dados.get("resultado", "")
            motivo = dados.get("motivo") or ""
            detalhe = f"Arquivo {dados.get('indice', '')} de {dados.get('total_arquivos', '')} | Resultado: {self._traduzir_resultado(resultado)}"
            if motivo:
                detalhe += f" | {motivo}"
            self._definir_atividade(arquivo, detalhe)
            self._registrar_resultado_arquivo(resultado)
        elif etapa == "finalizando":
            self.erro_fatal = dados.get("erro_fatal")
            self._definir_atividade(
                "Finalizando execução",
                dados.get("mensagem", "Finalizando a execução."),
            )

    def _finalizar_processamento(self, log_path, erro):
        self.processando = False
        try:
            self.progress.stop()
        except Exception:
            pass
        self.progress.grid_remove()
        self.label_atividade_titulo.configure(text="ÚLTIMA ATIVIDADE")

        if log_path:
            self.log_path = log_path

        if erro:
            self._definir_status(
                "Erro ao processar",
                COR_ERRO,
                f"{erro}. Consulte o log para mais detalhes.",
            )
        elif self.erro_fatal:
            self._definir_status(
                "Erro ao processar",
                COR_ERRO,
                f"{self.erro_fatal}. Consulte o log para mais detalhes.",
            )
        else:
            detalhe = f"Log: {self.log_path}" if self.log_path else "Processamento concluído."
            self._definir_status("Execução concluída", COR_SUCESSO, detalhe)

        self._atualizar_rodape()

        if self.auto_start:
            self.root.after(AUTO_CLOSE_MS, self.root.destroy)
        else:
            self._set_processing_state(False)

    def _atualizar_rodape(self):
        self.label_destino.configure(
            text=(
                f"DESTINO  ·  {self.pasta_temp}\n"
                f"LOG  ·  {self.log_path}"
            )
        )

    def abrir_logs(self):
        try:
            os.makedirs(APP_CONFIG_DIR, exist_ok=True)
            os.startfile(APP_CONFIG_DIR)
        except Exception as erro:
            messagebox.showerror(
                "Abrir logs",
                f"Não foi possível abrir a pasta de logs.\n\n{erro}",
            )

    def mostra_guia(self):
        janela = ctk.CTkToplevel(self.root)
        janela.title("Como usar")
        janela.geometry("680x480")
        janela.minsize(560, 400)
        janela.configure(fg_color=COR_FUNDO)
        janela.transient(self.root)
        janela.grid_rowconfigure(1, weight=1)
        janela.grid_columnconfigure(0, weight=1)

        titulo = ctk.CTkLabel(
            janela,
            text="COMO USAR O ROBÔ",
            font=("Segoe UI", 20, "bold"),
            text_color=COR_VERDE,
            anchor="w",
        )
        titulo.grid(row=0, column=0, sticky="w", padx=24, pady=(22, 12))

        texto = ctk.CTkTextbox(
            janela,
            fg_color=COR_CARD,
            border_width=1,
            border_color=COR_BORDA,
            text_color=COR_TEXTO,
            font=("Segoe UI", 13),
            corner_radius=0,
        )
        texto.grid(row=1, column=0, sticky="nsew", padx=24, pady=(0, 16))
        texto.insert(
            "1.0",
            "Os documentos da pasta selecionada devem seguir o seguinte padrão de nomenclatura:\n\n"
            "- Número das lojas_cidade_código do sindicato_vigência\n\n"
            "Exemplos:\n"
            "- 160,310_LAJEADO_SIND0088_2026\n"
            "- 642,285,187,259_VIAMAO_SIND0173_2024-2026\n\n"
            "Observações:\n"
            "- Se necessário, pode-se adicionar complementos após a vigência, como: "
            "_ADITIVOS, _FERIADOS, _TRABALHO NO DIA X, etc.\n\n"
            "  Exemplo:\n"
            "- 127,139,140,155_LAJEADO_SIND0077_2025-2027_ADITIVOS\n"
            "- 503,631,659,671_UMUARAMA_SIND0180_2025-2027_TRABALHO NO DIA 19.04.2025\n\n"
            "No ComuniQ eles serão salvos com o padrão:\n"
            "- CCT_vigência\n"
            "- CCT_vigência_complemento (caso exista complemento)",
        )
        texto.configure(state="disabled")

        botao_fechar = ctk.CTkButton(
            janela,
            text="FECHAR",
            width=120,
            height=44,
            fg_color="transparent",
            hover_color=COR_CARD_HOVER,
            border_width=1,
            border_color=COR_BORDA,
            text_color=COR_TEXTO,
            font=("Segoe UI", 11, "bold"),
            corner_radius=0,
            command=janela.destroy,
        )
        botao_fechar.grid(row=2, column=0, sticky="e", padx=24, pady=(0, 20))

        janela.grab_set()
        janela.after(60, janela.lift)

    def abrir_janela_login(self):
        if self.processando:
            return

        self.login_window = ctk.CTkToplevel(self.root)
        self.login_window.title("Editar Credenciais de Login")
        self.login_window.geometry("460x360")
        self.login_window.minsize(420, 340)
        self.login_window.configure(fg_color=COR_FUNDO)
        self.login_window.transient(self.root)
        self.login_window.grab_set()
        self.login_window.after(60, self.login_window.lift)
        self.login_window.grid_columnconfigure(0, weight=1)

        titulo = ctk.CTkLabel(
            self.login_window,
            text="CREDENCIAIS DO COMUNIQ",
            font=("Segoe UI", 18, "bold"),
            text_color=COR_VERDE,
            anchor="w",
        )
        titulo.grid(row=0, column=0, sticky="w", padx=28, pady=(24, 4))

        descricao = ctk.CTkLabel(
            self.login_window,
            text="Usuário e senha usados pelo robô para acessar o ComuniQ.",
            font=("Segoe UI", 12),
            text_color=COR_TEXTO_SECUNDARIO,
            anchor="w",
            justify="left",
            wraplength=390,
        )
        descricao.grid(row=1, column=0, sticky="w", padx=28, pady=(0, 16))

        form = ctk.CTkFrame(self.login_window, fg_color="transparent")
        form.grid(row=2, column=0, sticky="ew", padx=28)
        form.grid_columnconfigure(0, weight=1)

        label_usuario = ctk.CTkLabel(
            form,
            text="USUÁRIO (E-MAIL)",
            font=("Segoe UI", 10, "bold"),
            text_color=COR_TEXTO_SECUNDARIO,
            anchor="w",
        )
        label_usuario.grid(row=0, column=0, sticky="w")

        self.usuario_entry = ctk.CTkEntry(
            form,
            height=42,
            fg_color=COR_CARD,
            border_width=1,
            border_color=COR_BORDA,
            text_color=COR_TEXTO,
            font=("Segoe UI", 12),
            corner_radius=0,
        )
        self.usuario_entry.insert(0, self.usuario)
        self.usuario_entry.grid(row=1, column=0, sticky="ew", pady=(6, 14))

        label_senha = ctk.CTkLabel(
            form,
            text="SENHA",
            font=("Segoe UI", 10, "bold"),
            text_color=COR_TEXTO_SECUNDARIO,
            anchor="w",
        )
        label_senha.grid(row=2, column=0, sticky="w")

        self.senha_entry = ctk.CTkEntry(
            form,
            height=42,
            show="*",
            fg_color=COR_CARD,
            border_width=1,
            border_color=COR_BORDA,
            text_color=COR_TEXTO,
            font=("Segoe UI", 12),
            corner_radius=0,
        )
        self.senha_entry.insert(0, self.senha)
        self.senha_entry.grid(row=3, column=0, sticky="ew", pady=(6, 12))

        self.show_password_checkbox = ctk.CTkCheckBox(
            form,
            text="Mostrar senha",
            font=("Segoe UI", 11),
            checkbox_width=18,
            checkbox_height=18,
            corner_radius=0,
            fg_color=COR_VERDE,
            hover_color=COR_VERDE_HOVER,
            command=self.toggle_password_visibility,
        )
        self.show_password_checkbox.grid(row=4, column=0, sticky="w")

        actions = ctk.CTkFrame(self.login_window, fg_color="transparent")
        actions.grid(row=3, column=0, sticky="ew", padx=28, pady=(22, 20))
        actions.grid_columnconfigure(0, weight=1)

        botao_salvar = ctk.CTkButton(
            actions,
            text="SALVAR",
            width=150,
            height=44,
            fg_color=COR_VERDE,
            hover_color=COR_VERDE_HOVER,
            text_color=COR_ESCURO,
            font=("Segoe UI", 11, "bold"),
            corner_radius=0,
            command=self.salvar_e_gravar_credenciais,
        )
        botao_salvar.grid(row=0, column=0, sticky="w")

        botao_cancelar = ctk.CTkButton(
            actions,
            text="CANCELAR",
            width=130,
            height=44,
            fg_color="transparent",
            hover_color=COR_CARD_HOVER,
            border_width=1,
            border_color=COR_BORDA,
            text_color=COR_TEXTO,
            font=("Segoe UI", 11, "bold"),
            corner_radius=0,
            command=self.login_window.destroy,
        )
        botao_cancelar.grid(row=0, column=1, sticky="e")

    def toggle_password_visibility(self):
        if self.show_password_checkbox.get() == 1:
            self.senha_entry.configure(show="")
        else:
            self.senha_entry.configure(show="*")

    def salvar_e_gravar_credenciais(self):
        novo_usuario = self.usuario_entry.get().strip()
        nova_senha = self.senha_entry.get()
        if not novo_usuario or not nova_senha:
            messagebox.showwarning(
                "Aviso",
                "Usuário e senha não podem estar em branco.",
                parent=self.login_window,
            )
            return

        self.usuario = novo_usuario
        self.senha = nova_senha
        try:
            salvar_credenciais(self.usuario, self.senha)
        except Exception as erro:
            messagebox.showerror(
                "Erro",
                "Não foi possível salvar as credenciais no arquivo de configuração.\n"
                f"Erro: {erro}",
                parent=self.login_window,
            )
            return

        messagebox.showinfo(
            "Sucesso",
            "Credenciais salvas com sucesso!",
            parent=self.login_window,
        )
        self.login_window.destroy()

    def on_closing(self):
        if self.processando:
            if not messagebox.askokcancel(
                "Encerrar",
                "O processamento está em andamento. Deseja interromper e sair?",
            ):
                return
            self.parar_event.set()
        elif not messagebox.askokcancel("Sair", "Deseja realmente encerrar o robô?"):
            return

        self._cancelar_tooltip_caminho()
        if self.thread and self.thread.is_alive():
            self.thread.join(timeout=5)
        self.root.destroy()


if __name__ == "__main__":
    ctk.set_appearance_mode("dark")
    ctk.set_default_color_theme(resource_path("marsh.json"))

    root = ctk.CTk()
    app = App(root, auto_start="--auto-start" in sys.argv)
    root.mainloop()
