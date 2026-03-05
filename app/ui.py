import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk
from tkinter.scrolledtext import ScrolledText

import pandas as pd

from .backend import ChamadosBackend

CONFIG_PATH = Path(__file__).resolve().parent.parent / "config_local.json"
PLANILHA_FILETYPES = [("Planilhas", "*.ods *.xlsx *.xls")]


class ImportadorGLPIApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Importador de Chamados GLPI")
        self.root.geometry("1100x730")
        self.root.minsize(1000, 650)

        self.backend = ChamadosBackend("", CONFIG_PATH)

        self._build_ui()
        self._carregar_config_local()

    def _build_ui(self):
        style = ttk.Style()
        style.theme_use("clam")
        style.configure("Title.TLabel", font=("Segoe UI", 16, "bold"))
        style.configure("SubTitle.TLabel", font=("Segoe UI", 10))
        style.configure("Info.TLabel", font=("Segoe UI", 10))
        style.configure("Primary.TButton", font=("Segoe UI", 10, "bold"))
        style.configure(
            "Green.Horizontal.TProgressbar",
            troughcolor="#E5E7EB",
            background="#16A34A",
            lightcolor="#16A34A",
            darkcolor="#16A34A",
            bordercolor="#E5E7EB",
        )

        self.canvas = tk.Canvas(self.root, highlightthickness=0)
        self.canvas.pack(side="left", fill="both", expand=True)
        self.scrollbar = ttk.Scrollbar(self.root, orient="vertical", command=self.canvas.yview)
        self.scrollbar.pack(side="right", fill="y")
        self.canvas.configure(yscrollcommand=self.scrollbar.set)

        self.content = ttk.Frame(self.canvas)
        self.canvas_window = self.canvas.create_window((0, 0), window=self.content, anchor="nw")
        self._configurar_scroll()

        header = ttk.Frame(self.content, padding=(16, 14))
        header.pack(fill="x")
        ttk.Label(header, text="Importacao de Chamados para o GLPI", style="Title.TLabel").pack(anchor="w")
        ttk.Label(
            header,
            text="Selecione a planilha, revise os dados e execute a importacao com acompanhamento em tempo real.",
            style="SubTitle.TLabel",
        ).pack(anchor="w", pady=(4, 0))

        auth_frame = ttk.LabelFrame(self.content, text="1) Autenticacao", padding=12)
        auth_frame.pack(fill="x", padx=16, pady=(0, 10))
        self.url_var = tk.StringVar(value="")
        self.user_token_var = tk.StringVar(value="")
        self.app_token_var = tk.StringVar(value="")
        self.status_auth_var = tk.StringVar(value="Nao autenticado.")

        ttk.Label(auth_frame, text="URL API").grid(row=0, column=0, sticky="w")
        ttk.Entry(auth_frame, textvariable=self.url_var).grid(row=1, column=0, sticky="ew", padx=(0, 10))
        ttk.Label(auth_frame, text="User Token").grid(row=0, column=1, sticky="w")
        ttk.Entry(auth_frame, textvariable=self.user_token_var).grid(row=1, column=1, sticky="ew", padx=(0, 10))
        ttk.Label(auth_frame, text="App Token").grid(row=0, column=2, sticky="w")
        ttk.Entry(auth_frame, textvariable=self.app_token_var).grid(row=1, column=2, sticky="ew", padx=(0, 10))
        ttk.Button(auth_frame, text="Autenticar", style="Primary.TButton", command=self.autenticar).grid(
            row=1, column=3, sticky="ew"
        )
        ttk.Label(auth_frame, textvariable=self.status_auth_var, style="Info.TLabel").grid(
            row=2, column=0, columnspan=4, sticky="w", pady=(8, 0)
        )
        self.salvar_tokens_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(
            auth_frame,
            text="Salvar tokens localmente",
            variable=self.salvar_tokens_var,
        ).grid(row=3, column=0, columnspan=4, sticky="w", pady=(8, 0))
        auth_frame.columnconfigure(0, weight=2)
        auth_frame.columnconfigure(1, weight=2)
        auth_frame.columnconfigure(2, weight=2)
        auth_frame.columnconfigure(3, weight=1)

        arquivo_frame = ttk.LabelFrame(self.content, text="2) Arquivo", padding=12)
        arquivo_frame.pack(fill="x", padx=16, pady=(0, 10))
        self.caminho_var = tk.StringVar(value="Nenhum arquivo selecionado.")
        ttk.Entry(arquivo_frame, textvariable=self.caminho_var, state="readonly").pack(
            side="left", fill="x", expand=True, padx=(0, 10)
        )
        self.selecionar_btn = ttk.Button(
            arquivo_frame, text="Selecionar planilha", style="Primary.TButton", command=self.selecionar_planilha
        )
        self.selecionar_btn.pack(side="left")
        self.selecionar_btn.configure(state="disabled")

        resumo_frame = ttk.LabelFrame(self.content, text="3) Validacao", padding=12)
        resumo_frame.pack(fill="x", padx=16, pady=(0, 10))
        self.status_colunas_var = tk.StringVar(value="Colunas: aguardando arquivo.")
        self.status_api_var = tk.StringVar(value="Validacao API: aguardando.")
        self.total_var = tk.StringVar(value="Total de linhas: -")
        self.validas_var = tk.StringVar(value="Linhas validas: -")
        self.invalidas_var = tk.StringVar(value="Linhas invalidas: -")
        ttk.Label(resumo_frame, textvariable=self.status_colunas_var, style="Info.TLabel").grid(
            row=0, column=0, sticky="w", padx=(0, 20), pady=2
        )
        ttk.Label(resumo_frame, textvariable=self.total_var, style="Info.TLabel").grid(
            row=0, column=1, sticky="w", padx=(0, 20), pady=2
        )
        ttk.Label(resumo_frame, textvariable=self.validas_var, style="Info.TLabel").grid(
            row=0, column=2, sticky="w", padx=(0, 20), pady=2
        )
        ttk.Label(resumo_frame, textvariable=self.invalidas_var, style="Info.TLabel").grid(
            row=0, column=3, sticky="w", pady=2
        )
        ttk.Label(resumo_frame, textvariable=self.status_api_var, style="Info.TLabel").grid(
            row=1, column=0, columnspan=4, sticky="w", pady=(4, 0)
        )

        preview_frame = ttk.LabelFrame(self.content, text="4) Pre-visualizacao (primeiras 100 linhas)", padding=12)
        preview_frame.pack(fill="both", expand=True, padx=16, pady=(0, 10))
        self.tree = ttk.Treeview(preview_frame, show="headings")
        vsb = ttk.Scrollbar(preview_frame, orient="vertical", command=self.tree.yview)
        hsb = ttk.Scrollbar(preview_frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        self.tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        preview_frame.columnconfigure(0, weight=1)
        preview_frame.rowconfigure(0, weight=1)

        import_frame = ttk.LabelFrame(self.content, text="5) Importacao", padding=12)
        import_frame.pack(fill="x", padx=16, pady=(0, 10))
        self.importar_btn = ttk.Button(
            import_frame,
            text="Importar Chamados",
            style="Primary.TButton",
            state="disabled",
            command=self.importar_chamados,
        )
        self.importar_btn.pack(side="left")
        self.fechar_btn = ttk.Button(
            import_frame,
            text="Fechar Chamados (planilha)",
            style="Primary.TButton",
            state="disabled",
            command=self.fechar_chamados_planilha,
        )
        self.fechar_btn.pack(side="left", padx=(10, 0))
        self.solucionar_btn = ttk.Button(
            import_frame,
            text="Solucionar Chamados (planilha)",
            style="Primary.TButton",
            state="disabled",
            command=self.solucionar_chamados_planilha,
        )
        self.solucionar_btn.pack(side="left", padx=(10, 0))
        self.validar_api_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(
            import_frame,
            text="Validar tecnico/requerente/categoria/localizacao na API",
            variable=self.validar_api_var,
        ).pack(side="left", padx=(12, 0))
        self.usar_html_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(import_frame, text="Preservar formatacao em HTML", variable=self.usar_html_var).pack(
            side="left", padx=(12, 0)
        )
        self.buscar_nomes_btn = ttk.Button(
            import_frame,
            text="Mostrar nomes (API)",
            style="Primary.TButton",
            state="disabled",
            command=self.buscar_nomes_api,
        )
        self.buscar_nomes_btn.pack(side="left", padx=(12, 0))
        self.status_importacao_var = tk.StringVar(value="Aguardando arquivo valido para liberar a importacao.")
        ttk.Label(import_frame, textvariable=self.status_importacao_var, style="Info.TLabel").pack(side="left", padx=(12, 0))

        progresso_frame = ttk.Frame(self.content, padding=(16, 0, 16, 0))
        progresso_frame.pack(fill="x")
        self.progresso = ttk.Progressbar(progresso_frame, mode="determinate", style="Green.Horizontal.TProgressbar")
        self.progresso.pack(fill="x")
        self.progresso_var = tk.StringVar(value="0%")
        self.contador_var = tk.StringVar(value="0 de 0")
        ttk.Label(progresso_frame, textvariable=self.progresso_var, style="Info.TLabel").pack(anchor="w", pady=(4, 0))
        ttk.Label(progresso_frame, textvariable=self.contador_var, style="Info.TLabel").pack(anchor="w")

        log_frame = ttk.LabelFrame(self.content, text="6) Log", padding=12)
        log_frame.pack(fill="both", expand=False, padx=16, pady=(10, 14))
        self.log = ScrolledText(log_frame, height=10, font=("Consolas", 10))
        self.log.pack(fill="both", expand=True)
        self.log.tag_configure("erro", foreground="#B91C1C", font=("Consolas", 10, "bold"))
        self.log.configure(state="disabled")
        self.log_msg("[INFO] Informe URL, user token e app token para autenticar.")

    def log_msg(self, mensagem):
        self.log.configure(state="normal")
        tag = "erro" if "[ERRO]" in mensagem else None
        if tag:
            self.log.insert(tk.END, mensagem + "\n", tag)
        else:
            self.log.insert(tk.END, mensagem + "\n")
        self.log.see(tk.END)
        self.log.configure(state="disabled")
        self.root.update_idletasks()

    def _configurar_scroll(self):
        self.content.bind("<Configure>", lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all")))
        self.canvas.bind("<Configure>", lambda e: self.canvas.itemconfigure(self.canvas_window, width=e.width))
        self.canvas.bind_all("<MouseWheel>", self._on_mousewheel)
        self.canvas.bind_all("<Button-4>", self._on_mousewheel)
        self.canvas.bind_all("<Button-5>", self._on_mousewheel)

    def _on_mousewheel(self, event):
        if event.num == 4:
            self.canvas.yview_scroll(-1, "units")
        elif event.num == 5:
            self.canvas.yview_scroll(1, "units")
        elif getattr(event, "delta", 0):
            self.canvas.yview_scroll(int(-event.delta / 120), "units")

    def _set_botoes_operacao(self, habilitado):
        estado = "normal" if habilitado else "disabled"
        self.importar_btn.configure(state=estado)
        self.fechar_btn.configure(state=estado)
        self.solucionar_btn.configure(state=estado)

    def _reset_progresso(self, total):
        self.progresso.configure(maximum=total, value=0)
        self.progresso_var.set("0%")
        self.contador_var.set(f"0 de {total}")
        self.root.update_idletasks()

    def _preencher_tree_com_dataframe(self, dataframe):
        self.tree.delete(*self.tree.get_children())
        colunas = list(dataframe.columns)
        self.tree["columns"] = colunas
        for col in colunas:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=150, stretch=True)
        for _, row in dataframe.iterrows():
            valores = ["" if pd.isna(v) else str(v) for v in row.tolist()]
            self.tree.insert("", "end", values=valores)

    def preencher_preview(self):
        preview_df = self.backend.construir_preview_df(limite=100)
        if preview_df is not None:
            self._preencher_tree_com_dataframe(preview_df)

    def _atualizar_progresso(self, atual, total):
        self.progresso.configure(value=atual)
        percentual = int((atual / total) * 100) if total else 0
        self.progresso_var.set(f"{percentual}%")
        self.contador_var.set(f"{atual} de {total}")
        self.root.update_idletasks()

    def selecionar_planilha(self):
        if not self.backend.autenticado:
            messagebox.showwarning("Autenticacao", "Autentique-se antes de selecionar a planilha.")
            return
        caminho = filedialog.askopenfilename(title="Selecione a planilha", filetypes=PLANILHA_FILETYPES)
        if not caminho:
            return

        self.caminho_var.set(caminho)
        try:
            self.backend.carregar_planilha_importacao(caminho)
        except Exception as e:
            messagebox.showerror("Erro ao ler planilha", str(e))
            self.log_msg(f"[ERRO] Falha ao ler planilha: {e}")
            return

        self.preencher_preview()
        self.validar_planilha()
        self.buscar_nomes_btn.configure(state="normal")

    def validar_planilha(self):
        resultado = self.backend.validar_planilha_atual()
        if resultado is None:
            return

        self.total_var.set(f"Total de linhas: {resultado['total']}")
        self.validas_var.set(f"Linhas validas: {resultado['validas']}")
        self.invalidas_var.set(f"Linhas invalidas: {resultado['invalidas']}")
        self.status_api_var.set("Validacao API: pendente (importe ou clique em 'Mostrar nomes (API)').")

        if self.backend.colunas_faltantes:
            faltantes = ", ".join(self.backend.colunas_faltantes)
            self.status_colunas_var.set(f"Colunas: faltando {faltantes}")
            self.status_importacao_var.set("Importacao bloqueada: faltam colunas obrigatorias.")
            self.importar_btn.configure(state="disabled")
            self.log_msg(f"[ERRO] Colunas obrigatorias ausentes: {faltantes}")
            return

        self.status_colunas_var.set("Colunas: OK")
        if resultado["validas"] == 0:
            self.status_importacao_var.set("Importacao bloqueada: nenhuma linha valida.")
            self.importar_btn.configure(state="disabled")
            self.log_msg("[ALERTA] Planilha sem linhas validas para importar.")
            return

        if not self.backend.autenticado:
            self.status_importacao_var.set("Importacao bloqueada: autentique-se primeiro.")
            self.importar_btn.configure(state="disabled")
            return

        self.status_importacao_var.set("Planilha valida. Pronta para importacao.")
        self.importar_btn.configure(state="normal")
        self.log_msg(
            f"[OK] Planilha validada: {resultado['validas']} validas e {resultado['invalidas']} invalidas."
        )
        for linha, motivo in self.backend.linhas_invalidas[:10]:
            self.log_msg(f"[AVISO] Linha {linha}: {motivo}")
        if len(self.backend.linhas_invalidas) > 10:
            restante = len(self.backend.linhas_invalidas) - 10
            self.log_msg(f"[AVISO] ... e mais {restante} linhas invalidas.")

    def autenticar(self):
        url = self.url_var.get().strip().rstrip("/")
        user_token = self.user_token_var.get().strip()
        app_token = self.app_token_var.get().strip()
        if not url or not user_token or not app_token:
            messagebox.showerror("Autenticacao", "Preencha URL API, User Token e App Token.")
            return

        self.status_auth_var.set("Autenticando...")
        self.root.update_idletasks()
        try:
            self.backend.autenticar(url, user_token, app_token)
        except Exception as e:
            self.status_auth_var.set("Falha na autenticacao.")
            self.selecionar_btn.configure(state="disabled")
            self.buscar_nomes_btn.configure(state="disabled")
            self.importar_btn.configure(state="disabled")
            self.fechar_btn.configure(state="disabled")
            self.solucionar_btn.configure(state="disabled")
            self.status_importacao_var.set("Importacao bloqueada: autentique-se primeiro.")
            self.log_msg(f"[ERRO] Falha de autenticacao: {e}")
            messagebox.showerror("Autenticacao", f"Nao foi possivel autenticar no GLPI.\n{e}")
            return

        self.status_auth_var.set("Autenticado com sucesso.")
        self.selecionar_btn.configure(state="normal")
        self.fechar_btn.configure(state="normal")
        self.solucionar_btn.configure(state="normal")
        if self.backend.df is not None:
            self.buscar_nomes_btn.configure(state="normal")
        self._salvar_config_local()
        self.log_msg("[OK] Autenticacao validada com sucesso.")
        if self.backend.df is not None:
            self.validar_planilha()
        else:
            self.status_importacao_var.set("Autenticado. Selecione uma planilha para continuar.")

    def importar_chamados(self):
        if not self.backend.autenticado or self.backend.cliente is None:
            messagebox.showwarning("Autenticacao", "Autentique-se antes de importar.")
            return
        if self.backend.df is None:
            return
        if self.backend.colunas_faltantes:
            messagebox.showerror("Planilha invalida", "Existem colunas obrigatorias faltando.")
            return

        linhas_validas = len(self.backend.df) - len(self.backend.linhas_invalidas)
        confirmar = messagebox.askyesno(
            "Confirmar importacao",
            f"Deseja importar {linhas_validas} chamados validos?\n"
            f"{len(self.backend.linhas_invalidas)} linhas invalidas serao ignoradas.",
        )
        if not confirmar:
            return

        self._set_botoes_operacao(habilitado=False)
        self.log_msg("[INFO] Iniciando importacao...")
        if self.usar_html_var.get():
            self.log_msg("[INFO] Descricao sera enviada em HTML basico para preservar formatacao.")
        self.status_importacao_var.set("Importacao em andamento...")
        self._reset_progresso(len(self.backend.df))

        try:
            resultado = self.backend.importar_chamados(
                validar_api=self.validar_api_var.get(),
                usar_html=self.usar_html_var.get(),
                log_cb=self.log_msg,
                progresso_cb=self._atualizar_progresso,
            )
        except Exception as e:
            self.status_importacao_var.set("Falha ao importar.")
            self.log_msg(f"[ERRO] {e}")
            messagebox.showerror("Erro", f"Nao foi possivel concluir importacao.\n{e}")
            self._set_botoes_operacao(habilitado=True)
            return

        if self.validar_api_var.get():
            self.status_api_var.set(self.backend.resumo_referencias_api())
            self.preencher_preview()
        else:
            self.status_api_var.set("Validacao API: desativada para esta importacao.")

        self.status_importacao_var.set(resultado["resumo"])
        self.log_msg("[INFO] " + resultado["resumo"])
        try:
            log_texto = self.log.get("1.0", tk.END).strip() + "\n"
            caminho_xlsx, caminho_log = self.backend.salvar_relatorio_importacao(resultado, log_texto)
            self.log_msg(f"[OK] Planilha importada com coluna ticket_id salva em: {caminho_xlsx}")
            self.log_msg(f"[OK] Log salvo em: {caminho_log}")
        except Exception as e:
            self.log_msg(f"[AVISO] Nao foi possivel salvar relatorio/log da importacao: {e}")
        messagebox.showinfo("Importacao finalizada", resultado["resumo"])
        self._set_botoes_operacao(habilitado=True)

    def fechar_chamados_planilha(self):
        if not self.backend.autenticado or self.backend.cliente is None:
            messagebox.showwarning("Autenticacao", "Autentique-se antes de fechar chamados.")
            return
        caminho = filedialog.askopenfilename(title="Selecione a planilha de fechamento", filetypes=PLANILHA_FILETYPES)
        if not caminho:
            return

        try:
            df = self.backend.preparar_planilha_fechamento(caminho)
        except ValueError as e:
            messagebox.showerror("Planilha invalida", str(e))
            self.log_msg(f"[ERRO] {e}")
            return
        except Exception as e:
            self.log_msg(f"[ERRO] Falha ao ler planilha de fechamento: {e}")
            messagebox.showerror("Erro", f"Nao foi possivel ler a planilha.\n{e}")
            return

        if len(df) == 0:
            messagebox.showwarning("Planilha vazia", "A planilha de fechamento nao possui linhas.")
            return

        self._preencher_tree_com_dataframe(df.head(100))
        self.status_importacao_var.set("Pre-visualizacao de fechamento carregada.")
        self.log_msg("[INFO] Pre-visualizacao da planilha de fechamento exibida (primeiras 100 linhas).")

        confirmar = messagebox.askyesno(
            "Confirmar fechamento",
            f"Deseja processar fechamento de {len(df)} linha(s)?\nA pre-visualizacao foi atualizada.",
        )
        if not confirmar:
            return

        self._set_botoes_operacao(habilitado=False)
        self.status_importacao_var.set("Fechamento em andamento...")
        self.log_msg(f"[INFO] Iniciando fechamento em lote ({len(df)} linha(s))...")
        if self.usar_html_var.get() and "solucao" in df.columns:
            self.log_msg("[INFO] Solucao sera enviada em HTML basico para preservar formatacao.")
        self._reset_progresso(len(df))

        try:
            resultado = self.backend.fechar_chamados(
                df=df,
                usar_html=self.usar_html_var.get(),
                log_cb=self.log_msg,
                progresso_cb=self._atualizar_progresso,
            )
        except Exception as e:
            self.log_msg(f"[ERRO] Nao foi possivel concluir fechamento: {e}")
            messagebox.showerror("Erro", f"Nao foi possivel concluir fechamento.\n{e}")
            self.status_importacao_var.set("Falha ao fechar chamados.")
            self._set_botoes_operacao(habilitado=True)
            return

        self.status_importacao_var.set(resultado["resumo"])
        self.log_msg("[INFO] " + resultado["resumo"])
        messagebox.showinfo("Fechamento finalizado", resultado["resumo"])
        self._set_botoes_operacao(habilitado=True)

    def solucionar_chamados_planilha(self):
        if not self.backend.autenticado or self.backend.cliente is None:
            messagebox.showwarning("Autenticacao", "Autentique-se antes de solucionar chamados.")
            return
        caminho = filedialog.askopenfilename(title="Selecione a planilha de solucao", filetypes=PLANILHA_FILETYPES)
        if not caminho:
            return

        try:
            df = self.backend.preparar_planilha_solucao(caminho)
        except ValueError as e:
            messagebox.showerror("Planilha invalida", str(e))
            self.log_msg(f"[ERRO] {e}")
            return
        except Exception as e:
            self.log_msg(f"[ERRO] Falha ao ler planilha de solucao: {e}")
            messagebox.showerror("Erro", f"Nao foi possivel ler a planilha.\n{e}")
            return

        if len(df) == 0:
            messagebox.showwarning("Planilha vazia", "A planilha de solucao nao possui linhas.")
            return

        self._preencher_tree_com_dataframe(df.head(100))
        self.status_importacao_var.set("Pre-visualizacao de solucao carregada.")
        self.log_msg("[INFO] Pre-visualizacao da planilha de solucao exibida (primeiras 100 linhas).")

        confirmar = messagebox.askyesno(
            "Confirmar solucoes",
            f"Deseja processar inclusao de solucao em {len(df)} linha(s)?\nA pre-visualizacao foi atualizada.",
        )
        if not confirmar:
            return

        self._set_botoes_operacao(habilitado=False)
        self.status_importacao_var.set("Inclusao de solucoes em andamento...")
        self.log_msg(f"[INFO] Iniciando inclusao de solucoes em lote ({len(df)} linha(s))...")
        if self.usar_html_var.get():
            self.log_msg("[INFO] Solucao sera enviada em HTML basico para preservar formatacao.")
        self._reset_progresso(len(df))

        try:
            resultado = self.backend.solucionar_chamados(
                df=df,
                usar_html=self.usar_html_var.get(),
                log_cb=self.log_msg,
                progresso_cb=self._atualizar_progresso,
            )
        except Exception as e:
            self.log_msg(f"[ERRO] Nao foi possivel concluir inclusao de solucoes: {e}")
            messagebox.showerror("Erro", f"Nao foi possivel concluir inclusao de solucoes.\n{e}")
            self.status_importacao_var.set("Falha ao incluir solucoes.")
            self._set_botoes_operacao(habilitado=True)
            return

        self.status_importacao_var.set(resultado["resumo"])
        self.log_msg("[INFO] " + resultado["resumo"])
        messagebox.showinfo("Solucoes finalizadas", resultado["resumo"])
        self._set_botoes_operacao(habilitado=True)

    def buscar_nomes_api(self):
        if not self.backend.autenticado or self.backend.cliente is None:
            messagebox.showwarning("Autenticacao", "Autentique-se antes de consultar nomes na API.")
            return
        if self.backend.df is None:
            messagebox.showwarning("Planilha", "Selecione uma planilha antes de consultar nomes.")
            return

        self.status_importacao_var.set("Consultando nomes na API...")
        self.root.update_idletasks()
        try:
            resumo = self.backend.consultar_nomes_api(log_cb=self.log_msg)
            self.preencher_preview()
            self.status_api_var.set(resumo)
            self.log_msg("[OK] Nomes de tecnico/requerente/categoria/localizacao atualizados na pre-visualizacao.")
        except Exception as e:
            self.log_msg(f"[ERRO] Falha ao consultar nomes na API: {e}")
            messagebox.showerror("Erro", f"Nao foi possivel consultar nomes na API.\n{e}")
        finally:
            self.status_importacao_var.set("Pronto para importar.")
            self.root.update_idletasks()

    def _carregar_config_local(self):
        cfg = self.backend.carregar_config_local()
        if not cfg:
            return
        url = str(cfg.get("api_url", "")).strip()
        user_token = str(cfg.get("user_token", "")).strip()
        app_token = str(cfg.get("app_token", "")).strip()
        salvar_token = bool(cfg.get("salvar_tokens", cfg.get("salvar_user_token", True)))

        if url:
            self.url_var.set(url)
        if user_token:
            self.user_token_var.set(user_token)
            self.log_msg("[INFO] User token carregado do arquivo local.")
        if app_token:
            self.app_token_var.set(app_token)
            self.log_msg("[INFO] App token carregado do arquivo local.")
        self.salvar_tokens_var.set(salvar_token)

    def _salvar_config_local(self):
        try:
            self.backend.salvar_config_local(
                api_url=self.url_var.get(),
                user_token=self.user_token_var.get(),
                app_token=self.app_token_var.get(),
                salvar_tokens=self.salvar_tokens_var.get(),
            )
            if self.salvar_tokens_var.get():
                self.log_msg("[INFO] Tokens salvos localmente.")
            else:
                self.log_msg("[INFO] Salvamento de tokens desativado.")
        except Exception as e:
            self.log_msg(f"[AVISO] Nao foi possivel salvar configuracao local: {e}")


def main():
    root = tk.Tk()
    ImportadorGLPIApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()

