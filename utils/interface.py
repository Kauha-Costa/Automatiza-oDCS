import tkinter as tk
from tkinter import messagebox, filedialog
import ttkbootstrap as ttk
from ttkbootstrap.constants import *
import threading
import os
import sys
from datetime import datetime

from utils.editor import atualizar_certificados_em_pasta, listar_certificados
from utils.impressao import imprimir_certificados_em_pasta, listar_impressoras
from utils.config import carregar_config, salvar_config


def recurso(relativo: str) -> str:
    """Resolve caminho de recurso, compatível com PyInstaller."""
    if hasattr(sys, "_MEIPASS"):
        return os.path.join(sys._MEIPASS, relativo)
    return os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), relativo)


class App:
    def __init__(self):
        self.config = carregar_config()
        self.cancelar_flag = [False]
        self.processando = False

        # ── Janela principal ───────────────────────────────────────────
        self.root = ttk.Window(themename="cyborg")
        self.root.title("Automação de Certificados DCS")
        self.root.geometry("720x640")
        self.root.resizable(False, False)
        self.root.protocol("WM_DELETE_WINDOW", self._ao_fechar)

        try:
            self.root.iconbitmap(recurso("icone.ico"))
        except Exception:
            pass

        self._construir_interface()
        self._carregar_estado_salvo()
        self.root.mainloop()

    # ──────────────────────────────────────────────────────────────────
    #  Construção da interface
    # ──────────────────────────────────────────────────────────────────

    def _construir_interface(self):
        root = self.root

        # Padding externo
        outer = ttk.Frame(root, padding=(20, 16, 20, 10))
        outer.pack(fill="both", expand=True)

        # Título
        ttk.Label(
            outer,
            text="Automação de Certificados",
            font=("Segoe UI", 18, "bold"),
            bootstyle="inverse-dark"
        ).pack(pady=(0, 14))

        # ── Seção: Pasta ───────────────────────────────────────────────
        grp_pasta = ttk.LabelFrame(outer, text=" 📁  Pasta dos Certificados ", padding=10)
        grp_pasta.pack(fill="x", pady=(0, 10))

        pasta_row = ttk.Frame(grp_pasta)
        pasta_row.pack(fill="x")

        self.var_pasta = tk.StringVar()
        self.entry_pasta = ttk.Entry(pasta_row, textvariable=self.var_pasta, font=("Segoe UI", 10))
        self.entry_pasta.pack(side="left", fill="x", expand=True, padx=(0, 8))

        ttk.Button(
            pasta_row,
            text="Selecionar...",
            bootstyle="secondary-outline",
            command=self._selecionar_pasta
        ).pack(side="left")

        # Contador de arquivos
        self.lbl_contador = ttk.Label(grp_pasta, text="", font=("Segoe UI", 9), bootstyle="secondary")
        self.lbl_contador.pack(anchor="w", pady=(4, 0))

        self.var_pasta.trace_add("write", lambda *_: self._atualizar_contador())

        # ── Seção: Data ────────────────────────────────────────────────
        grp_data = ttk.LabelFrame(outer, text=" 📅  Nova Data ", padding=10)
        grp_data.pack(fill="x", pady=(0, 10))

        data_row = ttk.Frame(grp_data)
        data_row.pack(anchor="w")

        ttk.Label(data_row, text="Data (dd/mm/aaaa):", font=("Segoe UI", 10)).pack(side="left", padx=(0, 10))

        self.var_data = tk.StringVar()
        self.entry_data = ttk.Entry(data_row, textvariable=self.var_data, font=("Segoe UI", 13, "bold"), width=14)
        self.entry_data.pack(side="left")
        self.entry_data.bind("<KeyRelease>", self._formatar_data)

        self.lbl_data_status = ttk.Label(data_row, text="", font=("Segoe UI", 9))
        self.lbl_data_status.pack(side="left", padx=(10, 0))

        self.var_data.trace_add("write", lambda *_: self._validar_data())

        # ── Seção: Impressão ───────────────────────────────────────────
        grp_imp = ttk.LabelFrame(outer, text=" 🖨️  Impressão ", padding=10)
        grp_imp.pack(fill="x", pady=(0, 10))

        self.var_imprimir = tk.BooleanVar(value=self.config.get("imprimir_apos", False))
        chk = ttk.Checkbutton(
            grp_imp,
            text="Imprimir certificados após atualizar",
            variable=self.var_imprimir,
            bootstyle="round-toggle",
            command=self._toggle_impressora
        )
        chk.pack(anchor="w")

        imp_row = ttk.Frame(grp_imp)
        imp_row.pack(fill="x", pady=(8, 0))

        ttk.Label(imp_row, text="Impressora:", font=("Segoe UI", 10)).pack(side="left", padx=(0, 10))

        self.var_impressora = tk.StringVar()
        self.cmb_impressora = ttk.Combobox(
            imp_row,
            textvariable=self.var_impressora,
            state="readonly",
            font=("Segoe UI", 10),
            width=34
        )
        self.cmb_impressora.pack(side="left")

        ttk.Button(
            imp_row,
            text="↻",
            width=3,
            bootstyle="secondary-outline",
            command=self._recarregar_impressoras
        ).pack(side="left", padx=(6, 0))

        self._recarregar_impressoras()
        self._toggle_impressora()

        # ── Botões de ação ─────────────────────────────────────────────
        btn_row = ttk.Frame(outer)
        btn_row.pack(fill="x", pady=(4, 8))

        self.btn_executar = ttk.Button(
            btn_row,
            text="▶  Atualizar Certificados",
            bootstyle=SUCCESS,
            width=26,
            command=self._iniciar_processamento
        )
        self.btn_executar.pack(side="left")

        self.btn_cancelar = ttk.Button(
            btn_row,
            text="⛔  Cancelar",
            bootstyle="danger-outline",
            width=14,
            command=self._cancelar,
            state="disabled"
        )
        self.btn_cancelar.pack(side="left", padx=(10, 0))

        self.btn_limpar = ttk.Button(
            btn_row,
            text="🗑  Limpar Log",
            bootstyle="secondary-outline",
            width=14,
            command=self._limpar_log
        )
        self.btn_limpar.pack(side="right")

        # ── Barra de progresso ─────────────────────────────────────────
        self.progressbar = ttk.Progressbar(outer, mode="indeterminate", bootstyle="success-striped")
        self.progressbar.pack(fill="x", pady=(0, 6))

        # ── Log ────────────────────────────────────────────────────────
        log_frame = ttk.Frame(outer)
        log_frame.pack(fill="both", expand=True)

        self.text_log = tk.Text(
            log_frame,
            height=12,
            wrap="word",
            state="disabled",
            bg="#0a0a0a",
            fg="#e0e0e0",
            font=("Consolas", 9),
            insertbackground="white",
            relief="flat",
            bd=0
        )
        scroll = ttk.Scrollbar(log_frame, orient="vertical", command=self.text_log.yview)
        self.text_log.configure(yscrollcommand=scroll.set)
        scroll.pack(side="right", fill="y")
        self.text_log.pack(fill="both", expand=True)

        # Tags de cor no log
        self.text_log.tag_configure("ok",    foreground="#4caf50")
        self.text_log.tag_configure("erro",  foreground="#f44336")
        self.text_log.tag_configure("aviso", foreground="#ff9800")
        self.text_log.tag_configure("info",  foreground="#90caf9")
        self.text_log.tag_configure("dim",   foreground="#888888")

        # Rodapé
        ttk.Label(
            outer,
            text="AutoDCS — Automação de Certificados",
            font=("Segoe UI", 8),
            bootstyle="secondary"
        ).pack(side="right", pady=(4, 0))

    # ──────────────────────────────────────────────────────────────────
    #  Helpers de interface
    # ──────────────────────────────────────────────────────────────────

    def _carregar_estado_salvo(self):
        pasta = self.config.get("pasta_certificados", "")
        if pasta:
            self.var_pasta.set(pasta)

        impressora = self.config.get("impressora", "")
        if impressora and impressora in self.cmb_impressora["values"]:
            self.var_impressora.set(impressora)

    def _salvar_estado(self):
        self.config["pasta_certificados"] = self.var_pasta.get()
        self.config["imprimir_apos"] = self.var_imprimir.get()
        self.config["impressora"] = self.var_impressora.get()
        salvar_config(self.config)

    def _selecionar_pasta(self):
        pasta = filedialog.askdirectory(title="Selecione a pasta dos certificados")
        if pasta:
            self.var_pasta.set(pasta)

    def _atualizar_contador(self):
        pasta = self.var_pasta.get()
        arquivos = listar_certificados(pasta)
        if not pasta:
            self.lbl_contador.config(text="")
        elif not arquivos:
            self.lbl_contador.config(text="⚠️  Nenhum .docx encontrado nesta pasta", bootstyle="warning")
        else:
            self.lbl_contador.config(
                text=f"✅  {len(arquivos)} arquivo(s) encontrado(s): {', '.join(arquivos[:3])}{'...' if len(arquivos) > 3 else ''}",
                bootstyle="success"
            )

    def _formatar_data(self, event=None):
        """Formata automaticamente dd/mm/aaaa enquanto digita."""
        if event and event.keysym in (
            "BackSpace", "Delete", "Left", "Right", "Home", "End",
            "Shift_L", "Shift_R", "Control_L", "Control_R", "Tab"
        ):
            return

        texto = self.var_data.get()
        digitos = "".join(c for c in texto if c.isdigit())[:8]

        novo = ""
        for i, c in enumerate(digitos):
            if i == 2 or i == 4:
                novo += "/"
            novo += c

        if novo != texto:
            self.var_data.set(novo)
            self.entry_data.after(0, lambda: self.entry_data.icursor(tk.END))

    def _validar_data(self) -> bool:
        data = self.var_data.get()
        if not data:
            self.lbl_data_status.config(text="")
            return False
        try:
            datetime.strptime(data, "%d/%m/%Y")
            self.lbl_data_status.config(text="✓", bootstyle="success", foreground="#4caf50")
            return True
        except ValueError:
            self.lbl_data_status.config(text="Data inválida", bootstyle="danger", foreground="#f44336")
            return False

    def _toggle_impressora(self):
        estado = "readonly" if self.var_imprimir.get() else "disabled"
        self.cmb_impressora.config(state=estado)

    def _recarregar_impressoras(self):
        impressoras = listar_impressoras()
        self.cmb_impressora["values"] = impressoras
        if impressoras and not self.var_impressora.get():
            self.var_impressora.set(impressoras[0])

    def _limpar_log(self):
        self.text_log.config(state="normal")
        self.text_log.delete("1.0", tk.END)
        self.text_log.config(state="disabled")

    # ──────────────────────────────────────────────────────────────────
    #  Log colorido
    # ──────────────────────────────────────────────────────────────────

    def _log(self, msg: str):
        self.text_log.config(state="normal")

        # Determina tag pela primeira palavra/emoji
        tag = "info"
        if msg.startswith("✅") or msg.startswith("🖨"):
            tag = "ok"
        elif msg.startswith("❌"):
            tag = "erro"
        elif msg.startswith("⚠") or msg.startswith("⛔"):
            tag = "aviso"
        elif msg.startswith("ℹ") or msg.startswith("⏳"):
            tag = "dim"

        self.text_log.insert(tk.END, msg + "\n", tag)
        self.text_log.see(tk.END)
        self.text_log.config(state="disabled")
        self.root.update_idletasks()

    # ──────────────────────────────────────────────────────────────────
    #  Lógica de execução
    # ──────────────────────────────────────────────────────────────────

    def _iniciar_processamento(self):
        if self.processando:
            return

        pasta = self.var_pasta.get().strip()
        data = self.var_data.get().strip()

        # Validações
        if not pasta:
            messagebox.showerror("Erro", "Selecione a pasta dos certificados.")
            return
        if not os.path.exists(pasta):
            messagebox.showerror("Erro", f"Pasta não encontrada:\n{pasta}")
            return
        if not self._validar_data():
            messagebox.showerror("Erro", "Digite uma data válida no formato dd/mm/aaaa.")
            self.entry_data.focus()
            return

        arquivos = listar_certificados(pasta)
        if not arquivos:
            messagebox.showwarning("Aviso", "Nenhum arquivo .docx encontrado na pasta selecionada.")
            return

        # Confirma
        msg = f"Atualizar {len(arquivos)} certificado(s) com a data {data}?"
        if self.var_imprimir.get():
            msg += f"\n\nDepois irá imprimir na impressora:\n{self.var_impressora.get() or 'Padrão do sistema'}"
        if not messagebox.askyesno("Confirmar", msg):
            return

        self._salvar_estado()
        self._set_processando(True)
        self.cancelar_flag = [False]

        thread = threading.Thread(target=self._executar_em_thread, args=(pasta, data), daemon=True)
        thread.start()

    def _executar_em_thread(self, pasta: str, data: str):
        try:
            self._log(f"\n{'─'*50}")
            self._log(f"⏳ Iniciando atualização → {data}")
            self._log(f"📁 Pasta: {pasta}")
            self._log(f"{'─'*50}")

            atualizados, sem_data = atualizar_certificados_em_pasta(
                pasta, data,
                log_callback=self._log,
                cancelar_flag=self.cancelar_flag
            )

            if self.cancelar_flag[0]:
                self._log("\n⛔ Processo interrompido.")
                return

            self._log(f"\n✅ Atualização concluída — {atualizados} atualizado(s), {sem_data} sem data.")

            if self.var_imprimir.get() and not self.cancelar_flag[0]:
                impressora = self.var_impressora.get()
                self._log(f"\n{'─'*50}")
                self._log(f"🖨️  Iniciando impressão...")
                if impressora:
                    self._log(f"🖨️  Impressora: {impressora}")
                self._log(f"{'─'*50}")

                impressos = imprimir_certificados_em_pasta(
                    pasta,
                    impressora=impressora,
                    log_callback=self._log,
                    cancelar_flag=self.cancelar_flag
                )
                self._log(f"\n✅ Impressão concluída — {impressos} arquivo(s) enviado(s).")

            self.root.after(0, lambda: messagebox.showinfo("Concluído", "Processo finalizado com sucesso!"))

        except Exception as e:
            self._log(f"\n❌ Erro inesperado: {e}")
            self.root.after(0, lambda: messagebox.showerror("Erro", f"Ocorreu um erro:\n{e}"))

        finally:
            self.root.after(0, lambda: self._set_processando(False))

    def _cancelar(self):
        if self.processando:
            self.cancelar_flag[0] = True
            self._log("⛔ Cancelamento solicitado...")

    def _set_processando(self, valor: bool):
        self.processando = valor
        if valor:
            self.btn_executar.config(state="disabled")
            self.btn_cancelar.config(state="normal")
            self.progressbar.start(10)
        else:
            self.btn_executar.config(state="normal")
            self.btn_cancelar.config(state="disabled")
            self.progressbar.stop()

    def _ao_fechar(self):
        if self.processando:
            if not messagebox.askyesno("Sair", "Há um processo em andamento. Deseja cancelar e sair?"):
                return
            self.cancelar_flag[0] = True
        self._salvar_estado()
        self.root.destroy()


def iniciar_interface():
    App()
