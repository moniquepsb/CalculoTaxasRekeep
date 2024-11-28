import pandas as pd
from tkinter import Tk, filedialog, Label, Button, Entry, Listbox, END, messagebox, Toplevel, Frame
from tkinter import ttk
import threading
import os
import gc
import warnings
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font, Alignment

warnings.filterwarnings("ignore", category=pd.errors.PerformanceWarning)


class PlanilhaUnificadoraApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Unificador de Planilhas_v2.1")
        self.root.geometry("500x600")

        # Toggle para escolher entre um ou mais arquivos
        Label(root, text="Modo de operação:").pack(pady=5)
        self.toggle_mode = ttk.Combobox(root, values=["Apenas um arquivo", "Mais de um arquivo"], state="readonly")
        self.toggle_mode.pack(pady=5)
        self.toggle_mode.set("Mais de um arquivo")  # Valor padrão

        Label(root, text="Número de linhas do cabeçalho para remover:").pack(pady=5)
        self.entry_rows = Entry(root)
        self.entry_rows.pack(pady=5)
        self.entry_rows.insert(0, "7")

        Label(root, text="Selecione o arquivo principal (Planilha Mãe):").pack(pady=5)
        self.label_arquivo_principal = Label(root, text="Nenhum arquivo selecionado", fg="gray")
        self.label_arquivo_principal.pack(pady=5)
        Button(root, text="Upload Planilha Mãe", command=self.selecionar_planilha_mae).pack(pady=5)

        Label(root, text="Selecione as planilhas adicionais:").pack(pady=5)
        self.listbox_arquivos_adicionais = Listbox(root, selectmode="multiple", width=50, height=10)
        self.listbox_arquivos_adicionais.pack(pady=5)
        Button(root, text="Upload Planilhas Adicionais", command=self.selecionar_planilhas_adicionais).pack(pady=5)

        button_frame = Frame(root)
        button_frame.pack(pady=10)

        Button(button_frame, text="Iniciar Merge", command=self.iniciar_merge_thread).pack(side="left", padx=5)
        Button(button_frame, text="Limpar", command=self.limpar_inputs).pack(side="left", padx=5)

        self.label_status = Label(root, text="")
        self.label_status.pack(pady=5)

        self.arquivo_principal = None
        self.arquivos_adicionais = []

    def selecionar_planilha_mae(self):
        self.arquivo_principal = filedialog.askopenfilename(title="Selecione o arquivo principal", filetypes=[("Excel files", "*.xlsx")])
        if self.arquivo_principal:
            self.label_arquivo_principal.config(text=self.arquivo_principal.split("/")[-1])

    def selecionar_planilhas_adicionais(self):
        arquivos = filedialog.askopenfilenames(title="Selecione as planilhas adicionais", filetypes=[("Excel files", "*.xlsx")])
        if arquivos:
            self.arquivos_adicionais.extend(arquivos)
            self.listbox_arquivos_adicionais.delete(0, END)
            for arquivo in self.arquivos_adicionais:
                self.listbox_arquivos_adicionais.insert(END, arquivo.split("/")[-1])

    def iniciar_merge_thread(self):
        self.exibir_alerta_progresso()
        merge_thread = threading.Thread(target=self.unificar_planilhas)
        merge_thread.start()

    def exibir_alerta_progresso(self):
        self.alerta_progresso = Toplevel(self.root)
        self.alerta_progresso.title("Processando")
        self.alerta_progresso.geometry("300x150")
        self.alerta_progresso.transient(self.root)
        self.alerta_progresso.grab_set()

        Label(self.alerta_progresso, text="Aguarde, estamos juntando os arquivos...").pack(pady=10)
        
        self.progress_bar = ttk.Progressbar(self.alerta_progresso, orient="horizontal", length=250, mode="determinate")
        self.progress_bar.pack(pady=10)

    def fechar_alerta_progresso(self):
        self.alerta_progresso.destroy()

    def salvar_com_nome_incremental(self, nome_base):
        contador = 1
        nome_arquivo = f"{nome_base}.xlsx"
        while os.path.exists(nome_arquivo):
            nome_arquivo = f"{nome_base} ({contador}).xlsx"
            contador += 1
        return nome_arquivo

    def carregar_planilha(self, arquivo, rows_to_skip):
        try:
            df = pd.read_excel(arquivo, engine='openpyxl', skiprows=rows_to_skip)
            df = df.loc[:, ~df.columns.str.contains('^Unnamed:')]
            return df
        except Exception as e:
            messagebox.showerror(
                "Erro ao Processar Arquivo",
                f"Erro ao processar o arquivo. Confira se a Planilha Mãe está vazia."
            )
            return None

    def unificar_planilhas(self):
        modo = self.toggle_mode.get()

        if not self.arquivo_principal:
            messagebox.showerror("Erro", "Selecione uma planilha mãe.")
            self.fechar_alerta_progresso()
            return

        try:
            rows_to_skip = int(self.entry_rows.get())
        except ValueError:
            messagebox.showerror("Erro", "Insira um número válido de linhas para remover.")
            self.fechar_alerta_progresso()
            return

        if modo == "Apenas um arquivo":
            df_principal = self.carregar_planilha(self.arquivo_principal, rows_to_skip)
            if df_principal is not None:
                try:
                    nome_arquivo = self.salvar_com_nome_incremental("planilha_unificada")
                    df_principal.to_excel(nome_arquivo, index=False)
                    self.label_status.config(text=f"Arquivo processado com sucesso! Salvo como '{nome_arquivo}'.")
                except Exception as e:
                    messagebox.showerror("Erro", f"Erro ao salvar o arquivo: {e}")
            self.fechar_alerta_progresso()
            return

        if not self.arquivos_adicionais:
            messagebox.showerror("Erro", "Selecione pelo menos uma planilha adicional.")
            self.fechar_alerta_progresso()
            return

        df_principal = self.carregar_planilha(self.arquivo_principal, rows_to_skip)
        if df_principal is None:
            self.fechar_alerta_progresso()
            return
        common_columns = df_principal.columns.tolist()
        lista_dfs = [df_principal]

        total_arquivos = len(self.arquivos_adicionais) + 1
        self.progress_bar['maximum'] = total_arquivos
        self.progress_bar['value'] = 0
        self.root.update_idletasks()

        for i, arquivo in enumerate(self.arquivos_adicionais, start=1):
            df = self.carregar_planilha(arquivo, rows_to_skip)
            if df is not None:
                df_common = df.reindex(columns=common_columns)
                lista_dfs.append(df_common)
            self.progress_bar['value'] = i
            self.alerta_progresso.update_idletasks()

        try:
            df_unificado = pd.concat(lista_dfs, ignore_index=True)
            nome_arquivo = self.salvar_com_nome_incremental("planilha_unificada")
            df_unificado.to_excel(nome_arquivo, index=False)
            self.label_status.config(text=f"Planilhas unificadas com sucesso! Arquivo salvo como '{nome_arquivo}'.")
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao salvar o arquivo unificado: {e}")

        del lista_dfs
        gc.collect()
        self.fechar_alerta_progresso()

    def limpar_inputs(self):
        self.arquivo_principal = None
        self.arquivos_adicionais = []
        self.entry_rows.delete(0, END)
        self.entry_rows.insert(0, "7")
        self.label_arquivo_principal.config(text="Nenhum arquivo selecionado")
        self.listbox_arquivos_adicionais.delete(0, END)
        self.label_status.config(text="")


root = Tk()
app = PlanilhaUnificadoraApp(root)
root.mainloop()
