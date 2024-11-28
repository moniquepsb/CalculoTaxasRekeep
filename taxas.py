import pandas as pd
from tkinter import Tk, filedialog, Label, Button, messagebox, ttk
from openpyxl import load_workbook
from openpyxl.styles import Font, Alignment, PatternFill
import os


class PlanilhaVendasETaxasApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Processador de Vendas e Taxas v1.0")
        self.root.geometry("500x400")

        # Seção de Upload da Planilha de Vendas
        Label(root, text="Selecione a planilha de vendas:").pack(pady=20)
        self.label_vendas = Label(root, text="Nenhum arquivo selecionado", fg="gray")
        self.label_vendas.pack(pady=10)
        Button(root, text="Upload Planilha de Vendas", command=self.selecionar_planilha_vendas).pack(pady=5)

        # Seção de Upload da Planilha de Taxas
        Label(root, text="Selecione a planilha de taxas:").pack(pady=10)
        self.label_taxas = Label(root, text="Nenhum arquivo selecionado", fg="gray")
        self.label_taxas.pack(pady=5)
        Button(root, text="Upload Planilha de Taxas", command=self.selecionar_planilha_taxas).pack(pady=5)

        # Barra de Progresso
        self.label_status = Label(root, text="", fg="green", wraplength=400, justify="center")
        self.label_status.pack(pady=10)

        self.progress_bar = ttk.Progressbar(root, orient="horizontal", length=400, mode="determinate")
        self.progress_bar.pack(pady=10)

        # Botão para Processar as Planilhas
        Button(root, text="Iniciar Processamento", command=self.processar_planilha).pack(pady=10)

        # Inicializando variáveis dos arquivos
        self.arquivo_vendas = None
        self.arquivo_taxas = None

    def selecionar_planilha_vendas(self):
        self.arquivo_vendas = filedialog.askopenfilename(
            title="Selecione a planilha de vendas", filetypes=[("Excel files", "*.xlsx")]
        )
        if self.arquivo_vendas:
            self.label_vendas.config(text=self.arquivo_vendas.split("/")[-1])

    def selecionar_planilha_taxas(self):
        self.arquivo_taxas = filedialog.askopenfilename(
            title="Selecione a planilha de taxas", filetypes=[("Excel files", "*.xlsx")]
        )
        if self.arquivo_taxas:
            self.label_taxas.config(text=self.arquivo_taxas.split("/")[-1])

    def processar_planilha(self):
        if not self.arquivo_vendas or not self.arquivo_taxas:
            messagebox.showerror("Erro", "Selecione ambas as planilhas antes de iniciar o processamento.")
            return

        try:
            # Exibindo mensagem de processamento
            self.label_status.config(text="Processando...")
            self.progress_bar["value"] = 0
            self.progress_bar["maximum"] = 100
            self.root.update_idletasks()

            # Carregar as planilhas
            df_vendas = pd.read_excel(self.arquivo_vendas, nrows=100000)
            df_taxas = pd.read_excel(self.arquivo_taxas)

            # Normalizar datas
            df_vendas["Data"] = pd.to_datetime(df_vendas["Data"], errors="coerce")
            df_taxas["Data Inicial"] = pd.to_datetime(df_taxas["Data Inicial"], errors="coerce")
            df_taxas["Data Final"] = pd.to_datetime(df_taxas["Data Final"], errors="coerce")

            # Realizar o merge
            df_merged = self.merge_planilhas(
                df_vendas, df_taxas,
                col_cartoes_vendas="Cartões", col_cartoes_taxas="Cartão",
                col_data_vendas="Data", col_data_inicial="Data Inicial", col_data_final="Data Final",
                col_parcelas_vendas="Parcelas", col_parcelas_taxas="Parcelas",
                col_taxa="Taxa"
            )

            # Atualizar a barra de progresso
            self.progress_bar["value"] = 50
            self.root.update_idletasks()

            # Calcular o Valor Bruto-Líquido
            df_merged["Valor Bruto-Líquido"] = df_merged["Valor Bruto"] - df_merged["Valor Líquido"]

            # Gerar nome incremental para o arquivo de saída
            nome_arquivo = self.salvar_com_nome_incremental("planilha_calculo")

            # Salvando o arquivo processado com formatação
            self.salvar_com_formatacao(df_merged, nome_arquivo)

            # Finalizando o processamento
            self.progress_bar["value"] = 100
            self.label_status.config(
                text=f"Sucesso! Salvo como:\n'{os.path.abspath(nome_arquivo)}'",
                fg="green"
            )
            # Exibir mensagem de sucesso e chamar o método para limpar
            messagebox.showinfo("Sucesso", f"Arquivo processado com sucesso! Salvo como '{nome_arquivo}'.")
            self.zerar_estado()
        except ValueError as e:
            messagebox.showerror("Erro", f"Erro ao processar as planilhas: {e}")
        except PermissionError:
            messagebox.showerror("Erro", "O arquivo de saída está aberto. Feche o arquivo e tente novamente.")
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao processar as planilhas: {e}")

    def merge_planilhas(self, df_vendas, df_taxas, col_cartoes_vendas, col_cartoes_taxas, col_data_vendas,
                        col_data_inicial, col_data_final, col_parcelas_vendas, col_parcelas_taxas, col_taxa):
        """Realiza o merge entre as planilhas de vendas e taxas."""
        df = df_vendas.merge(
            df_taxas,
            left_on=col_cartoes_vendas,
            right_on=col_cartoes_taxas,
            how="inner"
        )
        df = df[
            (df[col_data_vendas] >= df[col_data_inicial]) &
            (df[col_data_vendas] <= df[col_data_final]) &
            (df[col_parcelas_vendas] == df[col_parcelas_taxas])
        ]
        return df[[*df_vendas.columns, col_taxa]]  # Somente as colunas da planilha de vendas + coluna "Taxa"

    def salvar_com_nome_incremental(self, nome_base):
        """Gera um nome incremental para o arquivo de saída."""
        contador = 1
        nome_arquivo = f"{nome_base}({contador}).xlsx"
        while os.path.exists(nome_arquivo):
            contador += 1
            nome_arquivo = f"{nome_base}({contador}).xlsx"
        return nome_arquivo

    def salvar_com_formatacao(self, dataframe, output_file):
        dataframe.to_excel(output_file, index=False, engine="openpyxl")

        # Carregar o arquivo salvo para adicionar formatação
        wb = load_workbook(output_file)
        ws = wb.active

        # Formatar a altura da linha 1
        ws.row_dimensions[1].height = 48

        # Formatar a coluna D como dd/mm/yyyy
        for cell in ws["D"]:
            if cell.row > 1:  # Ignorar o cabeçalho
                cell.number_format = "dd/mm/yyyy"

        # Ajustar a largura das colunas automaticamente
        for column in ws.columns:
            max_length = 0
            column_letter = column[0].column_letter  # Letra da coluna
            for cell in column:
                try:
                    if cell.value:
                        max_length = max(max_length, len(str(cell.value)))
                except Exception:
                    pass
            adjusted_width = max_length + 2
            ws.column_dimensions[column_letter].width = adjusted_width

        # Aplicar formatação na linha 1 (cabeçalho)
        for cell in ws[1]:
            cell.font = Font(bold=True)
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.fill = PatternFill(start_color="7FD4DE", end_color="7FD4DE", fill_type="solid")

        # Salvar novamente com a formatação aplicada
        wb.save(output_file)

    def zerar_estado(self):
        """Reseta o programa para estado inicial."""
        self.arquivo_vendas = None
        self.arquivo_taxas = None
        self.label_vendas.config(text="Nenhum arquivo selecionado")
        self.label_taxas.config(text="Nenhum arquivo selecionado")
        self.label_status.config(text="")
        self.progress_bar["value"] = 0


root = Tk()
app = PlanilhaVendasETaxasApp(root)
root.mainloop()
