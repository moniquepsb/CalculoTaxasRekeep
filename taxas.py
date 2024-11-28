import pandas as pd
from tkinter import Tk, filedialog, Label, Button, messagebox, ttk
from openpyxl import load_workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
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
            self.label_vendas.config(text=os.path.basename(self.arquivo_vendas))

    def selecionar_planilha_taxas(self):
        self.arquivo_taxas = filedialog.askopenfilename(
            title="Selecione a planilha de taxas", filetypes=[("Excel files", "*.xlsx")]
        )
        if self.arquivo_taxas:
            self.label_taxas.config(text=os.path.basename(self.arquivo_taxas))

    def processar_planilha(self):
        if not self.arquivo_vendas or not self.arquivo_taxas:
            messagebox.showerror("Erro", "Selecione ambas as planilhas antes de iniciar o processamento.")
            return

        try:
            self.label_status.config(text="Processando...")
            self.progress_bar["value"] = 0
            self.root.update_idletasks()

            # Carregar as planilhas
            df_vendas = pd.read_excel(self.arquivo_vendas)
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

            self.progress_bar["value"] = 50
            self.root.update_idletasks()

            # Gerar nome incremental para o arquivo de saída
            nome_arquivo = self.salvar_com_nome_incremental("planilha_calculo")

            # Salvando o arquivo processado com formatação
            self.salvar_com_formatacao(df_merged, nome_arquivo)

            self.progress_bar["value"] = 100
            self.label_status.config(
                text=f"Sucesso! Salvo como:\n'{os.path.abspath(nome_arquivo)}'",
                fg="green"
            )
            messagebox.showinfo("Sucesso", f"Arquivo processado com sucesso! Salvo como '{nome_arquivo}'.")
            self.zerar_estado()
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao processar as planilhas: {e}")
            print(f"Erro ao processar as planilhas: {e}")

    def merge_planilhas(self, df_vendas, df_taxas, col_cartoes_vendas, col_cartoes_taxas, col_data_vendas,
                        col_data_inicial, col_data_final, col_parcelas_vendas, col_parcelas_taxas, col_taxa):
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
        return df[[*df_vendas.columns, col_taxa]]

    def salvar_com_nome_incremental(self, nome_base):
        contador = 1
        nome_arquivo = f"{nome_base}({contador}).xlsx"
        while os.path.exists(nome_arquivo):
            contador += 1
            nome_arquivo = f"{nome_base}({contador}).xlsx"
        return nome_arquivo

    def salvar_com_formatacao(self, dataframe, output_file):
        dataframe.to_excel(output_file, index=False, engine="openpyxl")
        wb = load_workbook(output_file)
        ws = wb.active

        # Formatar a altura da linha 1 (cabeçalho)
        ws.row_dimensions[1].height = 48

        # Adicionar fórmula "=M-O" na coluna correspondente
        col_valor_bruto = "M"  # Supondo que "Valor Bruto" está na coluna M
        col_valor_liquido = "O"  # Supondo que "Valor Líquido" está na coluna O
        col_valor_resultado = "P"  # Próxima coluna para o resultado (ajuste conforme necessário)

        ws[f"{col_valor_resultado}1"] = "Retenção"  # Cabeçalho da nova coluna

        for row in range(2, ws.max_row + 1):  # Começa na linha 2 para evitar o cabeçalho
            ws[f"{col_valor_resultado}{row}"].value = f"={col_valor_bruto}{row}-{col_valor_liquido}{row}"

        # Nova Coluna Comissão Aplicada

        col_valor_bruto = "M"  # Supondo que "Valor Bruto" está na coluna M
        col_retenção = "P"  # Supondo que "Retenção" está na coluna P
        col_comissão_aplicada = "R"  # Nova coluna para comissão aplicada

        # Adicionando o cabeçalho da nova coluna
        ws[f"{col_comissão_aplicada}1"] = "Comissão Aplicada"  # Cabeçalho da nova coluna

        # Preenchendo a fórmula em cada linha da coluna col_comissão_aplicada
        for row in range(2, ws.max_row + 1):  # Começa na linha 2 para evitar o cabeçalho
            ws[f"{col_comissão_aplicada}{row}"].value = f"={col_retenção}{row}*100/{col_valor_bruto}{row}"

        # Nova Coluna Valor Líquido Contratado
        col_taxa = "Q"  # Supondo que "Taxa" está na coluna Q
        col_valor_bruto = "M"  # Supondo que "Valor Bruto" está na coluna M
        col_valor_liquido_contratado = "S"

        # Adicionando o cabeçalho da nova coluna
        ws[f"{col_valor_liquido_contratado}1"] = "Valor Líquido Contatado"  # Cabeçalho da nova coluna

        # Preenchendo a fórmula em cada linha da coluna col_valor_liquido_contratado
        for row in range(2, ws.max_row + 1):  # Começa na linha 2 para evitar o cabeçalho
            ws[f"{col_valor_liquido_contratado}{row}"].value = f"=100-{col_taxa}{row}/100*{col_valor_bruto}{row}"

        # Nova Coluna Diferença
        col_valor_liquido_contratado = "S"  # Valor do Valor Liquido Contratado
        col_valor_liquido = "O"  # Supondo que "Valor Bruto" está na coluna M
        col_diferença = "T"

        # Adicionando o cabeçalho da nova coluna
        ws[f"{col_diferença}1"] = "Diferença"  # Cabeçalho da nova coluna

        # Preenchendo a fórmula da coluna "Diferença"
        for row in range(2, ws.max_row + 1):  # Começa na linha 2 para evitar o cabeçalho
            ws[f"{col_diferença}{row}"].value = f"={col_valor_liquido_contratado}{row}-{col_valor_liquido}{row}"

        # Nova Coluna Indébito
        col_indébito = "U"

        # Adicionando o cabeçalho da nova coluna
        ws[f"{col_indébito}1"] = "Indébito"  # Cabeçalho da nova coluna

        # Preenchendo a fórmula condicional na coluna "Indébito"
        for row in range(2, ws.max_row + 1):  # Começa na linha 2 para evitar o cabeçalho
            ws[f"{col_indébito}{row}"].value = f"=IF({col_diferença}{row}<=0,0,{col_diferença}{row})"


        # Ajustar a largura das colunas automaticamente
        for column in ws.columns:
            max_length = 0
            column_letter = column[0].column_letter
            for cell in column:
                if cell.value:
                    max_length = max(max_length, len(str(cell.value)))
            ws.column_dimensions[column_letter].width = max_length + 2

        # Aplicar formatação na linha 1 (cabeçalho)
        for cell in ws[1]:
            cell.font = Font(bold=True)
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.fill = PatternFill(start_color="7FD4DE", end_color="7FD4DE", fill_type="solid")

        # Criando o estilo de borda preta
        thin_border = Border(
            left=Side(style="thin", color="000000"),
            right=Side(style="thin", color="000000"),
            top=Side(style="thin", color="000000"),
            bottom=Side(style="thin", color="000000")
        )

        # Aplicando bordas aos cabeçalhos (linha 1)
        for cell in ws[1]:
            cell.border = Border(
                left=Side(style="thin", color="000000"),  # Borda esquerda
                right=Side(style="thin", color="000000"),  # Borda direita
                top=Side(style="thin", color="000000"),  # Borda superior
                bottom=Side(style="thin", color="000000")  # Borda inferior
            )

        wb.save(output_file)

    def zerar_estado(self):
        self.arquivo_vendas = None
        self.arquivo_taxas = None
        self.label_vendas.config(text="Nenhum arquivo selecionado")
        self.label_taxas.config(text="Nenhum arquivo selecionado")
        self.label_status.config(text="")
        self.progress_bar["value"] = 0


root = Tk()
app = PlanilhaVendasETaxasApp(root)
root.mainloop()