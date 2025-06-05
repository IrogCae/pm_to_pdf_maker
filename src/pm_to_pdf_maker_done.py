import os
import pandas as pd
from fpdf import FPDF

# Caminho do arquivo Excel enviado
excel_path = "C:\\Users\\irogc\\OneDrive\\Documentos\\Python Scripts\\pm_to_pdf_maker\\data\\PROJECT MANAGEMENT.xlsx"

# Lê apenas a aba 'ENTREGAS', considerando que o cabeçalho está na linha 3 (index 2)
df_entregas = pd.read_excel(excel_path, sheet_name='ENTREGAS', header=2)

# Cria a pasta principal onde os PDFs serão salvos
main_folder = "C:\\Users\\irogc\\OneDrive\\Documentos\\Python Scripts\\pm_to_pdf_maker"     # Define o caminho da pasta principal
os.makedirs(main_folder, exist_ok=True)     # Cria a pasta se não existir

# Define a classe para criação de PDF
class PDF(FPDF):
    def header(self):                                           # Cria o cabeçalho do PDF
        self.set_font("Arial", "B", 12)                         # Define a fonte do cabeçalho
        self.cell(0, 10, "Relatório de Entregas", 0, 1, "C")    # Adiciona o título ao cabeçalho

    def add_row_data(self, row):
        self.set_font("Arial", "", 10)      # Define a fonte para os dados
        for col, value in row.items():      # Itera sobre as colunas e valores da linha
            if pd.isna(value):               # Verifica se o valor é NaN
                value = "N/A"                # Substitui NaN por "N/A"
            value_str = str(value)          # Converte para string
            safe_value = value_str.encode('latin-1', 'ignore').decode('latin-1')    # Remove caracteres não suportados pelo latin-1
            self.cell(0, 8, f"{col}: {safe_value}", ln=True)                        # Adiciona os dados da linha ao PDF

# Itera sobre as linhas do DataFrame e cria um PDF para cada uma
pdf_files = []      # Lista para armazenar os caminhos dos arquivos PDF criados
for idx, row in df_entregas.iterrows():     # Itera sobre cada linha do DataFrame
    pdf = PDF()     # Cria uma instância do PDF
    pdf.add_page()      # Adiciona uma nova página ao PDF
    pdf.add_row_data(row)       # Adiciona os dados da linha ao PDF
    filename = f"entrega_{idx + 1}.pdf"     # Define o nome do arquivo PDF
    filepath = os.path.join(main_folder, filename)      # Cria o caminho completo do arquivo PDF
    pdf.output(filepath)        # Salva o PDF no caminho especificado
    pdf_files.append(filepath)  # Adiciona o caminho do arquivo PDF à lista

# Imprime a lista de arquivos PDF criados
print("PDFs criados:")      # Mensagem de cabeçalho
for file in pdf_files:      # Itera sobre cada arquivo PDF criado
    print(file)             # Imprime o caminho do arquivo PDF
