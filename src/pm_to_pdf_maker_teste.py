import os
import pandas as pd
from fpdf import FPDF

# Caminho do arquivo Excel enviado
excel_path = "/mnt/data/PROJECT MANAGEMENT.xlsx"

# Lê apenas a aba 'ENTREGAS', considerando que o cabeçalho está na linha 3 (index 2)
df_entregas = pd.read_excel(excel_path, sheet_name='ENTREGAS', header=2)

# Cria a pasta principal onde os PDFs serão salvos
main_folder = "/mnt/data/pm_to_pdf_maker"
os.makedirs(main_folder, exist_ok=True)

# Define a classe para criação de PDF
class PDF(FPDF):
    def header(self):
        self.set_font("Arial", "B", 12)
        self.cell(0, 10, "Relatório de Entregas", 0, 1, "C")

    def add_row_data(self, row):
        self.set_font("Arial", "", 10)
        for col, value in row.items():
            # Converte para string e remove caracteres não suportados pelo latin-1
            value_str = str(value)
            safe_value = value_str.encode('latin-1', 'ignore').decode('latin-1')
            self.cell(0, 8, f"{col}: {safe_value}", ln=True)

# Itera sobre as linhas do DataFrame e cria um PDF para cada uma
pdf_files = []
for idx, row in df_entregas.iterrows():
    pdf = PDF()
    pdf.add_page()
    pdf.add_row_data(row)
    
    filename = f"entrega_{idx + 1}.pdf"
    filepath = os.path.join(main_folder, filename)
    pdf.output(filepath)
    pdf_files.append(filepath)

# Imprime a lista de arquivos PDF criados
print("PDFs criados:")
for file in pdf_files:
    print(file)
