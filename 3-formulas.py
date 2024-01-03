from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

# Pega a nova planilha e transforma em um workbook
wb = load_workbook("data/Grafico.xlsx")
sheet = wb["Relatorio"]

# Configura coluna e linha maxima e minima
min_column = wb.active.min_column
max_column = wb.active.max_column
min_row = wb.active.min_row
max_row = wb.active.max_row

# Cria medida de soma das colunas
for i in range(min_column + 1, max_column + 1):
    letter = get_column_letter(i)
    sheet[f"{letter}{max_row+1}"] = f"=SUM({letter}{min_row+1}:{letter}{max_row})"
    sheet[f"{letter}{max_row+1}"].style = "Currency"

# Salva em uma nova tabela
wb.save("data/Final.xlsx")


