import pandas as pd

#Importa o arquivo Excel
data = pd.read_excel('data/VendaCarros.xlsx')

# Pega as colunas principais para o projeto
df = data[["Fabricante", "ValorVenda", "Ano"]]

# Faz a tabela de Valor da Venda do Fabricante por Ano
pivot_table = df.pivot_table(
    index="Ano",
    columns="Fabricante",
    values="ValorVenda",
    aggfunc="sum"
)

#Cria nova tabela no excel
pivot_table.to_excel("data/ValorVendaFabricantePorAno.xlsx", "Relatorio")