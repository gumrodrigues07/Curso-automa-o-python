import pandas as pd
import win32com.client as win32

# importar base de dados
tabela_vendas = pd.read_excel('vendas.xlsx')

# vizualisar base de dados
pd.set_option('display.max_columns', None)

# faturamento por loja
faturamento = tabela_vendas[['ID Loja','Valor Final']].groupby('ID Loja').sum()
print(faturamento)

# quantidade de produtos vendidos por loja
quantidade = tabela_vendas[['ID Loja', 'Quantidade']].groupby('ID Loja').sum()
print(quantidade)
print('-' * 50)

# ticket médio produto vendido por loja
ticket_medio = (faturamento["Valor Final"] / quantidade["Quantidade"]).to_frame()
print(ticket_medio)

# enviar email com relatório
mail.outlook = win32.Dispatch('outlook.application') # Créditos: Stack Flow
mail = outlook.CreateItem(0)
mail.To = 'nofanafarofa@gmail.com'
mail.Subject = 'Relatório de Vendas'
mail.HTMLBody = '''
Prezados,

segue o Relatório de Vendas por Loja.

Faturamento:
{}

Quantidade de Vendas:
{}

Ticket Médio:
{}

Qualquer dúvida estou à disposição

Att.,
G. R.
'''

mail.Send()

print('E-mail Enviado')