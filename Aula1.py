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

# ticket médio produto vendido por loja
ticket_medio = (faturamento["Valor Final"] / quantidade["Quantidade"]).to_frame()
ticket_medio = ticket_medio.rename(columns={0: 'Ticket Médio'})
print(ticket_medio)

# enviar email com relatório
mail.outlook = win32.Dispatch('outlook.application')
mail = outlook.CreateItem(0)
mail.To = 'Seu email'
mail.Subject = 'Relatório de Vendas'
mail.HTMLBody = f'''
<p>Prezados,</p>

<p>segue o Relatório de Vendas por Loja.</p>

<p>Faturamento:</p>
{faturamento.to_html(formatters={'Valor Final': 'R${:,.2f}'.format})}

<p>Quantidade de Vendas:</p>
{quantidade.to_html}

<p>Ticket Médio:</p>
{ticket_medio.to_html(formatters={'Ticket Médio': 'R${:,.2f}'.format})}

<p>Qualquer dúvida estou à disposição</p>

<p>Att.,</p>
<p>G. R.</p>
'''

mail.Send()

print('E-mail Enviado')