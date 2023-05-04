import pandas as pd
import win32com.client as win32


# importar a base de dados
tabela_vendas = pd.read_excel('Vendas.xlsx')

# Visualizar a base de dados
pd.set_option('display.max_columns', None)
# print(tabela_vendas)

# Faturamento por loja
faturamento_loja = tabela_vendas[['ID Loja',
                                  'Valor Final']].groupby('ID Loja').sum()
print(faturamento_loja)
print('-' * 50)

# Quatidade de produtos por loja

quantidade_produto = tabela_vendas[[
    'ID Loja', 'Quantidade']].groupby('ID Loja').sum()


print(quantidade_produto)
print('-' * 50)

# Ticket médio por produto por loja

ticket_medio = (faturamento_loja['Valor Final'] /
                quantidade_produto['Quantidade']).to_frame()

ticket_medio = ticket_medio.rename(columns={0: 'Ticket Médio'})

print(ticket_medio)

# enviar um email com relatório
outlook = win32.Dispatch('outlook.application')
mail = outlook.CreateItem(0)
mail.To = 'pablovilasboas02@gmail.com'
mail.Subject = 'Relatório de vendas por loja.'
mail.HTMLBody = f'''
<p>Prezados,</p>

<p>Segue o Relatório de vendas por cada loja.</p>
<p>Faturamento:</p> 
{faturamento_loja.to_html(formatters={'Valor Final': 'R${:,.2f}'.format})}


<p>Quantidade Vendida:</p>
{quantidade_produto.to_html()}

<p>Ticked Médio por Prodtuo em cada loja:</p>
{ticket_medio.to_html(formatters={'Ticket Médio': 'R${:,.2f}'.format})}

<p>Qualquer dúvida estou a disposição.</p>
<p>Att.</p>
<p>Pablo</p>

'''

mail.Send()
print('Email Enviado...')
