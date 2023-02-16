import pandas as pd
import win32com.client as win32
import locale

# import the db
table_sales = pd.read_excel('Vendas.xlsx')

# vizualise the data base
pd.set_option('display.max_columns', None)

# turnover per store
turnover = table_sales[['ID Loja','Valor Final']].groupby('ID Loja').sum()
print(turnover)

# quantity of products sold per store
quantity = table_sales[['ID Loja','Quantidade']].groupby('ID Loja').sum()
print(quantity)

# avarage ticket per product on wich store
avarage_ticket = (turnover['Valor Final'] / quantity['Quantidade']).to_frame()
avarage_ticket = avarage_ticket.rename(columns={0: 'Ticket Médio'})
print(avarage_ticket)

# define the location of the brazil
locale.setlocale(locale.LC_ALL, 'pt_BR.utf-8')

# creating a function the format the collumn
def format_collum(value):
    return locale.currency(value, grouping=True)

# send email with report
outlook = win32.Dispatch('outlook.application')
mail = outlook.CreateItem(0)
mail.To = 'email@hotmail.com'
mail.Subject = 'Relatório de Vendas'
mail.HTMLBody = f'''
<p>Prezados</p>

<p>Segue o relátorio de vendas por cada loja.</p>

<p>Faturamento:</p>
{turnover.to_html(formatters={'Valor Final': format_collum})}


<p>Quantidade vendida:</p>
{quantity.to_html()}


<p>Ticket médio dos produtos em cada loja:</p>
{avarage_ticket.to_html(formatters={'Ticket Médio': format_collum})}


<p>Qualquer dúvida estou a disposição</p>

<p>Att.,</p>
<p>Wagner</p>
'''

mail.Send()

print('Email enviado')