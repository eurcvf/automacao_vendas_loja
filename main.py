import pandas as pd  # pandas é uma biblioteca para conseguir importar base de dados
# win32 é uma biblioteca para conseguir enviar e-mail
import win32com.client as win32

# importar base de dados
tabela_vendas = pd.read_excel('Vendas.xlsx')

# visualizar a base de dados
pd.set_option('display.max_columns', None)
print(tabela_vendas[['ID Loja', 'Valor Final']])

# faturamento loja
faturamento = tabela_vendas[['ID Loja', 'Valor Final']].groupby(
    'ID Loja').sum()
print(faturamento)

print('-' * 50)
# qtd produto vendido loja
quantidade = tabela_vendas[['ID Loja', 'Quantidade']].groupby(
    'ID Loja').sum()
print(quantidade)

print('-' * 50)
# ticket médio loja
ticket_medio = (faturamento['Valor Final'] /
                quantidade['Quantidade']).to_frame()
ticket_medio = ticket_medio.rename(columns={0: 'Ticket Médio'})
print(ticket_medio)

# enviar um email com o relatório
outlook = win32.Dispatch('outlook.application')
mail = outlook.CreateItem(0)
mail.To = 'eurcvf@hotmail.com'
mail.Subject = 'Relatório de Vendas | Lojas'
mail.HTMLBody = f'''
<p>Prezado,</p>

<p>Segue o Relatório de Vendas por Loja.</p>

<p>Faturamento:</p>
{faturamento.to_html(formatters={'Valor Final': 'R$ {:,.2f}'.format})}

<p>Qtde Vendida:</p>
{quantidade.to_html()}

<p>Ticket Médio (Produtos)</p>
{ticket_medio.to_html(formatters={'Valor Final': 'R$ {:,.2f}'.format})}

<p>Qualquer dúvida estou à disposição.</p>

<p>Att,</p>
Roberto
'''
mail.Send()

print('E-mail enviado com sucesso, em instantes você receberá!')
