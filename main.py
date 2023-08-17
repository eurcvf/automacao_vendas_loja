import pandas as pd  # pandas é uma biblioteca para conseguir importar base de dados
# win32 é uma biblioteca para conseguir enviar e-mail
import win32com.client as win32

# importar base de dados
tabela_vendas = pd.read_excel('Vendas.xlsx')

# visualizar a base de dados
pd.set_option('display.max_columns', None)

# faturamento loja
faturamento = tabela_vendas[['ID Loja', 'Valor Final']].groupby(
    'ID Loja').sum()

# qtd produto vendido loja
quantidade = tabela_vendas[['ID Loja', 'Quantidade']].groupby(
    'ID Loja').sum()

# ticket médio loja
ticket_medio = (faturamento['Valor Final'] /
                quantidade['Quantidade']).to_frame()
ticket_medio = ticket_medio.rename(columns={0: 'Ticket Médio'})

# tabela completa com todos os dados
tabela_completa = faturamento.join(quantidade).join(ticket_medio)

# enviar um email com o relatório
outlook = win32.Dispatch('outlook.application')
mail = outlook.CreateItem(0)
mail.To = 'eurcvf@hotmail.com'
mail.Subject = 'Relatório de Vendas | Análise Completa'
mail.HTMLBody = f'''
<p>Prezado,</p>

<p>Segue o Relatório de Vendas por Loja.</p>

<p>Análise Completa:</p>
{tabela_completa.to_html(formatters={'Valor Final': 'R$ {:,.2f}'.format,'Quantidade': '{:,.0f}'.format,'Ticket Médio': 'R$ {:,.2f}'.format})}

<p>Qualquer dúvida estou à disposição.</p>

<p>Att,</p>
Roberto
'''
mail.Send()

print('E-mail enviado com sucesso, em instantes você receberá!')
