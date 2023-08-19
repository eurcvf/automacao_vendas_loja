import pandas as pd  # pandas é uma biblioteca para conseguir importar base de dados
import os
import time 
import win32com.client as win32 # win32 é uma biblioteca para conseguir enviar e-mail

# guarde o nome do diretor da empresa
diretor = input('Informe o nome do diretor(a) da empresa: ')

# guarde o e-mail do diretor da empresa
email_diretor = input('Informe o e-mail do diretor(a) da empresa: ')

# guarde o subtitulo do e-mail para o diretor
subtitulo_email_diretor = input('Informe um subtítulo para o e-mail que será enviado para o(a) diretor(a) (Ex: Relatório de Vendas): ')

# guarde o nome do funcionário utilizando o sistema
funcionario = input('Informe seu nome: ')

# guarde o cargo do funcionário que está utilizando a aplicação
cargo = input('Informe seu cargo: ')

if diretor and email_diretor and subtitulo_email_diretor and funcionario and cargo:
    # informe que as informações foram enviadas e seu sistema fará a análise dos dados e enviará um e-mail para e-mail informado anteriormente
    print(f'\nOlá, {funcionario}. Estamos fazendo a análise dos dados informados e enviaremos um e-mail para o(a) diretor(a) assim que estiver concluído!\nAguarde alguns segundos...\n')
    
    try:
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
        mail.To = email_diretor
        mail.Subject = subtitulo_email_diretor
        mail.HTMLBody = f'''
        <p>Prezado, Sr(a) {diretor}.</p>

        <p>Segue a análise completa com os valores, quantidades e ticket médio de todas as empresas.</p>

        <p>Análise Completa:</p>
        {tabela_completa.to_html(formatters={'Valor Final': 'R${:,.2f}'.format,'Quantidade': '{:,.0f}'.format,'Ticket Médio': 'R${:,.2f}'.format})}

        <p>Qualquer dúvida estou à disposição.</p>

        <p>Att,</p>
        {funcionario} | {cargo}
        '''
        mail.Send()

        print('E-mail enviado com sucesso, em instantes você receberá!')
    except:
        print('Houve um erro ao enviar o e-mail!')
else:
    os.system('cls' or 'clear')
    print('Preencha todos os campos!')
    time.sleep(3)
    os.system('cls' or 'clear')
    
    
    
    

