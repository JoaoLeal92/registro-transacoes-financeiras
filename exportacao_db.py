from db_access import conexao_db, params
import pandas as pd
import re
import difflib
import dropbox

# Importando o arquivo texto de referência (do Dropbox)
dbx = dropbox.Dropbox(params.dropbox_access)
md, res = dbx.files_download(params.txt_ref)
gastos = res.content.decode("utf-8")

# Separando o texto em linhas e dividindo as colunas para transformar em dataframe
gastos = gastos.split('\n')
gastos_split = [re.split('\s*-\s*', gasto) for gasto in gastos]

print('Arquivo de referências lido com sucesso')

# Pede ao usuário o mês e ano de cada transação
mes = input("Insira o mês da transação: ")
ano = input("Insira o ano da transação: ")

# ----------------------------- FORMATAÇÕES DO DATAFRAME ---------------------------------
# Criando o dataframe
headers = ['Data', 'Local', 'Valor', 'Categoria', 'Descrição']
df = pd.DataFrame(gastos_split, columns=headers)

# Substitui virgulas por pontos no campo 'Valor':
df['Valor'] = [data.replace(',', '.') for data in df['Valor']]

# Transformando em números caso haja alguma expressão do tipo x+y na coluna 'Valor'
df['Valor'] = df.eval(df['Valor'])

# Transformando a primeira letra em maiúscula no campo 'Conta'
df['Categoria'] = df['Categoria'].str.title()

# Removendo grupos com o padrão [v] (já foram incluídos na planilha previamente
patternDel = "([v])"
filtro = df['Data'].str.contains(patternDel)
df = df[~filtro]

# Resetando os índices
df = df.reset_index(drop=True)

# Verificando campos que não estão identificados corretamente e substituindo pelos rótulos corretos

# Lista de categorias:
categorias = ['Contas Residenciais', 'Transporte', 'Compras', 'Desenvolvimento Pessoal', 'Lazer',
              'Independência Financeira',
              'Despesas de longo prazo', 'Poupança Transporte', 'Poupança Desenvolvimento', 'Poupança Lazer',
              'Poupança Despesas', 'Poupança Independência']

for categoria in df['Categoria']:
    if categoria not in categorias:
        ind = df[df['Categoria'] == categoria].index[0]
        repl = difflib.get_close_matches(categoria, categorias)[0]
        df.iloc[ind, 3] = repl

        print('Categoria \'' + str(categoria) + '\' foi substituída por \'' + str(repl) + '\'')

# Formatando as datas
datas = df['Data'].tolist()
datas = [data.replace('[ ', '').replace(' ]', '') for data in datas]

mes_ano = '/' + mes + '/' + ano
datas_final = [data + mes_ano for data in datas]

df['Data'] = datas_final
df['Data'] = pd.to_datetime(df['Data'], dayfirst=True).dt.strftime('%d/%m/%Y')

# Estabelece a conexão com o banco de dados
banco = conexao_db.BancoDeDados(user=params.user,
                                password=params.password,
                                host=params.host,
                                port=params.port,
                                database=params.database)

# Consulta banco para verificar id da última transação adicionada
ultimo_id = banco.get_last_id('gastos_diarios')

transacoes = [(ultimo_id + (row[0]+1), *row[1:]) for row in df.itertuples()]

banco.add_rows(dados=transacoes, tabela='gastos_diarios')

banco.close_connection()