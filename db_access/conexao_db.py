import psycopg2


class BancoDeDados:
    """
    Arquivo definindo rotinas para fazer a conexão e queries de select, imports e deletes de dados no banco de dados

    Atributos:
        user: nome do usuário de acesso ao banco de dados
        password: senha de acesso ao banco de dados
        database: nome do banco de dados acessado
        host: endereço de IP do host (localhost para bancos locais)
        port: número da porta de acesso
    """

    def __init__(self, user, password, database, host, port):
        self.user = user
        self.password = password
        self.database = database
        self.host = host
        self.port = port

        # Estabelece conexão com o banco
        self.connection = psycopg2.connect(user=self.user, password=self.password, database=self.database,
                                           host=self.host, port=self.port)
        self.cursor = self.connection.cursor()

    def get_params(self):
        """
        Retorna para o usuário os parâmetros da conexão com o banco de dados

        :return: dicionário contendo dados da conexão
        """

        return self.connection.get_dsn_parameters()

    def add_rows(self, dados, tabela):
        """
        Adiciona linhas à tabela de interesse no banco de dados

        :param dados: lista de tuplas contendo os dados a serem inseridos na tabela
        :param tabela: nome da tabela onde os dados serão inseridos
        """

        len_inputs = len(dados[0])
        self.cursor.executemany(f'INSERT INTO {tabela} VALUES ({"%s, " * (len_inputs - 1)}%s)', dados)
        self.connection.commit()
        print("Dados inseridos na tabela")

    def delete_rows(self, ids, tabela):
        """
        Deleta linhas da tabela com base em seu id

        :param ids: lista contendo os ids das transações a serem deletadas
        :param tabela: nome da tabela de onde os dados serão removidos
        """

        # Formata os ids para poderem ser consumidos pela função
        dados = [(id_,) for id_ in ids]

        self.cursor.executemany(f'DELETE FROM {tabela} WHERE "id" = %s', dados)
        self.connection.commit()
        print("Dados removidos da tabela")

    def select(self, tabela, cols=None):
        """
        Realiza select nas tabelas do banco de dados para consulta

        :param tabela: nome da tabela a ser consultada
        :param cols: colunas de interesse da tabela. Caso não seja fornecido, serão retornadas todas as colunas
        :return: dados da tabela consultada
        """

        if cols is None:
            self.cursor.execute(f'SELECT * FROM {tabela}')
        else:
            assert (type(cols == list)), 'Argumentos para SELECT devem ser passados em formato de lista de strings'

            colunas = ', '.join(cols)
            self.cursor.execute(f'SELECT {colunas} FROM {tabela}')

        return self.cursor.fetchall()

    def select_date(self, tabela, data, cols=None):
        """
        Realiza select para uma data específica nas tabelas do banco de dados para consulta

        :param tabela: nome da tabela a ser consultada
        :param cols: colunas de interesse da tabela. Caso não seja fornecido, serão retornadas todas as colunas
        :param data: data a ser pesquisada na tabela
        :return: dados da tabela consultada
        """

        if cols is None:
            self.cursor.execute(f'SELECT * FROM {tabela} WHERE data LIKE \'%{data}\'')
        else:
            assert (type(cols == list)), 'Argumentos para SELECT devem ser passados em formato de lista de strings'

            colunas = ', '.join(cols)
            self.cursor.execute(f'SELECT {colunas} FROM {tabela}')

        return self.cursor.fetchall()

    def get_last_id(self, tabela):
        """
        Busca o id da última transação lançada no banco
        :param tabela: nome da tabela a ser consultada
        :return: valor do id no banco (caso exista)
        """
        self.cursor.execute(f'SELECT * FROM {tabela} ORDER BY id DESC LIMIT 1')
        last_id = self.cursor.fetchall()[0][0]
        return last_id

    def close_connection(self):
        """
        Encerra a conexão com o banco de dados
        """

        self.cursor.close()
        self.connection.close()
        print('Conexão encerrada')
