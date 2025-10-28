# -*- coding: utf-8 -*-
import sqlite3
import os
from datetime import datetime
from validate_docbr import CPF
import re

class Database:
    def __init__(self, db_file='servicos.db'):
        # Garantir que o diretório existe
        os.makedirs(os.path.dirname(os.path.abspath(db_file)), exist_ok=True)

        self.db_file = db_file
        self.cpf_validator = CPF()

        # Inicializar o banco de dados
        with self.get_connection() as conn:
            self.create_tables(conn)
            self.create_indexes(conn)

    def get_connection(self):
        """
        Retorna uma conexão com o banco de dados usando context manager
        """
        return DatabaseConnection(self.db_file)

    def create_tables(self, conn):
        """Cria as tabelas necessárias se não existirem"""
        cursor = conn.cursor()
        cursor.execute('''
        CREATE TABLE IF NOT EXISTS servicos (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            data_solicitacao TEXT NOT NULL,
            cpf TEXT NOT NULL,
            nome TEXT NOT NULL,
            inscricao_municipal TEXT,
            telefone TEXT NOT NULL,
            bairro TEXT NOT NULL,
            rua TEXT NOT NULL,
            numero TEXT NOT NULL,
            referencia TEXT,
            quadra TEXT,
            lote TEXT,
            numero_fossas INTEGER,
            status TEXT DEFAULT 'Pendente',
            data_chegada TEXT,
            data_saida TEXT,
            data_conclusao TEXT,
            placa_veiculo TEXT,
            motorista TEXT,
            ajudante TEXT,
            observacao_empresa TEXT
        )
        ''')
        conn.commit()

    def create_indexes(self, conn):
        """Cria índices para melhorar a performance das consultas frequentes"""
        cursor = conn.cursor()
        cursor.execute('CREATE INDEX IF NOT EXISTS idx_cpf ON servicos (cpf)')
        cursor.execute('CREATE INDEX IF NOT EXISTS idx_endereco ON servicos (bairro, rua, numero, quadra, lote)')
        cursor.execute('CREATE INDEX IF NOT EXISTS idx_status ON servicos (status)')
        cursor.execute('CREATE INDEX IF NOT EXISTS idx_data ON servicos (data_solicitacao)')
        conn.commit()

    def validar_cpf(self, cpf):
        """Valida o CPF usando a biblioteca validate-docbr"""
        return self.cpf_validator.validate(cpf)

    def formatar_cpf(self, cpf):
        """Formata o CPF com pontos e traço"""
        return self.cpf_validator.mask(cpf)

    def verificar_endereco_duplicado(self, bairro, rua, numero, quadra=None, lote=None, id_atual=None):
        """
        Verifica se já existe um serviço cadastrado para o mesmo endereço
        Retorna o ID do serviço duplicado ou None se não houver duplicidade
        """
        with self.get_connection() as conn:
            cursor = conn.cursor()
            query = '''
            SELECT id FROM servicos 
            WHERE bairro = ? AND rua = ? AND numero = ?
            '''
            params = [bairro, rua, numero]

            if quadra:
                query += ' AND quadra = ?'
                params.append(quadra)
            else:
                query += ' AND (quadra IS NULL OR quadra = "")'

            if lote:
                query += ' AND lote = ?'
                params.append(lote)
            else:
                query += ' AND (lote IS NULL OR lote = "")'

            if id_atual:
                query += ' AND id != ?'
                params.append(id_atual)

            cursor.execute(query, params)
            result = cursor.fetchone()
            return result['id'] if result else None

    def inserir_servico(self, dados):
        """
        Insere um novo serviço no banco de dados
        Retorna o ID do serviço inserido
        """
        with self.get_connection() as conn:
            cursor = conn.cursor()
            if 'cpf' in dados and dados['cpf']:
                cpf_limpo = re.sub(r'\D', '', dados['cpf'])
                dados['cpf'] = self.formatar_cpf(cpf_limpo)

            if 'data_solicitacao' not in dados or not dados['data_solicitacao']:
                dados['data_solicitacao'] = datetime.now().strftime('%Y-%m-%d %H:%M:%S')

            campos = ', '.join(dados.keys())
            placeholders = ', '.join(['?'] * len(dados))
            query = f"INSERT INTO servicos ({campos}) VALUES ({placeholders})"
            valores = list(dados.values())

            cursor.execute(query, valores)
            conn.commit()
            return cursor.lastrowid

    def atualizar_servico(self, id_servico, dados):
        """
        Atualiza um serviço existente no banco de dados
        Retorna True se a atualização foi bem-sucedida
        """
        with self.get_connection() as conn:
            cursor = conn.cursor()
            if 'cpf' in dados and dados['cpf']:
                cpf_limpo = re.sub(r'\D', '', dados['cpf'])
                dados['cpf'] = self.formatar_cpf(cpf_limpo)

            atualizacoes = [f"{campo} = ?" for campo in dados.keys()]
            query = f"UPDATE servicos SET {', '.join(atualizacoes)} WHERE id = ?"
            valores = list(dados.values()) + [id_servico]

            cursor.execute(query, valores)
            conn.commit()
            return cursor.rowcount > 0

    def excluir_servico(self, id_servico):
        """
        Exclui um serviço do banco de dados
        Retorna True se a exclusão foi bem-sucedida
        """
        with self.get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute("DELETE FROM servicos WHERE id = ?", (id_servico,))
            conn.commit()
            return cursor.rowcount > 0

    def obter_servico(self, id_servico):
        """
        Obtém um serviço pelo ID
        Retorna um dicionário com os dados do serviço ou None se não encontrado
        """
        with self.get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute("SELECT * FROM servicos WHERE id = ?", (id_servico,))
            row = cursor.fetchone()
            return dict(row) if row else None

    def listar_servicos(self, filtros=None, ordem="data_solicitacao DESC", pagina=1, itens_por_pagina=None):
        """
        Lista todos os serviços, opcionalmente filtrados e paginados
        Retorna uma lista de dicionários com os dados dos serviços e o total de registros
        """
        with self.get_connection() as conn:
            cursor = conn.cursor()
            query = "SELECT * FROM servicos"
            count_query = "SELECT COUNT(*) as total FROM servicos"
            params = []

            if filtros:
                conditions = []
                for campo, valor in filtros.items():
                    campo_sanitizado = campo.lower()
                    if campo_sanitizado in ['id', 'cpf', 'nome', 'telefone', 'bairro', 'rua', 'numero', 'status']:
                        conditions.append(f"{campo_sanitizado} LIKE ?")
                        params.append(f"%{valor}%")
                if conditions:
                    where_clause = " WHERE " + " AND ".join(conditions)
                    query += where_clause
                    count_query += where_clause

            cursor.execute(count_query, params)
            total_registros = cursor.fetchone()['total']

            query += f" ORDER BY {ordem}"

            if itens_por_pagina:
                offset = (pagina - 1) * itens_por_pagina
                query += f" LIMIT {itens_por_pagina} OFFSET {offset}"

            cursor.execute(query, params)
            rows = cursor.fetchall()
            servicos = [dict(row) for row in rows]
            return servicos, total_registros

    def buscar_por_cpf(self, cpf):
        """
        Busca serviços por CPF
        Retorna uma lista de serviços que correspondem ao CPF
        """
        with self.get_connection() as conn:
            cursor = conn.cursor()
            cpf_formatado = self.formatar_cpf(re.sub(r'\D', '', cpf))
            cursor.execute("SELECT * FROM servicos WHERE cpf = ?", (cpf_formatado,))
            rows = cursor.fetchall()
            return [dict(row) for row in rows]


class DatabaseConnection:
    """
    Classe para gerenciar conexões com o banco de dados usando context manager
    """
    def __init__(self, db_file):
        self.db_file = db_file
        self.connection = None

    def __enter__(self):
        self.connection = sqlite3.connect(self.db_file)
        self.connection.row_factory = sqlite3.Row
        return self.connection

    def __exit__(self, exc_type, exc_val, exc_tb):
        if self.connection:
            if exc_type is None:
                self.connection.commit()
            else:
                self.connection.rollback()
            self.connection.close()
        return False