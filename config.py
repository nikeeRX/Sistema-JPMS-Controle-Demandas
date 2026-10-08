import os
import psycopg2

# O Railway vai injetar a URL do Postgres automaticamente aqui
DATABASE_URL = os.environ.get("DATABASE_URL", "")

def conectar_db():
    try:
        conn = psycopg2.connect(DATABASE_URL)
        return conn
    except Exception as e:
        print(f"Erro ao conectar no banco Web: {e}")
        return None

def setup_db():
    conn = conectar_db()
    if not conn: return
    c = conn.cursor()
    c.execute('''CREATE TABLE IF NOT EXISTS users (id SERIAL PRIMARY KEY, username VARCHAR(255) UNIQUE, password VARCHAR(255), role VARCHAR(50), first_login INTEGER DEFAULT 1)''')
    c.execute('''CREATE TABLE IF NOT EXISTS demands (id SERIAL PRIMARY KEY, filename VARCHAR(255) UNIQUE, type VARCHAR(100), status VARCHAR(100), assigned_to VARCHAR(100), prestador TEXT, cnpj VARCHAR(50), municipio VARCHAR(100), uf VARCHAR(10), data_entrada VARCHAR(50), filepath TEXT, tipo_negociacao VARCHAR(100), alcada VARCHAR(100), impacto_perc VARCHAR(50), impacto_rs VARCHAR(50), nup_edoc VARCHAR(100), tipo_processo VARCHAR(100), fop_autorizacao VARCHAR(100), data_aprovacao VARCHAR(50), data_finalizacao VARCHAR(50), sla_geral INTEGER, sla_gered INTEGER, observacao TEXT)''')
    c.execute('''CREATE TABLE IF NOT EXISTS demandas_avulsas (id SERIAL PRIMARY KEY, tipo_demanda VARCHAR(100), assunto TEXT, solicitante VARCHAR(100), data_entrada VARCHAR(50), prazo VARCHAR(50), status VARCHAR(50), responsavel VARCHAR(100), data_conclusao VARCHAR(50), observacao TEXT)''')
    
    c.execute("SELECT * FROM users WHERE username='admin'")
    if not c.fetchone(): 
        c.execute("INSERT INTO users (username, password, role, first_login) VALUES ('admin', 'admin', 'admin', 0)")
    conn.commit()
    conn.close()

def replace_placeholders(query):
    return query.replace("?", "%s")
