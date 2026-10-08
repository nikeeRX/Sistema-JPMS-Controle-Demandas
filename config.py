import os
import sqlite3
import psycopg2
from tkinter import messagebox
from dotenv import load_dotenv # Adicione esta linha

# Carrega as variáveis do arquivo .env (se ele existir na sua máquina)
load_dotenv()

# ==========================================
# LISTAS GLOBAIS (CONTROLE DE DEMANDAS)
# ==========================================
TIPOS_ROTINA = ["Reajuste", "Novo Contrato", "Inclusão", "Exclusão", "Ajuste", "Extensão", "Descredenciamento", "Não Identificada"]
TIPOS_AVULSA = ['Credenciamento', 'Descredenciamento', 'Análise de Vulnerabilidade', 'Parametrização', 'Resposta de E-mail', 'RPS', 'Lote de Erros', 'Atendimento de Chamado', 'Atualização Cadastral', 'Outro']
STATUS_AVULSA = ['Pendente', 'Recebida', 'Em Análise', 'Aguardando Retorno', 'Aguardando Área Demandante', 'Em andamento', 'Concluído', 'Cancelada']
SOLICITANTES = ['GEREL/COATE', 'GEREL/COCAD', 'UAR-AM', 'NAR-AC', 'NAR-RR', 'UAR-BA', 'NAR-AL', 'NAR-SE', 'UAR-CE', 'NAR-MA', 'NAR-PI', 'UAR-DF', 'NAR-GO', 'NAR-TO', 'UAR-MG', 'UAR-MS', 'NAR-MT', 'NAR-RO', 'UAR-PA', 'NAR-AP', 'UAR-PE', 'NAR-PB', 'NAR-RN', 'UAR-RJ', 'NAR-ES', 'UAR-RS', 'NAR-SC', 'UAR-SPI', 'UAR-SPM', 'NAR-PR', 'GEREM/COANS', 'GERIN', 'GEASA/COICS', 'GERED/COCAP', 'GEREG/COPAR', 'GERIA/CONFA', 'GEASA/COSER', 'GERED/CODER', 'GERED/CONEC', 'GEREG/CORAS', 'GERIA/COPME', 'GECOF/COCOT', 'GECOF/COPAG', 'GECOF/COREB', 'GEFAC/COANF', 'GEFAC/COAUS', 'GEFAC/COEFA', 'GEPES/COGEP', 'GERAD/COBES', 'GERAD/COPOP', 'GETEC/COSIS', 'GETEC/COINF', 'GETEC/COIDE', 'GECOM', 'GECRI', 'GEJUR/CONSU', 'GEJUR/CONTE', 'GEPRO', 'OUVID', 'SEGER', 'ASSESSORIA', 'COEDI', 'AUDIN']

# ==========================================
# MOTOR DO BANCO DE DADOS (PostgreSQL / SQLite Fallback)
# ==========================================

# O sistema vai puxar a URL direto das engrenagens do Railway de forma invisível!
DATABASE_URL_RAILWAY = os.environ.get("DATABASE_URL")

# Variável global para definir se estamos no Postgres ou SQLite
MOTOR_ATUAL = "postgres" 

def conectar_db():
    global MOTOR_ATUAL
    
    # Só tenta a nuvem se a URL existir no ambiente
    if DATABASE_URL_RAILWAY:
        try:
            # TENTA CONECTAR NA NUVEM (RAILWAY)
            conn = psycopg2.connect(DATABASE_URL_RAILWAY)
            MOTOR_ATUAL = "postgres"
            return conn
        except Exception as e_nuvem:
            print(f"Erro na Nuvem (Postgres): {e_nuvem}. Tentando Rede/Local...")
            
    MOTOR_ATUAL = "sqlite"
    # TENTA REDE OU LOCAL (FALLBACK)
    pasta_rede = r'S:\DIOPE\GENEG\COCAP\1. DEMANDAS GERENCIAIS\Sistema de Controle de Demandas\BANCO DE DADOS'
    try:
        if not os.path.exists(pasta_rede): os.makedirs(pasta_rede, exist_ok=True)
        caminho_banco = os.path.join(pasta_rede, 'demandas.db')
    except:
        pasta_local = os.path.join(os.getcwd(), 'BANCO_LOCAL')
        if not os.path.exists(pasta_local): os.makedirs(pasta_local, exist_ok=True)
        caminho_banco = os.path.join(pasta_local, 'demandas.db')
        
    conn = sqlite3.connect(caminho_banco, timeout=20.0)
    conn.execute('PRAGMA journal_mode=WAL')
    return conn

def setup_db():
    conn = conectar_db()
    c = conn.cursor()
    
    if MOTOR_ATUAL == "postgres":
        # Sintaxe PostgreSQL
        c.execute('''CREATE TABLE IF NOT EXISTS users (id SERIAL PRIMARY KEY, username VARCHAR(255) UNIQUE, password VARCHAR(255), role VARCHAR(50), first_login INTEGER DEFAULT 1)''')
        c.execute('''CREATE TABLE IF NOT EXISTS demands (id SERIAL PRIMARY KEY, filename VARCHAR(255) UNIQUE, type VARCHAR(100), status VARCHAR(100), assigned_to VARCHAR(100), prestador TEXT, cnpj VARCHAR(50), municipio VARCHAR(100), uf VARCHAR(10), data_entrada VARCHAR(50), filepath TEXT, tipo_negociacao VARCHAR(100), alcada VARCHAR(100), impacto_perc VARCHAR(50), impacto_rs VARCHAR(50), nup_edoc VARCHAR(100), tipo_processo VARCHAR(100), fop_autorizacao VARCHAR(100), data_aprovacao VARCHAR(50), data_finalizacao VARCHAR(50), sla_geral INTEGER, sla_gered INTEGER, observacao TEXT)''')
        c.execute('''CREATE TABLE IF NOT EXISTS demandas_avulsas (id SERIAL PRIMARY KEY, tipo_demanda VARCHAR(100), assunto TEXT, solicitante VARCHAR(100), data_entrada VARCHAR(50), prazo VARCHAR(50), status VARCHAR(50), responsavel VARCHAR(100), data_conclusao VARCHAR(50), observacao TEXT)''')
        
        # Garante que admin exista
        c.execute("SELECT * FROM users WHERE username='admin'")
        if not c.fetchone(): 
            c.execute("INSERT INTO users (username, password, role, first_login) VALUES ('admin', 'admin', 'admin', 0)")
            
    else:
        # Sintaxe SQLite Original (Fallback)
        c.execute('''CREATE TABLE IF NOT EXISTS users (id INTEGER PRIMARY KEY AUTOINCREMENT, username TEXT UNIQUE, password TEXT, role TEXT, first_login INTEGER DEFAULT 1)''')
        c.execute('''CREATE TABLE IF NOT EXISTS demands (id INTEGER PRIMARY KEY AUTOINCREMENT, filename TEXT UNIQUE, type TEXT, status TEXT, assigned_to TEXT, prestador TEXT, cnpj TEXT, municipio TEXT, uf TEXT, data_entrada TEXT, filepath TEXT, tipo_negociacao TEXT, alcada TEXT, impacto_perc TEXT, impacto_rs TEXT, nup_edoc TEXT, tipo_processo TEXT, fop_autorizacao TEXT, data_aprovacao TEXT, data_finalizacao TEXT, sla_geral INTEGER, sla_gered INTEGER, observacao TEXT)''')
        c.execute('''CREATE TABLE IF NOT EXISTS demandas_avulsas (id INTEGER PRIMARY KEY AUTOINCREMENT, tipo_demanda TEXT, assunto TEXT, solicitante TEXT, data_entrada TEXT, prazo TEXT, status TEXT, responsavel TEXT, data_conclusao TEXT, observacao TEXT)''')
        
        c.execute("SELECT * FROM users WHERE username='admin'")
        if not c.fetchone(): 
            c.execute("INSERT INTO users (username, password, role, first_login) VALUES ('admin', 'admin', 'admin', 0)")

    conn.commit()
    conn.close()

# Função auxiliar para lidar com curingas em queries dinâmicas SQL (?) vs Postgres (%s)
def replace_placeholders(query):
    if MOTOR_ATUAL == "postgres":
        return query.replace("?", "%s")
    return query
