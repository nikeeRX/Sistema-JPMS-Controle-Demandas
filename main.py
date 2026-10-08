from flask import Flask, request, redirect, url_for, session, flash, render_template_string, send_file
from config import setup_db, conectar_db
import os
import PyPDF2
import re
from datetime import datetime
import base64
from io import BytesIO
import pandas as pd

# Configuração CRÍTICA para o Matplotlib rodar em servidores Web (Sem tela)
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt

app = Flask(__name__)
app.secret_key = 'chave_super_secreta_da_postal_saude'

# Inicia o banco de dados
setup_db()

# Função auxiliar para compatibilidade entre Postgres e SQLite
def get_ph(conn):
    return "%s" if "psycopg2" in str(type(conn)) else "?"

# ==========================================
# LISTAS GLOBAIS
# ==========================================
TIPOS_ROTINA = ["Reajuste", "Novo Contrato", "Inclusão", "Exclusão", "Ajuste", "Extensão", "Descredenciamento", "Não Identificada"]
TIPOS_AVULSA = ['Credenciamento', 'Descredenciamento', 'Análise de Vulnerabilidade', 'Parametrização', 'Resposta de E-mail', 'RPS', 'Lote de Erros', 'Atendimento de Chamado', 'Atualização Cadastral', 'Outro']
STATUS_AVULSA = ['Pendente', 'Recebida', 'Em Análise', 'Aguardando Retorno', 'Aguardando Área Demandante', 'Em andamento', 'Concluído', 'Cancelada']
SOLICITANTES = ['GEREL/COATE', 'GEREL/COCAD', 'UAR-AM', 'NAR-AC', 'NAR-RR', 'UAR-BA', 'NAR-AL', 'NAR-SE', 'UAR-CE', 'NAR-MA', 'NAR-PI', 'UAR-DF', 'NAR-GO', 'NAR-TO', 'UAR-MG', 'UAR-MS', 'NAR-MT', 'NAR-RO', 'UAR-PA', 'NAR-AP', 'UAR-PE', 'NAR-PB', 'NAR-RN', 'UAR-RJ', 'NAR-ES', 'UAR-RS', 'NAR-SC', 'UAR-SPI', 'UAR-SPM', 'NAR-PR', 'GEREM/COANS', 'GERIN', 'GEASA/COICS', 'GERED/COCAP']

# ==========================================
# CÓDIGOS HTML EMBUTIDOS
# ==========================================

TELA_LOGIN = """
<!DOCTYPE html>
<html lang="pt-BR">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Login - COCAP</title>
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet">
    <style>
        body { background-color: #f4f7f6; display: flex; align-items: center; justify-content: center; height: 100vh; margin: 0; }
        .login-card { border-radius: 15px; box-shadow: 0 10px 30px rgba(0,0,0,0.1); padding: 40px; background: white; width: 100%; max-width: 400px; }
        .btn-postal { background-color: #004b87; color: white; font-weight: bold; }
        .btn-postal:hover { background-color: #003666; color: white; }
    </style>
</head>
<body>
<div class="login-card">
    <div class="text-center mb-4">
        <h3 style="color: #004b87; font-weight: 800;">POSTAL SAÚDE</h3>
        <p class="text-muted">Central de Operações (COCAP)</p>
    </div>
    {% with messages = get_flashed_messages(with_categories=true) %}
      {% if messages %}
        {% for category, message in messages %}
          <div class="alert alert-{{ category }}">{{ message }}</div>
        {% endfor %}
      {% endif %}
    {% endwith %}
    <form method="POST" action="/">
        <div class="mb-3">
            <label class="form-label font-weight-bold">Usuário</label>
            <input type="text" name="username" class="form-control" required>
        </div>
        <div class="mb-4">
            <label class="form-label font-weight-bold">Senha</label>
            <input type="password" name="password" class="form-control" required>
        </div>
        <button type="submit" class="btn btn-postal w-100 py-2">Entrar no Sistema</button>
    </form>
</div>
</body>
</html>
"""

TELA_DASHBOARD = """
<!DOCTYPE html>
<html lang="pt-BR">
<head>
    <meta charset="UTF-8">
    <title>Portal COCAP</title>
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet">
</head>
<body style="background-color: #f0f2f5;">
<nav class="navbar navbar-expand-lg navbar-dark" style="background-color: #004b87;">
  <div class="container-fluid">
    <a class="navbar-brand fw-bold" href="#">COCAP Web</a>
    <div class="d-flex text-white align-items-center">
      <span class="me-3">Olá, <strong>{{ user }}</strong> ({{ role }})</span>
      <a href="/logout" class="btn btn-danger btn-sm">Sair</a>
    </div>
  </div>
</nav>

<div class="container mt-5 text-center">
    <h1 style="color: #004b87; font-weight: 800; margin-bottom: 2rem;">PORTAL CENTRAL DE OPERAÇÕES</h1>
    <div class="row justify-content-center gap-4">
        <div class="col-md-5">
            <div class="card shadow-sm h-100 p-4 border-0">
                <h1 class="text-info mb-3" style="font-size: 3rem;">📊</h1>
                <h4>Controle de Demandas</h4>
                <p class="text-muted mb-4">Gerencie Rotinas, PDFs, Avulsas e acesse o Dashboard.</p>
                <a href="/demandas" class="btn btn-info text-white w-100 fw-bold py-2">Acessar Módulo</a>
            </div>
        </div>
        <div class="col-md-5">
            <div class="card shadow-sm h-100 p-4 border-0">
                <h1 class="text-warning mb-3" style="font-size: 3rem;">🏥</h1>
                <h4>Pente Fino (RN 665)</h4>
                <p class="text-muted mb-4">Substituição de Prestadores e análise de vulnerabilidade de rede.</p>
                <a href="/pentefino" class="btn btn-warning text-white w-100 fw-bold py-2">Acessar Módulo</a>
            </div>
        </div>
    </div>
    {% if role == 'admin' %}
    <div class="row justify-content-center mt-4">
        <div class="col-md-5">
            <a href="/admin" class="btn btn-secondary w-100 fw-bold py-3 shadow-sm border-0">⚙️ Painel Administrativo (BD & Usuários)</a>
        </div>
    </div>
    {% endif %}
</div>
</body>
</html>
"""

TELA_ADMIN = """
<!DOCTYPE html>
<html lang="pt-BR">
<head>
    <meta charset="UTF-8">
    <title>Painel Administrativo</title>
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet">
</head>
<body style="background-color: #f0f2f5;">
<nav class="navbar navbar-dark" style="background-color: #333;">
  <div class="container-fluid">
    <a href="/dashboard" class="btn btn-outline-light btn-sm">⬅ Voltar ao Portal</a>
    <span class="navbar-text text-white fw-bold">Painel Administrativo: BD e Usuários</span>
  </div>
</nav>

<div class="container mt-4">
    {% with messages = get_flashed_messages(with_categories=true) %}
      {% if messages %}
        {% for category, message in messages %}
          <div class="alert alert-{{ category }}">{{ message }}</div>
        {% endfor %}
      {% endif %}
    {% endwith %}

    <div class="row">
        <!-- ÁREA DE BANCO DE DADOS -->
        <div class="col-md-12 mb-4">
            <div class="card shadow-sm border-0">
                <div class="card-header bg-dark text-white fw-bold">💾 Gestão do Banco de Dados</div>
                <div class="card-body">
                    <div class="row text-center">
                        <div class="col-md-6 border-end">
                            <h5 class="text-success">Exportar Backup</h5>
                            <p class="text-muted small">Baixe um arquivo Excel com todas as demandas e usuários.</p>
                            <a href="/admin/exportar" class="btn btn-success fw-bold w-75">⬇ Baixar Backup (Excel)</a>
                        </div>
                        <div class="col-md-6">
                            <h5 class="text-danger">Restaurar ou Migrar Banco</h5>
                            <p class="text-muted small">Faça upload do Excel (.xlsx) ou do Banco Antigo (.db)</p>
                            <form action="/admin/importar" method="POST" enctype="multipart/form-data" class="d-flex flex-column align-items-center">
                                <input type="file" name="file_backup" class="form-control form-control-sm w-75 mb-2" accept=".xlsx,.db" required>
                                <button type="submit" class="btn btn-danger fw-bold w-75" onclick="return confirm('ATENÇÃO: Todas as informações atuais da Nuvem serão APAGADAS e substituídas por este arquivo. Deseja continuar?');">⬆ Importar Arquivo</button>
                            </form>
                        </div>
                    </div>
                </div>
            </div>
        </div>

        <!-- ÁREA DE USUÁRIOS -->
        <div class="col-md-12">
            <div class="card shadow-sm mb-4 border-0">
                <div class="card-body">
                    <h5 class="card-title mb-3">Criar Novo Usuário</h5>
                    <form method="POST" action="/admin">
                        <input type="hidden" name="acao" value="adicionar">
                        <div class="row g-2 align-items-end">
                            <div class="col-md-4">
                                <label class="form-label small fw-bold">Usuário</label>
                                <input type="text" name="username" class="form-control" required>
                            </div>
                            <div class="col-md-3">
                                <label class="form-label small fw-bold">Senha</label>
                                <input type="password" name="password" class="form-control" required>
                            </div>
                            <div class="col-md-3">
                                <label class="form-label small fw-bold">Nível</label>
                                <select name="role" class="form-select">
                                    <option value="user">User</option>
                                    <option value="admin">Admin</option>
                                </select>
                            </div>
                            <div class="col-md-2">
                                <button type="submit" class="btn btn-primary w-100 fw-bold">Salvar</button>
                            </div>
                        </div>
                    </form>
                </div>
            </div>

            <div class="card shadow-sm border-0">
                <div class="card-body">
                    <h5 class="card-title mb-3">Lista de Usuários</h5>
                    <div class="table-responsive">
                        <table class="table table-hover align-middle">
                            <thead class="table-light">
                                <tr><th>ID</th><th>Usuário</th><th>Nível</th><th>Ações</th></tr>
                            </thead>
                            <tbody>
                                {% for u in usuarios %}
                                <tr>
                                    <td>{{ u[0] }}</td>
                                    <td>{{ u[1] }}</td>
                                    <td><span class="badge bg-{{ 'dark' if u[2] == 'admin' else 'secondary' }}">{{ u[2] | upper }}</span></td>
                                    <td>
                                        {% if u[1] != 'admin' and u[1] != user %}
                                        <form method="POST" action="/admin" style="display:inline;">
                                            <input type="hidden" name="acao" value="eliminar">
                                            <input type="hidden" name="user_id" value="{{ u[0] }}">
                                            <button type="submit" class="btn btn-outline-danger btn-sm" onclick="return confirm('Eliminar o usuário {{ u[1] }}?');">Excluir</button>
                                        </form>
                                        {% else %}
                                        <span class="text-muted small">Sistema</span>
                                        {% endif %}
                                    </td>
                                </tr>
                                {% endfor %}
                            </tbody>
                        </table>
                    </div>
                </div>
            </div>
        </div>
    </div>
</div>
</body>
</html>
"""

TELA_HUB_DEMANDAS = """
<!DOCTYPE html>
<html lang="pt-BR">
<head>
    <title>Módulo de Demandas</title>
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet">
</head>
<body style="background-color: #f0f2f5;">
<nav class="navbar navbar-dark bg-dark">
  <div class="container-fluid">
    <a href="/dashboard" class="btn btn-outline-light btn-sm">⬅ Voltar ao Portal</a>
    <span class="navbar-text text-white fw-bold">Gestão de Demandas</span>
  </div>
</nav>
<div class="container mt-5 text-center">
    <div class="row justify-content-center">
        <div class="col-md-4 mb-4">
            <div class="card shadow-sm border-0 h-100 p-4">
                <h1 class="mb-3">📄</h1>
                <a href="/demandas/rotinas" class="btn btn-primary w-100 py-3 fw-bold fs-6">Demandas de Rotina (PDFs)</a>
            </div>
        </div>
        <div class="col-md-4 mb-4">
            <div class="card shadow-sm border-0 h-100 p-4">
                <h1 class="mb-3">📝</h1>
                <a href="/demandas/avulsas" class="btn btn-warning text-white w-100 py-3 fw-bold fs-6">Demandas Avulsas</a>
            </div>
        </div>
        <div class="col-md-4 mb-4">
            <div class="card shadow-sm border-0 h-100 p-4">
                <h1 class="mb-3">📊</h1>
                <a href="/demandas/dashboard_graficos" class="btn btn-info text-white w-100 py-3 fw-bold fs-6">Dashboard Gerencial</a>
            </div>
        </div>
    </div>
</div>
</body>
</html>
"""

TELA_AVULSAS = """
<!DOCTYPE html>
<html lang="pt-BR">
<head>
    <title>Demandas Avulsas</title>
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet">
</head>
<body style="background-color: #f0f2f5;">
<nav class="navbar navbar-dark bg-dark">
  <div class="container-fluid"><a href="/demandas" class="btn btn-outline-light btn-sm">⬅ Voltar</a><span class="text-white fw-bold">Demandas Avulsas</span></div>
</nav>

<div class="container-fluid mt-4">
    {% with messages = get_flashed_messages(with_categories=true) %}
      {% if messages %}
        {% for category, message in messages %}<div class="alert alert-{{ category }}">{{ message }}</div>{% endfor %}
      {% endif %}
    {% endwith %}

    <div class="card shadow-sm mb-4 border-0">
        <div class="card-body">
            <form method="POST" action="/demandas/avulsas">
                <input type="hidden" name="acao" value="nova">
                <div class="row g-2">
                    <div class="col-md-2"><select name="tipo" class="form-select" required><option value="">Tipo...</option>{% for t in tipos %}<option value="{{t}}">{{t}}</option>{% endfor %}</select></div>
                    <div class="col-md-3"><input type="text" name="assunto" class="form-control" placeholder="Assunto/Demanda" required></div>
                    <div class="col-md-2"><select name="solicitante" class="form-select" required><option value="">Solicitante...</option>{% for s in sols %}<option value="{{s}}">{{s}}</option>{% endfor %}</select></div>
                    <div class="col-md-2"><input type="date" name="prazo" class="form-control" required></div>
                    <div class="col-md-2"><button type="submit" class="btn btn-success w-100 fw-bold">+ Nova Demanda</button></div>
                </div>
            </form>
        </div>
    </div>

    <div class="card shadow-sm border-0">
        <div class="card-body">
            <div class="table-responsive">
                <table class="table table-hover table-bordered align-middle text-center" style="font-size: 0.9em;">
                    <thead class="table-dark">
                        <tr><th>ID</th><th>Tipo</th><th>Assunto</th><th>Solicitante</th><th>Entrada</th><th>Prazo</th><th>Status</th><th>Responsável</th><th>Conclusão</th><th>Ações</th></tr>
                    </thead>
                    <tbody>
                        {% for d in demandas %}
                        <tr>
                            <td>{{ d[0] }}</td><td>{{ d[1] }}</td><td class="text-start">{{ d[2] }}</td><td>{{ d[3] }}</td><td>{{ d[4] }}</td><td>{{ d[5] }}</td>
                            <td><span class="badge bg-{{ 'warning text-dark' if d[6]=='Pendente' else 'success' if d[6]=='Concluído' else 'info' }}">{{ d[6] }}</span></td>
                            <td>{{ d[7] }}</td><td>{{ d[8] }}</td>
                            <td>
                                <form method="POST" action="/demandas/avulsas" style="display:inline;">
                                    <input type="hidden" name="id" value="{{ d[0] }}">
                                    <button type="submit" name="acao" value="assumir" class="btn btn-warning btn-sm fw-bold" {% if d[6] == 'Concluído' %}disabled{% endif %}>Assumir</button>
                                    <button type="submit" name="acao" value="concluir" class="btn btn-success btn-sm fw-bold" {% if d[6] == 'Concluído' %}disabled{% endif %}>✔</button>
                                </form>
                            </td>
                        </tr>
                        {% endfor %}
                    </tbody>
                </table>
            </div>
        </div>
    </div>
</div>
</body>
</html>
"""

TELA_ROTINAS = """
<!DOCTYPE html>
<html lang="pt-BR">
<head>
    <title>Rotinas (PDFs)</title>
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet">
</head>
<body style="background-color: #f0f2f5;">
<nav class="navbar navbar-dark bg-dark">
  <div class="container-fluid"><a href="/demandas" class="btn btn-outline-light btn-sm">⬅ Voltar</a><span class="text-white fw-bold">Processamento de PDFs</span></div>
</nav>

<div class="container-fluid mt-4">
    {% with messages = get_flashed_messages(with_categories=true) %}
      {% if messages %}
        {% for category, message in messages %}<div class="alert alert-{{ category }}">{{ message }}</div>{% endfor %}
      {% endif %}
    {% endwith %}

    <div class="card shadow-sm mb-4 border-0">
        <div class="card-body text-center bg-light rounded">
            <h5 class="card-title text-primary mb-3">Importar Múltiplos PDFs</h5>
            <form method="POST" action="/demandas/rotinas/upload" enctype="multipart/form-data">
                <input type="file" name="pdfs" class="form-control mb-3 w-50 mx-auto" multiple accept=".pdf" required>
                <button type="submit" class="btn btn-primary fw-bold px-5">Processar Arquivos</button>
            </form>
        </div>
    </div>

    <div class="card shadow-sm border-0">
        <div class="card-body">
            <div class="table-responsive">
                <table class="table table-hover table-bordered align-middle text-center" style="font-size: 0.85em;">
                    <thead class="table-dark">
                        <tr><th>ID</th><th>Data Entrada</th><th>Prestador</th><th>CNPJ</th><th>Mun/UF</th><th>Demanda</th><th>Status</th><th>Resp.</th><th>Ações</th></tr>
                    </thead>
                    <tbody>
                        {% for d in rotinas %}
                        <tr>
                            <td>{{ d[0] }}</td><td>{{ d[9] }}</td><td class="text-start fw-bold">{{ d[5] }}</td><td>{{ d[6] }}</td><td>{{ d[7] }}/{{ d[8] }}</td><td>{{ d[2] }}</td>
                            <td><span class="badge bg-{{ 'warning text-dark' if d[3]=='Pendente' else 'success' if d[3]=='Finalizada' else 'info' }}">{{ d[3] }}</span></td>
                            <td>{{ d[4] }}</td>
                            <td>
                                <form method="POST" action="/demandas/rotinas/acao" style="display:inline;">
                                    <input type="hidden" name="id" value="{{ d[0] }}">
                                    <button type="submit" name="acao" value="assumir" class="btn btn-warning btn-sm fw-bold" {% if d[3] == 'Finalizada' %}disabled{% endif %}>Assumir</button>
                                    <button type="submit" name="acao" value="finalizar" class="btn btn-success btn-sm fw-bold" {% if d[3] == 'Finalizada' %}disabled{% endif %}>Finalizar</button>
                                </form>
                            </td>
                        </tr>
                        {% endfor %}
                    </tbody>
                </table>
            </div>
        </div>
    </div>
</div>
</body>
</html>
"""

TELA_DASHBOARD_GRAFICOS = """
<!DOCTYPE html>
<html lang="pt-BR">
<head>
    <title>Dashboard Gerencial</title>
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet">
</head>
<body style="background-color: #f0f2f5;">
<nav class="navbar navbar-dark bg-dark">
  <div class="container-fluid"><a href="/demandas" class="btn btn-outline-light btn-sm">⬅ Voltar</a><span class="text-white fw-bold">Dashboard Geral</span></div>
</nav>

<div class="container mt-4">
    <div class="row text-center mb-4">
        <div class="col-md-3"><div class="card shadow-sm border-0 bg-primary text-white p-3"><h5 class="mb-1">Total Demandas</h5><h2 class="fw-bold">{{ d_total }}</h2></div></div>
        <div class="col-md-3"><div class="card shadow-sm border-0 bg-warning text-dark p-3"><h5 class="mb-1">Pendentes</h5><h2 class="fw-bold">{{ d_pend }}</h2></div></div>
        <div class="col-md-3"><div class="card shadow-sm border-0 bg-info text-white p-3"><h5 class="mb-1">Em Análise</h5><h2 class="fw-bold">{{ d_ana }}</h2></div></div>
        <div class="col-md-3"><div class="card shadow-sm border-0 bg-success text-white p-3"><h5 class="mb-1">Concluídas</h5><h2 class="fw-bold">{{ d_conc }}</h2></div></div>
    </div>
    
    <div class="row">
        <div class="col-md-6 mb-4">
            <div class="card shadow-sm border-0 h-100"><div class="card-body text-center">
                <h5 class="fw-bold text-secondary mb-3">Produtividade por Colaborador</h5>
                <img src="data:image/png;base64,{{ chart_prod }}" class="img-fluid rounded">
            </div></div>
        </div>
        <div class="col-md-6 mb-4">
            <div class="card shadow-sm border-0 h-100"><div class="card-body text-center">
                <h5 class="fw-bold text-secondary mb-3">Distribuição por Tipo de Demanda</h5>
                <img src="data:image/png;base64,{{ chart_tipo }}" class="img-fluid rounded">
            </div></div>
        </div>
    </div>
</div>
</body>
</html>
"""

TELA_CONSTRUCAO = """
<!DOCTYPE html>
<html lang="pt-BR"><head><title>Pente Fino</title><link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet"></head>
<body style="display: flex; flex-direction: column; align-items: center; justify-content: center; height: 100vh; background-color: #f0f2f5;">
    <h1 style="font-size: 5rem;">🚧</h1><h2 class="mt-3 text-secondary fw-bold">Módulo: Pente Fino (RN 665)</h2><p class="text-muted">Na próxima etapa, injetaremos o código de upload das planilhas!</p>
    <a href="/dashboard" class="btn btn-primary mt-4 fw-bold px-4 py-2">⬅ Voltar ao Portal</a>
</body></html>
"""

# ==========================================
# ROTAS DO SERVIDOR WEB
# ==========================================

@app.route('/', methods=['GET', 'POST'])
def login():
    if request.method == 'POST':
        user = request.form.get('username')
        password = request.form.get('password')
        conn = conectar_db()
        if not conn:
            flash("Erro de conexão com o banco de dados.", "danger")
            return render_template_string(TELA_LOGIN)
            
        c = conn.cursor()
        ph = get_ph(conn)
        c.execute(f"SELECT id, role FROM users WHERE username={ph} AND password={ph}", (user, password))
        res = c.fetchone()
        conn.close()
        
        if res:
            session['user'] = user
            session['role'] = res[1]
            return redirect(url_for('dashboard'))
        else:
            flash("Usuário ou senha incorretos!", "danger")
    return render_template_string(TELA_LOGIN)

@app.route('/dashboard')
def dashboard():
    if 'user' not in session: return redirect(url_for('login'))
    return render_template_string(TELA_DASHBOARD, user=session['user'], role=session['role'])

@app.route('/logout')
def logout():
    session.clear()
    return redirect(url_for('login'))

# -------- ROTA: ADMINISTRAÇÃO & BANCO DE DADOS --------
@app.route('/admin', methods=['GET', 'POST'])
def admin():
    if 'user' not in session or session.get('role') != 'admin':
        return redirect(url_for('dashboard'))

    conn = conectar_db(); c = conn.cursor(); ph = get_ph(conn)

    if request.method == 'POST':
        acao = request.form.get('acao')
        if acao == 'adicionar':
            u = request.form.get('username'); p = request.form.get('password'); r = request.form.get('role')
            try:
                c.execute(f"INSERT INTO users (username, password, role, first_login) VALUES ({ph}, {ph}, {ph}, 1)", (u, p, r))
                conn.commit()
                flash(f"Usuário '{u}' criado com sucesso!", "success")
            except Exception:
                conn.rollback(); flash("Erro: Nome de usuário já existe!", "danger")
        elif acao == 'eliminar':
            uid = request.form.get('user_id')
            c.execute(f"DELETE FROM users WHERE id={ph}", (uid,))
            conn.commit(); flash("Usuário excluído com sucesso!", "warning")

    c.execute("SELECT id, username, role, first_login FROM users ORDER BY id ASC")
    usuarios = c.fetchall()
    conn.close()
    return render_template_string(TELA_ADMIN, user=session['user'], role=session['role'], usuarios=usuarios)

@app.route('/admin/exportar')
def exportar_backup():
    if 'user' not in session or session.get('role') != 'admin': return redirect(url_for('dashboard'))
    try:
        conn = conectar_db()
        df_demands = pd.read_sql_query("SELECT * FROM demands", conn)
        df_avulsas = pd.read_sql_query("SELECT * FROM demandas_avulsas", conn)
        df_users = pd.read_sql_query("SELECT * FROM users", conn) 
        conn.close()
        
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df_demands.to_excel(writer, sheet_name='Rotina_PDFs', index=False)
            df_avulsas.to_excel(writer, sheet_name='Demandas_Avulsas', index=False)
            df_users.to_excel(writer, sheet_name='Usuarios', index=False)
        
        output.seek(0)
        nome_arquivo = f"Backup_COCAP_{datetime.now().strftime('%Y_%m_%d_%Hh%Mm')}.xlsx"
        return send_file(output, download_name=nome_arquivo, as_attachment=True)
    except Exception as e:
        flash(f"Erro ao gerar backup: {e}", "danger")
        return redirect(url_for('admin'))

@app.route('/admin/importar', methods=['POST'])
def importar_backup():
    if 'user' not in session or session.get('role') != 'admin': return redirect(url_for('dashboard'))
    
    file = request.files.get('file_backup')
    if not file:
        flash("Nenhum arquivo selecionado.", "danger")
        return redirect(url_for('admin'))
        
    filename = file.filename.lower()
    if not (filename.endswith('.xlsx') or filename.endswith('.db')):
        flash("Selecione um arquivo Excel (.xlsx) ou Banco SQLite (.db) válido.", "danger")
        return redirect(url_for('admin'))
        
    try:
        conn = conectar_db(); c = conn.cursor()
        is_postgres = "psycopg2" in str(type(conn))
        
        if filename.endswith('.xlsx'):
            xls = pd.ExcelFile(file)
            sheets = xls.sheet_names
            mapeamento = {'Rotina_PDFs': 'demands', 'Demandas_Avulsas': 'demandas_avulsas', 'Usuarios': 'users'}
            
            for aba, tabela in mapeamento.items():
                if aba in sheets:
                    df = pd.read_excel(xls, sheet_name=aba)
                    c.execute(f"DELETE FROM {tabela}") 
                    if not df.empty:
                        df = df.where(pd.notnull(df), None)
                        cols = ", ".join(df.columns)
                        placeholders = ", ".join(["%s" if is_postgres else "?"] * len(df.columns))
                        q = f"INSERT INTO {tabela} ({cols}) VALUES ({placeholders})"
                        c.executemany(q, df.values.tolist())
                        if is_postgres:
                            try: c.execute(f"SELECT setval('{tabela}_id_seq', COALESCE((SELECT MAX(id)+1 FROM {tabela}), 1), false)")
                            except: pass

        elif filename.endswith('.db'):
            import sqlite3
            temp_db = "temp_migration.db"
            file.save(temp_db) 
            sqlite_conn = sqlite3.connect(temp_db)
            
            tabelas_db = ['demands', 'demandas_avulsas', 'users']
            for tabela in tabelas_db:
                try:
                    df = pd.read_sql_query(f"SELECT * FROM {tabela}", sqlite_conn)
                    c.execute(f"DELETE FROM {tabela}") 
                    if not df.empty:
                        df = df.where(pd.notnull(df), None)
                        cols = ", ".join(df.columns)
                        placeholders = ", ".join(["%s" if is_postgres else "?"] * len(df.columns))
                        q = f"INSERT INTO {tabela} ({cols}) VALUES ({placeholders})"
                        c.executemany(q, df.values.tolist())
                        if is_postgres:
                            try: c.execute(f"SELECT setval('{tabela}_id_seq', COALESCE((SELECT MAX(id)+1 FROM {tabela}), 1), false)")
                            except: pass
                except Exception as e_tab:
                    print(f"Erro ao migrar tabela {tabela}: {e_tab}")
                    
            sqlite_conn.close()
            os.remove(temp_db)
            
        conn.commit(); conn.close()
        session.clear()
        flash("Banco de Dados importado com sucesso! Os dados antigos foram migrados. Faça login novamente.", "success")
        return redirect(url_for('login'))
        
    except Exception as e:
        flash(f"Falha ao restaurar banco. Formato incorreto ou arquivo corrompido. Erro: {e}", "danger")
        return redirect(url_for('admin'))

# -------- ROTAS DO MÓDULO DEMANDAS --------
@app.route('/demandas')
def hub_demandas():
    if 'user' not in session: return redirect(url_for('login'))
    return render_template_string(TELA_HUB_DEMANDAS)

@app.route('/demandas/avulsas', methods=['GET', 'POST'])
def demandas_avulsas():
    if 'user' not in session: return redirect(url_for('login'))
    conn = conectar_db(); c = conn.cursor(); ph = get_ph(conn)
    
    if request.method == 'POST':
        acao = request.form.get('acao')
        if acao == 'nova':
            td = request.form.get('tipo'); assunto = request.form.get('assunto')
            sol = request.form.get('solicitante'); prazo = request.form.get('prazo')
            dt_ent = datetime.now().strftime("%d/%m/%Y %H:%M")
            c.execute(f"INSERT INTO demandas_avulsas (tipo_demanda, assunto, solicitante, data_entrada, prazo, status, responsavel, observacao) VALUES ({ph}, {ph}, {ph}, {ph}, {ph}, 'Pendente', 'Nenhum', '')", (td, assunto, sol, dt_ent, prazo))
            flash("Demanda Avulsa criada com sucesso!", "success")
        elif acao == 'assumir':
            d_id = request.form.get('id')
            c.execute(f"UPDATE demandas_avulsas SET status='Em Análise', responsavel={ph} WHERE id={ph}", (session['user'], d_id))
        elif acao == 'concluir':
            d_id = request.form.get('id'); dt_c = datetime.now().strftime("%d/%m/%Y")
            c.execute(f"UPDATE demandas_avulsas SET status='Concluído', data_conclusao={ph} WHERE id={ph}", (dt_c, d_id))
        conn.commit()

    c.execute("SELECT * FROM demandas_avulsas ORDER BY id DESC")
    demandas = c.fetchall()
    conn.close()
    return render_template_string(TELA_AVULSAS, demandas=demandas, tipos=TIPOS_AVULSA, sols=SOLICITANTES)

@app.route('/demandas/rotinas')
def demandas_rotinas():
    if 'user' not in session: return redirect(url_for('login'))
    conn = conectar_db(); c = conn.cursor()
    c.execute("SELECT * FROM demands ORDER BY id DESC")
    rotinas = c.fetchall()
    conn.close()
    return render_template_string(TELA_ROTINAS, rotinas=rotinas)

@app.route('/demandas/rotinas/upload', methods=['POST'])
def upload_rotinas():
    if 'user' not in session: return redirect(url_for('login'))
    arquivos = request.files.getlist('pdfs')
    if not arquivos or arquivos[0].filename == '':
        flash("Nenhum arquivo selecionado.", "warning")
        return redirect(url_for('demandas_rotinas'))

    kw = {"REAJUSTE": "Reajuste", "NOVO CONTRATO": "Novo Contrato", "INCLUSÃO": "Inclusão", "EXCLUSÃO": "Exclusão", "AJUSTE": "Ajuste", "EXTENSÃO": "Extensão", "DESCREDENCIAMENTO": "Descredenciamento"}
    valid_ufs = {'AC', 'AL', 'AM', 'AP', 'BA', 'CE', 'DF', 'ES', 'GO', 'MA', 'MT', 'MS', 'MG', 'PA', 'PB', 'PR', 'PE', 'PI', 'RJ', 'RN', 'RS', 'RO', 'RR', 'SC', 'SP', 'SE', 'TO'}
    
    conn = conectar_db(); c = conn.cursor(); ph = get_ph(conn)
    count = 0
    
    for file in arquivos:
        f_name = file.filename
        c.execute(f"SELECT id FROM demands WHERE filename={ph}", (f_name,))
        if c.fetchone(): continue
        
        fu = f_name.upper()
        dem = "Não Identificada"
        for k, v in kw.items():
            if k in fu: dem = v; break
                
        f_clean = re.sub(r'\.pdf$', '', fu, flags=re.IGNORECASE)
        docs_brutos = re.findall(r'\d{2}[.\s_-]*\d{3}[.\s_-]*\d{3}[.\s/-]*\d{4}[-\s_]*\d{2}|\d{3}[.\s_-]*\d{3}[.\s_-]*\d{3}[-\s_]*\d{2}|\b\d{14}\b|\b\d{11}\b', f_clean)
        cnpj_p = "Não encontrado"
        
        for c_ in docs_brutos:
            cn = re.sub(r'\D', '', c_)
            if len(cn) == 14: cnpj_p = f"{cn[:2]}.{cn[2:5]}.{cn[5:8]}/{cn[8:12]}-{cn[12:]}"; f_clean = f_clean.replace(c_, ''); break
            elif len(cn) == 11: cnpj_p = f"{cn[:3]}.{cn[3:6]}.{cn[6:9]}-{cn[9:]}"; f_clean = f_clean.replace(c_, ''); break

        mun, uf, prest = "-", "-", "Não encontrado"
        for k in kw.keys(): f_clean = re.sub(k, '', f_clean, flags=re.IGNORECASE)
        parts = [p.strip() for p in f_clean.split('-') if p.strip()]
        
        if len(parts) >= 2:
            p1 = parts[-1].upper(); p2 = parts[-2].upper() if len(parts) > 1 else ""
            if p1 in valid_ufs: uf = p1; mun = p2 if len(parts) >= 3 else "-"; prest = " ".join(parts[:-2]).strip() if len(parts) >= 3 else p2
            elif p2 in valid_ufs: uf = p2; mun = p1; prest = " ".join(parts[:-2]).strip()
            else:
                m_uf_end = re.search(r'\b([A-Z]{2})\b$', p1)
                if m_uf_end and m_uf_end.group(1) in valid_ufs: uf = m_uf_end.group(1); mun = re.sub(r'\b' + uf + r'\b$', '', p1).strip(); prest = " ".join(parts[:-1]).strip()
                else: prest = " ".join(parts).strip()
        elif len(parts) == 1:
            m_uf_end = re.search(r'\b([A-Z]{2})\b$', parts[0].upper())
            if m_uf_end and m_uf_end.group(1) in valid_ufs: uf = m_uf_end.group(1); prest = re.sub(r'\b' + uf + r'\b$', '', parts[0]).strip()
            else: prest = parts[0]
            
        prest = re.sub(r'^[-_\s]+|[-_\s]+$', '', prest); prest = re.sub(r'\s+', ' ', prest).strip()
        if not prest or len(prest) <= 3: prest = "Não encontrado"
        
        if prest == "Não encontrado" or uf == "-" or cnpj_p == "Não encontrado":
            try:
                reader = PyPDF2.PdfReader(file.stream)
                full_txt = ""
                for page in reader.pages: full_txt += (page.extract_text() or "") + " "
                
                if cnpj_p == "Não encontrado":
                    docs_text = re.findall(r'\d{2}[.\s]*\d{3}[.\s]*\d{3}[/\s]*\d{4}[-\s]*\d{2}|\b\d{14}\b', full_txt)
                    for c_ in docs_text:
                        cn = re.sub(r'\D', '', c_)
                        if len(cn) == 14: cnpj_p = f"{cn[:2]}.{cn[2:5]}.{cn[5:8]}/{cn[8:12]}-{cn[12:]}"; break
            except: pass

        c.execute(f"INSERT INTO demands (filename, filepath, type, status, assigned_to, prestador, cnpj, municipio, uf, data_entrada) VALUES ({ph}, 'WEB_UPLOAD', {ph}, 'Pendente', 'Nenhum', {ph}, {ph}, {ph}, {ph}, {ph})", 
                     (f_name, dem, prest, cnpj_p, mun, uf, datetime.now().strftime("%d/%m/%Y %H:%M")))
        count += 1
        
    conn.commit(); conn.close()
    flash(f"{count} PDFs processados com sucesso!", "success")
    return redirect(url_for('demandas_rotinas'))

@app.route('/demandas/rotinas/acao', methods=['POST'])
def acao_rotinas():
    d_id = request.form.get('id')
    acao = request.form.get('acao')
    conn = conectar_db(); c = conn.cursor(); ph = get_ph(conn)
    
    if acao == 'assumir':
        c.execute(f"UPDATE demands SET status='Em Análise', assigned_to={ph} WHERE id={ph}", (session['user'], d_id))
    elif acao == 'finalizar':
        dt_c = datetime.now().strftime("%d/%m/%Y %H:%M")
        c.execute(f"UPDATE demands SET status='Finalizada', data_finalizacao={ph} WHERE id={ph}", (dt_c, d_id))
        
    conn.commit(); conn.close()
    return redirect(url_for('demandas_rotinas'))

@app.route('/demandas/dashboard_graficos')
def dashboard_graficos():
    if 'user' not in session: return redirect(url_for('login'))
    conn = conectar_db(); c = conn.cursor()
    
    c.execute("SELECT COUNT(*) FROM demands")
    t_rotina = c.fetchone()[0]
    c.execute("SELECT COUNT(*) FROM demandas_avulsas")
    t_avulsa = c.fetchone()[0]
    total = t_rotina + t_avulsa
    
    c.execute("SELECT COUNT(*) FROM demands WHERE status IN ('Pendente', 'Em Análise')")
    p_rot = c.fetchone()[0]
    c.execute("SELECT COUNT(*) FROM demandas_avulsas WHERE status IN ('Pendente', 'Em Análise')")
    p_av = c.fetchone()[0]
    pend = p_rot + p_av
    
    c.execute("SELECT COUNT(*) FROM demands WHERE status='Finalizada'")
    c_rot = c.fetchone()[0]
    c.execute("SELECT COUNT(*) FROM demandas_avulsas WHERE status='Concluído'")
    c_av = c.fetchone()[0]
    conc = c_rot + c_av
    
    # Gráfico Produtividade
    fig1, ax1 = plt.subplots(figsize=(6,4))
    c.execute("SELECT assigned_to, COUNT(*) FROM demands WHERE status='Finalizada' AND assigned_to != 'Nenhum' GROUP BY assigned_to")
    dados_prod = c.fetchall()
    if dados_prod:
        ax1.bar([r[0] for r in dados_prod], [r[1] for r in dados_prod], color='#5cb85c')
        ax1.set_xticklabels([r[0] for r in dados_prod], rotation=45, ha='right')
    buf1 = BytesIO()
    fig1.tight_layout(); fig1.savefig(buf1, format="png"); buf1.seek(0)
    chart_prod = base64.b64encode(buf1.read()).decode('utf-8')
    plt.close(fig1)

    # Gráfico Tipos
    fig2, ax2 = plt.subplots(figsize=(6,4))
    c.execute("SELECT type, COUNT(*) FROM demands GROUP BY type")
    dados_tipo = c.fetchall()
    if dados_tipo:
        ax2.bar([r[0] for r in dados_tipo], [r[1] for r in dados_tipo], color='#008CBA')
        ax2.set_xticklabels([r[0] for r in dados_tipo], rotation=45, ha='right')
    buf2 = BytesIO()
    fig2.tight_layout(); fig2.savefig(buf2, format="png"); buf2.seek(0)
    chart_tipo = base64.b64encode(buf2.read()).decode('utf-8')
    plt.close(fig2)

    conn.close()
    return render_template_string(TELA_DASHBOARD_GRAFICOS, d_total=total, d_pend=pend, d_ana=0, d_conc=conc, chart_prod=chart_prod, chart_tipo=chart_tipo)

@app.route('/pentefino')
def pentefino():
    if 'user' not in session: return redirect(url_for('login'))
    return render_template_string(TELA_CONSTRUCAO)

if __name__ == '__main__':
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)
