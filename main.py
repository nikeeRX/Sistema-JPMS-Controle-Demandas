from flask import Flask, request, redirect, url_for, session, flash, render_template_string
from config import setup_db, conectar_db
import os

app = Flask(__name__)
app.secret_key = 'chave_super_secreta_da_postal_saude'

# Garante que o banco de dados está criado e atualizado quando o servidor liga
setup_db()

# Função auxiliar para manter compatibilidade entre Postgres (Nuvem) e SQLite (Local)
def get_ph(conn):
    return "%s" if "psycopg2" in str(type(conn)) else "?"

# ==========================================
# CÓDIGOS HTML/CSS EMBUTIDOS (FULL PYTHON)
# ==========================================

TELA_LOGIN = """
<!DOCTYPE html>
<html lang="pt-BR">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Login - Postal Saúde COCAP</title>
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
            <input type="text" name="username" class="form-control" placeholder="Digite seu usuário" required>
        </div>
        <div class="mb-4">
            <label class="form-label font-weight-bold">Senha</label>
            <input type="password" name="password" class="form-control" placeholder="Digite sua senha" required>
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

<div class="container mt-5">
    <div class="text-center mb-5">
        <h1 style="color: #004b87; font-weight: 800;">PORTAL CENTRAL DE OPERAÇÕES</h1>
        <p class="text-muted">Selecione o módulo desejado abaixo</p>
    </div>

    <div class="row justify-content-center gap-4">
        <!-- Módulo Demandas -->
        <div class="col-md-5">
            <div class="card shadow-sm border-0 h-100">
                <div class="card-body text-center p-5">
                    <h1 class="card-title text-info mb-3">📊</h1>
                    <h4 class="card-title mb-3">Controle de Demandas</h4>
                    <p class="card-text text-muted mb-4">Gerencie Rotinas, PDFs, Avulsas e acesse o Dashboard Gerencial.</p>
                    <a href="/demandas" class="btn btn-info text-white w-100 fw-bold">Acessar Módulo</a>
                </div>
            </div>
        </div>

        <!-- Módulo Pente Fino -->
        <div class="col-md-5">
            <div class="card shadow-sm border-0 h-100">
                <div class="card-body text-center p-5">
                    <h1 class="card-title text-warning mb-3">🏥</h1>
                    <h4 class="card-title mb-3">Pente Fino (RN 665)</h4>
                    <p class="card-text text-muted mb-4">Substituição de Prestadores, cruzamento regional e busca sniper.</p>
                    <a href="/pentefino" class="btn btn-warning text-white w-100 fw-bold">Acessar Módulo</a>
                </div>
            </div>
        </div>
    </div>
    
    <!-- Módulo Admin (Apenas para Administradores) -->
    {% if role == 'admin' %}
    <div class="row justify-content-center mt-4">
        <div class="col-md-5">
            <div class="card shadow-sm border-0 h-100" style="background-color: #e9ecef;">
                <div class="card-body text-center p-4">
                    <h2 class="card-title text-secondary mb-2">⚙️</h2>
                    <h5 class="card-title mb-3">Painel Administrativo</h5>
                    <a href="/admin" class="btn btn-secondary w-100 fw-bold">Gerir Utilizadores</a>
                </div>
            </div>
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
    <span class="navbar-text text-white fw-bold">Gestão de Utilizadores</span>
  </div>
</nav>

<div class="container mt-5">
    {% with messages = get_flashed_messages(with_categories=true) %}
      {% if messages %}
        {% for category, message in messages %}
          <div class="alert alert-{{ category }}">{{ message }}</div>
        {% endfor %}
      {% endif %}
    {% endwith %}

    <div class="card shadow-sm mb-4">
        <div class="card-body">
            <h5 class="card-title mb-4">Criar Novo Utilizador</h5>
            <form method="POST" action="/admin">
                <input type="hidden" name="acao" value="adicionar">
                <div class="row g-3 align-items-end">
                    <div class="col-md-4">
                        <label class="form-label">Usuário</label>
                        <input type="text" name="username" class="form-control" required>
                    </div>
                    <div class="col-md-3">
                        <label class="form-label">Senha</label>
                        <input type="password" name="password" class="form-control" required>
                    </div>
                    <div class="col-md-3">
                        <label class="form-label">Nível</label>
                        <select name="role" class="form-select">
                            <option value="user">User</option>
                            <option value="admin">Admin</option>
                        </select>
                    </div>
                    <div class="col-md-2">
                        <button type="submit" class="btn btn-success w-100 fw-bold">Criar</button>
                    </div>
                </div>
            </form>
        </div>
    </div>

    <div class="card shadow-sm">
        <div class="card-body">
            <h5 class="card-title mb-4">Lista de Utilizadores</h5>
            <div class="table-responsive">
                <table class="table table-hover table-bordered align-middle">
                    <thead class="table-dark">
                        <tr>
                            <th>ID</th>
                            <th>Usuário</th>
                            <th>Nível</th>
                            <th>Status 1º Acesso</th>
                            <th class="text-center">Ações</th>
                        </tr>
                    </thead>
                    <tbody>
                        {% for u in usuarios %}
                        <tr>
                            <td>{{ u[0] }}</td>
                            <td>{{ u[1] }}</td>
                            <td><span class="badge bg-{{ 'primary' if u[2] == 'admin' else 'secondary' }}">{{ u[2] | upper }}</span></td>
                            <td>{{ 'Pendente' if u[3] == 1 else 'OK' }}</td>
                            <td class="text-center">
                                {% if u[1] != 'admin' and u[1] != user %}
                                <form method="POST" action="/admin" style="display:inline;">
                                    <input type="hidden" name="acao" value="eliminar">
                                    <input type="hidden" name="user_id" value="{{ u[0] }}">
                                    <button type="submit" class="btn btn-danger btn-sm" onclick="return confirm('Tem a certeza que deseja eliminar o utilizador {{ u[1] }}?');">Eliminar</button>
                                </form>
                                {% else %}
                                <span class="text-muted" style="font-size: 0.85em;">Protegido</span>
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
</body>
</html>
"""

TELA_CONSTRUCAO = """
<!DOCTYPE html>
<html lang="pt-BR">
<head>
    <meta charset="UTF-8">
    <title>{{ titulo }}</title>
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet">
</head>
<body style="background-color: #f0f2f5; display: flex; flex-direction: column; align-items: center; justify-content: center; height: 100vh;">
    <h1 style="font-size: 4rem;">🚧</h1>
    <h2 class="mt-3 text-secondary">{{ titulo }}</h2>
    <p class="text-muted">A preparar o motor de processamento Web de ficheiros...</p>
    <a href="/dashboard" class="btn btn-primary mt-4 fw-bold">⬅ Voltar ao Portal</a>
</body>
</html>
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
            flash("Erro de ligação com a base de dados da nuvem.", "danger")
            return render_template_string(TELA_LOGIN)
            
        c = conn.cursor()
        ph = get_ph(conn)
        c.execute(f"SELECT id, role, first_login FROM users WHERE username={ph} AND password={ph}", (user, password))
        res = c.fetchone()
        conn.close()
        
        if res:
            session['user'] = user
            session['role'] = res[1]
            return redirect(url_for('dashboard'))
        else:
            flash("Utilizador ou senha incorretos!", "danger")
            
    return render_template_string(TELA_LOGIN)

@app.route('/dashboard')
def dashboard():
    if 'user' not in session:
        return redirect(url_for('login'))
    
    return render_template_string(TELA_DASHBOARD, user=session['user'], role=session['role'])

@app.route('/logout')
def logout():
    session.clear()
    return redirect(url_for('login'))

# ROTA: PAINEL DE ADMINISTRAÇÃO 
@app.route('/admin', methods=['GET', 'POST'])
def admin():
    if 'user' not in session or session.get('role') != 'admin':
        return redirect(url_for('dashboard'))

    conn = conectar_db()
    c = conn.cursor()
    ph = get_ph(conn)

    if request.method == 'POST':
        acao = request.form.get('acao')
        
        if acao == 'adicionar':
            u = request.form.get('username')
            p = request.form.get('password')
            r = request.form.get('role')
            try:
                c.execute(f"INSERT INTO users (username, password, role, first_login) VALUES ({ph}, {ph}, {ph}, 1)", (u, p, r))
                conn.commit()
                flash(f"Utilizador '{u}' criado com sucesso!", "success")
            except Exception as e:
                conn.rollback()
                flash("Erro: O nome de utilizador já existe!", "danger")
                
        elif acao == 'eliminar':
            uid = request.form.get('user_id')
            c.execute(f"DELETE FROM users WHERE id={ph}", (uid,))
            conn.commit()
            flash("Utilizador eliminado com sucesso!", "warning")

    # Atualiza a lista de utilizadores para mostrar no ecrã
    c.execute("SELECT id, username, role, first_login FROM users ORDER BY id ASC")
    usuarios = c.fetchall()
    conn.close()

    return render_template_string(TELA_ADMIN, user=session['user'], role=session['role'], usuarios=usuarios)

# ROTA: MÓDULO DEMANDAS (Estrutura)
@app.route('/demandas')
def demandas():
    if 'user' not in session: return redirect(url_for('login'))
    return render_template_string(TELA_CONSTRUCAO, titulo="Módulo: Controlo de Demandas")

# ROTA: MÓDULO PENTE FINO RN 665 (Estrutura)
@app.route('/pentefino')
def pentefino():
    if 'user' not in session: return redirect(url_for('login'))
    return render_template_string(TELA_CONSTRUCAO, titulo="Módulo: Pente Fino (RN 665)")

if __name__ == '__main__':
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)
