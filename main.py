from flask import Flask, request, redirect, url_for, session, flash, render_template_string
from config import setup_db, conectar_db
import os

app = Flask(__name__)
app.secret_key = 'chave_super_secreta_da_postal_saude'

# Garante que o banco está criado quando o app liga
setup_db()

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
        <div class="col-md-5">
            <div class="card shadow-sm border-0 h-100">
                <div class="card-body text-center p-5">
                    <h1 class="card-title text-info mb-3">📊</h1>
                    <h4 class="card-title mb-3">Controle de Demandas</h4>
                    <p class="card-text text-muted mb-4">Gerencie Rotinas, PDFs, Avulsas e acesse o Dashboard Gerencial.</p>
                    <button class="btn btn-info text-white w-100 fw-bold">Acessar Módulo</button>
                </div>
            </div>
        </div>

        <div class="col-md-5">
            <div class="card shadow-sm border-0 h-100">
                <div class="card-body text-center p-5">
                    <h1 class="card-title text-warning mb-3">🏥</h1>
                    <h4 class="card-title mb-3">Pente Fino (RN 665)</h4>
                    <p class="card-text text-muted mb-4">Substituição de Prestadores, cruzamento regional e busca sniper.</p>
                    <button class="btn btn-warning text-white w-100 fw-bold">Acessar Módulo</button>
                </div>
            </div>
        </div>
    </div>
</div>
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
            flash("Erro de conexão com o banco de dados da nuvem.", "danger")
            return render_template_string(TELA_LOGIN)
            
        c = conn.cursor()
        c.execute("SELECT id, role, first_login FROM users WHERE username=%s AND password=%s", (user, password))
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
    if 'user' not in session:
        return redirect(url_for('login'))
    
    return render_template_string(TELA_DASHBOARD, user=session['user'], role=session['role'])

@app.route('/logout')
def logout():
    session.clear()
    return redirect(url_for('login'))

if __name__ == '__main__':
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)
