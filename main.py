from flask import Flask, request, redirect, url_for, session, flash, render_template_string, send_file
from config import setup_db, conectar_db
import os
import PyPDF2
import re
from datetime import datetime
import base64
from io import BytesIO
import pandas as pd
import unicodedata
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side

import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt

app = Flask(__name__)
app.secret_key = 'chave_super_secreta_da_postal_saude'

setup_db()

def get_ph(conn):
    return "%s" if "psycopg2" in str(type(conn)) else "?"

# ==========================================
# LISTAS GLOBAIS E CONSTANTES
# ==========================================
TIPOS_ROTINA = ["Reajuste", "Novo Contrato", "Inclusão", "Exclusão", "Ajuste", "Extensão", "Descredenciamento", "Não Identificada"]
TIPOS_AVULSA = ['Credenciamento', 'Descredenciamento', 'Análise de Vulnerabilidade', 'Parametrização', 'Resposta de E-mail', 'RPS', 'Lote de Erros', 'Atendimento de Chamado', 'Atualização Cadastral', 'Outro']
STATUS_AVULSA = ['Pendente', 'Recebida', 'Em Análise', 'Aguardando Retorno', 'Aguardando Área Demandante', 'Em andamento', 'Concluído', 'Cancelada']
SOLICITANTES = ['GEREL/COATE', 'GEREL/COCAD', 'UAR-AM', 'NAR-AC', 'NAR-RR', 'UAR-BA', 'NAR-AL', 'NAR-SE', 'UAR-CE', 'NAR-MA', 'NAR-PI', 'UAR-DF', 'NAR-GO', 'NAR-TO', 'UAR-MG', 'UAR-MS', 'NAR-MT', 'NAR-RO', 'UAR-PA', 'NAR-AP', 'UAR-PE', 'NAR-PB', 'NAR-RN', 'UAR-RJ', 'NAR-ES', 'UAR-RS', 'NAR-SC', 'UAR-SPI', 'UAR-SPM', 'NAR-PR', 'GEREM/COANS', 'GERIN', 'GEASA/COICS', 'GERED/COCAP']

IGNORADAS = {"DIARIAS", "PACOTE", "TAXAS E GASES", "MATERIAIS E OPME", "MEDICAMENTOS", "FORNECEDOR DE MEDICAMENTO ONCOLOGICO", "INSTRUMENTADOR CIRURGICO", "HOME CARE - ATENDIMENTO DOMICILIAR", "REMOCAO", "MEDICINA LEGAL E PERICIA MEDICA", "MEDICO HIPERBARISTA"}
DE_PARA = {"MEDICO HEMOTERAPEUTA": "HEMATOLOGIA E HEMOTERAPIA", "GENETICA MEDICA": "MEDICO GENETICISTA", "MEDICO CANCEROLOGISTA CIRURGICO": "CIRURGIA ONCOLOGICA", "MEDICO CANCEROLOGISTA PEDIATRICO": "CANCEROLOGIA", "MEDICO DE FAMILIA E COMUNIDADE": "MEDICINA DE FAMILIA E COMUNIDADE", "MEDICO GENERALISTA": "CLINICA MEDICA", "MEDICO PATOLOGISTA": "PATOLOGIA CLINICA/MEDICINA LABORATORIAL", "PATOLOGIA": "PATOLOGIA CLINICA/MEDICINA LABORATORIAL", "MEDICO NEUROFISIOLOGISTA": "NEUROFISIOLOGIA CLINICA", "NEUROPSICOLOGO": "PSICOLOGIA", "PSICOPEDAGOGO": "PSICOLOGIA", "PSICOMOTRICISTA": "FISIOTERAPIA", "MUSICOTERAPEUTA": "TERAPIA OCUPACIONAL", "ORTOPEDISTA": "ORTOPEDIA E TRAUMATOLOGIA", "BIOMEDICO": "PATOLOGIA CLINICA/MEDICINA LABORATORIAL", "FONOAUDIOLOGO EDUCACIONAL": "FONOAUDIOLOGIA", "FONOAUDIOLOGO EM AUDIOLOGIA": "FONOAUDIOLOGIA", "FONOAUDIOLOGO EM DISFAGIA": "FONOAUDIOLOGIA", "FONOAUDIOLOGO EM LINGUAGEM": "FONOAUDIOLOGIA", "FONOAUDIOLOGO EM MOTRICIDADE OROFACIAL": "FONOAUDIOLOGIA", "FONOAUDIOLOGO EM VOZ": "FONOAUDIOLOGIA", "FISIOTERAPEUTA NEUROFUNCIONAL": "FISIOTERAPIA", "FISIOTERAPEUTA OSTEOPATA": "FISIOTERAPIA", "FISIOTERAPEUTA RESPIRATORIA": "FISIOTERAPIA", "FISIOTERAPEUTA TRAUMATO-ORTOPEDICA FUNCIONAL": "FISIOTERAPIA", "ENFERMEIRO DA ESTRATEGIA DE SAUDE DA FAMILIA": "ENFERMEIRO", "CIRURGIAO DENTISTA - TRAUMATOLOGISTA BUCOMAXILOFAC": "CIRURGIA E TRAUMATOLOGIA BUCO-MAXILO-FACIAL", "CLINICA GERAL - ODONTOLOGIA": "ODONTOLOGIA", "CIRURGIAO DENTISTA - DENTISTICA": "ODONTOLOGIA", "CIRURGIAO DENTISTA - DISFUNCAO TEMPOROMANDIBULAR E": "ODONTOLOGIA", "DENTISTICA RESTAURADORA": "ODONTOLOGIA", "DISFUNCAO TEMPOROMANDIBULAR E DOR OROFACIAL": "ODONTOLOGIA", "ENDODONTIA": "ODONTOLOGIA", "ESTOMATOLOGIA": "ODONTOLOGIA", "ODONTOLOGIA - AUDITORIA INICIAL E FINAL": "ODONTOLOGIA", "ODONTOLOGIA DO TRABALHO": "ODONTOLOGIA", "ODONTOLOGIA P/ PACIENTE COM NECESSIDADE ESPECIAL": "ODONTOLOGIA", "ODONTOPEDIATRIA": "ODONTOLOGIA", "PERIODONTIA": "ODONTOLOGIA", "PROTESE DENTARIA": "ODONTOLOGIA", "RADIOLOGIA ODONTOLOGICA E IMAGINOLOGIA - RX ODONTO": "RADIOLOGIA E DIAGNOSTICO POR IMAGEM", "NEUROLIGIA": "NEUROLOGIA"}

# ==========================================
# CÓDIGOS HTML EMBUTIDOS
# ==========================================

TELA_LOGIN = """<!DOCTYPE html><html lang="pt-BR"><head><meta charset="UTF-8"><title>Login - COCAP</title><link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet"></head>
<body style="background-color: #f4f7f6; display: flex; align-items: center; justify-content: center; height: 100vh;">
<div style="border-radius: 15px; box-shadow: 0 10px 30px rgba(0,0,0,0.1); padding: 40px; background: white; width: 100%; max-width: 400px;">
    <div class="text-center mb-4"><h3 style="color: #004b87; font-weight: 800;">POSTAL SAÚDE</h3><p class="text-muted">Central de Operações (COCAP)</p></div>
    {% with messages = get_flashed_messages(with_categories=true) %}{% if messages %}{% for category, message in messages %}<div class="alert alert-{{ category }}">{{ message }}</div>{% endfor %}{% endif %}{% endwith %}
    <form method="POST" action="/">
        <div class="mb-3"><label class="form-label fw-bold">Utilizador</label><input type="text" name="username" class="form-control" required></div>
        <div class="mb-4"><label class="form-label fw-bold">Senha</label><input type="password" name="password" class="form-control" required></div>
        <button type="submit" class="btn w-100 py-2" style="background-color: #004b87; color: white; font-weight: bold;">Entrar no Sistema</button>
    </form>
</div></body></html>"""

TELA_DASHBOARD = """<!DOCTYPE html><html lang="pt-BR"><head><meta charset="UTF-8"><title>Portal COCAP</title><link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet"></head>
<body style="background-color: #f0f2f5;">
<nav class="navbar navbar-expand-lg navbar-dark" style="background-color: #004b87;">
  <div class="container-fluid"><a class="navbar-brand fw-bold" href="#">COCAP Web</a>
    <div class="d-flex text-white align-items-center"><span class="me-3">Olá, <strong>{{ user }}</strong> ({{ role }})</span><a href="/logout" class="btn btn-danger btn-sm">Sair</a></div>
  </div>
</nav>
<div class="container mt-5 text-center">
    <h1 style="color: #004b87; font-weight: 800; margin-bottom: 2rem;">PORTAL CENTRAL DE OPERAÇÕES</h1>
    <div class="row justify-content-center gap-4">
        <div class="col-md-5">
            <div class="card shadow-sm h-100 p-4 border-0">
                <h1 class="text-info mb-3" style="font-size: 3rem;">📊</h1><h4>Controlo de Demandas</h4><p class="text-muted mb-4">Gira Rotinas, PDFs, Avulsas e aceda ao Dashboard.</p>
                <a href="/demandas" class="btn btn-info text-white w-100 fw-bold py-2">Aceder ao Módulo</a>
            </div>
        </div>
        <div class="col-md-5">
            <div class="card shadow-sm h-100 p-4 border-0">
                <h1 class="text-warning mb-3" style="font-size: 3rem;">🏥</h1><h4>Pente Fino (RN 665)</h4><p class="text-muted mb-4">Substituição de Prestadores e análise de rede.</p>
                <a href="/pentefino" class="btn btn-warning text-white w-100 fw-bold py-2">Aceder ao Módulo</a>
            </div>
        </div>
    </div>
    {% if role == 'admin' %}<div class="row justify-content-center mt-4"><div class="col-md-5"><a href="/admin" class="btn btn-secondary w-100 fw-bold py-3 shadow-sm border-0">⚙️ Painel Administrativo (BD & Utilizadores)</a></div></div>{% endif %}
</div></body></html>"""

TELA_ADMIN = """<!DOCTYPE html><html lang="pt-BR"><head><meta charset="UTF-8"><title>Painel Administrativo</title><link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet"></head>
<body style="background-color: #f0f2f5;">
<nav class="navbar navbar-dark" style="background-color: #333;"><div class="container-fluid"><a href="/dashboard" class="btn btn-outline-light btn-sm">⬅ Voltar ao Portal</a><span class="navbar-text text-white fw-bold">Painel Administrativo</span></div></nav>
<div class="container mt-4">
    {% with messages = get_flashed_messages(with_categories=true) %}{% if messages %}{% for category, message in messages %}<div class="alert alert-{{ category }}">{{ message }}</div>{% endfor %}{% endif %}{% endwith %}
    <div class="row">
        <div class="col-md-12 mb-4">
            <div class="card shadow-sm border-0">
                <div class="card-header bg-dark text-white fw-bold">💾 Gestão da Base de Dados</div>
                <div class="card-body">
                    <div class="row text-center">
                        <div class="col-md-6 border-end">
                            <h5 class="text-success">Exportar Backup</h5><a href="/admin/exportar" class="btn btn-success fw-bold w-75 mt-2">⬇ Descarregar Backup (Excel)</a>
                        </div>
                        <div class="col-md-6">
                            <h5 class="text-danger">Restaurar ou Migrar Base de Dados</h5>
                            <form action="/admin/importar" method="POST" enctype="multipart/form-data" class="d-flex flex-column align-items-center mt-2">
                                <input type="file" name="file_backup" class="form-control form-control-sm w-75 mb-2" accept=".xlsx,.db" required>
                                <button type="submit" class="btn btn-danger fw-bold w-75" onclick="return confirm('ATENÇÃO: Todas as informações atuais da Nuvem serão APAGADAS e substituídas por este ficheiro. Deseja continuar?');">⬆ Importar Ficheiro</button>
                            </form>
                        </div>
                    </div>
                </div>
            </div>
        </div>
        <div class="col-md-12">
            <div class="card shadow-sm mb-4 border-0">
                <div class="card-body"><h5 class="card-title mb-3">Criar Novo Utilizador</h5>
                    <form method="POST" action="/admin"><input type="hidden" name="acao" value="adicionar">
                        <div class="row g-2 align-items-end">
                            <div class="col-md-4"><label class="form-label small fw-bold">Utilizador</label><input type="text" name="username" class="form-control" required></div>
                            <div class="col-md-3"><label class="form-label small fw-bold">Senha</label><input type="password" name="password" class="form-control" required></div>
                            <div class="col-md-3"><label class="form-label small fw-bold">Nível</label><select name="role" class="form-select"><option value="user">User</option><option value="admin">Admin</option></select></div>
                            <div class="col-md-2"><button type="submit" class="btn btn-primary w-100 fw-bold">Gravar</button></div>
                        </div>
                    </form>
                </div>
            </div>
            <div class="card shadow-sm border-0">
                <div class="card-body">
                    <table class="table table-hover align-middle"><thead class="table-light"><tr><th>ID</th><th>Utilizador</th><th>Nível</th><th>Ações</th></tr></thead>
                        <tbody>{% for u in usuarios %}<tr><td>{{ u[0] }}</td><td>{{ u[1] }}</td><td><span class="badge bg-{{ 'dark' if u[2] == 'admin' else 'secondary' }}">{{ u[2] | upper }}</span></td><td>
                            {% if u[1] != 'admin' and u[1] != user %}<form method="POST" action="/admin" style="display:inline;"><input type="hidden" name="acao" value="eliminar"><input type="hidden" name="user_id" value="{{ u[0] }}"><button type="submit" class="btn btn-outline-danger btn-sm" onclick="return confirm('Eliminar o utilizador {{ u[1] }}?');">Excluir</button></form>
                            {% else %}<span class="text-muted small">Sistema</span>{% endif %}</td></tr>{% endfor %}
                        </tbody>
                    </table>
                </div>
            </div>
        </div>
    </div>
</div></body></html>"""

TELA_HUB_DEMANDAS = """<!DOCTYPE html><html lang="pt-BR"><head><title>Módulo de Demandas</title><link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet"></head>
<body style="background-color: #f0f2f5;">
<nav class="navbar navbar-dark bg-dark"><div class="container-fluid"><a href="/dashboard" class="btn btn-outline-light btn-sm">⬅ Voltar ao Portal</a><span class="navbar-text text-white fw-bold">Gestão de Demandas</span></div></nav>
<div class="container mt-5 text-center"><div class="row justify-content-center">
    <div class="col-md-4 mb-4"><div class="card shadow-sm border-0 h-100 p-4"><h1 class="mb-3">📄</h1><a href="/demandas/rotinas" class="btn btn-primary w-100 py-3 fw-bold fs-6">Demandas de Rotina (PDFs)</a></div></div>
    <div class="col-md-4 mb-4"><div class="card shadow-sm border-0 h-100 p-4"><h1 class="mb-3">📝</h1><a href="/demandas/avulsas" class="btn btn-warning text-white w-100 py-3 fw-bold fs-6">Demandas Avulsas</a></div></div>
    <div class="col-md-4 mb-4"><div class="card shadow-sm border-0 h-100 p-4"><h1 class="mb-3">📊</h1><a href="/demandas/dashboard_graficos" class="btn btn-info text-white w-100 py-3 fw-bold fs-6">Dashboard Gerencial</a></div></div>
</div></div></body></html>"""

TELA_AVULSAS = """<!DOCTYPE html><html lang="pt-BR"><head><title>Demandas Avulsas</title><link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet"></head>
<body style="background-color: #f0f2f5;">
<nav class="navbar navbar-dark bg-dark"><div class="container-fluid"><a href="/demandas" class="btn btn-outline-light btn-sm">⬅ Voltar</a><span class="text-white fw-bold">Demandas Avulsas</span></div></nav>
<div class="container-fluid mt-4">
    {% with messages = get_flashed_messages(with_categories=true) %}{% if messages %}{% for category, message in messages %}<div class="alert alert-{{ category }}">{{ message }}</div>{% endfor %}{% endif %}{% endwith %}
    
    <div class="card shadow-sm mb-3 border-0 bg-light">
        <div class="card-body py-2">
            <form method="GET" action="/demandas/avulsas" class="row g-2 align-items-center">
                <div class="col-md-3"><select name="status" class="form-select form-select-sm"><option value="Todos">Status (Todos)</option>{% for s in status_list %}<option value="{{s}}">{{s}}</option>{% endfor %}</select></div>
                <div class="col-md-3"><select name="tipo" class="form-select form-select-sm"><option value="Todos">Tipo (Todos)</option>{% for t in tipos %}<option value="{{t}}">{{t}}</option>{% endfor %}</select></div>
                <div class="col-md-3"><select name="resp" class="form-select form-select-sm"><option value="Todos">Responsável (Todos)</option>{% for u in usuarios %}<option value="{{u}}">{{u}}</option>{% endfor %}</select></div>
                <div class="col-md-3"><button type="submit" class="btn btn-secondary btn-sm w-100 fw-bold">🔍 Filtrar Resultados</button></div>
            </form>
        </div>
    </div>

    <div class="card shadow-sm mb-4 border-0">
        <div class="card-body">
            <form method="POST" action="/demandas/avulsas"><input type="hidden" name="acao" value="nova">
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
    <div class="card shadow-sm border-0"><div class="card-body"><div class="table-responsive">
        <table class="table table-hover table-bordered align-middle text-center" style="font-size: 0.9em;">
            <thead class="table-dark"><tr><th>ID</th><th>Tipo</th><th>Assunto</th><th>Solicitante</th><th>Entrada</th><th>Prazo</th><th>Status</th><th>Resp.</th><th>Conclusão</th><th>Ações</th></tr></thead>
            <tbody>{% for d in demandas %}<tr>
                <td>{{ d[0] }}</td><td>{{ d[1] }}</td><td class="text-start">{{ d[2] }}</td><td>{{ d[3] }}</td><td>{{ d[4] }}</td><td>{{ d[5] }}</td>
                <td><span class="badge bg-{{ 'warning text-dark' if d[6]=='Pendente' else 'success' if d[6] in ('Concluído', 'Finalizada') else 'info' }}">{{ d[6] }}</span></td>
                <td>{{ d[7] }}</td><td>{{ d[8] }}</td>
                <td>
                    <form method="POST" action="/demandas/avulsas" style="display:inline;"><input type="hidden" name="id" value="{{ d[0] }}">
                        <button type="submit" name="acao" value="assumir" class="btn btn-warning btn-sm fw-bold" {% if d[6] == 'Concluído' %}disabled{% endif %}>Assumir</button>
                        <button type="submit" name="acao" value="concluir" class="btn btn-success btn-sm fw-bold" {% if d[6] == 'Concluído' %}disabled{% endif %}>✔</button>
                    </form>
                </td>
            </tr>{% endfor %}</tbody>
        </table>
    </div></div></div>
</div></body></html>"""

TELA_ROTINAS = """<!DOCTYPE html><html lang="pt-BR"><head><title>Rotinas (PDFs)</title><link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet"></head>
<body style="background-color: #f0f2f5;">
<nav class="navbar navbar-dark bg-dark"><div class="container-fluid"><a href="/demandas" class="btn btn-outline-light btn-sm">⬅ Voltar</a><span class="text-white fw-bold">Processamento de PDFs</span></div></nav>
<div class="container-fluid mt-4">
    {% with messages = get_flashed_messages(with_categories=true) %}{% if messages %}{% for category, message in messages %}<div class="alert alert-{{ category }}">{{ message }}</div>{% endfor %}{% endif %}{% endwith %}

    <div class="card shadow-sm mb-3 border-0 bg-light">
        <div class="card-body py-2">
            <form method="GET" action="/demandas/rotinas" class="row g-2 align-items-center">
                <div class="col-md-3"><select name="status" class="form-select form-select-sm"><option value="Todos">Status (Todos)</option><option value="Pendente">Pendente</option><option value="Em Análise">Em Análise</option><option value="Finalizada">Finalizada</option></select></div>
                <div class="col-md-3"><select name="tipo" class="form-select form-select-sm"><option value="Todos">Tipo (Todos)</option>{% for t in tipos %}<option value="{{t}}">{{t}}</option>{% endfor %}</select></div>
                <div class="col-md-3"><select name="resp" class="form-select form-select-sm"><option value="Todos">Responsável (Todos)</option>{% for u in usuarios %}<option value="{{u}}">{{u}}</option>{% endfor %}</select></div>
                <div class="col-md-3"><button type="submit" class="btn btn-secondary btn-sm w-100 fw-bold">🔍 Filtrar Resultados</button></div>
            </form>
        </div>
    </div>

    <div class="card shadow-sm mb-4 border-0">
        <div class="card-body text-center bg-light rounded">
            <form method="POST" action="/demandas/rotinas/upload" enctype="multipart/form-data" class="d-flex justify-content-center align-items-center gap-3">
                <span class="text-primary fw-bold">Importar PDFs:</span>
                <input type="file" name="pdfs" class="form-control w-25" multiple accept=".pdf" required>
                <button type="submit" class="btn btn-primary fw-bold px-4">Processar Ficheiros</button>
            </form>
        </div>
    </div>
    <div class="card shadow-sm border-0"><div class="card-body"><div class="table-responsive">
        <table class="table table-hover table-bordered align-middle text-center" style="font-size: 0.85em;">
            <thead class="table-dark"><tr><th>ID</th><th>Data Entrada</th><th>Prestador</th><th>CNPJ</th><th>Mun/UF</th><th>Demanda</th><th>Status</th><th>Resp.</th><th>Ações</th></tr></thead>
            <tbody>{% for d in rotinas %}<tr>
                <td>{{ d[0] }}</td><td>{{ d[9] }}</td><td class="text-start fw-bold">{{ d[5] }}</td><td>{{ d[6] }}</td><td>{{ d[7] }}/{{ d[8] }}</td><td>{{ d[2] }}</td>
                <td><span class="badge bg-{{ 'warning text-dark' if d[3]=='Pendente' else 'success' if d[3]=='Finalizada' else 'info' }}">{{ d[3] }}</span></td>
                <td>{{ d[4] }}</td>
                <td>
                    <form method="POST" action="/demandas/rotinas/acao" style="display:inline;"><input type="hidden" name="id" value="{{ d[0] }}">
                        <button type="submit" name="acao" value="assumir" class="btn btn-warning btn-sm fw-bold" {% if d[3] == 'Finalizada' %}disabled{% endif %}>Assumir</button>
                        <button type="submit" name="acao" value="finalizar" class="btn btn-success btn-sm fw-bold" {% if d[3] == 'Finalizada' %}disabled{% endif %}>Finalizar</button>
                    </form>
                </td>
            </tr>{% endfor %}</tbody>
        </table>
    </div></div></div>
</div></body></html>"""

TELA_DASHBOARD_GRAFICOS = """<!DOCTYPE html><html lang="pt-BR"><head><title>Dashboard Gerencial</title><link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet"></head>
<body style="background-color: #f0f2f5;">
<nav class="navbar navbar-dark bg-dark"><div class="container-fluid"><a href="/demandas" class="btn btn-outline-light btn-sm">⬅ Voltar</a><span class="text-white fw-bold">Dashboard Geral</span></div></nav>
<div class="container mt-4">
    <div class="row text-center mb-4">
        <div class="col-md-3"><div class="card shadow-sm border-0 bg-primary text-white p-3"><h5 class="mb-1">Total Demandas</h5><h2 class="fw-bold">{{ d_total }}</h2></div></div>
        <div class="col-md-3"><div class="card shadow-sm border-0 bg-warning text-dark p-3"><h5 class="mb-1">Pendentes</h5><h2 class="fw-bold">{{ d_pend }}</h2></div></div>
        <div class="col-md-3"><div class="card shadow-sm border-0 bg-info text-white p-3"><h5 class="mb-1">Em Análise</h5><h2 class="fw-bold">{{ d_ana }}</h2></div></div>
        <div class="col-md-3"><div class="card shadow-sm border-0 bg-success text-white p-3"><h5 class="mb-1">Concluídas</h5><h2 class="fw-bold">{{ d_conc }}</h2></div></div>
    </div>
    <div class="row">
        <div class="col-md-6 mb-4"><div class="card shadow-sm border-0 h-100"><div class="card-body text-center"><h5 class="fw-bold text-secondary mb-3">Produtividade por Colaborador</h5><img src="data:image/png;base64,{{ chart_prod }}" class="img-fluid rounded"></div></div></div>
        <div class="col-md-6 mb-4"><div class="card shadow-sm border-0 h-100"><div class="card-body text-center"><h5 class="fw-bold text-secondary mb-3">Distribuição por Tipo de Demanda</h5><img src="data:image/png;base64,{{ chart_tipo }}" class="img-fluid rounded"></div></div></div>
    </div>
</div></body></html>"""

TELA_HUB_PENTEFINO = """<!DOCTYPE html><html lang="pt-BR"><head><title>Pente Fino RN 665</title><link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet"></head>
<body style="background-color: #f0f2f5;">
<nav class="navbar navbar-dark bg-dark"><div class="container-fluid"><a href="/dashboard" class="btn btn-outline-light btn-sm">⬅ Voltar ao Portal</a><span class="navbar-text text-white fw-bold">Pente Fino (RN 665)</span></div></nav>
<div class="container mt-5 text-center">
    <div class="row justify-content-center">
        <div class="col-md-5 mb-4">
            <div class="card shadow-sm border-0 h-100 p-4"><h1 class="mb-3 text-warning" style="font-size: 3rem;">🌍</h1><h4 class="fw-bold text-secondary">Cruzamento Regional em Massa</h4><p class="text-muted">Análise de cobertura completa comparando a Postal com a Operadora Intermediária.</p><a href="/pentefino/regional" class="btn btn-warning text-white w-100 py-3 fw-bold fs-6">Aceder</a></div>
        </div>
        <div class="col-md-5 mb-4">
            <div class="card shadow-sm border-0 h-100 p-4"><h1 class="mb-3 text-danger" style="font-size: 3rem;">🎯</h1><h4 class="fw-bold text-secondary">Busca Individual (Sniper)</h4><p class="text-muted">Procure substitutos internamente usando Nome ou CNPJ.</p><a href="/pentefino/individual" class="btn btn-danger text-white w-100 py-3 fw-bold fs-6">Aceder</a></div>
        </div>
    </div>
</div></body></html>"""

# TELA COM JAVASCRIPT (SHEETJS) INJETADO PARA LER CABEÇALHOS ANTES DO UPLOAD
TELA_PENTEFINO_REGIONAL = """<!DOCTYPE html><html lang="pt-BR"><head><title>Cruzamento Regional</title><link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet">
<script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script></head>
<body style="background-color: #f0f2f5;">
<nav class="navbar navbar-dark bg-dark"><div class="container-fluid"><a href="/pentefino" class="btn btn-outline-light btn-sm">⬅ Voltar</a><span class="text-white fw-bold">Cruzamento Regional (Em Massa)</span></div></nav>
<div class="container mt-4">
    {% with messages = get_flashed_messages(with_categories=true) %}{% if messages %}{% for category, message in messages %}<div class="alert alert-{{ category }}">{{ message }}</div>{% endfor %}{% endif %}{% endwith %}
    <div class="card shadow-sm border-0">
        <div class="card-header bg-warning text-white fw-bold">🌍 Mapeamento e Upload das Bases</div>
        <div class="card-body">
            <form method="POST" action="/pentefino/regional" enctype="multipart/form-data">
                <div class="row mb-3 bg-light p-3 rounded">
                    <div class="col-md-6 border-end">
                        <label class="form-label fw-bold text-primary">1. Base Postal Saúde</label>
                        <input type="file" id="f_postal" name="f_postal" class="form-control mb-2" accept=".xlsx,.xls,.csv" required onchange="lerCabecalho(this, 'col_tipo_pos', 'col_esp_pos')">
                        <label class="form-label small">Coluna: Tipo Prestador</label>
                        <select id="col_tipo_pos" name="col_tipo_pos" class="form-select form-select-sm mb-2"><option value="TIPO PRESTADOR">TIPO PRESTADOR</option></select>
                        <label class="form-label small">Coluna: Especialidade</label>
                        <select id="col_esp_pos" name="col_esp_pos" class="form-select form-select-sm"><option value="ESPECIALIDADE">ESPECIALIDADE</option></select>
                    </div>
                    <div class="col-md-6">
                        <label class="form-label fw-bold text-success">2. Base Operadora Intermediária</label>
                        <input type="file" id="f_op" name="f_op" class="form-control mb-2" accept=".xlsx,.xls,.csv" required onchange="lerCabecalho(this, 'col_tipo_op', 'col_esp_op')">
                        <label class="form-label small">Coluna: Tipo Prestador</label>
                        <select id="col_tipo_op" name="col_tipo_op" class="form-select form-select-sm mb-2"><option value="TIPO PRESTADOR">TIPO PRESTADOR</option></select>
                        <label class="form-label small">Coluna: Especialidade</label>
                        <select id="col_esp_op" name="col_esp_op" class="form-select form-select-sm"><option value="ESPECIALIDADE">ESPECIALIDADE</option></select>
                    </div>
                </div>
                <div class="mb-4 text-center">
                    <label class="form-label fw-bold">3. Base Geográfica (IBGE)</label>
                    <input type="file" name="f_ibge" class="form-control w-50 mx-auto" accept=".xlsx,.xls,.csv" required>
                </div>
                <button type="submit" class="btn btn-success w-100 py-3 fw-bold fs-5" onclick="this.innerHTML='A processar... Aguarde! Pode demorar alguns minutos.'; this.style.opacity='0.7';">Processar e Gerar Relatório Excel</button>
            </form>
        </div>
    </div>
</div>
<script>
function lerCabecalho(input, selectId1, selectId2) {
    if (!input.files || input.files.length === 0) return;
    let file = input.files[0];
    let reader = new FileReader();
    reader.onload = function(e) {
        let data = new Uint8Array(e.target.result);
        let workbook = XLSX.read(data, {type: 'array'});
        let firstSheet = workbook.Sheets[workbook.SheetNames[0]];
        let headers = XLSX.utils.sheet_to_json(firstSheet, {header: 1})[0];
        if(headers) {
            let s1 = document.getElementById(selectId1); let s2 = document.getElementById(selectId2);
            s1.innerHTML = ''; s2.innerHTML = '';
            headers.forEach(h => {
                let text = (h || '').toString().trim().toUpperCase();
                if(text) {
                    s1.options.add(new Option(text, text)); s2.options.add(new Option(text, text));
                }
            });
        }
    };
    reader.readAsArrayBuffer(file);
}
</script>
</body></html>"""

TELA_PENTEFINO_INDIVIDUAL = """<!DOCTYPE html><html lang="pt-BR"><head><title>Busca Sniper Individual</title><link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet"></head>
<body style="background-color: #f0f2f5;">
<nav class="navbar navbar-dark bg-dark"><div class="container-fluid"><a href="/pentefino" class="btn btn-outline-light btn-sm">⬅ Voltar</a><span class="text-white fw-bold">Busca Individual (Sniper)</span></div></nav>
<div class="container mt-4">
    {% with messages = get_flashed_messages(with_categories=true) %}{% if messages %}{% for category, message in messages %}<div class="alert alert-{{ category }}">{{ message }}</div>{% endfor %}{% endif %}{% endwith %}
    <div class="card shadow-sm border-0">
        <div class="card-header bg-danger text-white fw-bold">🎯 Parâmetros de Substituição</div>
        <div class="card-body">
            <form method="POST" action="/pentefino/individual" enctype="multipart/form-data">
                <div class="row mb-3 bg-light p-3 rounded">
                    <div class="col-md-6 border-end">
                        <label class="form-label fw-bold text-primary">1. Base Postal Saúde</label>
                        <input type="file" name="f_postal" class="form-control mb-2" accept=".xlsx,.xls,.csv" required>
                    </div>
                    <div class="col-md-6">
                        <label class="form-label fw-bold text-success">2. Base Geográfica (IBGE)</label>
                        <input type="file" name="f_ibge" class="form-control mb-2" accept=".xlsx,.xls,.csv" required>
                    </div>
                </div>
                <hr>
                <div class="row mt-4">
                    <div class="col-md-6 mb-3">
                        <label class="form-label fw-bold text-danger">Pesquisar Alvo (Nome OU CNPJ)</label>
                        <input type="text" name="alvo" class="form-control border-danger" placeholder="Ex: CLINICA SANTA HELENA ou 13086053000119" required>
                    </div>
                    <div class="col-md-6 mb-3">
                        <label class="form-label fw-bold">Filtro de Tipo Prestador (Opcional)</label>
                        <input type="text" name="filtro_tipo" class="form-control" placeholder="Deixe em branco para procurar todos os tipos">
                    </div>
                </div>
                <button type="submit" class="btn btn-danger w-100 py-3 fw-bold fs-5 mt-3" onclick="this.innerHTML='A processar... Aguarde!'; this.style.opacity='0.7';">Procurar Substitutos e Descarregar Relatório (Excel)</button>
            </form>
        </div>
    </div>
</div></body></html>"""


# ==========================================
# FUNÇÕES DE APOIO PENTE FINO (PANDAS)
# ==========================================
def normalizar_texto(texto):
    if pd.isna(texto) or str(texto).lower() in ['nan', 'none', '']: return ""
    return ''.join(c for c in unicodedata.normalize('NFD', str(texto)) if unicodedata.category(c) != 'Mn').upper().strip()

def cacador_de_colunas(df, palavras_chave):
    for k in palavras_chave:
        for col in df.columns:
            if k == str(col).strip().upper(): return col
    for k in palavras_chave:
        for col in df.columns:
            if k in str(col).strip().upper(): return col
    return None

def limpar_ibge(val):
    if pd.isna(val): return ""
    v = ''.join(filter(str.isdigit, str(val).split('.')[0].strip()))
    return v[:6] if len(v) >= 6 else v

def extrair_especialidades(series):
    specs = set()
    for val in series.dropna():
        for s in str(val).replace(';', ',').split(','):
            v = normalizar_texto(s)
            if v:
                if v in IGNORADAS: continue
                specs.add(DE_PARA.get(v, v))
    return specs

def ler_arquivo(file_obj):
    filename = file_obj.filename.lower()
    if filename.endswith('.csv'):
        try: return pd.read_csv(file_obj, sep=';', encoding='latin1', dtype=str)
        except: 
            file_obj.seek(0)
            return pd.read_csv(file_obj, sep=',', encoding='utf-8', dtype=str)
    else: return pd.read_excel(file_obj, dtype=str)


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
            flash("Erro de ligação com a base de dados.", "danger")
            return render_template_string(TELA_LOGIN)
            
        c = conn.cursor(); ph = get_ph(conn)
        c.execute(f"SELECT id, role FROM users WHERE username={ph} AND password={ph}", (user, password))
        res = c.fetchone()
        conn.close()
        
        if res:
            session['user'] = user; session['role'] = res[1]
            return redirect(url_for('dashboard'))
        else: flash("Utilizador ou senha incorretos!", "danger")
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
    if 'user' not in session or session.get('role') != 'admin': return redirect(url_for('dashboard'))
    conn = conectar_db(); c = conn.cursor(); ph = get_ph(conn)

    if request.method == 'POST':
        acao = request.form.get('acao')
        if acao == 'adicionar':
            u = request.form.get('username'); p = request.form.get('password'); r = request.form.get('role')
            try:
                c.execute(f"INSERT INTO users (username, password, role, first_login) VALUES ({ph}, {ph}, {ph}, 1)", (u, p, r))
                conn.commit(); flash(f"Utilizador '{u}' criado com sucesso!", "success")
            except Exception: conn.rollback(); flash("Erro: Nome de utilizador já existe!", "danger")
        elif acao == 'eliminar':
            uid = request.form.get('user_id')
            c.execute(f"DELETE FROM users WHERE id={ph}", (uid,))
            conn.commit(); flash("Utilizador excluído com sucesso!", "warning")

    c.execute("SELECT id, username, role, first_login FROM users ORDER BY id ASC")
    usuarios = c.fetchall(); conn.close()
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
        flash(f"Erro ao gerar backup: {e}", "danger"); return redirect(url_for('admin'))

@app.route('/admin/importar', methods=['POST'])
def importar_backup():
    if 'user' not in session or session.get('role') != 'admin': return redirect(url_for('dashboard'))
    file = request.files.get('file_backup')
    if not file:
        flash("Nenhum ficheiro selecionado.", "danger"); return redirect(url_for('admin'))
        
    filename = file.filename.lower()
    if not (filename.endswith('.xlsx') or filename.endswith('.db')):
        flash("Selecione um ficheiro Excel (.xlsx) ou Base de Dados SQLite (.db) válido.", "danger")
        return redirect(url_for('admin'))
        
    try:
        conn = conectar_db(); c = conn.cursor(); is_postgres = "psycopg2" in str(type(conn))
        if filename.endswith('.xlsx'):
            xls = pd.ExcelFile(file); sheets = xls.sheet_names
            mapeamento = {'Rotina_PDFs': 'demands', 'Demandas_Avulsas': 'demandas_avulsas', 'Usuarios': 'users'}
            for aba, tabela in mapeamento.items():
                if aba in sheets:
                    df = pd.read_excel(xls, sheet_name=aba); c.execute(f"DELETE FROM {tabela}") 
                    if not df.empty:
                        df = df.where(pd.notnull(df), None)
                        cols = ", ".join(df.columns); placeholders = ", ".join(["%s" if is_postgres else "?"] * len(df.columns))
                        c.executemany(f"INSERT INTO {tabela} ({cols}) VALUES ({placeholders})", df.values.tolist())
                        if is_postgres:
                            try: c.execute(f"SELECT setval('{tabela}_id_seq', COALESCE((SELECT MAX(id)+1 FROM {tabela}), 1), false)")
                            except: pass

        elif filename.endswith('.db'):
            import sqlite3, tempfile
            temp_db = os.path.join(tempfile.gettempdir(), "temp_migration.db")
            file.save(temp_db) 
            sqlite_conn = sqlite3.connect(temp_db); sqlite_c = sqlite_conn.cursor()
            for tabela in ['demands', 'demandas_avulsas', 'users']:
                try:
                    sqlite_c.execute(f"SELECT * FROM {tabela}")
                    rows = sqlite_c.fetchall()
                    if rows:
                        col_names = [description[0] for description in sqlite_c.description]
                        cols = ", ".join(col_names); placeholders = ", ".join(["%s" if is_postgres else "?"] * len(col_names))
                        c.execute(f"DELETE FROM {tabela}") 
                        c.executemany(f"INSERT INTO {tabela} ({cols}) VALUES ({placeholders})", rows)
                        if is_postgres:
                            try: c.execute(f"SELECT setval('{tabela}_id_seq', COALESCE((SELECT MAX(id)+1 FROM {tabela}), 1), false)")
                            except: pass
                except Exception as e_tab: flash(f"Aviso: Tabela {tabela} ignorada. {e_tab}", "warning")
            sqlite_conn.close(); os.remove(temp_db)
            
        conn.commit(); conn.close(); session.clear()
        flash("Base de Dados importada com sucesso! Inicie sessão novamente.", "success")
        return redirect(url_for('login'))
        
    except Exception as e:
        flash(f"Falha ao restaurar banco. Erro: {e}", "danger"); return redirect(url_for('admin'))


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

    # FILTROS GET
    s_filtro = request.args.get('status', 'Todos'); t_filtro = request.args.get('tipo', 'Todos'); r_filtro = request.args.get('resp', 'Todos')
    q = "SELECT * FROM demandas_avulsas WHERE 1=1"; p = []
    if s_filtro != 'Todos': q += f" AND status={ph}"; p.append(s_filtro)
    if t_filtro != 'Todos': q += f" AND tipo_demanda={ph}"; p.append(t_filtro)
    if r_filtro != 'Todos': q += f" AND responsavel={ph}"; p.append(r_filtro)
    q += " ORDER BY id DESC"
    
    c.execute(q, tuple(p)); demandas = c.fetchall()
    c.execute("SELECT username FROM users"); usuarios = ['Todos', 'Nenhum'] + [r[0] for r in c.fetchall()]
    conn.close()
    return render_template_string(TELA_AVULSAS, demandas=demandas, tipos=TIPOS_AVULSA, sols=SOLICITANTES, usuarios=usuarios, status_list=STATUS_AVULSA)

@app.route('/demandas/rotinas')
def demandas_rotinas():
    if 'user' not in session: return redirect(url_for('login'))
    conn = conectar_db(); c = conn.cursor(); ph = get_ph(conn)
    
    # FILTROS GET
    s_filtro = request.args.get('status', 'Todos'); t_filtro = request.args.get('tipo', 'Todos'); r_filtro = request.args.get('resp', 'Todos')
    q = "SELECT * FROM demands WHERE 1=1"; p = []
    if s_filtro != 'Todos': q += f" AND status={ph}"; p.append(s_filtro)
    if t_filtro != 'Todos': q += f" AND type={ph}"; p.append(t_filtro)
    if r_filtro != 'Todos': q += f" AND assigned_to={ph}"; p.append(r_filtro)
    q += " ORDER BY id DESC"
    
    c.execute(q, tuple(p)); rotinas = c.fetchall()
    c.execute("SELECT username FROM users"); usuarios = ['Todos', 'Nenhum'] + [r[0] for r in c.fetchall()]
    conn.close()
    return render_template_string(TELA_ROTINAS, rotinas=rotinas, tipos=TIPOS_ROTINA, usuarios=usuarios)

@app.route('/demandas/rotinas/upload', methods=['POST'])
def upload_rotinas():
    if 'user' not in session: return redirect(url_for('login'))
    arquivos = request.files.getlist('pdfs')
    if not arquivos or arquivos[0].filename == '': flash("Nenhum ficheiro selecionado.", "warning"); return redirect(url_for('demandas_rotinas'))

    kw = {"REAJUSTE": "Reajuste", "NOVO CONTRATO": "Novo Contrato", "INCLUSÃO": "Inclusão", "EXCLUSÃO": "Exclusão", "AJUSTE": "Ajuste", "EXTENSÃO": "Extensão", "DESCREDENCIAMENTO": "Descredenciamento"}
    valid_ufs = {'AC', 'AL', 'AM', 'AP', 'BA', 'CE', 'DF', 'ES', 'GO', 'MA', 'MT', 'MS', 'MG', 'PA', 'PB', 'PR', 'PE', 'PI', 'RJ', 'RN', 'RS', 'RO', 'RR', 'SC', 'SP', 'SE', 'TO'}
    conn = conectar_db(); c = conn.cursor(); ph = get_ph(conn); count = 0
    
    for file in arquivos:
        f_name = file.filename
        c.execute(f"SELECT id FROM demands WHERE filename={ph}", (f_name,))
        if c.fetchone(): continue
        
        fu = f_name.upper(); dem = "Não Identificada"
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
                reader = PyPDF2.PdfReader(file.stream); full_txt = ""
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
    d_id = request.form.get('id'); acao = request.form.get('acao')
    conn = conectar_db(); c = conn.cursor(); ph = get_ph(conn)
    if acao == 'assumir': c.execute(f"UPDATE demands SET status='Em Análise', assigned_to={ph} WHERE id={ph}", (session['user'], d_id))
    elif acao == 'finalizar':
        dt_c = datetime.now().strftime("%d/%m/%Y %H:%M")
        c.execute(f"UPDATE demands SET status='Finalizada', data_finalizacao={ph} WHERE id={ph}", (dt_c, d_id))
    conn.commit(); conn.close()
    return redirect(url_for('demandas_rotinas'))

@app.route('/demandas/dashboard_graficos')
def dashboard_graficos():
    if 'user' not in session: return redirect(url_for('login'))
    conn = conectar_db(); c = conn.cursor()
    c.execute("SELECT COUNT(*) FROM demands"); t_rotina = c.fetchone()[0]
    c.execute("SELECT COUNT(*) FROM demandas_avulsas"); t_avulsa = c.fetchone()[0]
    total = t_rotina + t_avulsa
    c.execute("SELECT COUNT(*) FROM demands WHERE status IN ('Pendente', 'Em Análise')"); p_rot = c.fetchone()[0]
    c.execute("SELECT COUNT(*) FROM demandas_avulsas WHERE status IN ('Pendente', 'Em Análise')"); p_av = c.fetchone()[0]
    pend = p_rot + p_av
    c.execute("SELECT COUNT(*) FROM demands WHERE status='Finalizada'"); c_rot = c.fetchone()[0]
    c.execute("SELECT COUNT(*) FROM demandas_avulsas WHERE status='Concluído'"); c_av = c.fetchone()[0]
    conc = c_rot + c_av
    
    fig1, ax1 = plt.subplots(figsize=(6,4))
    c.execute("SELECT assigned_to, COUNT(*) FROM demands WHERE status='Finalizada' AND assigned_to != 'Nenhum' GROUP BY assigned_to")
    dados_prod = c.fetchall()
    if dados_prod:
        ax1.bar([r[0] for r in dados_prod], [r[1] for r in dados_prod], color='#5cb85c')
        ax1.set_xticklabels([r[0] for r in dados_prod], rotation=45, ha='right')
    buf1 = BytesIO(); fig1.tight_layout(); fig1.savefig(buf1, format="png"); buf1.seek(0)
    chart_prod = base64.b64encode(buf1.read()).decode('utf-8'); plt.close(fig1)

    fig2, ax2 = plt.subplots(figsize=(6,4))
    c.execute("SELECT type, COUNT(*) FROM demands GROUP BY type")
    dados_tipo = c.fetchall()
    if dados_tipo:
        ax2.bar([r[0] for r in dados_tipo], [r[1] for r in dados_tipo], color='#008CBA')
        ax2.set_xticklabels([r[0] for r in dados_tipo], rotation=45, ha='right')
    buf2 = BytesIO(); fig2.tight_layout(); fig2.savefig(buf2, format="png"); buf2.seek(0)
    chart_tipo = base64.b64encode(buf2.read()).decode('utf-8'); plt.close(fig2)
    conn.close()
    return render_template_string(TELA_DASHBOARD_GRAFICOS, d_total=total, d_pend=pend, d_ana=0, d_conc=conc, chart_prod=chart_prod, chart_tipo=chart_tipo)

# -------- ROTAS DO MÓDULO PENTE FINO (RN 665) --------
@app.route('/pentefino')
def hub_pentefino():
    if 'user' not in session: return redirect(url_for('login'))
    return render_template_string(TELA_HUB_PENTEFINO)

@app.route('/pentefino/regional', methods=['GET', 'POST'])
def pentefino_regional():
    if 'user' not in session: return redirect(url_for('login'))
    if request.method == 'GET': return render_template_string(TELA_PENTEFINO_REGIONAL)
        
    try:
        f_postal = request.files.get('f_postal'); f_op = request.files.get('f_op'); f_ibge = request.files.get('f_ibge')
        df_postal = ler_arquivo(f_postal); df_operadora = ler_arquivo(f_op); df_ibge = ler_arquivo(f_ibge)
        
        df_postal.columns = [str(c).strip().upper() for c in df_postal.columns]
        df_operadora.columns = [str(c).strip().upper() for c in df_operadora.columns]
        df_ibge.columns = [str(c).strip().upper() for c in df_ibge.columns]

        col_ibge_7 = cacador_de_colunas(df_ibge, ['COMPLETO']); col_ibge_6 = cacador_de_colunas(df_ibge, ['AJUSTADO', 'CÓD. MUNIC'])
        col_regiao_csv = cacador_de_colunas(df_ibge, ['REGIÃO', 'REGIAO DE SAUDE', 'NOME DA REGIÃO DE SAÚDE'])
        col_mun_csv = cacador_de_colunas(df_ibge, ['MUNICÍPIO', 'MUNICIPIO']); col_uf_csv = cacador_de_colunas(df_ibge, ['UF'])

        map_ibge_regiao = {}
        if col_ibge_7: map_ibge_regiao.update(dict(zip(df_ibge[col_ibge_7].apply(limpar_ibge), df_ibge[col_regiao_csv])))
        if col_ibge_6: map_ibge_regiao.update(dict(zip(df_ibge[col_ibge_6].apply(limpar_ibge), df_ibge[col_regiao_csv])))
        df_ibge['CHAVE_GEO'] = df_ibge[col_mun_csv].apply(normalizar_texto) + "_" + df_ibge[col_uf_csv].apply(normalizar_texto)
        map_geo_regiao = dict(zip(df_ibge['CHAVE_GEO'], df_ibge[col_regiao_csv]))

        dics_processados = []
        for df in [df_postal, df_operadora]:
            col_ibge = cacador_de_colunas(df, ['IBGE', 'CÓDIGO IBGE'])
            col_mun = cacador_de_colunas(df, ['MUNICÍPIO', 'MUNICIPIO', 'CIDADE'])
            col_uf = cacador_de_colunas(df, ['UF', 'ESTADO'])
            if col_ibge: df['REG_IBGE'] = df[col_ibge].apply(limpar_ibge).map(map_ibge_regiao)
            else: df['REG_IBGE'] = None
            if col_mun and col_uf: df['REG_GEO'] = (df[col_mun].apply(normalizar_texto) + "_" + df[col_uf].apply(normalizar_texto)).map(map_geo_regiao)
            else: df['REG_GEO'] = None
            df['REGIAO_SAUDE'] = df['REG_IBGE'].fillna(df['REG_GEO']).fillna("REGIÃO NÃO IDENTIFICADA")
            dics_processados.append(df)

        df_postal, df_operadora = dics_processados[0], dics_processados[1]

        v_tipo_pos = request.form.get('col_tipo_pos', '').strip().upper()
        v_esp_pos = request.form.get('col_esp_pos', '').strip().upper()
        v_tipo_op = request.form.get('col_tipo_op', '').strip().upper()
        v_esp_op = request.form.get('col_esp_op', '').strip().upper()

        col_cnpj_pos = cacador_de_colunas(df_postal, ['CNPJ', 'CPFCNPJ'])
        col_nome_pos = cacador_de_colunas(df_postal, ['NOME', 'RAZAO', 'PRESTADOR', 'FANTASIA'])
        col_mun_pos = cacador_de_colunas(df_postal, ['MUNICÍPIO', 'MUNICIPIO', 'CIDADE'])
        
        col_esp_pos = v_esp_pos if v_esp_pos in df_postal.columns else cacador_de_colunas(df_postal, ['ESPECIALIDADE', 'ESPECIALIDADES'])
        col_tipo_pos = v_tipo_pos if v_tipo_pos in df_postal.columns else cacador_de_colunas(df_postal, ['TIPO PRESTADOR', 'TIPO_PRESTADOR', 'TIPOPRESTADOR', 'TIPO'])
        col_esp_op = v_esp_op if v_esp_op in df_operadora.columns else cacador_de_colunas(df_operadora, ['ESPECIALIDADE', 'ESPECIALIDADES'])

        if not col_cnpj_pos: df_postal['CNPJ_TEMP'] = "S/CNPJ"; col_cnpj_pos = 'CNPJ_TEMP'
        if not col_nome_pos: df_postal['NOME_TEMP'] = "S/NOME"; col_nome_pos = 'NOME_TEMP'
        if not col_mun_pos: df_postal['MUN_TEMP'] = "S/MUNICIPIO"; col_mun_pos = 'MUN_TEMP'
        df_postal['CNPJ_LIMPO'] = df_postal[col_cnpj_pos].astype(str).str.replace(r'\D', '', regex=True)

        dict_agg = {}
        if col_esp_pos: dict_agg[col_esp_pos] = lambda s: ", ".join(sorted(list(extrair_especialidades(s))))
        if col_tipo_pos: dict_agg[col_tipo_pos] = lambda x: " | ".join(x.dropna().astype(str).unique())
        df_postal_agrupado = df_postal.groupby(['REGIAO_SAUDE', 'CNPJ_LIMPO', col_nome_pos, col_mun_pos], dropna=False).agg(dict_agg).reset_index()

        regioes_operadora = set(df_operadora['REGIAO_SAUDE'].dropna().unique())
        regioes_operadora.discard("REGIÃO NÃO IDENTIFICADA")
        df_postal_alvos = df_postal_agrupado[df_postal_agrupado['REGIAO_SAUDE'].isin(regioes_operadora)].copy()

        piscina_operadora = {}
        for regiao, grupo in df_operadora.groupby('REGIAO_SAUDE'):
            if regiao == "NÃO ENCONTRADO": continue
            piscina_operadora[regiao] = extrair_especialidades(grupo[col_esp_op]) if col_esp_op else set()

        res_reg = {}
        for _, row in df_postal_alvos.iterrows():
            regiao = row.get('REGIAO_SAUDE')
            if regiao not in res_reg: res_reg[regiao] = []
            espec_str = str(row.get(col_esp_pos, '')) if col_esp_pos else ""
            if not espec_str.strip() or espec_str.upper() == 'NAN': obs = "Vulnerabilidade: Especialidades não informadas na base da Postal."; espec_alvo = set()
            else: espec_alvo = set([e.strip() for e in espec_str.split(',') if e.strip()]); obs = ""
            
            if not obs:
                faltantes = espec_alvo - piscina_operadora.get(regiao, set())
                obs = "Vulnerabilidade: " + ", ".join(sorted(list(faltantes))) if faltantes else "Cobertura Total"
            res_reg[regiao].append({"CNPJ": str(row.get('CNPJ_LIMPO', 'N/A')), "Prestador": str(row.get(col_nome_pos, 'N/A')), "Município": str(row.get(col_mun_pos, 'N/A')), "OBSERVAÇÃO": obs})

        for regiao in regioes_operadora:
            if regiao not in res_reg: res_reg[regiao] = [{"CNPJ": "", "Prestador": "", "Município": "", "OBSERVAÇÃO": "Sem prestadores postal"}]

        output = BytesIO()
        wb = Workbook(); ws = wb.active; ws.title = "Analise Equivalencia"
        fill_regiao = PatternFill(start_color="005580", end_color="005580", fill_type="solid"); font_regiao = Font(color="FFFFFF", bold=True, size=11)
        fill_colunas = PatternFill(start_color="005580", end_color="005580", fill_type="solid"); font_colunas = Font(color="FFFFFF", bold=True)
        align_center = Alignment(horizontal="center", vertical="center"); align_left = Alignment(horizontal="left", vertical="center", wrap_text=True)
        thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))

        linha_atual = 1
        for regiao in sorted(res_reg.keys()):
            ws.merge_cells(start_row=linha_atual, start_column=1, end_row=linha_atual, end_column=4)
            c_reg = ws.cell(row=linha_atual, column=1, value=str(regiao).upper()); c_reg.fill = fill_regiao; c_reg.font = font_regiao; c_reg.alignment = align_center; linha_atual += 1
            for col_idx, texto in enumerate(["CNPJ", "Prestador", "Município", "OBSERVAÇÃO"], 1):
                cell = ws.cell(row=linha_atual, column=col_idx, value=texto); cell.fill = fill_colunas; cell.font = font_colunas; cell.alignment = align_left; cell.border = thin_border
            linha_atual += 1
            for item in res_reg[regiao]:
                c1 = ws.cell(row=linha_atual, column=1, value=item['CNPJ']); c2 = ws.cell(row=linha_atual, column=2, value=item['Prestador'])
                c3 = ws.cell(row=linha_atual, column=3, value=item['Município']); c4 = ws.cell(row=linha_atual, column=4, value=item['OBSERVAÇÃO'])
                for c in [c1, c2, c3, c4]: c.alignment = align_left; c.border = thin_border
                linha_atual += 1
            linha_atual += 1

        ws.column_dimensions['A'].width = 20; ws.column_dimensions['B'].width = 45; ws.column_dimensions['C'].width = 25; ws.column_dimensions['D'].width = 65
        wb.save(output); output.seek(0)
        nome_arquivo = f"Relatorio_Regional_{datetime.now().strftime('%Hh%Mm')}.xlsx"
        return send_file(output, download_name=nome_arquivo, as_attachment=True)

    except Exception as e:
        flash(f"Erro ao processar as planilhas: {e}", "danger")
        return redirect(url_for('pentefino_regional'))

@app.route('/pentefino/individual', methods=['GET', 'POST'])
def pentefino_individual():
    if 'user' not in session: return redirect(url_for('login'))
    if request.method == 'GET': return render_template_string(TELA_PENTEFINO_INDIVIDUAL)
        
    try:
        f_postal = request.files.get('f_postal'); f_ibge = request.files.get('f_ibge')
        
        # BUSCA INTELIGENTE: NOME OU CNPJ
        alvo_bruto = request.form.get('alvo', '').strip().upper()
        filtro_tipo = request.form.get('filtro_tipo', '').strip().upper()
        
        df_ind = ler_arquivo(f_postal); df_ibge = ler_arquivo(f_ibge)
        df_ind.columns = [str(c).strip().upper() for c in df_ind.columns]
        df_ibge.columns = [str(c).strip().upper() for c in df_ibge.columns]

        col_cnpj_ind = cacador_de_colunas(df_ind, ['CNPJ', 'CPFCNPJ'])
        col_nome_ind = cacador_de_colunas(df_ind, ['NOME', 'RAZAO', 'PRESTADOR', 'FANTASIA'])
        col_mun_ind = cacador_de_colunas(df_ind, ['MUNICÍPIO', 'MUNICIPIO', 'CIDADE'])
        col_uf_ind = cacador_de_colunas(df_ind, ['UF', 'ESTADO'])
        col_tipo_ind = cacador_de_colunas(df_ind, ['TIPO PRESTADOR', 'TIPO_PRESTADOR', 'TIPOPRESTADOR', 'TIPO'])
        col_esp_ind = cacador_de_colunas(df_ind, ['ESPECIALIDADE', 'ESPECIALIDADES'])
        col_ibge_ind = cacador_de_colunas(df_ind, ['IBGE', 'CÓDIGO IBGE'])
        
        if not col_cnpj_ind or not col_nome_ind:
            flash("Colunas básicas (Nome/CNPJ) não encontradas na base da Postal.", "danger"); return redirect(url_for('pentefino_individual'))
            
        # Converte a coluna de CNPJ para string e remove não números para facilitar a busca
        df_ind[col_cnpj_ind] = df_ind[col_cnpj_ind].astype(str).str.replace(r'\D', '', regex=True)
        
        # MOTOR DE BUSCA SNIPER (Nome ou CNPJ)
        alvo_numeros = re.sub(r'\D', '', alvo_bruto)
        # Cria uma máscara verificando se o alvo_bruto está no Nome OU se o alvo_numeros está no CNPJ
        mask = df_ind[col_nome_ind].fillna('').astype(str).str.upper().str.contains(alvo_bruto, na=False)
        if len(alvo_numeros) > 3: 
            mask = mask | (df_ind[col_cnpj_ind].str.contains(alvo_numeros, na=False))
        
        linhas_alvo = df_ind[mask]
        if linhas_alvo.empty:
            flash(f"O Alvo '{alvo_bruto}' não foi encontrado na base.", "danger"); return redirect(url_for('pentefino_individual'))
            
        row_alvo = linhas_alvo.iloc[0]
        cnpj_alvo = str(row_alvo.get(col_cnpj_ind, '')).strip()
        nome_alvo_real = str(row_alvo.get(col_nome_ind, '')).strip()

        dict_agg = {}
        if col_esp_ind: dict_agg[col_esp_ind] = lambda s: ", ".join(sorted(list(extrair_especialidades(s))))
        if col_tipo_ind: dict_agg[col_tipo_ind] = lambda x: " | ".join(x.dropna().astype(str).unique())
        group_cols = [c for c in [col_cnpj_ind, col_nome_ind, col_mun_ind, col_uf_ind, col_ibge_ind] if c]
        
        # Agrupamento para remover duplicatas
        df_ind = df_ind.groupby(group_cols, dropna=False).agg(dict_agg).reset_index()

        c_ibge_7 = cacador_de_colunas(df_ibge, ['COMPLETO']); c_ibge_6 = cacador_de_colunas(df_ibge, ['AJUSTADO', 'CÓD. MUNIC'])
        c_regiao = cacador_de_colunas(df_ibge, ['REGIÃO', 'REGIAO DE SAUDE', 'NOME DA REGIÃO DE SAÚDE'])
        c_mun = cacador_de_colunas(df_ibge, ['MUNICÍPIO', 'MUNICIPIO']); c_uf = cacador_de_colunas(df_ibge, ['UF'])

        map_ibge_regiao = {}
        if c_ibge_7: map_ibge_regiao.update(dict(zip(df_ibge[c_ibge_7].apply(limpar_ibge), df_ibge[c_regiao])))
        if c_ibge_6: map_ibge_regiao.update(dict(zip(df_ibge[c_ibge_6].apply(limpar_ibge), df_ibge[c_regiao])))
        df_ibge['CHAVE_GEO'] = df_ibge[c_mun].apply(normalizar_texto) + "_" + df_ibge[c_uf].apply(normalizar_texto)
        map_geo_regiao = dict(zip(df_ibge['CHAVE_GEO'], df_ibge[c_regiao]))

        # Atualizando a linha alvo após agrupamento
        linhas_alvo_agrupadas = df_ind[df_ind[col_cnpj_ind] == cnpj_alvo]
        if not linhas_alvo_agrupadas.empty: row_alvo = linhas_alvo_agrupadas.iloc[0]

        mun_alvo = str(row_alvo.get(col_mun_ind, '')).strip().upper(); uf_alvo = str(row_alvo.get(col_uf_ind, '')).strip().upper()
        chave_geo_alvo = normalizar_texto(mun_alvo) + "_" + normalizar_texto(uf_alvo)
        tipo_alvo = str(row_alvo.get(col_tipo_ind, '')).strip().upper() if col_tipo_ind else ""
        
        ibge_alvo = str(row_alvo.get(col_ibge_ind, '')) if col_ibge_ind else None
        regiao_alvo = None
        if ibge_alvo and ibge_alvo != 'NAN': regiao_alvo = map_ibge_regiao.get(limpar_ibge(ibge_alvo))
        if not regiao_alvo: regiao_alvo = map_geo_regiao.get(chave_geo_alvo, "NÃO ENCONTRADA")
        
        esp_alvo_series = pd.Series([row_alvo.get(col_esp_ind, '')]) if col_esp_ind else pd.Series()
        esp_alvo_set = extrair_especialidades(esp_alvo_series)
        
        resultados = []; piscina_regiao = set() 
        for _, row in df_ind.iterrows():
            cnpj_cand = str(row.get(col_cnpj_ind, '')).strip()
            if cnpj_cand == cnpj_alvo or cnpj_cand == 'NAN' or not cnpj_cand: continue
            
            mun_cand = str(row.get(col_mun_ind, '')).strip().upper(); uf_cand = str(row.get(col_uf_ind, '')).strip().upper()
            chave_geo_cand = normalizar_texto(mun_cand) + "_" + normalizar_texto(uf_cand)
            
            ibge_cand = str(row.get(col_ibge_ind, '')) if col_ibge_ind else None
            regiao_cand = None
            if ibge_cand and ibge_cand != 'NAN': regiao_cand = map_ibge_regiao.get(limpar_ibge(ibge_cand))
            if not regiao_cand: regiao_cand = map_geo_regiao.get(chave_geo_cand, "NÃO ENCONTRADA")
            
            prox_str = ""; ordem_prox = 2
            if mun_cand == mun_alvo and uf_cand == uf_alvo: prox_str = "1 - Mesmo Município"; ordem_prox = 0
            elif regiao_cand == regiao_alvo and regiao_alvo != "NÃO ENCONTRADA": prox_str = "2 - Mesma Região de Saúde"; ordem_prox = 1
            else: continue
            
            tipo_cand = str(row.get(col_tipo_ind, '')).strip().upper() if col_tipo_ind else "-"
            if filtro_tipo and tipo_cand != filtro_tipo: continue
                
            esp_cand_series = pd.Series([row.get(col_esp_ind, '')]) if col_esp_ind else pd.Series()
            esp_cand_set = extrair_especialidades(esp_cand_series); piscina_regiao.update(esp_cand_set)
            
            faltante = esp_alvo_set - esp_cand_set; nome_cand = str(row.get(col_nome_ind, '')).strip()
            if len(faltante) == 0: status_cob = "100% Coberto (Sem Vuln)"; txt_faltante = "Nenhuma"
            else: status_cob = f"Vulnerabilidade: Faltam {len(faltante)}"; txt_faltante = "\n".join([f"• {f}" for f in sorted(faltante)])
            resultados.append((ordem_prox, len(faltante), prox_str, cnpj_cand, nome_cand, tipo_cand, f"{mun_cand}/{uf_cand}", status_cob, txt_faltante))
            
        resultados.sort(key=lambda x: (x[0], x[1], x[4])) 
        faltante_pool = esp_alvo_set - piscina_regiao
        if len(faltante_pool) == 0: status_pool = "100% Coberto pela Rede"; txt_faltante_pool = "A soma de todos os prestadores da região cobre as especialidades."
        else: status_pool = f"Vulnerabilidade na Rede: Faltam {len(faltante_pool)}"; txt_faltante_pool = "\n".join([f"• {f}" for f in sorted(faltante_pool)])
            
        pool_row = ("☆ POOL DA REGIÃO ☆", "MÚLTIPLOS", "SOMA DOS PRESTADORES VÁLIDOS DA REGIÃO", filtro_tipo, regiao_alvo, status_pool, txt_faltante_pool)

        output = BytesIO()
        wb = Workbook(); ws = wb.active; ws.title = "Substituicao Individual"
        fill_header = PatternFill(start_color="002060", end_color="002060", fill_type="solid")
        fill_cols = PatternFill(start_color="e0e0e0", end_color="e0e0e0", fill_type="solid")
        fill_100 = PatternFill(start_color="e2efda", end_color="e2efda", fill_type="solid"); fill_pool = PatternFill(start_color="FFD966", end_color="FFD966", fill_type="solid") 
        font_white = Font(color="FFFFFF", bold=True); font_black = Font(color="000000", bold=False)
        align_c = Alignment(horizontal="center", vertical="center"); align_l = Alignment(horizontal="left", vertical="center", wrap_text=True)
        thin_border = Border(left=Side(style='thin', color='A0A0A0'), right=Side(style='thin', color='A0A0A0'), top=Side(style='thin', color='A0A0A0'), bottom=Side(style='thin', color='A0A0A0'))

        ws.merge_cells("A1:G1")
        c_title = ws.cell(row=1, column=1, value=f"RELATÓRIO DE SUBSTITUIÇÃO: {cnpj_alvo} - {nome_alvo_real} ({mun_alvo}/{uf_alvo})")
        c_title.fill = fill_header; c_title.font = font_white; c_title.alignment = align_c
        
        headers = ["Proximidade", "CNPJ", "Prestador Candidato", "Tipo Prestador", "Município/UF", "Status Cobertura", "OBSERVAÇÃO (O que falta)"]
        for col_idx, texto in enumerate(headers, 1):
            cell = ws.cell(row=2, column=col_idx, value=texto); cell.fill = fill_cols; cell.font = Font(color="000000", bold=True); cell.alignment = align_c; cell.border = thin_border
            
        for i, val in enumerate(pool_row, 1):
            cp = ws.cell(row=3, column=i, value=val); cp.fill = fill_pool; cp.font = Font(color="000000", bold=True); cp.alignment = align_l; cp.border = thin_border

        linha = 4
        for r in resultados:
            c1 = ws.cell(row=linha, column=1, value=r[2]); c2 = ws.cell(row=linha, column=2, value=r[3])
            c3 = ws.cell(row=linha, column=3, value=r[4]); c4 = ws.cell(row=linha, column=4, value=r[5])
            c5 = ws.cell(row=linha, column=5, value=r[6]); c6 = ws.cell(row=linha, column=6, value=r[7]); c7 = ws.cell(row=linha, column=7, value=r[8])
            for c in [c1, c2, c3, c4, c5, c6, c7]:
                c.alignment = align_l; c.border = thin_border
                if r[1] == 0: c.fill = fill_100
                else: c.font = font_black
            linha += 1

        ws.column_dimensions['A'].width = 25; ws.column_dimensions['B'].width = 20; ws.column_dimensions['C'].width = 50; ws.column_dimensions['D'].width = 20
        ws.column_dimensions['E'].width = 20; ws.column_dimensions['F'].width = 25; ws.column_dimensions['G'].width = 65

        wb.save(output); output.seek(0)
        nome_arquivo = f"Relatorio_Substituicao_{cnpj_alvo}.xlsx"
        return send_file(output, download_name=nome_arquivo, as_attachment=True)

    except Exception as e:
        flash(f"Erro ao processar as planilhas: {e}", "danger")
        return redirect(url_for('pentefino_individual'))

if __name__ == '__main__':
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)
