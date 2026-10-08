import os
import platform
import subprocess
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import PyPDF2
import re
import sqlite3
import psycopg2
from datetime import datetime

from matplotlib.figure import Figure
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from matplotlib.backends.backend_pdf import PdfPages
from config import conectar_db, replace_placeholders, TIPOS_ROTINA, TIPOS_AVULSA, STATUS_AVULSA, SOLICITANTES

class ControleDemandasFrame(tk.Frame):
    def __init__(self, parent, controller):
        super().__init__(parent, bg="#f4f4f9"); self.controller = controller
        tb = tk.Frame(self, bg="#333", height=50); tb.pack(fill=tk.X)
        tk.Button(tb, text="⬅ Voltar ao Portal Principal", command=lambda: self.controller.show_frame("Portal"), bg="#777", fg="white", font=("Helvetica", 10, "bold"), relief=tk.FLAT).pack(side=tk.LEFT, padx=20, pady=10)
        self.lui = tk.Label(tb, text="", bg="#333", fg="white", font=("Helvetica", 10, "bold")); self.lui.pack(side=tk.RIGHT, padx=20, pady=10)
        self.ca = tk.Frame(self, bg="#f4f4f9"); self.ca.pack(fill=tk.BOTH, expand=True)
        self.fh = tk.Frame(self.ca, bg="#f4f4f9"); self.fr = tk.Frame(self.ca, bg="#f4f4f9"); self.fa = tk.Frame(self.ca, bg="#f4f4f9"); self.fd = tk.Frame(self.ca, bg="#f4f4f9")
        self.setup_home_menu(); self.setup_rotina(); self.setup_avulsas(); self.setup_dashboard()

    def refresh_all(self): 
        self.lui.config(text=f"Usuário: {self.controller.current_user} | Nível: {self.controller.current_role.upper()}")
        self.atualizar_usuarios_combos()
        self.show_module("Home")

    def criar_btn_voltar(self, pf):
        h = tk.Frame(pf, bg="#f4f4f9", pady=10); h.pack(fill=tk.X, padx=10)
        tk.Button(h, text="⬅ Voltar ao Menu de Demandas", command=lambda: self.show_module("Home"), bg="#777", fg="white", font=("Helvetica", 10, "bold"), relief=tk.FLAT).pack(side=tk.LEFT)
        return h

    def atualizar_usuarios_combos(self):
        conn = conectar_db(); c = conn.cursor()
        c.execute("SELECT username FROM users"); u = ["Todos"] + [r[0] for r in c.fetchall()]
        conn.close()
        
        vu_rotina = self.cbo_rotina_resp.get(); self.cbo_rotina_resp['values'] = u
        if vu_rotina in u: self.cbo_rotina_resp.set(vu_rotina)
        else: self.cbo_rotina_resp.current(0)
            
        vu_avulsa = self.cbo_av_resp.get(); self.cbo_av_resp['values'] = u
        if vu_avulsa in u: self.cbo_av_resp.set(vu_avulsa)
        else: self.cbo_av_resp.current(0)

    def show_module(self, m_name):
        for f in (self.fh, self.fr, self.fa, self.fd): f.pack_forget()
        if m_name == "Home": self.fh.pack(fill=tk.BOTH, expand=True)
        elif m_name == "Rotina":
            self.fr.pack(fill=tk.BOTH, expand=True); self.load_rotina()
            if self.controller.current_role == 'admin': self.btn_excluir_rotina.pack(side=tk.RIGHT, padx=5)
            else: self.btn_excluir_rotina.pack_forget()
        elif m_name == "Avulsas":
            self.fa.pack(fill=tk.BOTH, expand=True); self.load_avulsas()
            if self.controller.current_role == 'admin': self.btn_excluir_avulsa.pack(side=tk.RIGHT, padx=5)
            else: self.btn_excluir_avulsa.pack_forget()
        elif m_name == "Dashboard": self.fd.pack(fill=tk.BOTH, expand=True); self.atualizar_filtros_dash(); self.load_dashboard()

    def setup_home_menu(self):
        tk.Label(self.fh, text="Gestão de Demandas", font=("Helvetica", 20, "bold"), bg="#f4f4f9", fg="#333").pack(pady=(60, 40))
        mc = tk.Frame(self.fh, bg="#f4f4f9"); mc.pack()
        tk.Button(mc, text="📄 Demandas de Rotina (Contratos e PDFs)", command=lambda: self.show_module("Rotina"), bg="#008CBA", fg="white", font=("Helvetica", 14, "bold"), relief=tk.FLAT, width=40, pady=15).pack(pady=10)
        tk.Button(mc, text="📝 Demandas Avulsas (Estudos e Reuniões)", command=lambda: self.show_module("Avulsas"), bg="#f0ad4e", fg="white", font=("Helvetica", 14, "bold"), relief=tk.FLAT, width=40, pady=15).pack(pady=10)
        tk.Button(mc, text="📊 Demonstrativo Gerencial (Dashboard)", command=lambda: self.show_module("Dashboard"), bg="#5bc0de", fg="white", font=("Helvetica", 14, "bold"), relief=tk.FLAT, width=40, pady=15).pack(pady=10)

    # ------- DASHBOARD CONTROLE DEMANDAS -------
    def setup_dashboard(self):
        self.criar_btn_voltar(self.fd)
        ff = tk.Frame(self.fd, bg="#f4f4f9", pady=2); ff.pack(fill=tk.X, padx=20)
        tk.Label(ff, text="Origem:", font=("Helvetica", 9, "bold"), bg="#f4f4f9").pack(side=tk.LEFT)
        self.cbo_origem = ttk.Combobox(ff, values=["Ambas", "Rotina (PDFs)", "Avulsa"], state="readonly", width=12); self.cbo_origem.current(0); self.cbo_origem.pack(side=tk.LEFT, padx=(2, 10)); self.cbo_origem.bind("<<ComboboxSelected>>", lambda e: self.load_dashboard())
        tk.Label(ff, text="Tipo:", font=("Helvetica", 9, "bold"), bg="#f4f4f9").pack(side=tk.LEFT)
        self.cbo_tipo = ttk.Combobox(ff, values=["Todos"] + TIPOS_ROTINA + TIPOS_AVULSA, state="readonly", width=15); self.cbo_tipo.current(0); self.cbo_tipo.pack(side=tk.LEFT, padx=(2, 10)); self.cbo_tipo.bind("<<ComboboxSelected>>", lambda e: self.load_dashboard())
        tk.Label(ff, text="Colab.:", font=("Helvetica", 9, "bold"), bg="#f4f4f9").pack(side=tk.LEFT)
        self.cbo_colab = ttk.Combobox(ff, state="readonly", width=12); self.cbo_colab.pack(side=tk.LEFT, padx=(2, 10)); self.cbo_colab.bind("<<ComboboxSelected>>", lambda e: self.load_dashboard())
        tk.Label(ff, text="UF:", font=("Helvetica", 9, "bold"), bg="#f4f4f9").pack(side=tk.LEFT)
        self.cbo_uf = ttk.Combobox(ff, state="readonly", width=6); self.cbo_uf.pack(side=tk.LEFT, padx=(2, 10)); self.cbo_uf.bind("<<ComboboxSelected>>", lambda e: self.load_dashboard())
        tk.Button(ff, text="Exportar Dash PDF", command=self.exportar_dashboard_pdf, bg="#5bc0de", fg="white", font=("Helvetica", 9, "bold"), relief=tk.FLAT).pack(side=tk.RIGHT)
        
        fc = tk.Frame(self.fd, bg="#f4f4f9"); fc.pack(fill=tk.X, padx=20, pady=2)
        fst = tk.Frame(fc, bg="#f4f4f9"); fst.pack(side=tk.TOP, fill=tk.X)
        self.lbl_total = self.cc(fst, "Total Demandas", "#008CBA", 0); self.lbl_pendente = self.cc(fst, "Pendentes", "#f0ad4e", 1)
        self.lbl_analise = self.cc(fst, "Em Análise", "#5bc0de", 2); self.lbl_devolvida = self.cc(fst, "Dev/Cancelada", "#8e44ad", 3); self.lbl_finalizada = self.cc(fst, "Finalizadas", "#5cb85c", 4)
        fsla = tk.Frame(fc, bg="#f4f4f9"); fsla.pack(side=tk.TOP, fill=tk.X, pady=(2, 0))
        tk.Label(fsla, text="⚠️ Alertas de SLA GERED (Rotinas Pendentes):", font=("Helvetica", 9, "bold"), bg="#f4f4f9", fg="#d9534f").pack(side=tk.LEFT, padx=(5, 10))
        self.lbl_s10 = self.ccm(fsla, "10-14 Dias", "#e0a800"); self.lbl_s15 = self.ccm(fsla, "15-19 Dias", "#fd7e14"); self.lbl_s20 = self.ccm(fsla, "20+ Dias", "#c9302c")
        
        fg = tk.Frame(self.fd, bg="#f4f4f9", pady=2); fg.pack(fill=tk.BOTH, expand=True, padx=20)
        self.fig = Figure(figsize=(14, 7), dpi=100); self.fig.patch.set_facecolor('#f4f4f9')
        self.ax1 = self.fig.add_subplot(221); self.ax2 = self.fig.add_subplot(222); self.ax3 = self.fig.add_subplot(223); self.ax4 = self.fig.add_subplot(224)
        self.fig.subplots_adjust(left=0.06, right=0.96, top=0.90, bottom=0.25, hspace=0.6, wspace=0.3)
        self.canvas = FigureCanvasTkAgg(self.fig, master=fg); self.canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)

    def cc(self, p, t, c, col):
        cd = tk.Frame(p, bg=c, width=180, height=60, relief=tk.RAISED, bd=2); cd.grid(row=0, column=col, padx=5, pady=2, sticky="nsew"); cd.grid_propagate(False); p.grid_columnconfigure(col, weight=1)
        tk.Label(cd, text=t, font=("Helvetica", 9, "bold"), bg=c, fg="white").pack(pady=(5, 0))
        l = tk.Label(cd, text="0", font=("Helvetica", 18, "bold"), bg=c, fg="white"); l.pack(); return l

    def ccm(self, p, t, c):
        cd = tk.Frame(p, bg=c, width=120, height=45, relief=tk.RAISED, bd=2); cd.pack(side=tk.LEFT, padx=5, pady=2); cd.pack_propagate(False)
        tk.Label(cd, text=t, font=("Helvetica", 8, "bold"), bg=c, fg="white").pack(pady=(2, 0))
        l = tk.Label(cd, text="0", font=("Helvetica", 14, "bold"), bg=c, fg="white"); l.pack(); return l

    def base_sql(self): return " FROM (SELECT type as tipo, status, assigned_to as responsavel, data_entrada, data_finalizacao, 'Rotina' as origem, uf FROM demands UNION ALL SELECT tipo_demanda as tipo, status, responsavel, data_entrada, data_conclusao as data_finalizacao, 'Avulsa' as origem, '-' as uf FROM demandas_avulsas) AS tabela_unida WHERE 1=1 "

    def fq(self, q, p):
        o, t, c, u = self.cbo_origem.get(), self.cbo_tipo.get(), self.cbo_colab.get(), self.cbo_uf.get()
        if o == "Rotina (PDFs)": q += " AND origem='Rotina'"
        elif o == "Avulsa": q += " AND origem='Avulsa'"
        if t != "Todos": q += " AND tipo=?"; p.append(t)
        if c != "Todos": q += " AND responsavel=?"; p.append(c)
        if u != "Todos" and u != "": q += " AND uf=?"; p.append(u)
        return replace_placeholders(q), p

    def dd(self, ax, dv, dl, t, c=None):
        if not dv: ax.text(0.5, 0.5, "Sem dados", ha='center', va='center', color='gray'); ax.set_title(t, fontsize=11, fontweight='bold'); return
        w, txt, at = ax.pie(dv, autopct='%1.1f%%', pctdistance=0.75, colors=c, wedgeprops=dict(width=0.4, edgecolor='w'), textprops={'fontsize': 9})
        ax.legend(w, dl, loc="center left", bbox_to_anchor=(0.95, 0.5), fontsize=8); ax.set_title(t, fontsize=11, fontweight='bold')

    def atualizar_filtros_dash(self):
        conn = conectar_db(); c = conn.cursor()
        c.execute("SELECT username FROM users"); u = ["Todos"] + [r[0] for r in c.fetchall()]
        vu = self.cbo_colab.get(); self.cbo_colab['values'] = u
        if vu in u: self.cbo_colab.set(vu)
        else: self.cbo_colab.current(0)
        c.execute("SELECT DISTINCT uf FROM demands WHERE uf != '-' AND uf != ''")
        f = ["Todos"] + [r[0] for r in c.fetchall() if r[0]]
        vf = self.cbo_uf.get(); self.cbo_uf['values'] = f
        if vf in f: self.cbo_uf.set(vf)
        else: self.cbo_uf.current(0)
        conn.close()

    def load_dashboard(self):
        conn = conectar_db(); c = conn.cursor(); vs = self.base_sql()
        qb, pb = self.fq(f"SELECT COUNT(*) {vs}", [])
        c.execute(qb, tuple(pb)); self.lbl_total.config(text=str(c.fetchone()[0]))
        
        qb_pend, pb_pend = self.fq(f"SELECT COUNT(*) {vs} AND status IN ('Pendente', 'Recebida', 'Aguardando Retorno', 'Aguardando Área Demandante')", [])
        c.execute(qb_pend, tuple(pb_pend)); self.lbl_pendente.config(text=str(c.fetchone()[0]))
        
        qb_ana, pb_ana = self.fq(f"SELECT COUNT(*) {vs} AND status IN ('Em Análise', 'Em andamento')", [])
        c.execute(qb_ana, tuple(pb_ana)); self.lbl_analise.config(text=str(c.fetchone()[0]))
        
        qb_dev, pb_dev = self.fq(f"SELECT COUNT(*) {vs} AND status IN ('Devolvido', 'Cancelada')", [])
        c.execute(qb_dev, tuple(pb_dev)); self.lbl_devolvida.config(text=str(c.fetchone()[0]))
        
        qb_fin, pb_fin = self.fq(f"SELECT COUNT(*) {vs} AND status IN ('Finalizada', 'Concluído')", [])
        c.execute(qb_fin, tuple(pb_fin)); self.lbl_finalizada.config(text=str(c.fetchone()[0]))
        
        qs, ps = self.fq(f"SELECT data_entrada {vs} AND status NOT IN ('Finalizada', 'Concluído', 'Devolvido', 'Cancelada')", [])
        c.execute(qs, tuple(ps)); h, c10, c15, c20 = datetime.now().date(), 0, 0, 0
        for r in c.fetchall():
            try:
                d = (h - datetime.strptime(r[0], "%d/%m/%Y %H:%M").date()).days
                if d >= 20: c20 += 1
                elif d >= 15: c15 += 1
                elif d >= 10: c10 += 1
            except: pass
        self.lbl_s10.config(text=str(c10)); self.lbl_s15.config(text=str(c15)); self.lbl_s20.config(text=str(c20))
        
        self.ax1.clear(); self.ax2.clear(); self.ax3.clear(); self.ax4.clear()
        
        q1, p1 = self.fq(f"SELECT responsavel, COUNT(*) {vs} AND status IN ('Finalizada', 'Concluído', 'Devolvido', 'Cancelada') AND responsavel != 'Nenhum'", [])
        c.execute(q1 + " GROUP BY responsavel", tuple(p1)); d1 = c.fetchall()
        if d1:
            self.ax1.bar([r[0] for r in d1], [r[1] for r in d1], color="#5cb85c")
            self.ax1.set_title("Produtividade por Colab.", fontsize=11, fontweight='bold'); self.ax1.set_yticks(range(0, max([r[1] for r in d1]) + 2))
            self.ax1.set_xticklabels([r[0] for r in d1], rotation=45, ha='right', fontsize=8)
            
        q2, p2 = self.fq(f"SELECT tipo, COUNT(*) {vs}", [])
        c.execute(q2 + " GROUP BY tipo", tuple(p2)); d2 = c.fetchall()
        if d2:
            tps = [r[0] for r in d2]; qts = [r[1] for r in d2]
            cor = ['#008CBA', '#f0ad4e', '#5bc0de', '#5cb85c', '#d9534f', '#8e44ad', '#34495e', '#16a085', '#27ae60', '#c0392b'] * 3
            self.ax2.bar(tps, qts, color=cor[:len(tps)]); self.ax2.set_title("Distribuição por Tipo", fontsize=11, fontweight='bold')
            self.ax2.set_yticks(range(0, max(qts) + 2)); self.ax2.set_xticks(range(len(tps))); self.ax2.set_xticklabels(tps, rotation=45, ha='right', fontsize=8)
            
        q3, p3 = self.fq(f"SELECT status, COUNT(*) {vs}", [])
        c.execute(q3 + " GROUP BY status", tuple(p3)); d3 = c.fetchall()
        if d3:
            cm = {'Pendente': '#f0ad4e', 'Recebida': '#f0ad4e', 'Aguardando Retorno': '#f0ad4e', 'Aguardando Área Demandante': '#f0ad4e', 'Em Análise': '#5bc0de', 'Em andamento': '#5bc0de', 'Devolvido': '#8e44ad', 'Cancelada': '#8e44ad', 'Finalizada': '#5cb85c', 'Concluído': '#5cb85c'}
            self.dd(self.ax3, [r[1] for r in d3], [r[0] for r in d3], "Status Geral", colors=[cm.get(s[0], '#ccc') for s in d3])
            
        q4, p4 = self.fq(f"SELECT uf, COUNT(*) {vs} AND uf != '-' AND uf != ''", [])
        c.execute(q4 + " GROUP BY uf", tuple(p4)); d4 = c.fetchall()
        if d4:
            self.ax4.barh([r[0] for r in d4], [r[1] for r in d4], color="#008CBA"); self.ax4.set_title("Demandas de Rotina por UF", fontsize=11, fontweight='bold')
            
        self.canvas.draw(); conn.close()

    def exportar_dashboard_pdf(self):
        fp = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF files", "*.pdf")], title="Salvar Relatório")
        if not fp: return
        try:
            with PdfPages(fp) as pdf:
                fig = Figure(figsize=(8.27, 11.69), dpi=100)
                fig.text(0.5, 0.95, "RELATÓRIO GERENCIAL UNIFICADO", ha='center', fontsize=16, fontweight='bold')
                fig.text(0.5, 0.92, f"Data de Emissão: {datetime.now().strftime('%d/%m/%Y %H:%M')}", ha='center', fontsize=10)
                fig.text(0.5, 0.89, f"Filtros - Origem: {self.cbo_origem.get()} | Tipo: {self.cbo_tipo.get()} | Colab: {self.cbo_colab.get()} | UF: {self.cbo_uf.get()}", ha='center', fontsize=10, style='italic')
                fig.text(0.1, 0.81, f"TOTAL\n{self.lbl_total.cget('text')}", ha='center', fontsize=10, fontweight='bold', color='#008CBA')
                fig.text(0.3, 0.81, f"PENDENTES\n{self.lbl_pendente.cget('text')}", ha='center', fontsize=10, fontweight='bold', color='#f0ad4e')
                fig.text(0.5, 0.81, f"EM ANÁLISE\n{self.lbl_analise.cget('text')}", ha='center', fontsize=10, fontweight='bold', color='#5bc0de')
                fig.text(0.7, 0.81, f"DEV/CANCEL\n{self.lbl_devolvida.cget('text')}", ha='center', fontsize=10, fontweight='bold', color='#8e44ad')
                fig.text(0.9, 0.81, f"CONCLUÍDAS\n{self.lbl_finalizada.cget('text')}", ha='center', fontsize=10, fontweight='bold', color='#5cb85c')
                fig.text(0.5, 0.74, "_"*80, ha='center', color='gray')
                fig.text(0.5, 0.69, "ALERTAS DE SLA (NA FILA)", ha='center', fontsize=12, fontweight='bold', color='#d9534f')
                fig.text(0.3, 0.62, f"10 A 14 DIAS\n{self.lbl_s10.cget('text')}", ha='center', fontsize=12, fontweight='bold', color='#e0a800')
                fig.text(0.5, 0.62, f"15 A 19 DIAS\n{self.lbl_s15.cget('text')}", ha='center', fontsize=12, fontweight='bold', color='#fd7e14')
                fig.text(0.7, 0.62, f"20+ DIAS\n{self.lbl_s20.cget('text')}", ha='center', fontsize=12, fontweight='bold', color='#c9302c')
                fig.text(0.5, 0.54, "_"*80, ha='center', color='gray')
                
                conn = conectar_db(); c = conn.cursor(); vs = self.base_sql()
                q1, p1 = self.fq(f"SELECT responsavel, COUNT(*) {vs} AND status IN ('Finalizada', 'Concluído', 'Devolvido', 'Cancelada') AND responsavel != 'Nenhum'", [])
                c.execute(q1 + " GROUP BY responsavel", tuple(p1)); d1 = c.fetchall()
                a1 = fig.add_axes([0.10, 0.35, 0.35, 0.18])
                if d1: a1.bar([r[0] for r in d1], [r[1] for r in d1], color="#5cb85c"); a1.set_title("Produtividade", fontweight='bold', fontsize=9)
                
                q2, p2 = self.fq(f"SELECT tipo, COUNT(*) {vs}", [])
                c.execute(q2 + " GROUP BY tipo", tuple(p2)); d2 = c.fetchall()
                a2 = fig.add_axes([0.55, 0.35, 0.35, 0.18])
                if d2: 
                    a2.bar([r[0] for r in d2], [r[1] for r in d2], color='#008CBA'); a2.set_title("Tipos de Demanda", fontweight='bold', fontsize=9)
                    a2.set_xticks(range(len([r[0] for r in d2]))); a2.set_xticklabels([r[0] for r in d2], rotation=30, ha='right', fontsize=6)
                    
                q3, p3 = self.fq(f"SELECT status, COUNT(*) {vs}", [])
                c.execute(q3 + " GROUP BY status", tuple(p3)); d3 = c.fetchall()
                a3 = fig.add_axes([0.10, 0.05, 0.35, 0.22])
                if d3:
                    cm = {'Pendente': '#f0ad4e', 'Recebida': '#f0ad4e', 'Aguardando Retorno': '#f0ad4e', 'Aguardando Área Demandante': '#f0ad4e', 'Em Análise': '#5bc0de', 'Em andamento': '#5bc0de', 'Devolvido': '#8e44ad', 'Cancelada': '#8e44ad', 'Finalizada': '#5cb85c', 'Concluído': '#5cb85c'}
                    w, t, at = a3.pie([r[1] for r in d3], autopct='%1.1f%%', pctdistance=0.75, colors=[cm.get(x[0], '#ccc') for x in d3], wedgeprops=dict(width=0.4, edgecolor='w'))
                    a3.legend(w, [r[0] for r in d3], loc="upper center", bbox_to_anchor=(0.5, -0.1), fontsize=7, ncol=2); a3.set_title("Status Geral", fontweight='bold', fontsize=9)
                
                q4, p4 = self.fq(f"SELECT uf, COUNT(*) {vs} AND uf != '-' AND uf != ''", [])
                c.execute(q4 + " GROUP BY uf", tuple(p4)); d4 = c.fetchall()
                a4 = fig.add_axes([0.55, 0.05, 0.35, 0.18])
                if d4: a4.barh([r[0] for r in d4], [r[1] for r in d4], color="#008CBA"); a4.set_title("Demandas por UF", fontweight='bold', fontsize=9)

                conn.close(); pdf.savefig(fig)
            messagebox.showinfo("Sucesso", "Relatório PDF gerado com sucesso!")
            if platform.system() == 'Windows': os.startfile(fp)
            else: subprocess.call(('open', fp))
        except Exception as e: messagebox.showerror("Erro", f"Erro ao gerar o PDF:\n{e}")

    # ------- ROTINAS -------
    def setup_rotina(self):
        self.criar_btn_voltar(self.fr)
        fi = tk.Frame(self.fr, bg="#e9ecef", pady=15, padx=20, relief=tk.GROOVE, bd=1); fi.pack(fill=tk.X, padx=10, pady=5)
        tk.Label(fi, text="Importar PDFs:", font=("Helvetica", 11, "bold"), bg="#e9ecef").pack(side=tk.LEFT, padx=(0,10))
        self.btn_select_pasta = tk.Button(fi, text="Selecionar Pasta", command=self.select_folder_rotina, bg="#008CBA", fg="white", relief=tk.FLAT, padx=10); self.btn_select_pasta.pack(side=tk.LEFT)
        self.lbl_folder_rotina = tk.Label(fi, text="Nenhuma pasta", font=("Helvetica", 9), bg="#e9ecef", fg="#666"); self.lbl_folder_rotina.pack(side=tk.LEFT, padx=10)
        self.btn_analyze_rotina = tk.Button(fi, text="Processar PDFs", command=self.process_pdfs, bg="#4CAF50", fg="white", relief=tk.FLAT, padx=10, state=tk.DISABLED); self.btn_analyze_rotina.pack(side=tk.LEFT)
        self.folder_path_rotina = ""

        f_filtros = tk.Frame(self.fr, bg="#f4f4f9", pady=10); f_filtros.pack(fill=tk.X, padx=10)
        tk.Label(f_filtros, text="Tipo:", bg="#f4f4f9", font=("Helvetica", 10, "bold")).pack(side=tk.LEFT)
        self.cbo_rotina_tipo = ttk.Combobox(f_filtros, values=["Todos"] + TIPOS_ROTINA, state="readonly", width=15); self.cbo_rotina_tipo.current(0); self.cbo_rotina_tipo.pack(side=tk.LEFT, padx=5)
        tk.Label(f_filtros, text="Status:", bg="#f4f4f9", font=("Helvetica", 10, "bold")).pack(side=tk.LEFT, padx=(10,0))
        self.cbo_rotina_status = ttk.Combobox(f_filtros, values=["Todos", "Pendente", "Em Análise", "Devolvido", "Finalizada"], state="readonly", width=15); self.cbo_rotina_status.current(0); self.cbo_rotina_status.pack(side=tk.LEFT, padx=5)
        tk.Label(f_filtros, text="Resp.:", bg="#f4f4f9", font=("Helvetica", 10, "bold")).pack(side=tk.LEFT, padx=(10,0))
        self.cbo_rotina_resp = ttk.Combobox(f_filtros, state="readonly", width=15); self.cbo_rotina_resp.pack(side=tk.LEFT, padx=5)
        tk.Button(f_filtros, text="Aplicar Filtro", command=self.load_rotina, bg="#008CBA", fg="white", relief=tk.FLAT).pack(side=tk.LEFT, padx=10)
        
        fc = tk.Frame(self.fr, bg="#f4f4f9"); fc.pack(fill=tk.X, padx=10, pady=(0, 10))
        tk.Button(fc, text="Assumir", command=self.assumir_rotina, bg="#f0ad4e", fg="white", font=("Helvetica", 10, "bold"), relief=tk.FLAT).pack(side=tk.LEFT, padx=5)
        tk.Button(fc, text="Finalizar", command=lambda: self.marcar_rotina('Finalizada'), bg="#5cb85c", fg="white", font=("Helvetica", 10, "bold"), relief=tk.FLAT).pack(side=tk.LEFT, padx=5)
        tk.Button(fc, text="Devolver", command=lambda: self.marcar_rotina('Devolvido'), bg="#8e44ad", fg="white", font=("Helvetica", 10, "bold"), relief=tk.FLAT).pack(side=tk.LEFT, padx=5)
        self.btn_excluir_rotina = tk.Button(fc, text="Excluir", command=self.excluir_rotina, bg="#d9534f", fg="white", font=("Helvetica", 10, "bold"), relief=tk.FLAT)
        
        cols = ("ID", "Data Entrada", "Arquivo", "Prestador", "CNPJ", "Município", "UF", "Demanda", "Tipo Neg.", "Alçada", "Imp. %", "Imp. R$", "NUP e-DOC", "FOP Aut.", "Data Aprov.", "Status", "Resp.", "SLA Geral", "SLA Gered")
        tf = tk.Frame(self.fr); tf.pack(fill=tk.BOTH, expand=True, padx=10, pady=(0, 10))
        sy = ttk.Scrollbar(tf, orient="vertical"); sy.pack(side=tk.RIGHT, fill=tk.Y)
        sx = ttk.Scrollbar(tf, orient="horizontal"); sx.pack(side=tk.BOTTOM, fill=tk.X)
        self.tree_rotina = ttk.Treeview(tf, columns=cols, show="headings", height=15, yscrollcommand=sy.set, xscrollcommand=sx.set)
        
        for c in cols:
            self.tree_rotina.heading(c, text=c)
            w = 110
            if c == "ID": w = 40
            elif c in ("Arquivo", "Prestador"): w = 250
            elif c in ("Demanda", "Data Entrada"): w = 120
            elif c in ("Município", "UF", "SLA Geral", "SLA Gered"): w = 90
            self.tree_rotina.column(c, width=w, stretch=False, anchor=tk.CENTER if c not in ["Arquivo", "Prestador", "Município"] else tk.W)
            
        self.tree_rotina.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        sy.config(command=self.tree_rotina.yview); sx.config(command=self.tree_rotina.xview)
        self.tree_rotina.bind("<Double-1>", lambda event: self.abrir_detalhes_rotina())

    def load_rotina(self):
        for item in self.tree_rotina.get_children(): self.tree_rotina.delete(item)
        query = "SELECT id, data_entrada, filename, prestador, cnpj, municipio, uf, type, tipo_negociacao, alcada, impacto_perc, impacto_rs, nup_edoc, fop_autorizacao, data_aprovacao, status, assigned_to, sla_geral, sla_gered FROM demands WHERE 1=1"
        params = []
        if self.cbo_rotina_tipo.get() != "Todos": query += " AND type=?"; params.append(self.cbo_rotina_tipo.get())
        if self.cbo_rotina_status.get() != "Todos": query += " AND status=?"; params.append(self.cbo_rotina_status.get())
        if self.cbo_rotina_resp.get() != "Todos" and self.cbo_rotina_resp.get() != "": query += " AND assigned_to=?"; params.append(self.cbo_rotina_resp.get())
        query += " ORDER BY id DESC"
        
        conn = conectar_db(); c = conn.cursor()
        c.execute(replace_placeholders(query), tuple(params)); hoje = datetime.now().date()
        for row in c.fetchall():
            rl = list(row)
            if rl[15] not in ('Finalizada', 'Devolvido'):
                try: rl[18] = (hoje - datetime.strptime(rl[1], "%d/%m/%Y %H:%M").date()).days
                except: rl[18] = "-"
                try: rl[17] = (hoje - datetime.strptime(rl[14], "%d/%m/%Y").date()).days if rl[14] else "-"
                except: rl[17] = "-"
            rl = ["-" if v is None or str(v).strip() == "" else v for v in rl]
            self.tree_rotina.insert("", tk.END, values=rl)
        conn.close()

    def select_folder_rotina(self):
        self.folder_path_rotina = filedialog.askdirectory(title="Selecione a pasta")
        if self.folder_path_rotina: self.lbl_folder_rotina.config(text=f"Pasta: {self.folder_path_rotina}"); self.btn_analyze_rotina.config(state=tk.NORMAL)

    def process_pdfs(self):
        pdf_files = [f for f in os.listdir(self.folder_path_rotina) if f.lower().endswith('.pdf') and not f.startswith('~$')]
        if not pdf_files: return messagebox.showinfo("Aviso", "Nenhum PDF válido encontrado na pasta.")
        
        kw = {"REAJUSTE": "Reajuste", "NOVO CONTRATO": "Novo Contrato", "INCLUSÃO": "Inclusão", "INCLUSAO": "Inclusão", "EXCLUSÃO": "Exclusão", "EXCLUSAO": "Exclusão", "AJUSTE": "Ajuste", "EXTENSÃO": "Extensão", "EXTENSAO": "Extensão", "DESCREDENCIAMENTO": "Descredenciamento", "DESCRECENCIAMENTO": "Descredenciamento"}
        valid_ufs = {'AC', 'AL', 'AM', 'AP', 'BA', 'CE', 'DF', 'ES', 'GO', 'MA', 'MT', 'MS', 'MG', 'PA', 'PB', 'PR', 'PE', 'PI', 'RJ', 'RN', 'RS', 'RO', 'RR', 'SC', 'SP', 'SE', 'TO'}
        orgaos_inv = {'SSP', 'SESP', 'CRM', 'CRO', 'COREN', 'CRF', 'PC', 'SDS', 'PM', 'IFP', 'DETRAN', 'DIC'}
        conn = conectar_db(); c = conn.cursor()
        c.execute("SELECT username FROM users WHERE role='user'"); colabs = [row[0] for row in c.fetchall()]
        if not colabs: c.execute("SELECT username FROM users"); colabs = [row[0] for row in c.fetchall()]
        
        count = 0
        for f_name in pdf_files:
            c.execute(replace_placeholders("SELECT id FROM demands WHERE filename=?"), (f_name,))
            if c.fetchone(): continue
            f_path = os.path.join(self.folder_path_rotina, f_name)
            fu = f_name.upper(); full_txt = ""
            dem = "Não Identificada"
            for k, v in kw.items():
                if k in fu: dem = v; break
                    
            f_clean = re.sub(r'\.pdf$', '', fu, flags=re.IGNORECASE)
            docs_brutos = re.findall(r'\d{2}[.\s_-]*\d{3}[.\s_-]*\d{3}[.\s/-]*\d{4}[-\s_]*\d{2}|\d{3}[.\s_-]*\d{3}[.\s_-]*\d{3}[-\s_]*\d{2}|\b\d{14}\b|\b\d{11}\b', f_clean)
            cnpj_p = "Não encontrado"
            
            for c_ in docs_brutos:
                cn = re.sub(r'\D', '', c_)
                if len(cn) == 14 and cn != "18275071000162":
                    cnpj_p = f"{cn[:2]}.{cn[2:5]}.{cn[5:8]}/{cn[8:12]}-{cn[12:]}"
                    f_clean = f_clean.replace(c_, ''); break
                elif len(cn) == 11:
                    cnpj_p = f"{cn[:3]}.{cn[3:6]}.{cn[6:9]}-{cn[9:]}"
                    f_clean = f_clean.replace(c_, ''); break
                    
            mun, uf, prest = "-", "-", "Não encontrado"
            for k in kw.keys(): f_clean = re.sub(k, '', f_clean, flags=re.IGNORECASE)
            parts = [p.strip() for p in f_clean.split('-') if p.strip()]
            
            if len(parts) >= 2:
                p1 = parts[-1].upper(); p2 = parts[-2].upper() if len(parts) > 1 else ""
                if p1 in valid_ufs:
                    uf = p1; mun = p2 if len(parts) >= 3 else "-"; prest = " ".join(parts[:-2]).strip() if len(parts) >= 3 else p2
                elif p2 in valid_ufs:
                    uf = p2; mun = p1; prest = " ".join(parts[:-2]).strip()
                else:
                    m_uf_end = re.search(r'\b([A-Z]{2})\b$', p1)
                    if m_uf_end and m_uf_end.group(1) in valid_ufs:
                        uf = m_uf_end.group(1); mun = re.sub(r'\b' + uf + r'\b$', '', p1).strip(); prest = " ".join(parts[:-1]).strip()
                    else: prest = " ".join(parts).strip()
            elif len(parts) == 1:
                m_uf_end = re.search(r'\b([A-Z]{2})\b$', parts[0].upper())
                if m_uf_end and m_uf_end.group(1) in valid_ufs:
                    uf = m_uf_end.group(1); prest = re.sub(r'\b' + uf + r'\b$', '', parts[0]).strip()
                else: prest = parts[0]
                
            prest = re.sub(r'^[-_\s]+|[-_\s]+$', '', prest); prest = re.sub(r'\s+', ' ', prest).strip()
            if not prest or len(prest) <= 3: prest = "Não encontrado"
            
            if prest == "Não encontrado" or uf == "-" or cnpj_p == "Não encontrado":
                try:
                    with open(f_path, 'rb') as f:
                        for p in PyPDF2.PdfReader(f).pages: full_txt += (p.extract_text() or "") + " "
                    
                    if cnpj_p == "Não encontrado":
                        docs_text = re.findall(r'\d{2}[.\s]*\d{3}[.\s]*\d{3}[/\s]*\d{4}[-\s]*\d{2}|\d{3}[.\s]*\d{3}[.\s]*\d{3}[-\s]*\d{2}|\b\d{14}\b|\b\d{11}\b', full_txt)
                        for c_ in docs_text:
                            cn = re.sub(r'\D', '', c_)
                            if len(cn) == 14 and cn != "18275071000162": 
                                cnpj_p = f"{cn[:2]}.{cn[2:5]}.{cn[5:8]}/{cn[8:12]}-{cn[12:]}"; break
                            elif len(cn) == 11:
                                cnpj_p = f"{cn[:3]}.{cn[3:6]}.{cn[6:9]}-{cn[9:]}"; break

                    if prest == "Não encontrado":
                        match_p = re.search(r'outro lado[,]?\s*(.*?)(?:[,]?\s+inscrita|[,]?\s+localizada|[,]?\s+nome fantasia|[,]?\s+CNPJ|[,]?\s+estabelecimento)', full_txt, re.IGNORECASE | re.DOTALL)
                        if match_p:
                            nome_s = match_p.group(1).replace('\n', ' '); nome_l = re.sub(r'\s+', ' ', nome_s).strip(',. ')
                            prest = nome_l[:-1].strip() if nome_l.endswith(',') else nome_l

                    if uf == "-":
                        ftc = re.sub(r'S\s+ESP', 'SESP', full_txt.upper()); ftc = re.sub(r'S\s+SP', 'SSP', ftc)
                        ftc = re.sub(r'\b(CONCLUI|TRATA|APLICA|REFERE|VERIFICA|DESTACA|OBSERVA|N[AÃ]O)[/-]\s*SE\b', '', ftc)
                        for loc in re.findall(r'([A-ZÀ-Úa-zà-ú\s]+)[/-]\s?([A-Z]{2})\b', ftc):
                            mb = loc[0].strip(); est = loc[1].strip(); mw = mb.split()
                            if len(mw) > 4: mb = " ".join(mw[-4:])
                            mun_c = mb.replace(',', '').strip(); up = mun_c.split()[-1] if mun_c.split() else ""
                            if est in valid_ufs and up not in orgaos_inv and up not in {'SE', 'CONCLUI', 'TRATA', 'NÃO', 'NAO'}:
                                if not (est == "DF" and "BRASÍLIA" in mun_c): mun, uf = mun_c, est; break
                except: pass
                            
            c.execute(replace_placeholders("""INSERT INTO demands (filename, filepath, type, status, assigned_to, prestador, cnpj, municipio, uf, data_entrada) VALUES (?, ?, ?, 'Pendente', ?, ?, ?, ?, ?, ?)"""), 
                         (f_name, f_path, dem, colabs[count % len(colabs)] if colabs else "Nenhum", prest, cnpj_p, mun, uf, datetime.now().strftime("%d/%m/%Y %H:%M")))
            count += 1
            
        conn.commit(); conn.close(); messagebox.showinfo("Concluído", f"{count} novos PDFs processados!")
        self.load_rotina()

    def assumir_rotina(self):
        sel = self.tree_rotina.selection()
        if not sel: return messagebox.showwarning("Aviso", "Selecione.")
        d_id, st = self.tree_rotina.item(sel[0])['values'][0], self.tree_rotina.item(sel[0])['values'][15] 
        if st in ('Finalizada', 'Devolvido'): return messagebox.showwarning("Aviso", "Concluída.")
        conn = conectar_db(); c = conn.cursor()
        c.execute(replace_placeholders("UPDATE demands SET status='Em Análise', assigned_to=? WHERE id=?"), (self.controller.current_user, d_id))
        conn.commit(); conn.close(); self.load_rotina()

    def marcar_rotina(self, nst):
        sel = self.tree_rotina.selection()
        if not sel: return messagebox.showwarning("Aviso", "Selecione.")
        d_id, resp = self.tree_rotina.item(sel[0])['values'][0], self.tree_rotina.item(sel[0])['values'][16] 
        if resp != self.controller.current_user and self.controller.current_role != 'admin': return messagebox.showerror("Erro", "Apenas a sua demanda.")
        conn = conectar_db(); c = conn.cursor(); c.execute(replace_placeholders("SELECT status, data_entrada, data_aprovacao FROM demands WHERE id=?"), (d_id,))
        row = c.fetchone()
        if row[0] in ('Finalizada', 'Devolvido'): conn.close(); return messagebox.showwarning("Aviso", f"Já está {row[0]}.")
        hj = datetime.now()
        try: d_e = datetime.strptime(row[1], "%d/%m/%Y %H:%M").date()
        except: d_e = None
        try: d_a = datetime.strptime(row[2], "%d/%m/%Y").date() if row[2] else None
        except: d_a = None
        c.execute(replace_placeholders("""UPDATE demands SET status=?, data_finalizacao=?, sla_geral=?, sla_gered=? WHERE id=?"""), (nst, hj.strftime("%d/%m/%Y %H:%M"), (hj.date() - d_a).days if d_a else 0, (hj.date() - d_e).days if d_e else 0, d_id))
        conn.commit(); conn.close(); self.load_rotina(); messagebox.showinfo("Sucesso", f"Marcada como {nst}!")

    def excluir_rotina(self):
        sel = self.tree_rotina.selection()
        if not sel: return
        if messagebox.askyesno("Confirmar", "Deseja excluir?"):
            conn = conectar_db(); c = conn.cursor()
            c.execute(replace_placeholders("DELETE FROM demands WHERE id=?"), (self.tree_rotina.item(sel[0])['values'][0],))
            conn.commit(); conn.close(); self.load_rotina()

    def abrir_detalhes_rotina(self):
        sel = self.tree_rotina.selection()
        if not sel: return
        d_id = self.tree_rotina.item(sel[0])['values'][0]
        conn = conectar_db(); c = conn.cursor(); c.execute(replace_placeholders("SELECT * FROM demands WHERE id=?"), (d_id,))
        row = c.fetchone(); cols = [desc[0] for desc in c.description]
        c.execute("SELECT username FROM users"); aus = [r[0] for r in c.fetchall()]; conn.close()
        if not row: return
        d_data = dict(zip(cols, row))
        
        m = tk.Toplevel(self); m.title(f"Rotina #{d_id}"); m.geometry("750x650")
        m.configure(bg="#f4f4f9"); m.transient(self.controller.root); m.grab_set() 
        tk.Label(m, text=f"Tipo: {d_data['type']} | Status Atual: {d_data['status']}", bg="#333", fg="white", font=("Helvetica", 12, "bold")).pack(fill=tk.X, pady=10)
        ff = tk.Frame(m, bg="#f4f4f9", pady=20, padx=20); ff.pack(fill=tk.BOTH, expand=True)
        
        v = {
            'tn': tk.StringVar(value=d_data.get('tipo_negociacao', '') or ''), 'al': tk.StringVar(value=d_data.get('alcada', '') or ''),
            'ip': tk.StringVar(value=d_data.get('impacto_perc', '') or ''), 'ir': tk.StringVar(value=d_data.get('impacto_rs', '') or ''),
            'nu': tk.StringVar(value=d_data.get('nup_edoc', '') or ''), 'fo': tk.StringVar(value=d_data.get('fop_autorizacao', '') or ''),
            'da': tk.StringVar(value=d_data.get('data_aprovacao', '') or ''), 'at': tk.StringVar(value=d_data.get('assigned_to', '') or ''),
            'st': tk.StringVar(value=d_data.get('status', ''))
        }
        
        def fd(e):
            if e.keysym in ('BackSpace', 'Delete', 'Left', 'Right', 'Tab'): return
            t = v['da'].get().replace('/', '')
            v['da'].set("".join([c + ("/" if i in [1, 3] else "") for i, c in enumerate(filter(str.isdigit, t))])[:10]); e.widget.icursor(tk.END)
            
        def fp(e):
            if e.keysym in ('BackSpace', 'Delete', 'Left', 'Right', 'Tab'): return
            t = v['ip'].get().replace('%', '').strip(); t = ''.join([c for c in t if c.isdigit() or c == ','])
            if t: v['ip'].set(t + '%'); e.widget.icursor(len(t))
            else: v['ip'].set('')
            
        def fr(e):
            if e.keysym in ('BackSpace', 'Delete', 'Left', 'Right', 'Tab'): return
            t = ''.join(filter(str.isdigit, v['ir'].get()))
            if not t: v['ir'].set(''); return
            v['ir'].set(f"R$ {float(t)/100:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.')); e.widget.icursor(tk.END)
        
        tk.Label(ff, text="Tipo Negociação:", bg="#f4f4f9").grid(row=0, column=0, sticky=tk.W, pady=5)
        ttk.Combobox(ff, textvariable=v['tn'], values=["", "ESTADO", "DIFERENCIADA", "MISTA"], state="readonly", width=25).grid(row=0, column=1, padx=10)
        tk.Label(ff, text="Alçada:", bg="#f4f4f9").grid(row=1, column=0, sticky=tk.W, pady=5)
        ttk.Combobox(ff, textvariable=v['al'], values=["", "GERED", "CODER", "REGIONAL", "CONEC", "DIOPE"], state="readonly", width=25).grid(row=1, column=1, padx=10)
        tk.Label(ff, text="Impacto %:", bg="#f4f4f9").grid(row=2, column=0, sticky=tk.W, pady=5)
        ep = ttk.Entry(ff, textvariable=v['ip'], width=28); ep.grid(row=2, column=1, padx=10); ep.bind('<KeyRelease>', fp)
        tk.Label(ff, text="Impacto R$:", bg="#f4f4f9").grid(row=3, column=0, sticky=tk.W, pady=5)
        er = ttk.Entry(ff, textvariable=v['ir'], width=28); er.grid(row=3, column=1, padx=10); er.bind('<KeyRelease>', fr)
        if d_data['type'] not in ('Reajuste', 'Ajuste'): ep.config(state=tk.DISABLED); er.config(state=tk.DISABLED)
            
        tk.Label(ff, text="NUP e-DOC:", bg="#f4f4f9").grid(row=0, column=2, sticky=tk.W, pady=5, padx=(20,0))
        ttk.Entry(ff, textvariable=v['nu'], width=28).grid(row=0, column=3, padx=10)
        tk.Label(ff, text="FOP Autorização:", bg="#f4f4f9").grid(row=1, column=2, sticky=tk.W, pady=5, padx=(20,0))
        ttk.Entry(ff, textvariable=v['fo'], width=28).grid(row=1, column=3, padx=10)
        tk.Label(ff, text="Data Aprov. (DD/MM/AAAA):", bg="#f4f4f9").grid(row=2, column=2, sticky=tk.W, pady=5, padx=(20,0))
        edt = ttk.Entry(ff, textvariable=v['da'], width=28); edt.grid(row=2, column=3, padx=10); edt.bind('<KeyRelease>', fd)
        tk.Label(ff, text="Colab. Responsável:", bg="#f4f4f9", fg="#d9534f", font=("Helvetica", 9, "bold")).grid(row=3, column=2, sticky=tk.W, pady=5, padx=(20,0))
        cb_resp = ttk.Combobox(ff, textvariable=v['at'], values=aus, state="readonly", width=25); cb_resp.grid(row=3, column=3, padx=10)
        if self.controller.current_role != 'admin': cb_resp.config(state=tk.DISABLED)

        tk.Label(ff, text="Mudar Status:", bg="#f4f4f9", font=("Helvetica", 9, "bold")).grid(row=4, column=0, sticky=tk.W, pady=(15,5))
        ttk.Combobox(ff, textvariable=v['st'], values=["Pendente", "Em Análise", "Devolvido", "Finalizada"], state="readonly", width=25).grid(row=4, column=1, padx=10, pady=(15,5))
        tk.Label(ff, text="Observação:", bg="#f4f4f9", font=("Helvetica", 9, "bold")).grid(row=5, column=0, sticky=tk.NW, pady=5)
        to = tk.Text(ff, height=4, width=60, font=("Helvetica", 9)); to.grid(row=5, column=1, columnspan=3, padx=10, pady=5, sticky=tk.W); to.insert("1.0", d_data.get('observacao', '') or '')
        
        tk.Label(ff, text=f"Prestador: {d_data.get('prestador','-')} | CNPJ: {d_data.get('cnpj','-')} | Data Entrada: {d_data.get('data_entrada','-')}", bg="#e9ecef").grid(row=6, column=0, columnspan=4, pady=15, sticky="we")
        
        fb = tk.Frame(m, bg="#f4f4f9"); fb.pack(fill=tk.X, pady=10)
        tk.Button(fb, text="Ver PDF", command=lambda: self.abrir_pdf_path(d_data.get('filepath')), bg="#5bc0de", fg="white", font=("Helvetica", 10, "bold"), relief=tk.FLAT, padx=10).pack(side=tk.LEFT, padx=20)
        
        def salvar():
            st = v['st'].get()
            conn = conectar_db(); c = conn.cursor()
            c.execute(replace_placeholders("SELECT status, data_entrada, data_aprovacao FROM demands WHERE id=?"), (d_id,))
            st_ant, d_e, d_a = c.fetchone()
            if st in ['Finalizada', 'Devolvido'] and st_ant not in ['Finalizada', 'Devolvido']:
                hj = datetime.now()
                try: e_dt = datetime.strptime(d_e, "%d/%m/%Y %H:%M").date()
                except: e_dt = None
                try: a_dt = datetime.strptime(d_a, "%d/%m/%Y").date() if d_a else None
                except: a_dt = None
                c.execute(replace_placeholders("UPDATE demands SET status=?, data_finalizacao=?, sla_geral=?, sla_gered=? WHERE id=?"), (st, hj.strftime("%d/%m/%Y %H:%M"), (hj.date() - a_dt).days if a_dt else 0, (hj.date() - e_dt).days if e_dt else 0, d_id))
            elif st not in ['Finalizada', 'Devolvido']:
                c.execute(replace_placeholders("UPDATE demands SET status=?, data_finalizacao=NULL, sla_geral=NULL, sla_gered=NULL WHERE id=?"), (st, d_id))
            c.execute(replace_placeholders("""UPDATE demands SET tipo_negociacao=?, alcada=?, impacto_perc=?, impacto_rs=?, nup_edoc=?, fop_autorizacao=?, data_aprovacao=?, assigned_to=?, observacao=? WHERE id=?"""), 
                      (v['tn'].get(), v['al'].get(), v['ip'].get(), v['ir'].get(), v['nu'].get(), v['fo'].get(), v['da'].get(), v['at'].get(), to.get("1.0", tk.END).strip(), d_id))
            conn.commit(); conn.close()
            m.destroy(); self.load_rotina(); messagebox.showinfo("Sucesso", "Atualizado!")
            
        tk.Button(fb, text="Salvar", command=salvar, bg="#4CAF50", fg="white", font=("Helvetica", 10, "bold"), relief=tk.FLAT, padx=10).pack(side=tk.RIGHT, padx=20)

    def abrir_pdf_path(self, filepath):
        if filepath and os.path.exists(filepath):
            try:
                if platform.system() == 'Windows': os.startfile(filepath)
                else: subprocess.call(('open' if platform.system() == 'Darwin' else 'xdg-open', filepath))
            except Exception as e: messagebox.showerror("Erro", f"Não foi possível abrir o arquivo.\n{e}")
        else: messagebox.showerror("Erro", "Arquivo PDF não encontrado.")

    # ------- AVULSAS -------
    def setup_avulsas(self):
        self.criar_btn_voltar(self.fa)
        ff = tk.Frame(self.fa, bg="#f4f4f9", pady=10); ff.pack(fill=tk.X, padx=10)
        tk.Label(ff, text="Status:", bg="#f4f4f9", font=("Helvetica", 10, "bold")).pack(side=tk.LEFT)
        self.cbo_av_st = ttk.Combobox(ff, values=["Todos"] + STATUS_AVULSA, state="readonly", width=15); self.cbo_av_st.current(0); self.cbo_av_st.pack(side=tk.LEFT, padx=5)
        tk.Label(ff, text="Resp.:", bg="#f4f4f9", font=("Helvetica", 10, "bold")).pack(side=tk.LEFT, padx=(10,0))
        self.cbo_av_resp = ttk.Combobox(ff, state="readonly", width=15); self.cbo_av_resp.pack(side=tk.LEFT, padx=5)
        
        tk.Button(ff, text="Aplicar Filtro", command=self.load_avulsas, bg="#008CBA", fg="white", relief=tk.FLAT).pack(side=tk.LEFT, padx=10)
        tk.Button(ff, text="+ Nova Demanda Avulsa", command=lambda: self.abrir_modal_avulsa(), bg="#4CAF50", fg="white", font=("Helvetica", 10, "bold"), relief=tk.FLAT).pack(side=tk.RIGHT, padx=10)
        
        fc = tk.Frame(self.fa, bg="#f4f4f9"); fc.pack(fill=tk.X, padx=10, pady=(0, 10))
        tk.Button(fc, text="Assumir", command=self.assumir_avulsa, bg="#f0ad4e", fg="white", font=("Helvetica", 10, "bold"), relief=tk.FLAT).pack(side=tk.LEFT, padx=5)
        tk.Button(fc, text="Concluir", command=lambda: self.marcar_avulsa('Concluído'), bg="#5cb85c", fg="white", font=("Helvetica", 10, "bold"), relief=tk.FLAT).pack(side=tk.LEFT, padx=5)
        tk.Button(fc, text="Cancelar", command=lambda: self.marcar_avulsa('Cancelada'), bg="#8e44ad", fg="white", font=("Helvetica", 10, "bold"), relief=tk.FLAT).pack(side=tk.LEFT, padx=5)
        self.btn_excluir_avulsa = tk.Button(fc, text="Excluir", command=self.excluir_avulsa, bg="#d9534f", fg="white", font=("Helvetica", 10, "bold"), relief=tk.FLAT)
        
        cols = ("ID", "TIPO DE DEMANDA", "DEMANDA/ASSUNTO", "SOLICITANTE", "DATA DE ENTRADA", "PRAZO", "STATUS", "RESPONSÁVEL", "DATA DE CONCLUSÃO", "OBSERVAÇÃO")
        tf = tk.Frame(self.fa); tf.pack(fill=tk.BOTH, expand=True, padx=10, pady=(0, 10))
        sy = ttk.Scrollbar(tf, orient="vertical"); sy.pack(side=tk.RIGHT, fill=tk.Y)
        sx = ttk.Scrollbar(tf, orient="horizontal"); sx.pack(side=tk.BOTTOM, fill=tk.X)
        self.tree_avulsas = ttk.Treeview(tf, columns=cols, show="headings", height=15, yscrollcommand=sy.set, xscrollcommand=sx.set)
        
        for c in cols:
            self.tree_avulsas.heading(c, text=c)
            w = 120
            if c == "ID": w = 40
            elif c == "DEMANDA/ASSUNTO": w = 250
            elif c == "OBSERVAÇÃO": w = 200
            self.tree_avulsas.column(c, width=w, stretch=False, anchor=tk.CENTER if c not in ["DEMANDA/ASSUNTO", "OBSERVAÇÃO"] else tk.W)
            
        self.tree_avulsas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        sy.config(command=self.tree_avulsas.yview); sx.config(command=self.tree_avulsas.xview)
        self.tree_avulsas.bind("<Double-1>", lambda e: self.abrir_modal_avulsa(True))

    def load_avulsas(self):
        for item in self.tree_avulsas.get_children(): self.tree_avulsas.delete(item)
        query = "SELECT * FROM demandas_avulsas WHERE 1=1"; params = []
        if self.cbo_av_st.get() != "Todos": query += " AND status=?"; params.append(self.cbo_av_st.get())
        if self.cbo_av_resp.get() != "Todos" and self.cbo_av_resp.get() != "": query += " AND responsavel=?"; params.append(self.cbo_av_resp.get())
        query += " ORDER BY id DESC"
        conn = conectar_db(); c = conn.cursor(); c.execute(replace_placeholders(query), tuple(params))
        for row in c.fetchall():
            self.tree_avulsas.insert("", tk.END, values=["-" if v is None or str(v).strip() == "" else v for v in list(row)])
        conn.close()

    def assumir_avulsa(self):
        sel = self.tree_avulsas.selection()
        if not sel: return messagebox.showwarning("Aviso", "Selecione.")
        d_id, st = self.tree_avulsas.item(sel[0])['values'][0], self.tree_avulsas.item(sel[0])['values'][6] 
        if st in ('Concluído', 'Cancelada'): return messagebox.showwarning("Aviso", "Já concluída.")
        conn = conectar_db(); c = conn.cursor()
        c.execute(replace_placeholders("UPDATE demandas_avulsas SET status='Em Análise', responsavel=? WHERE id=?"), (self.controller.current_user, d_id))
        conn.commit(); conn.close(); self.load_avulsas()

    def marcar_avulsa(self, nst):
        sel = self.tree_avulsas.selection()
        if not sel: return messagebox.showwarning("Aviso", "Selecione uma demanda.")
        d_id, resp = self.tree_avulsas.item(sel[0])['values'][0], self.tree_avulsas.item(sel[0])['values'][7] 
        if resp != self.controller.current_user and self.controller.current_role != 'admin': return messagebox.showerror("Erro", "Apenas a sua.")
        conn = conectar_db(); c = conn.cursor()
        c.execute(replace_placeholders("SELECT status FROM demandas_avulsas WHERE id=?"), (d_id,))
        if c.fetchone()[0] in ('Concluído', 'Cancelada'): conn.close(); return messagebox.showwarning("Aviso", "Demanda já concluída.")
        dt_c = datetime.now().strftime("%d/%m/%Y") if nst in ('Concluído', 'Cancelada') else ""
        c.execute(replace_placeholders("UPDATE demandas_avulsas SET status=?, data_conclusao=? WHERE id=?"), (nst, dt_c, d_id))
        conn.commit(); conn.close(); self.load_avulsas(); messagebox.showinfo("Sucesso", f"Marcada como {nst}!")

    def excluir_avulsa(self):
        if self.controller.current_role != 'admin': return
        sel = self.tree_avulsas.selection()
        if not sel: return
        if messagebox.askyesno("Confirmar Exclusão", "Tem certeza que deseja excluir?"):
            conn = conectar_db(); c = conn.cursor()
            c.execute(replace_placeholders("DELETE FROM demandas_avulsas WHERE id=?"), (self.tree_avulsas.item(sel[0])['values'][0],))
            conn.commit(); conn.close(); self.load_avulsas()

    def abrir_modal_avulsa(self, editar=False):
        conn = conectar_db(); c = conn.cursor()
        c.execute("SELECT username FROM users"); aus = [r[0] for r in c.fetchall()]; conn.close()
        
        d_id = None
        v = {'td': tk.StringVar(), 'as': tk.StringVar(), 'so': tk.StringVar(), 'pr': tk.StringVar(), 'st': tk.StringVar(value='Pendente'), 're': tk.StringVar(value='Nenhum')}; obs_t = ''
        
        if editar:
            sel = self.tree_avulsas.selection()
            if not sel: return
            d_id = self.tree_avulsas.item(sel[0])['values'][0]
            conn = conectar_db(); c = conn.cursor(); c.execute(replace_placeholders("SELECT * FROM demandas_avulsas WHERE id=?"), (d_id,))
            row = c.fetchone(); conn.close()
            if row:
                v['td'].set(row[1] or ''); v['as'].set(row[2] or ''); v['so'].set(row[3] or '')
                v['pr'].set(row[5] or ''); v['st'].set(row[6] or ''); v['re'].set(row[7] or ''); obs_t = row[9] or ''
        
        m = tk.Toplevel(self); m.title(f"Avulsa #{d_id if d_id else 'Nova'}"); m.geometry("650x550")
        m.configure(bg="#f4f4f9"); m.transient(self.controller.root); m.grab_set() 
        tk.Label(m, text="Formulário de Demanda Avulsa", bg="#333", fg="white", font=("Helvetica", 12, "bold")).pack(fill=tk.X, pady=10)
        ff = tk.Frame(m, bg="#f4f4f9", pady=20, padx=20); ff.pack(fill=tk.BOTH, expand=True)
        
        def fd(e):
            if e.keysym in ('BackSpace', 'Delete', 'Left', 'Right', 'Tab'): return
            t = v['pr'].get().replace('/', '')
            v['pr'].set("".join([c + ("/" if i in [1, 3] else "") for i, c in enumerate(filter(str.isdigit, t))])[:10]); e.widget.icursor(tk.END)
        
        tk.Label(ff, text="Tipo de Demanda:", bg="#f4f4f9").grid(row=0, column=0, sticky=tk.W, pady=5)
        ttk.Combobox(ff, textvariable=v['td'], values=TIPOS_AVULSA, width=35).grid(row=0, column=1, padx=10, sticky=tk.W)
        tk.Label(ff, text="Demanda/Assunto:", bg="#f4f4f9").grid(row=1, column=0, sticky=tk.W, pady=5)
        ttk.Entry(ff, textvariable=v['as'], width=50).grid(row=1, column=1, padx=10, sticky=tk.W)
        tk.Label(ff, text="Solicitante:", bg="#f4f4f9").grid(row=2, column=0, sticky=tk.W, pady=5)
        ttk.Combobox(ff, textvariable=v['so'], values=SOLICITANTES, width=35).grid(row=2, column=1, padx=10, sticky=tk.W)
        tk.Label(ff, text="Prazo (DD/MM/AAAA):", bg="#f4f4f9").grid(row=3, column=0, sticky=tk.W, pady=5)
        edt = ttk.Entry(ff, textvariable=v['pr'], width=15); edt.grid(row=3, column=1, padx=10, sticky=tk.W); edt.bind('<KeyRelease>', fd)
        tk.Label(ff, text="Responsável:", bg="#f4f4f9").grid(row=4, column=0, sticky=tk.W, pady=5)
        cb_r = ttk.Combobox(ff, textvariable=v['re'], values=aus, state="readonly", width=15); cb_r.grid(row=4, column=1, padx=10, sticky=tk.W)
        if self.controller.current_role != 'admin': cb_r.config(state=tk.DISABLED)
        tk.Label(ff, text="Status:", bg="#f4f4f9").grid(row=5, column=0, sticky=tk.W, pady=5)
        ttk.Combobox(ff, textvariable=v['st'], values=STATUS_AVULSA, state="readonly", width=15).grid(row=5, column=1, padx=10, sticky=tk.W)
        tk.Label(ff, text="Observação:", bg="#f4f4f9").grid(row=6, column=0, sticky=tk.NW, pady=5)
        txt_obs = tk.Text(ff, height=5, width=45, font=("Helvetica", 9)); txt_obs.grid(row=6, column=1, padx=10, pady=5, sticky=tk.W); txt_obs.insert("1.0", obs_t)
        
        def salvar():
            dt_ent = datetime.now().strftime("%d/%m/%Y")
            dt_c = dt_ent if v['st'].get() in ('Concluído', 'Cancelada') else ""
            conn = conectar_db(); c = conn.cursor()
            if editar:
                if v['st'].get() not in ('Concluído', 'Cancelada'): c.execute(replace_placeholders("SELECT data_conclusao FROM demandas_avulsas WHERE id=?"), (d_id,)); dt_c = c.fetchone()[0] or ""
                c.execute(replace_placeholders("""UPDATE demandas_avulsas SET tipo_demanda=?, assunto=?, solicitante=?, prazo=?, status=?, responsavel=?, data_conclusao=?, observacao=? WHERE id=?"""), 
                          (v['td'].get(), v['as'].get(), v['so'].get(), v['pr'].get(), v['st'].get(), v['re'].get(), dt_c, txt_obs.get("1.0", tk.END).strip(), d_id))
            else:
                c.execute(replace_placeholders("""INSERT INTO demandas_avulsas (tipo_demanda, assunto, solicitante, data_entrada, prazo, status, responsavel, data_conclusao, observacao) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)"""), 
                          (v['td'].get(), v['as'].get(), v['so'].get(), dt_ent, v['pr'].get(), v['st'].get(), v['re'].get(), dt_c, txt_obs.get("1.0", tk.END).strip()))
            conn.commit(); conn.close(); m.destroy(); self.load_avulsas(); messagebox.showinfo("Sucesso", "Salvo!")
            
        tk.Button(m, text="Salvar Informações", command=salvar, bg="#4CAF50", fg="white", font=("Helvetica", 10, "bold"), relief=tk.FLAT, pady=10).pack(fill=tk.X, padx=20, pady=10)
