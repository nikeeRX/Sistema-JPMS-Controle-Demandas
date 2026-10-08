import os
import platform
import subprocess
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import sqlite3
import psycopg2
import pandas as pd
from datetime import datetime
from config import conectar_db, replace_placeholders

class AdminFrame(tk.Frame):
    def __init__(self, parent, controller):
        super().__init__(parent, bg="#f4f4f9")
        self.controller = controller
        
        self.top_bar = tk.Frame(self, bg="#333", height=50)
        self.top_bar.pack(fill=tk.X)
        
        tk.Button(self.top_bar, text="⬅ Voltar ao Portal Principal", command=lambda: self.controller.show_frame("Portal"), bg="#777", fg="white", font=("Helvetica", 10, "bold"), relief=tk.FLAT).pack(side=tk.LEFT, padx=20, pady=10)
        
        self.lbl_user_info = tk.Label(self.top_bar, text="", bg="#333", fg="white", font=("Helvetica", 10, "bold"))
        self.lbl_user_info.pack(side=tk.RIGHT, padx=20, pady=10)
        
        self.setup_admin()

    def refresh_all(self):
        self.lbl_user_info.config(text=f"Usuário: {self.controller.current_user} | Nível: {self.controller.current_role.upper()}")
        self.load_users()

    def setup_admin(self):
        # FRAME DE CADASTRO DE USUÁRIOS
        frame_form = tk.Frame(self, bg="#f4f4f9", pady=10)
        frame_form.pack(fill=tk.X, padx=20)
        
        tk.Label(frame_form, text="Gestão Global de Colaboradores", font=("Helvetica", 18, "bold"), bg="#f4f4f9", fg="#333").grid(row=0, column=0, columnspan=4, pady=(20,25), sticky=tk.W)
        self.var_user_id = tk.StringVar(value="")
        
        tk.Label(frame_form, text="Usuário:", bg="#f4f4f9").grid(row=1, column=0, sticky=tk.W, pady=5)
        self.entry_new_user = ttk.Entry(frame_form, width=15)
        self.entry_new_user.grid(row=1, column=1, padx=5, pady=5)
        
        tk.Label(frame_form, text="Nova Senha:", bg="#f4f4f9").grid(row=1, column=2, sticky=tk.W, pady=5)
        self.entry_new_pass = ttk.Entry(frame_form, width=15)
        self.entry_new_pass.grid(row=1, column=3, padx=5, pady=5)
        
        tk.Label(frame_form, text="Nível:", bg="#f4f4f9").grid(row=1, column=4, sticky=tk.W, pady=5, padx=(10,0))
        self.combo_role = ttk.Combobox(frame_form, values=["user", "admin"], state="readonly", width=10)
        self.combo_role.current(0)
        self.combo_role.grid(row=1, column=5, padx=5, pady=5)
        
        tk.Button(frame_form, text="Criar Novo", command=self.add_user, bg="#008CBA", fg="white", relief=tk.FLAT).grid(row=1, column=6, padx=10)
        tk.Button(frame_form, text="Atualizar Selecionado", command=self.update_user, bg="#f0ad4e", fg="white", relief=tk.FLAT).grid(row=1, column=7, padx=5)
        tk.Button(frame_form, text="Excluir", command=self.delete_user, bg="#d9534f", fg="white", relief=tk.FLAT).grid(row=1, column=8, padx=5)
        
        # BOTÕES DE SISTEMA: BACKUP E RESTAURAÇÃO
        f_bkp = tk.Frame(self, bg="#f4f4f9")
        f_bkp.pack(fill=tk.X, padx=20, pady=(15,0))
        
        btn_backup = tk.Button(f_bkp, text="💾 EXPORTAR BANCO (Backup de Segurança)", command=self.gerar_backup, bg="#16a085", fg="white", font=("Segoe UI", 11, "bold"), relief=tk.FLAT, width=40)
        btn_backup.pack(side=tk.LEFT, padx=5)

        btn_restore = tk.Button(f_bkp, text="📥 IMPORTAR BANCO (Restaurar Backup)", command=self.restaurar_backup, bg="#8e44ad", fg="white", font=("Segoe UI", 11, "bold"), relief=tk.FLAT, width=40)
        btn_restore.pack(side=tk.LEFT, padx=5)

        # TABELA DE USUÁRIOS
        columns = ("ID", "Usuário", "Nível", "Status 1º Acesso")
        self.tree_users = ttk.Treeview(self, columns=columns, show="headings", height=20)
        for col in columns: 
            self.tree_users.heading(col, text=col)
            
        self.tree_users.column("ID", width=50, anchor=tk.CENTER)
        self.tree_users.column("Status 1º Acesso", width=120, anchor=tk.CENTER)
        self.tree_users.pack(fill=tk.BOTH, expand=True, padx=20, pady=(15,10))
        
        self.tree_users.bind("<Double-1>", lambda event: self.load_user_for_edit())

    def load_users(self):
        for item in self.tree_users.get_children(): 
            self.tree_users.delete(item)
            
        conn = conectar_db()
        c = conn.cursor()
        c.execute("SELECT id, username, role, first_login FROM users ORDER BY id ASC")
        
        for row in c.fetchall(): 
            self.tree_users.insert("", tk.END, values=(row[0], row[1], row[2], "Pendente" if row[3] == 1 else "OK"))
            
        conn.close()

    def load_user_for_edit(self):
        sel = self.tree_users.selection()
        if not sel: return
        i = self.tree_users.item(sel[0])
        self.var_user_id.set(i['values'][0])
        
        self.entry_new_user.delete(0, tk.END)
        self.entry_new_user.insert(0, i['values'][1])
        self.combo_role.set(i['values'][2])

    def add_user(self):
        u = self.entry_new_user.get()
        p = self.entry_new_pass.get()
        r = self.combo_role.get()
        
        if not u or not p: 
            return messagebox.showwarning("Aviso", "Preencha usuário e senha.")
            
        try:
            conn = conectar_db()
            c = conn.cursor()
            q = replace_placeholders("INSERT INTO users (username, password, role, first_login) VALUES (?, ?, ?, 1)")
            c.execute(q, (u, p, r))
            conn.commit()
            conn.close()
            
            self.entry_new_user.delete(0, tk.END)
            self.entry_new_pass.delete(0, tk.END)
            self.load_users()
            messagebox.showinfo("Sucesso", "Colaborador cadastrado!")
            
        except (sqlite3.IntegrityError, psycopg2.errors.UniqueViolation): 
            messagebox.showerror("Erro", "Nome já existe!")

    def update_user(self):
        uid = self.var_user_id.get()
        if not uid: 
            return messagebox.showwarning("Aviso", "Selecione um usuário na lista.")
            
        u = self.entry_new_user.get()
        p = self.entry_new_pass.get()
        r = self.combo_role.get()
        
        conn = conectar_db()
        c = conn.cursor()
        
        if p: 
            q = replace_placeholders("UPDATE users SET username=?, password=?, role=?, first_login=1 WHERE id=?")
            c.execute(q, (u, p, r, uid))
        else: 
            q = replace_placeholders("UPDATE users SET username=?, role=? WHERE id=?")
            c.execute(q, (u, r, uid))
            
        conn.commit()
        conn.close()
        
        self.load_users()
        self.entry_new_user.delete(0, tk.END)
        self.entry_new_pass.delete(0, tk.END)
        self.var_user_id.set("")
        messagebox.showinfo("Sucesso", "Atualizado!")

    def delete_user(self):
        sel = self.tree_users.selection()
        if not sel: 
            return messagebox.showwarning("Aviso", "Selecione um usuário para excluir.")
            
        item = self.tree_users.item(sel[0])
        
        if item['values'][1] in ('admin', self.controller.current_user): 
            return messagebox.showerror("Erro", "Ação negada.")
            
        conn = conectar_db()
        c = conn.cursor()
        q = replace_placeholders("DELETE FROM users WHERE id=?")
        c.execute(q, (item['values'][0],))
        conn.commit()
        conn.close()
        
        self.load_users()
        messagebox.showinfo("Sucesso", "Colaborador excluído!")

    # ==========================================
    # EXPORTAÇÃO (BACKUP) E IMPORTAÇÃO (RESTAURE)
    # ==========================================
    def gerar_backup(self):
        try:
            conn = conectar_db()
            # Puxa tudo. Até as senhas são exportadas agora para garantir a restauração completa.
            df_demands = pd.read_sql_query("SELECT * FROM demands", conn)
            df_avulsas = pd.read_sql_query("SELECT * FROM demandas_avulsas", conn)
            df_users = pd.read_sql_query("SELECT * FROM users", conn)
            conn.close()
            
            pasta_backup = r'S:\DIOPE\GENEG\COCAP\1. DEMANDAS GERENCIAIS\Sistema de Controle de Demandas\BACKUP_BANCO'
            if not os.path.exists(pasta_backup):
                try: os.makedirs(pasta_backup, exist_ok=True)
                except Exception:
                    pasta_backup = os.path.join(os.path.expanduser('~'), 'Desktop', 'BACKUP_COCAP')
                    os.makedirs(pasta_backup, exist_ok=True)
            
            data_hora = datetime.now().strftime("%Y_%m_%d_%Hh%Mm%Ss")
            nome_arquivo = f"Backup_Geral_{data_hora}.xlsx"
            caminho_completo = os.path.join(pasta_backup, nome_arquivo)
            
            with pd.ExcelWriter(caminho_completo, engine='openpyxl') as writer:
                df_demands.to_excel(writer, sheet_name='Rotina_PDFs', index=False)
                df_avulsas.to_excel(writer, sheet_name='Demandas_Avulsas', index=False)
                df_users.to_excel(writer, sheet_name='Usuarios', index=False)
            
            messagebox.showinfo("Backup Concluído", f"Base de Dados extraída com sucesso!\n\nSalvo em:\n{caminho_completo}")
            
            if platform.system() == 'Windows': os.startfile(pasta_backup)
            elif platform.system() == 'Darwin': subprocess.Popen(['open', pasta_backup])
                
        except Exception as e:
            messagebox.showerror("Erro de Backup", f"Ocorreu um erro ao extrair o banco:\n{str(e)}")

    def restaurar_backup(self):
        msg = "⚠️ ATENÇÃO: PERIGO ⚠️\n\nEssa ação vai APAGAR todas as demandas, avulsas e usuários atuais do banco de dados e substituir pelas informações do arquivo Excel.\n\nTem certeza absoluta que deseja continuar?"
        if not messagebox.askyesno("Restaurar Banco de Dados", msg, icon='warning'): 
            return
            
        filepath = filedialog.askopenfilename(filetypes=[("Arquivos Excel", "*.xlsx")], title="Selecione o arquivo de Backup")
        if not filepath: return
        
        try:
            xls = pd.ExcelFile(filepath)
            sheets = xls.sheet_names
            
            conn = conectar_db()
            c = conn.cursor()
            is_postgres = "psycopg2" in str(type(conn))
            
            # Mapeamento do Excel para o Banco de Dados
            mapeamento = {
                'Rotina_PDFs': 'demands',
                'Demandas_Avulsas': 'demandas_avulsas',
                'Usuarios': 'users'
            }
            
            for aba, tabela in mapeamento.items():
                if aba in sheets:
                    df = pd.read_excel(xls, sheet_name=aba)
                    
                    # Tática de Guerra: Limpa a tabela inteira!
                    c.execute(f"DELETE FROM {tabela}")
                    
                    if not df.empty:
                        # Converte os "NaN" vazios do Excel para "None" para o Banco de Dados não quebrar
                        df = df.where(pd.notnull(df), None)
                        
                        cols = ", ".join(df.columns)
                        placeholders = ", ".join(["?"] * len(df.columns))
                        q = replace_placeholders(f"INSERT INTO {tabela} ({cols}) VALUES ({placeholders})")
                        
                        # Injeta a planilha inteira de uma vez no banco
                        c.executemany(q, df.values.tolist())
                        
                        # PostgreSQL exige que as sequências (IDs automáticos) sejam arrumadas depois de injetar IDs antigos
                        if is_postgres:
                            try:
                                c.execute(f"SELECT setval('{tabela}_id_seq', COALESCE((SELECT MAX(id)+1 FROM {tabela}), 1), false)")
                            except Exception as e_seq:
                                print(f"Aviso de sequência ignorado para a tabela {tabela}: {e_seq}")
            
            conn.commit()
            conn.close()
            
            messagebox.showinfo("Restauração Concluída", "O Banco de Dados foi restaurado com sucesso!\n\nPor segurança, o sistema fará o logoff agora. Entre novamente com os usuários recuperados.")
            
            # Força o logoff para limpar a memória do usuário
            self.controller.current_user = None
            self.controller.current_role = None
            self.controller.show_frame("Login")
            
        except Exception as e:
            messagebox.showerror("Erro de Restauração", f"O arquivo selecionado pode ser inválido ou estar corrompido.\n\nDetalhes do Erro:\n{str(e)}")
