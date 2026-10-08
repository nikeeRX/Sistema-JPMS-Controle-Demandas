import tkinter as tk
from tkinter import ttk, messagebox
from config import conectar_db, setup_db
from modulo_admin import AdminFrame
from modulo_demandas import ControleDemandasFrame
from modulo_pentefino import PenteFinoFrame
import os

class AppGestaoDemandas:
    def __init__(self, root):
        self.root = root
        self.root.title("Postal Saúde - Central de Operações")
        self.root.geometry("1440x900")
        self.root.configure(bg="#f4f4f9")
        
        # Força Maximização de Tela
        try: self.root.state('zoomed')
        except: self.root.attributes('-zoomed', True)

        setup_db()
        self.current_user = None
        self.current_role = None
        
        self.container = tk.Frame(self.root, bg="#f4f4f9")
        self.container.pack(fill=tk.BOTH, expand=True)
        
        # Mapeamento de todas as telas mestres do sistema (Importadas)
        self.frames = {
            "Login": LoginFrame(self.container, self),
            "Portal": PortalFrame(self.container, self),
            "ControleDemandas": ControleDemandasFrame(self.container, self),
            "PenteFino": PenteFinoFrame(self.container, self),
            "Admin": AdminFrame(self.container, self)
        }
        self.show_frame("Login")

    def show_frame(self, page_name):
        for frame in self.frames.values():
            frame.pack_forget()
        frame = self.frames[page_name]
        frame.pack(fill=tk.BOTH, expand=True)
        
        # Atualiza os dados da tela antes de mostrar
        if hasattr(frame, "refresh_all"): 
            frame.refresh_all()
            
        if page_name == "Login": 
            frame.focus_login()


# ---------------------------------------------------------
# Telas Simples (Login e Portal) embutidas no Main
# ---------------------------------------------------------
class LoginFrame(tk.Frame):
    def __init__(self, parent, controller):
        super().__init__(parent, bg="#f4f4f9")
        self.controller = controller
        
        fl = tk.Frame(self, bg="#ffffff", padx=40, pady=40, relief=tk.RAISED, bd=2)
        fl.place(relx=0.5, rely=0.5, anchor=tk.CENTER)
        
        tk.Label(fl, text="LOGIN DO SISTEMA", font=("Helvetica", 16, "bold"), bg="#ffffff").pack(pady=(0, 20))
        tk.Label(fl, text="Usuário:", font=("Helvetica", 10), bg="#ffffff").pack(anchor=tk.W)
        self.entry_user = ttk.Entry(fl, width=30)
        self.entry_user.pack(pady=(0, 15))
        
        tk.Label(fl, text="Senha:", font=("Helvetica", 10), bg="#ffffff").pack(anchor=tk.W)
        self.entry_pass = ttk.Entry(fl, width=30, show="*")
        self.entry_pass.pack(pady=(0, 20))
        
        tk.Button(fl, text="Entrar", command=self.fazer_login, bg="#4CAF50", fg="white", font=("Helvetica", 12, "bold"), relief=tk.FLAT, pady=5).pack(fill=tk.X)
        
        self.entry_user.bind("<Return>", lambda event: self.entry_pass.focus())
        self.entry_pass.bind("<Return>", lambda event: self.fazer_login())

    def focus_login(self): 
        self.entry_user.focus()

    def fazer_login(self):
        user = self.entry_user.get()
        password = self.entry_pass.get()
        conn = conectar_db()
        c = conn.cursor()
        c.execute("SELECT id, role, first_login FROM users WHERE username=%s AND password=%s" if "psycopg2" in str(type(conn)) else "SELECT id, role, first_login FROM users WHERE username=? AND password=?", (user, password))
        res = c.fetchone()
        conn.close()
        
        if res:
            if res[2] == 1: 
                self.forcar_troca_senha(res[0], user, res[1])
            else: 
                self.concluir_login(user, res[1])
        else: 
            messagebox.showerror("Erro", "Usuário ou senha incorretos!")

    def forcar_troca_senha(self, uid, u, r):
        m = tk.Toplevel(self)
        m.title("Mudar Senha")
        m.geometry("400x300")
        m.configure(bg="#f4f4f9")
        m.transient(self.controller.root)
        m.grab_set()
        
        tk.Label(m, text="Bem-vindo! Este é seu primeiro acesso.\nDefina uma nova senha.", font=("Helvetica", 11, "bold"), bg="#f4f4f9", pady=20).pack()
        tk.Label(m, text="Nova Senha:", bg="#f4f4f9").pack(anchor=tk.W, padx=50)
        en = ttk.Entry(m, show="*", width=30)
        en.pack(pady=(0, 15))
        
        tk.Label(m, text="Confirmar Senha:", bg="#f4f4f9").pack(anchor=tk.W, padx=50)
        ec = ttk.Entry(m, show="*", width=30)
        ec.pack(pady=(0, 20))
        
        def salvar():
            if not en.get() or not ec.get(): 
                return messagebox.showwarning("Aviso", "Preencha tudo.")
            if en.get() != ec.get(): 
                return messagebox.showerror("Erro", "Senhas não coincidem!")
                
            conn = conectar_db()
            c = conn.cursor()
            query = "UPDATE users SET password=%s, first_login=0 WHERE id=%s" if "psycopg2" in str(type(conn)) else "UPDATE users SET password=?, first_login=0 WHERE id=?"
            c.execute(query, (en.get(), uid))
            conn.commit()
            conn.close()
            
            messagebox.showinfo("Sucesso", "Senha atualizada!")
            m.destroy()
            self.concluir_login(u, r)
            
        tk.Button(m, text="Salvar", command=salvar, bg="#4CAF50", fg="white", font=("Helvetica", 10, "bold"), relief=tk.FLAT).pack()

    def concluir_login(self, u, r):
        self.controller.current_user = u
        self.controller.current_role = r
        self.entry_user.delete(0, tk.END)
        self.entry_pass.delete(0, tk.END)
        self.controller.show_frame("Portal")


class PortalFrame(tk.Frame):
    def __init__(self, parent, controller):
        super().__init__(parent, bg="#f4f4f9")
        self.controller = controller
        
        tb = tk.Frame(self, bg="#333", height=50)
        tb.pack(fill=tk.X)
        self.lui = tk.Label(tb, text="", bg="#333", fg="white", font=("Helvetica", 10, "bold"))
        self.lui.pack(side=tk.LEFT, padx=20, pady=10)
        tk.Button(tb, text="Sair do Sistema", command=self.logout, bg="#d9534f", fg="white", relief=tk.FLAT).pack(side=tk.RIGHT, padx=20, pady=10)
        
        tk.Label(self, text="PORTAL CENTRAL DE OPERAÇÕES", font=("Segoe UI", 24, "bold"), bg="#f4f4f9", fg="#004b87").pack(pady=(80, 40))
        
        mc = tk.Frame(self, bg="#f4f4f9")
        mc.pack()
        
        tk.Button(mc, text="📊 SISTEMA DE CONTROLE DE DEMANDAS", command=lambda: self.controller.show_frame("ControleDemandas"), bg="#008CBA", fg="white", font=("Segoe UI", 16, "bold"), relief=tk.FLAT, width=45, pady=20).pack(pady=15)
        tk.Button(mc, text="🏥 SUBSTITUIÇÃO DE PRESTADORES (RN 665)", command=lambda: self.controller.show_frame("PenteFino"), bg="#f0ad4e", fg="white", font=("Segoe UI", 16, "bold"), relief=tk.FLAT, width=45, pady=20).pack(pady=15)
        
        self.ba = tk.Button(mc, text="⚙️ PAINEL ADMINISTRATIVO (Usuários)", command=lambda: self.controller.show_frame("Admin"), bg="#333333", fg="white", font=("Segoe UI", 14, "bold"), relief=tk.FLAT, width=45, pady=15)

    def refresh_all(self):
        self.lui.config(text=f"Usuário Logado: {self.controller.current_user} | Nível: {self.controller.current_role.upper()}")
        if self.controller.current_role == 'admin': 
            self.ba.pack(pady=15)
        else: 
            self.ba.pack_forget()

    def logout(self): 
        self.controller.current_user = None
        self.controller.current_role = None
        self.controller.show_frame("Login")

if __name__ == "__main__":
    root = tk.Tk()
    app = AppGestaoDemandas(root)
    root.mainloop()
