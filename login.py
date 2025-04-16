# login.py
import tkinter as tk
from tkinter import messagebox
import keyring
from typing import Optional, Tuple


class LoginManager:
    """Gerenciador de login seguro com armazenamento no keyring"""

    def __init__(self, app_name: str = "selenium_app"):
        self.app_name = app_name
        self.janela = None
        self.entry_login = None
        self.entry_senha = None

    def _setup_ui(self):
        """Configura a interface gráfica"""
        self.janela = tk.Tk()
        self.janela.title("Autenticação")
        self.janela.resizable(False, False)

        # Centraliza a janela na tela do computador
        self.janela.eval('tk::PlaceWindow . center')

        # Centraliza o frame dentro da janela
        frame = tk.Frame(self.janela, padx=20, pady=20)
        frame.place(relx=0.5, rely=0.5, anchor=tk.CENTER)
        # Frame principal
        frame.pack()

        # Widgets
        tk.Label(frame, text="Usuário:").grid(
            row=0, column=0, sticky="w", pady=5)
        self.entry_login = tk.Entry(frame, width=25)
        self.entry_login.grid(row=0, column=1, pady=5)

        tk.Label(frame, text="Senha:").grid(
            row=1, column=0, sticky="w", pady=5)
        self.entry_senha = tk.Entry(frame, width=25, show="•")
        self.entry_senha.grid(row=1, column=1, pady=5)

        btn_entrar = tk.Button(frame, text="Entrar", command=self._fazer_login)
        btn_entrar.grid(row=2, columnspan=2, pady=10)

        # Configurações
        self.entry_login.focus()
        self.janela.bind('<Return>', lambda e: self._fazer_login())

    def _fazer_login(self):
        """Processa o login e armazena as credenciais"""
        login = self.entry_login.get()
        senha = self.entry_senha.get()

        if not login or not senha:
            messagebox.showerror("Erro", "Preencha todos os campos!")
            return

        keyring.set_password(self.app_name, login, senha)
        self.janela.quit()

    def get_credentials(self) -> Optional[Tuple[str, str]]:
        """
        Exibe a janela de login e retorna as credenciais
        Retorna: (login, senha) ou None se cancelado
        """
        self._setup_ui()
        self.janela.mainloop()

        login = self.entry_login.get()
        senha = keyring.get_password(self.app_name, login) if login else None

        self.janela.destroy()

        return (login, senha) if (login and senha) else None

    @staticmethod
    def clear_credentials(app_name: str, login: str):
        """Remove credenciais armazenadas"""
        try:
            keyring.delete_password(app_name, login)
        except keyring.errors.PasswordDeleteError:
            pass


# Função de conveniência para uso rápido
def obter_credenciais(app_name: str = "selenium_app") -> Optional[Tuple[str, str]]:
    """
    Mostra a janela de login e retorna (usuário, senha)
    Exemplo de uso:
        from login import obter_credenciais
        usuario, senha = obter_credenciais()
    """
    return LoginManager(app_name).get_credentials()
