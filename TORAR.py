import tkinter as tk
import json
import os
import pyttsx3
import shutil
from openpyxl import Workbook
import sqlite3
import webbrowser
# Inicialize o mecanismo de fala uma vez no início do programa
engine = pyttsx3.init()

# Variáveis globais para armazenar os itens do almoxarifado e o nome do usuário logado
almoxarifado = {}
nome_usuario_logado = ""

# Inicialize a conexão com o banco de dados SQLite
conn = sqlite3.connect("usuarios.db")
cursor = conn.cursor()

# Crie a tabela de usuários se ela não existir
cursor.execute('''CREATE TABLE IF NOT EXISTS usuarios (
                  usuario TEXT PRIMARY KEY,
                  senha TEXT)''')
conn.commit()

# Variável global para armazenar a mensagem de erro
erro_label = None
def spk(a):
    engine.say(a)
    engine.runAndWait()


def fechar_programa():
    spk("encerrando programa")
    root.destroy()


def criar_arquivo_json(nome_arquivo):
    if not os.path.exists(nome_arquivo):
        with open(nome_arquivo, 'w') as file:
            json.dump({}, file)


def carregar_itens_iniciais(lista_itens, json_file):
    with open(json_file, 'r') as file:
        almoxarifado = json.load(file)

    for item, info in almoxarifado.items():
        nome = info["nome"]
        quantidade = info["quantidade"]
        lista_itens.insert(tk.END, f"{nome}: {quantidade}")


def adicionar_item(item_entry, quantidade_entry, lista_itens, json_file):
    nome_item = item_entry.get()
    quantidade = quantidade_entry.get()

    if nome_item and quantidade:
        with open(json_file, 'r') as file:
            almoxarifado = json.load(file)

        almoxarifado[nome_item] = {
            "nome": nome_item,
            "quantidade": int(quantidade)
        }

        with open(json_file, 'w') as file:
            json.dump(almoxarifado, file)

        item_entry.delete(0, tk.END)
        quantidade_entry.delete(0, tk.END)

        atualizar_lista(lista_itens, json_file)


def editar_item(item_entry, quantidade_entry, lista_itens, json_file):
    selecionado = lista_itens.curselection()
    if selecionado:
        indice = selecionado[0]
        novo_nome_item = item_entry.get()
        nova_quantidade = quantidade_entry.get()

        if novo_nome_item and nova_quantidade:
            with open(json_file, 'r') as file:
                almoxarifado = json.load(file)

            item_selecionado = lista_itens.get(indice)
            nome_item_selecionado = item_selecionado.split(":")[0].strip()

            almoxarifado[nome_item_selecionado]["quantidade"] = int(nova_quantidade)

            with open(json_file, 'w') as file:
                json.dump(almoxarifado, file)

            item_entry.delete(0, tk.END)
            quantidade_entry.delete(0, tk.END)

            atualizar_lista(lista_itens, json_file)


def remover_item(lista_itens, json_file):
    selecionado = lista_itens.curselection()
    if selecionado:
        indice = selecionado[0]

        with open(json_file, 'r') as file:
            almoxarifado = json.load(file)

        item_selecionado = lista_itens.get(indice)
        nome_item_selecionado = item_selecionado.split(":")[0].strip()

        del almoxarifado[nome_item_selecionado]

        with open(json_file, 'w') as file:
            json.dump(almoxarifado, file)

        atualizar_lista(lista_itens, json_file)


def atualizar_lista(lista_itens, json_file):
    lista_itens.delete(0, tk.END)
    carregar_itens_iniciais(lista_itens, json_file)


def salvar_em_excel(json_file, excel_file):
    # Carregue os dados do arquivo JSON
    with open(json_file, 'r') as file:
        almoxarifado = json.load(file)

    wb = Workbook()
    ws = wb.active

    # Escreva os cabeçalhos das colunas
    ws.append(["Item", "Quantidade"])

    # Preencha as células com os dados do almoxarifado
    for item, info in almoxarifado.items():
        nome = info["nome"]
        quantidade = info["quantidade"]
        ws.append([nome, quantidade])

    # Salve a planilha com um nome exclusivo (por exemplo, com o nome do usuário)
    wb.save(excel_file)

    spk(f"Os itens foram salvos na planilha {excel_file}")


def fazer_backup():
    # Defina o nome do arquivo de backup com base na data e hora atual
    backup_filename = f"usuarios_backup_{nome_usuario_logado}.db"

    # Copie o arquivo do banco de dados original para o arquivo de backup
    shutil.copy("usuarios.db", backup_filename)

    spk(f"Backup do banco de dados de usuários criado como {backup_filename}")


def verificar_credenciais(usuario_entry, senha_entry):
    global nome_usuario_logado, erro_label

    usuario = usuario_entry.get()
    senha = senha_entry.get()

    cursor.execute("SELECT senha FROM usuarios WHERE usuario=?", (usuario,))
    resultado = cursor.fetchone()

    if resultado is not None and resultado[0] == senha:
        login_window.destroy()
        nome_usuario_logado = usuario
        json_file = f"almoxarifado_{nome_usuario_logado}.json"
        excel_file = f"almoxarifado_{nome_usuario_logado}.xlsx"
        PROGRAMA(json_file, excel_file)
    else:
        erro_label.config(text="Credenciais inválidas. Tente novamente.")


def criar_conta(usuario, senha):
    if usuario and senha:
        cursor.execute("SELECT usuario FROM usuarios WHERE usuario=?", (usuario,))
        resultado = cursor.fetchone()

        if resultado is None:
            cursor.execute("INSERT INTO usuarios VALUES (?, ?)", (usuario, senha))
            conn.commit()
            cadastro_window.destroy()
            login()
        else:
            erro_cadastro_label.config(text="Usuário já existe. Escolha outro nome de usuário.")
    else:
        erro_cadastro_label.config(text="Preencha todos os campos.")


def login():
    global login_window, erro_label
    login_window = tk.Tk()
    login_window.title("Login")

    largura_janela = 400
    altura_janela = 300
    login_window.geometry(f"{largura_janela}x{altura_janela}")

    usuario_label = tk.Label(login_window, text="Usuário:")
    usuario_label.pack()

    usuario_entry = tk.Entry(login_window)
    usuario_entry.pack()

    senha_label = tk.Label(login_window, text="Senha:")
    senha_label.pack()

    senha_entry = tk.Entry(login_window, show="*")
    senha_entry.pack()

    login_button = tk.Button(login_window, text="Login", command=lambda: verificar_credenciais(usuario_entry, senha_entry))
    login_button.pack()
    criar_conta_button = tk.Button(login_window, text="Criar Conta", command=checa_admin)
    criar_conta_button.pack()
    erro_label = tk.Label(login_window, text="", fg="red")
    erro_label.pack()

    login_window.mainloop()


def criar_conta_window():
    global cadastro_window, erro_cadastro_label
    login_window.withdraw()
    cadastro_window = tk.Toplevel(login_window)
    cadastro_window.title("Cadastro de Usuário")

    largura_janela = 400
    altura_janela = 200
    cadastro_window.geometry(f"{largura_janela}x{altura_janela}")

    novo_usuario_label = tk.Label(cadastro_window, text="Novo Usuário:")
    novo_usuario_label.pack()

    novo_usuario_entry = tk.Entry(cadastro_window)
    novo_usuario_entry.pack()

    nova_senha_label = tk.Label(cadastro_window, text="Nova Senha:")
    nova_senha_label.pack()

    nova_senha_entry = tk.Entry(cadastro_window, show="*")
    nova_senha_entry.pack()

    criar_button = tk.Button(cadastro_window, text="Criar Conta", command=lambda: criar_conta(novo_usuario_entry.get(), nova_senha_entry.get()))
    criar_button.pack()

    erro_cadastro_label = tk.Label(cadastro_window, text="", fg="red")
    erro_cadastro_label.pack()

    cadastro_window.mainloop()


def checa_admin():
    def fazer_login():
        usuario = usuario_entry.get()
        senha = senha_entry.get()

        if usuario == "ADMIN" and senha == "@375p2tRm#@":
            resultado_label.config(text="Login bem-sucedido!")
            criar_conta_window()
        else:
            resultado_label.config(text="Login falhou. Verifique o nome de usuário e senha.")

    # Criar uma janela
    janela = tk.Tk()
    janela.title("Login")

    # Criar rótulos e campos de entrada de texto
    usuario_label = tk.Label(janela, text="Nome de Usuário:")
    usuario_label.pack()
    usuario_entry = tk.Entry(janela)
    usuario_entry.pack()

    senha_label = tk.Label(janela, text="Senha:")
    senha_label.pack()
    senha_entry = tk.Entry(janela, show="*")
    senha_entry.pack()

    # Botão de login
    login_button = tk.Button(janela, text="Login", command=fazer_login)
    login_button.pack()

    # Rótulo para exibir o resultado do login
    resultado_label = tk.Label(janela, text="")
    resultado_label.pack()
    janela.protocol("WM_DELETE_WINDOW")
    # Iniciar a interface gráfica
    janela.mainloop()


def PROGRAMA(json_file, excel_file):
    global root
    root = tk.Tk()
    root.title(f"TORAR (desenvolvido pelo soldado Jhonn 409-2023) - Usuário: {nome_usuario_logado}")
    root.configure(bg="OliveDrab")

    usuario_label = tk.Label(root, text=f"Usuário: {nome_usuario_logado}", font=("Arial", 14))
    usuario_label.grid(row=0, column=0, columnspan=2, pady=(20, 0), sticky="w")

    desenvolvido_label = tk.Label(root, text="Desenvolvido pelo Soldado Jhonn 409-2023")
    desenvolvido_label.grid(row=1, column=0, columnspan=2, pady=(0, 10), sticky="w")

    item_label = tk.Label(root, text="Nome do Item:")
    item_label.grid(row=2, column=0, sticky="w", padx=10)

    item_entry = tk.Entry(root)
    item_entry.grid(row=2, column=1, padx=10)

    quantidade_label = tk.Label(root, text="Quantidade:")
    quantidade_label.grid(row=3, column=0, sticky="w", padx=10)

    quantidade_entry = tk.Entry(root)
    quantidade_entry.grid(row=3, column=1, padx=10)

    adicionar_button = tk.Button(root, text="Adicionar Item", command=lambda: adicionar_item(item_entry, quantidade_entry, lista_itens, json_file))
    adicionar_button.grid(row=4, column=0, columnspan=2, pady=10)

    editar_button = tk.Button(root, text="Editar Item", command=lambda: editar_item(item_entry, quantidade_entry, lista_itens, json_file))
    editar_button.grid(row=5, column=0, columnspan=2, pady=5)

    lista_itens = tk.Listbox(root)
    lista_itens.grid(row=6, column=0, columnspan=2, padx=10, pady=10)

    criar_arquivo_json(json_file)
    carregar_itens_iniciais(lista_itens, json_file)

    remover_button = tk.Button(root, text="Remover Item", command=lambda: remover_item(lista_itens, json_file))
    remover_button.grid(row=7, column=0, columnspan=2, pady=5)

    salvar_button = tk.Button(root, text="Salvar em Excel", command=lambda: salvar_em_excel(json_file, excel_file))
    salvar_button.grid(row=8, column=0, columnspan=2, pady=10)

    fazer_backup_button = tk.Button(root, text="Fazer Backup", command=fazer_backup)
    fazer_backup_button.grid(row=9, column=0, columnspan=2, pady=10)

    root.protocol("WM_DELETE_WINDOW", fechar_programa)

    root.mainloop()


if __name__ == "__main__":
    spk("Iniciando TORAR")

    login()
