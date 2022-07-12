import pandas as pd
from pandastable import Table
from datetime import date
from tkinter import *
from tkinter import messagebox
from tkinter.ttk import Combobox
from tkcalendar import DateEntry
import sqlite3


def containsNumber(value):
    return any([char.isdigit() for char in value])


def containsSpecialChar(value):
    special_characters = "''!@#$%^&*()-+?_=,<>/"
    return any(c in special_characters for c in value)


def containsLetters(value):
    return any(c.isalpha() for c in value)


def verificacao_cadastro():
    if containsNumber(nome_e.get()) or containsSpecialChar(nome_e.get()):
        messagebox.showerror(
            title="Error", message="Nome só pode conter letras.")
        return False
    elif containsLetters(cpf_e.get()) or containsSpecialChar(cpf_e.get()):
        messagebox.showerror(
            title="Error", message="CPF só pode conter números.")
        return False
    elif containsLetters(matricula_e.get()) or containsSpecialChar(matricula_e.get()):
        messagebox.showerror(
            title="Error", message="Matrícula só pode conter números.")
        return False
    else:
        return True


def verificacao_login():
    if containsLetters(matricula_e.get()) or containsSpecialChar(matricula_e.get()):
        messagebox.showerror(
            title="Error", message="Matrícula só pode conter números.")
        return False
    else:
        return True


def verificacao_registro():
    if containsNumber(nome_r.get()) or containsSpecialChar(nome_r.get()):
        messagebox.showerror(
            title="Error", message="Nome só pode conter letras.")
        return False
    elif containsLetters(cpf_r.get()) or containsSpecialChar(cpf_r.get()):
        messagebox.showerror(
            title="Error", message="CPF só pode conter números.")
        return False
    elif containsLetters(rg_r.get()) or containsSpecialChar(rg_r.get()):
        messagebox.showerror(
            title="Error", message="RG só pode conter números.")
        return False
    elif len(desc_bem.get("1.0", 'end-1c')) == 0 or len(desc_serv.get("1.0", 'end-1c')) == 0 or len(str(cal.get_date())) == 0:
        messagebox.showerror(
            title="Error", message="Existem campos vazios.")
        return False
    elif containsLetters(tomb.get()) or containsSpecialChar(tomb.get()):
        messagebox.showerror(
            title="Error", message="Tombamento só pode conter números.")
        return False
    else:
        return True


def login():
    global matricula_entregador, nome_entregador

    if not verificacao_login():
        return ...

    try:
        senha_entregador = str(senha_e.get())
        matricula_entregador = int(matricula_e.get())
        cursor = conn.cursor()
        registro = cursor.execute(
            f"SELECT Senha FROM Entregadores WHERE Matrícula = {matricula_entregador}").fetchone()

        if registro != None and registro[0] == senha_entregador:
            cursor.close()
            return menu()

        return messagebox.showerror(
            title="Error", message="Usuário não cadastrado.")
    except ValueError:
        return messagebox.showerror(
            title="Error", message="Usuário não cadastrado.")


def cadastrar():
    global nome_entregador, cpf_entregador, matricula_entregador, senha_entregador

    if not verificacao_cadastro():
        return cadastro_interface()

    try:
        nome_entregador = nome_e.get()
        cpf_entregador = int(cpf_e.get())
        matricula_entregador = int(matricula_e.get())
        senha_entregador = str(senha_e.get())

        if len(senha_entregador) == 0:
            raise ValueError
        elif cpf_entregador == matricula_entregador:
            messagebox.showerror(
                title="Error", message="CPF e Matrícula não podem ser iguais!")
            return cadastro_interface()

        conn.execute(
            f"INSERT INTO Entregadores (Nome, CPF, Matrícula, Senha) VALUES ('{nome_entregador}', {cpf_entregador}, {matricula_entregador}, '{senha_entregador}')")
        conn.commit()
        messagebox.showinfo(
            title="Sucesso", message="Usuário Cadastrado com Sucesso.")
        window.destroy()
        return tela_inicial()
    except sqlite3.IntegrityError:
        messagebox.showerror(
            title="Error", message="CPF ou Matrícula já cadastrado!")
        window.destroy()
        return tela_inicial()
    except ValueError:
        messagebox.showerror(
            title="Error", message="CPF, Matrícula ou Senha estão vazios!")
        return cadastro_interface()


def logout():
    window_menu.destroy()
    return tela_inicial()


def cadastro_interface():
    global cpf_e, nome_e

    Label(
        frame,
        text='Nome:',
        font=("Times", "14")
    ).grid(row=3, column=0, pady=5)

    nome_e = Entry(frame, width=30)
    nome_e.grid(row=3, column=1)

    Label(
        frame,
        text='CPF:',
        font=("Times", "14")
    ).grid(row=4, column=0, pady=5)

    cpf_e = Entry(frame, width=30)
    cpf_e.grid(row=4, column=1)

    button_login["state"] = DISABLED
    button_cadastro["state"] = DISABLED

    button_cadastrar = Button(frame, text="Cadastrar", padx=20, pady=10,
                              relief=SOLID, command=cadastrar, font=("Times", "14", "bold"))
    button_cadastrar.grid(row=6, column=2, pady=20)


def exportar_excel():
    today = date.today()

    if data_de.get_date() != today:
        df = pd.read_sql(
            f"SELECT * FROM Recebedores WHERE Matricula_Entregador = {matricula_entregador} AND Data_Entrega >= '{str(data_de.get_date())}' AND Data_Entrega <= '{str(data_ate.get_date())}'", conn)
        df.to_excel(
            f"{str(data_de.get_date())}_to_{data_ate.get_date()} - {nome_entregador}.xlsx", index=False)
    else:
        df = pd.read_sql(
            f"SELECT * FROM Recebedores WHERE Matricula_Entregador = {matricula_entregador}", conn)
        df.to_excel(f"Registros - {nome_entregador}.xlsx", index=False)

    messagebox.showinfo(
        title="Sucesso", message="Exportado com Sucesso.")
    window_excel.destroy()


def exportar_excel_interface():
    global window_excel, data_de, data_ate
    window_excel = Tk()
    window_excel.title("Manutenção STI")
    window_excel.geometry("700x600")
    window_excel.resizable(width=False, height=False)

    frame_excel = Frame(window_excel, padx=20, pady=20)
    frame_excel.pack(expand=True)

    Label(
        frame_excel,
        text="Exportar para Excel",
        font=("Times", "24", "bold")
    ).grid(row=0, columnspan=3, pady=15)

    Label(
        frame_excel,
        text='De:',
        font=("Times", "14")
    ).grid(row=1, column=0, pady=5)

    Label(
        frame_excel,
        text='Até:',
        font=("Times", "14")
    ).grid(row=2, column=0, pady=5)

    data_de = DateEntry(frame_excel, selectmode="day")
    data_de.grid(row=1, column=1)

    data_ate = DateEntry(frame_excel, selectmode="day")
    data_ate.grid(row=2, column=1)

    button_login = Button(frame_excel, text="Exportar", padx=20, pady=10,
                          relief=SOLID, command=exportar_excel, font=("Times", "14", "bold"))
    button_login.grid(row=3, column=1, pady=20)


def add_registro():
    global window_patrimonio, desc_bem, desc_serv, cal, rep, tomb, nome_r, cpf_r, rg_r, matricula_entregador

    if not verificacao_registro():
        window_add.destroy()
        return registro_interface()

    nome_r = nome_r.get()
    cpf_r = int(cpf_r.get())
    rg_r = int(rg_r.get())
    desc_bem = desc_bem.get("1.0", 'end-1c')
    desc_serv = desc_serv.get("1.0", 'end-1c')
    cal = str(cal.get_date())
    rep = "Sim" if rep.get() == "Sim" else "Não"
    tomb = int(tomb.get())

    conn.execute(
        f"INSERT INTO Recebedores (Nome, CPF, RG, Matricula_Entregador, Desc_Bem, Tombamento, Data_Entrega, Reparado, Desc_Serviço) values ('{nome_r}', {cpf_r}, {rg_r}, {matricula_entregador}, '{desc_bem}', {tomb}, '{cal}', '{rep}', '{desc_serv}')")
    conn.commit()
    messagebox.showinfo(
        title="Sucesso", message="Registro Cadastrado.")
    window_add.destroy()


def registro_interface():
    global nome_r, cpf_r, rg_r, desc_bem, desc_serv, tomb, cal, rep, window_add
    window_add = Tk()
    window_add.title("Manutenção STI")
    window_add.geometry("1280x720")
    window_add.state('zoomed')
    window_add.resizable(width=False, height=False)

    frame_add = Frame(window_add, padx=20, pady=20)
    frame_add.pack(expand=True)

    Label(
        frame_add,
        text="Novo Registro",
        font=("Times", "24", "bold")
    ).grid(row=0, columnspan=3, pady=10)

    Label(
        frame_add,
        text='Nome',
        font=("Times", "14")
    ).grid(row=1, column=0, pady=5)

    Label(
        frame_add,
        text='CPF',
        font=("Times", "14")
    ).grid(row=2, column=0, pady=5)

    Label(
        frame_add,
        text='RG',
        font=("Times", "14")
    ).grid(row=3, column=0, pady=5)

    Label(
        frame_add,
        text='Descrição do Bem',
        font=("Times", "14")
    ).grid(row=4, column=0, pady=5)

    Label(
        frame_add,
        text='Tombamento',
        font=("Times", "14")
    ).grid(row=5, column=0, pady=5)

    Label(
        frame_add,
        text='Data de Entrega',
        font=("Times", "14")
    ).grid(row=6, column=0, pady=5)

    Label(
        frame_add,
        text='Reparado?',
        font=("Times", "14")
    ).grid(row=7, column=0, pady=5)

    Label(
        frame_add,
        text='Descrição do Serviço',
        font=("Times", "14")
    ).grid(row=8, column=0, pady=5)

    nome_r = Entry(frame_add, width=30)
    cpf_r = Entry(frame_add, width=30)
    rg_r = Entry(frame_add, width=30)
    desc_bem = Text(frame_add, width=60, height=2, font=("Times", "10"))
    tomb = Entry(frame_add, width=30)
    cal = DateEntry(frame_add, selectmode='day')
    rep = Combobox(frame_add, width=30, height=2, values=[
                   "Sim", "Não"], state="readonly")
    desc_serv = Text(frame_add, width=60, height=2, font=("Times", "10"))

    rep.current(0)

    nome_r.grid(row=1, column=1)
    cpf_r.grid(row=2, column=1)
    rg_r.grid(row=3, column=1)
    desc_bem.grid(row=4, column=1)
    tomb.grid(row=5, column=1)
    cal.grid(row=6, column=1)
    rep.grid(row=7, column=1)
    desc_serv.grid(row=8, column=1)

    button_concluir = Button(frame_add, text="Concluir", padx=20, pady=10,
                             relief=SOLID, command=add_registro, font=("Times", "14", "bold"))
    button_concluir.grid(row=9, column=1, pady=20)


def menu():
    global window_menu
    window.destroy()
    window_menu = Tk()
    window_menu.title("Manutenção STI")
    window_menu.state('zoomed')
    window_menu.resizable(width=False, height=False)

    frame_menu = Frame(window_menu, padx=20, pady=20)
    frame_menu.pack(expand=True)

    button_tabela = Button(frame_menu, text="Tabela de Recebedores", padx=20, pady=10,
                           relief=SOLID, command=abrir_tabela, font=("Times", "14", "bold"))
    button_tabela.grid(row=2, column=2, pady=20)

    button_excel = Button(frame_menu, text="Exportar para Excel", padx=20, pady=10,
                          relief=SOLID, command=exportar_excel_interface, font=("Times", "14", "bold"))
    button_excel.grid(row=3, column=2, pady=20)

    button_add = Button(frame_menu, text="Adicionar Registro", padx=20, pady=10,
                        relief=SOLID, command=registro_interface, font=("Times", "14", "bold"))
    button_add.grid(row=4, column=2, pady=20)

    button_deslogar = Button(frame_menu, text="Logout", padx=20, pady=10,
                             relief=SOLID, command=logout, font=("Times", "14", "bold"))
    button_deslogar.grid(row=5, column=2, pady=20)


def abrir_tabela():
    global matricula_entregador
    window_tabela = Tk()
    window_tabela.title("Manutenção STI")
    window_tabela.geometry("1280x720")
    window_tabela.state('zoomed')
    window_tabela.resizable(width=False, height=False)

    frame_tabela = Frame(window_tabela, padx=20, pady=20)
    frame_tabela.pack(expand=True)

    df = pd.read_sql(
        f"SELECT * FROM Recebedores WHERE Matricula_Entregador = {matricula_entregador}", conn)
    pt = Table(frame_tabela, dataframe=df, editable=False, width=1280,
               height=720)
    pt.show()


def tela_inicial():
    global window, nome_e, matricula_e, button_cadastro, button_login, frame, nome_entregador, matricula_entregador, cpf_entregador, senha_e, senha_entregador
    window = Tk()
    window.title("Manutenção STI")
    window.state("zoomed")
    window.resizable(width=False, height=False)
    window.iconphoto(False, PhotoImage(file="ufpi.png"))

    frame = Frame(window, padx=20, pady=20)
    frame.pack(expand=True)

    nome_entregador = None
    matricula_entregador = None
    cpf_entregador = None
    senha_entregador = None

    Label(
        frame,
        text="Entregador",
        font=("Times", "24", "bold")
    ).grid(row=0, columnspan=3, pady=10)

    Label(
        frame,
        text='Matrícula:',
        font=("Times", "14")
    ).grid(row=1, column=0, pady=5)

    Label(
        frame,
        text='Senha:',
        font=("Times", "14")
    ).grid(row=2, column=0, pady=5)

    matricula_e = Entry(frame, width=30)
    senha_e = Entry(frame, show="*", width=30)

    matricula_e.grid(row=1, column=1)
    senha_e.grid(row=2, column=1)

    button_login = Button(frame, text="Login", padx=20, pady=10,
                          relief=SOLID, command=login, font=("Times", "14", "bold"))
    button_login.grid(row=6, column=0, pady=20)

    button_cadastro = Button(frame, text="Cadastro", padx=20, pady=10,
                             relief=SOLID, command=cadastro_interface, font=("Times", "14", "bold"))
    button_cadastro.grid(row=6, column=2, pady=20)


if __name__ == "__main__":
    conn = sqlite3.connect("sti.db")
    tela_inicial()
    window.mainloop()
