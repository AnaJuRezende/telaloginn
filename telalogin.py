import re
from tkinter import *
from tkinter import Tk
from tkinter import messagebox
from openpyxl import load_workbook
import os
from dotenv import dotenv_values

env_vars = dotenv_values()
LOGIN = env_vars['LOGIN']
SENHA = env_vars['SENHA']
env_vars = dotenv_values('.env')

co0 = "#f0f3f5"  #Preta / black
co1 = "#feffff"  #branca / white 
co2 = "#3fb5a3"  #verde / green
co3 = "#38576b"  #valor / value
co4 = "#403d3d"  #letra / letters

def marcar_cliente():

    for widget in janela.winfo_children():
        widget.destroy()

    janela.title('Marcação Salão Espaço Vip')
    janela.geometry('500x500')
    janela.configure(background=co1)
    janela.resizable(width=FALSE, height=FALSE)

    frame_superior = Frame(janela, width=310, height=50, bg=co1, relief='flat')
    frame_superior.grid(row=0, column=0, pady=1, padx=0, sticky=NSEW)

    frame_inferior = Frame(janela, width=310, height=50, bg=co1, relief='flat')
    frame_inferior.grid(row=1, column=0, pady=1, padx=0, sticky=NSEW)

    l_nome = Label(frame_superior, text='MARCAÇÃO', anchor=NE, font=('Ivy 25'), bg=co1, fg=co4)
    l_nome.place(x=5, y=5)

    l_linha = Label(frame_superior, text='', width=275, anchor=NW, font=('Ivy 1'), bg=co2, fg=co4)
    l_linha.place(x=0, y=45)

    l_nome_cliente = Label(janela, text='Nome ', anchor=NW, font=('Ivy 10'), bg=co1, fg=co4)
    l_nome_cliente.place(x=10, y=60)
    
    e_nome_cliente = Entry(janela, width=20, justify='left', font=("", 15), highlightthickness=1, relief='solid')
    e_nome_cliente.place(x=10, y=85)

    l_horario = Label(janela, text='Horário ', anchor=NW, font=('Ivy 10'), bg=co1, fg=co4)
    l_horario.place(x=10, y=120)

    e_horario = Entry(janela, width=20, justify='left', font=('Arial 10', 15), highlightthickness=1, relief='solid')
    e_horario.place(x=10, y=145)

    l_procedimento = Label(janela, text='Procedimento ', anchor=NW, font=('Ivy 10'), bg=co1, fg=co4)
    l_procedimento.place(x=10, y=185)

    e_procedimento = Entry(janela, width=20, justify='left', font=('Arial 10', 15), highlightthickness=1, relief='solid')
    e_procedimento.place(x=10, y=210)

    l_celular = Label(janela, text = 'Celular ', anchor = NW, font=('Ivy 10'), bg=co1, fg=co4)
    l_celular.place(x=10, y=250)

    e_celular = Entry(janela, width=20, justify='left', font=('Arial 10', 15), highlightthickness=1, relief='solid')
    e_celular.place(x=10, y=275)

    l_data = Label(janela, text = 'Data ', anchor = NW, font=('Ivy 10'), bg=co1, fg=co4)
    l_data.place(x=10, y=310)

    e_data = Entry(janela, width=20, font=('', 15), highlightthickness=1, relief='solid')
    e_data.place(x=10, y= 330)

    l_valor = Label(janela, text='Valor', anchor = NW, font=('Ivy 10'), bg=co1, fg=co4)
    l_valor.place(x=10, y=370)

    e_valor = Entry(janela, width=20, font=('', 15), highlightthickness=1, relief='solid')
    e_valor.place(x=10, y=390)

    b_marcar = Button(janela, text='MARCAR', command=lambda: adicionar_dados_no_excel(e_nome_cliente.get(), e_procedimento.get(), e_celular.get(), e_data.get(), e_valor.get(), e_horario.get()), width=20, height=1, font=('Ivy 8 bold'), bg=co1, fg=co4)
    b_marcar.place(x=50, y=450)

def verif_marcacao(cliente, horario, data, celular, valor):
    if not re.match("^[a-zA-Z ]+$", cliente):
        messagebox.showerror("Erro", "Nome inválido")
        return False
    if not re.match("^\d{9,11}$", celular):
        messagebox.showerror("Erro", "Celular inválido")
        return False
    if not re.match("^\d{1,5}$", valor):
        messagebox.showerror("Erro", "Valor inválido")
        return False
    if not re.match("^\d{2}/\d{2}$", data):
        messagebox.showerror("Erro", "Data inválida")
        return False
    if not re.match("^\d{2}:\d{2}$", horario):
        messagebox.showerror("Erro", "Horario inválido")
        return False
    else:
        return True


def adicionar_dados_no_excel(nomcliente, procedimento, celular, data, valor, horario):
    if verif_marcacao(nomcliente, horario, data, celular, valor): 
        workbook = load_workbook("./ExcelClientes.xlsx")
        sheet = workbook.active
        sheet['A2'] = nomcliente
        sheet['B2'] = procedimento
        sheet['C2'] = celular
        sheet['D2'] = data
        sheet['E2'] = valor
        sheet['F2'] = horario
        workbook.save("./ExcelClientes.xlsx")
        messagebox.showinfo('MARCADO', 'Cliente Marcado!')
 
janela = Tk()
janela.title('Login Salão Espaço Vip')
janela.geometry('450x350')
janela.configure(background=co1)
janela.resizable(width=FALSE, height=FALSE)

credenciais = [LOGIN, SENHA] #mudar depois

l_login = Label(janela, text='LOGIN', anchor=NE, font=('Ivy 15'), bg=co1, fg=co4)
l_login.place(x=175, y=5)

l_linha = Label(janela, text='', width=275, font=('Ivy 1'), bg=co2, fg=co4)
l_linha.place(x=80, y=45)

l_usuario = Label(janela, text='Usuário ', width=20, font=('Ivy 10'), bg=co1, fg=co4)
l_usuario.place(x=130, y=80)

e_usuario = Entry(janela, width=20, font=('', 15), highlightthickness=1, relief='solid')
e_usuario.place(x=110, y=100)

l_senha = Label(janela, text='Senha', font=('Ivy 10'), bg=co1, fg=co4)
l_senha.place(x=190, y=150)

e_senha = Entry(janela, width=20, font=('Arial 10', 15), show='*', highlightthickness=1, relief='solid')
e_senha.place(x=110, y=170)

b_entrar = Button(janela, text='ENTRAR', command=lambda: verificar_login(e_usuario.get(), e_senha.get()), width=20, height=1, font=('Ivy 8 bold'), bg=co1, fg=co4)
b_entrar.place(x=150, y=260)


usuario = e_usuario.get()
senha = e_senha.get()

def verificar_login(usuario, senha):
    if credenciais[0] == usuario and credenciais[1] == senha:
        messagebox.showinfo('Login', 'Seja bem vindo, login realizado com sucesso!!')
        marcar_cliente()
    else:
        messagebox.showwarning('Erro', 'Verifique suas credenciais novamente.')


janela.mainloop()
