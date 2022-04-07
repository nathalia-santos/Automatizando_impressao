#!/usr/bin/env python
# coding: utf-8

# In[2]:


from openpyxl import load_workbook
from itertools import cycle
import time
import win32print
import win32api
import traceback
import os
from tkinter import *
from tkinter import ttk
from tkinter import font # * doesn't import font or messagebox
from tkinter import messagebox
from tkinter.filedialog import askopenfilename


root = Tk()
root.title('Canhoto')
root.geometry('700x700')
root.resizable(False, False)
root.tk.call('encoding', 'system', 'utf-8')


mainframe = Frame(root)
#mainframe.grid(column=0,row=0, sticky=(N,W,E,S) )
mainframe.grid(column=0,row=0, sticky=(N) )
mainframe.columnconfigure(0, weight = 1)
mainframe.rowconfigure(0, weight = 1)
mainframe.pack(pady = 10, padx = 0)

info = Label(mainframe, text='Sistema de geração de canhotos', fg='black', bg='#00E6FF', font='baskerville 20 bold')
info.grid(row=0, column=1, sticky="ew")

celulas_editaveis = ['C1', 'C2', 'C3', 'C4', 'C5', 'C6']
canhoto_wb = load_workbook('planilha.xlsx')
canhoto_ws = canhoto_wb.active
numero_lista = []


def gerar_canhotos():
    cont = 0
    numero_inicial = int(n1.get())
    numero_final = int(n2.get())
    for numero in range(numero_inicial, numero_final + 1):
        numero_lista.append(numero)
    for a, b in zip(cycle(celulas_editaveis), numero_lista):
        canhoto_ws[a] = str(f'Nº.:      {b}')
        time.sleep(2)
        cont += 1
        if cont > 5:
            canhoto_wb.save('planilha.xlsx')
            time.sleep(2)
            win32api.ShellExecute(0, "print", 'planilha.xlsx', None, 'caminho', 0)
            time.sleep(1)
            cont = 0

entrada_n1 = Label(mainframe, text='Insira o número inicial ', font='baskerville 15 bold')
entrada_n1.grid(row = 6, column = 1)
n1 = Entry(mainframe, highlightbackground='black', highlightthickness=2, width=40)
n1.grid(row = 6, column = 2)

entrada_n2 = Label(mainframe, text='Insira o número final', font='baskerville 15 bold')
entrada_n2.grid(row = 7, column = 1)
n2 = Entry(mainframe, highlightbackground='black', highlightthickness=2, width=40)
n2.grid(row = 7, column = 2)


botao = Button(text='Gerar canhoto', fg='white', bg='gray', highlightbackground='black', highlightthickness=3,
                  font='baskerville 12 bold', width=40, command=gerar_canhotos).pack()
root.mainloop()

