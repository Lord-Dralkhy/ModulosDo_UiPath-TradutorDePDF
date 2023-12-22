#v 1.0.0

import tkinter as tk
from docx import Document

from tkinter import simpledialog, messagebox

#------------------------------------------

# Cria uma janela principal oculta
root = tk.Tk()
root.withdraw()

# Solicita um texto ao usuário usando a caixa de diálogo
DOCXSelecionado = simpledialog.askstring("Escolher DOCX", "Digite o caminho do arquivo DOCX (com o formato):")

#------------------------------------------

if DOCXSelecionado == "Criado por":

    # Exibe uma mensagem na tela com um botão "OK"
    messagebox.showinfo("Criado por:", "Este programa foi feito por:\nLord Dralkhy - Magno da Silva Gomes")

else:

#------------------------------------------

    #---

    paragrafoAtual = 0
    paragrafoTotal = len(Document(DOCXSelecionado).paragraphs)

    messagebox.showinfo("Total de parágrafos", str(paragrafoTotal))

    while paragrafoAtual < paragrafoTotal:

        paragrafo = Document(DOCXSelecionado).paragraphs[paragrafoAtual].text

        messagebox.showinfo("Index", str(paragrafoAtual))

        messagebox.showinfo("Texto " + str(paragrafoAtual +1) + "/" + str(paragrafoTotal), paragrafo)

        paragrafoAtual += 1

    #---

#------------------------------------------

print('Feito')