#v 1.1.0

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

    # Contadores para o controle
    paragrafoAtual = 0
    paragrafoTotal = len(Document(DOCXSelecionado).paragraphs)

    # Exibe o total de parágrafos no documento
    messagebox.showinfo("Total de parágrafos", str(paragrafoTotal))

    while paragrafoAtual < paragrafoTotal:

        # Obtém o conteúdo do parágrafo selecionado
        paragrafo = Document(DOCXSelecionado).paragraphs[paragrafoAtual].text

        # Exibe o indice do parágrafo selecionado, para fins de controle
        messagebox.showinfo("Index", str(paragrafoAtual))

        # Exibe o conteúdo do parágrafo selecionado
        messagebox.showinfo("Texto " + str(paragrafoAtual +1) + "/" + str(paragrafoTotal), paragrafo)

        # Passa para o próximo parágrafo
        paragrafoAtual += 1

    #---

#------------------------------------------

print('Feito')
