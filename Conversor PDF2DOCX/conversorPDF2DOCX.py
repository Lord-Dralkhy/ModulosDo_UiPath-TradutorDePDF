#V 1.1.0

import tkinter as tk

from pdf2docx import Converter
from tkinter import simpledialog, messagebox

# Cria uma janela principal oculta
root = tk.Tk()
root.withdraw()

# Solicita um texto ao usuário usando a caixa de diálogo
PDFSelecionado = simpledialog.askstring("Escolher PDF", "Digite o caminho do PDF, com o nome do arquivo, e sem o formato:")

if PDFSelecionado == "Criado por":

    # Exibe uma mensagem na tela com um botão "OK"
    messagebox.showinfo("Criado por:", "Este programa foi feito por:\nLord Dralkhy - Magno da Silva Gomes")

else:

    # Salvar o diretório dos arquivos a ser trabalhados
    arquivoPDF = PDFSelecionado + '.pdf'
    arquivoDOCX = PDFSelecionado + '.docx'

    # Selecionar o arquivo PDF
    conv = Converter(arquivoPDF)

    # Converter para docx
    conv.convert(arquivoDOCX)

    # Encerrar o processo
    conv.close()

print('Feito')