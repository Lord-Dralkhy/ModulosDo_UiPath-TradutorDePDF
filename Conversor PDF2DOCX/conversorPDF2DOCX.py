"""
Conversor de PDF para DOCX v1.2.0

Descrição Geral:
Este é um programa simples que possibilita a conversão de arquivos PDF para o formato DOCX, tornando-os editáveis no Microsoft Word, utilizando a biblioteca PDF2DOCX.

O código-fonte está disponível no GitHub: https://github.com/Lord-Dralkhy/ModulosDo_UiPath-TradutorDePDF/tree/main/Conversor%20PDF2DOCX

Instruções de Uso:
1. Execute o programa, e uma interface gráfica será exibida.
2. Insira o caminho do arquivo PDF na caixa de diálogo, com o nome do arquivo, e sem o formato.
3. Clique em "OK" para iniciar a conversão.

Observação:
O arquivo será salvo no mesmo local do PDF.

Desenvolvido por:
Lord Dralkhy - Magno da Silva Gomes
GitHub: https://github.com/Lord-Dralkhy
"""

import tkinter as tk

from pdf2docx import Converter
from tkinter import simpledialog, messagebox

def obterCaminhoPDF():

    # Solicita o caminho do PDF ao usuário
    return simpledialog.askstring("Escolher PDF", "Digite o caminho do PDF, com o nome do arquivo, e sem o formato:")

def processarPDF(caminhoPDF):

    # Salva o diretório dos arquivos a serem trabalhados
    arquivoPDF = caminhoPDF + '.pdf'
    arquivoDOCX = caminhoPDF + '.docx'

    try:

        # Seleciona o arquivo PDF
        conv = Converter(arquivoPDF)

        # Converte para docx
        conv.convert(arquivoDOCX)

        # Encerra o processo
        conv.close()

        exibirMensagem("Concluído", f"PDF convertido com sucesso para: \n{arquivoDOCX}", "info")

    except Exception as e:

        exibirMensagem("Erro", f"Erro ao processar o arquivo PDF: {str(e)}", "error")

def exibirMensagem(titulo, mensagem, tipo):

    # Exibe uma mensagem na tela com um botão "OK"
    if tipo == "info":

        messagebox.showinfo(titulo, mensagem)

    elif tipo == "error":

        messagebox.showerror(titulo, mensagem)

def main():

    # Dando um feedback para o usuário
    print('Iniciando sistema')

    # Cria uma janela principal oculta
    root = tk.Tk()
    root.withdraw()

    # Dando um feedback para o usuário
    print('Obtendo PDF')

    PDFSelecionado = ""

    while PDFSelecionado == "":

        # Solicita o caminho do arquivo PDF
        PDFSelecionado = obterCaminhoPDF()

    if PDFSelecionado is None:

        exibirMensagem("Criado por:", "Este programa foi feito por:\nLord Dralkhy - Magno da Silva Gomes", "info")

    else:

        # Processa o PDF
        processarPDF(PDFSelecionado)

    # Dando um feedback para o usuário
    print('Feito')

if __name__ == "__main__":
    main()

