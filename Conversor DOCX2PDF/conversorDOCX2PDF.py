'''
Conversor de DOCX para PDF v1.0.0

Descrição Geral:
Este é um programa simples que possibilita a conversão de arquivos DOCX para o formato PDF, tornando-os visualizáveis em leitores de PDF, utilizando a biblioteca docx2pdf.

O código-fonte está disponível no GitHub: [link]

Instruções de Uso:
1. Execute o programa, e uma interface gráfica será exibida.
2. Insira o caminho do arquivo DOCX na caixa de diálogo, com o nome do arquivo e o formato.
3. Clique em "OK" para iniciar a conversão.

Observação:
O arquivo PDF será salvo no mesmo local do arquivo DOCX, com o sufixo "(Traduzido)" adicionado ao nome.

Desenvolvido por:
Lord Dralkhy - Magno da Silva Gomes
GitHub: https://github.com/Lord-Dralkhy

Aproveite a simplicidade e eficiência para converter um documento DOCX para um documento PDF!
'''

import tkinter as tk

from tkinter import simpledialog, messagebox
from docx2pdf import convert

def obterCaminhoDOCX():

    # Dando feedback para o usuário
    print("Obtendo arquivo")

    # Solicita o caminho do DOCX ao usuário
    return simpledialog.askstring("Escolher DOCX", "Digite o caminho do DOCX, com o nome do arquivo e o formato:")

def processarDOCX(caminhoDOCX):

    # Dando feedback para o usuário
    print("Convertendo arquivo")

    # Adiciona "(Traduzido)" ao nome do arquivo PDF
    caminhoPDF = caminhoDOCX + ' (Traduzido).pdf'

    try:

        # Converte DOCX para PDF - C:\Users\Magno\Documents\UiPath\UiPath Certified Professional - Automation Developer Associate Exam Description
        convert(caminhoDOCX, caminhoPDF)

        print(f"Arquivo PDF criado com sucesso: {caminhoPDF}")

        exibirMensagem("Concluído", f"DOCX convertido com sucesso para PDF.", "info")

    except Exception as e:

        # Dando feedback para o usuário
        print(f"Erro ao processar o arquivo DOCX: {str(e)}", "error")

        exibirMensagem("Erro", f"Erro ao processar o arquivo DOCX: {str(e)}", "error")

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

    DOCXSelecionado = ""

    while DOCXSelecionado == "":

        # Solicita o caminho do arquivo DOCX
        DOCXSelecionado = obterCaminhoDOCX()

    if DOCXSelecionado is None:

        exibirMensagem("Criado por:", "Este programa foi feito por:\nLord Dralkhy - Magno da Silva Gomes", "info")

    else:

        # Processa o DOCX
        processarDOCX(DOCXSelecionado)

    # Dando um feedback para o usuário
    print('Feito')

if __name__ == "__main__":
    main()
