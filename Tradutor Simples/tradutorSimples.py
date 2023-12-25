"""
Tradutor Simples v1.0.0

Este é um programa simples que permite ao usuário traduzir texto para diferentes idiomas usando a API Google Translate.
O código está disponível no GitHub: https://github.com/Lord-Dralkhy/ModulosDo_UiPath-TradutorDePDF/tree/main/Tradutor%20Simples

Instruções de Uso:
1. Certifique-se de ter o arquivo 'Linguagens.txt' com o pack de idiomas no mesmo diretório do executável.
2. O programa pode ser executado diretamente, e uma interface gráfica será exibida.
3. Digite o texto a ser traduzido na caixa de diálogo e escolha o idioma desejado.
4. Clique em "OK" para ver a tradução.

Observação: Caso ocorram erros durante a tradução, o programa tentará fornecer informações detalhadas na tradução resultante.

Desenvolvido por:
Lord Dralkhy - Magno da Silva Gomes
GitHub: https://github.com/Lord-Dralkhy
"""

import tkinter as tk

from googletrans import Translator
from tkinter import simpledialog, messagebox
from tkinter.ttk import Combobox

# Constantes
ERRO_TRADUCAO = "<@X45H>Erro durante a tradução: {erro} \nTexto original: {frase}"

# Função exclusiva para enviar o texto e receber sua tradução
def traduzirFrase(frase, idiomaDestino='pt'):

    # Dando um feedback para o usuário
    print("Iniciando tradução")

    # Iniciando o módulo do tradutor
    translator = Translator()

    # Enviando o texto original para traduzir, e verificar se ouve erro com a tradução
    try:

        traducao = translator.translate(frase, dest=idiomaDestino)

        # Dando um feedback para o usuário
        print("Texto traduzido com sucesso!")

        return traducao.text

    # Caso aja erro, indicar isso, e devolver o texto original, para facilitar a programação de um robô
    except Exception as e:

        print(f"Erro durante a tradução: {e}")

        return ERRO_TRADUCAO.format(erro=e, frase=frase)


# Função para obter idiomas disponíveis
def obterIdiomasDisponiveis():

    # Dando um feedback para o usuário
    print("Obtendo pack de idiomas")

    try:

        # Lê o arquivo com o peck de idiomas
        with open("Linguagens.txt", 'r', encoding='utf-8') as arquivo:
            conteudo = arquivo.read()

        dicionario = {}

        # Separar as linhas, criando uma lista de strings
        linhas = conteudo.split('\n')

        # Fazer o tratamento de cada linha, para atender os padrões do projeto
        for linha in linhas:

            if linha.strip():

                chave, valor = linha.strip().split(':')
                dicionario[chave.strip()] = valor.strip().replace('\'', '').replace(',', '')

        # Criar uma lista, com apenas o nome dos idiomas
        opcoesIdiomas = list(dicionario.values())

        return opcoesIdiomas

    # Caso dê algum problema com o arquivo, do pack de idiomas, retorne alguma coisa
    except Exception as e:

        print(f"Erro durante a leitura do arquivo: {e}")

        return ["portuguese"]


# Função para escolher idioma
def escolherIdioma():

    # Iniciando uma janela temporária para lidar com as escolhas de idioma.
    # Ela está invisível para não atrapalhar a ordem da seleção
    root = tk.Tk()
    root.withdraw()

    # Obtendo um pack com os idiomas disponíveis
    opcoesIdiomas = obterIdiomasDisponiveis()

    # Pedindo um idioma, caso seja um robô que esteja usando o programa, isso facilita a programação dele
    idiomaDestino = simpledialog.askstring("Selecione um idioma",
                                           "Escolha um idioma para tradução (deixe em branco se quiser ver a lista de idiomas compatíveis):",
                                           parent=root)

    # Verificando se o idioma digitado é válido
    if idiomaDestino not in opcoesIdiomas:

        # Dando um feedback para o usuário, caso ele não digite algum idioma válido
        print("Nenhum idioma compatível foi digitado")

        # Tratamento de erro, para caso der erro no arquivo com o pack de idiomas
        try:

            idiomaDestino = opcoesIdiomas[74]

        except Exception as e:

            print(f"Erro ao obter idioma padrão: {e}")
            idiomaDestino = "portuguese"

        # Ativando a janela da caixa de escolhas, para o usuário vê-la
        root.deiconify()

        # Configurando caixa de escolhas
        combobox = Combobox(root, values=opcoesIdiomas, state="readonly")
        combobox.set(idiomaDestino)
        combobox.pack()

        # Função do botão, para salvar a opção escolhida e dar continuidade ao programa
        def selecionarIdioma():
            nonlocal idiomaDestino
            idiomaDestino = combobox.get()
            root.destroy()

        # Configurar o botão
        botao = tk.Button(root, text="Selecionar", command=selecionarIdioma)
        botao.pack()

        # Aguardar o usuário escolher o idioma
        root.wait_window()

    else:

        # Destruir essa janela temporária
        root.destroy()

    # Dando um feedback para o usuário
    print("Idioma escolhido: " + idiomaDestino)

    return idiomaDestino

def main():

    # Dando um feedback para o usuário
    print("Iniciando sistema")

    idiomaDestino = escolherIdioma()

    # Cria uma janela principal oculta
    root = tk.Tk()
    root.withdraw()

    # Iniciando uma variável vazia
    textoOriginal = ""

    while not textoOriginal == "sair":

        # Bloco para tratamento de erros ou exceções
        try:

            # Solicita um texto ao usuário usando a caixa de diálogo
            textoOriginal = simpledialog.askstring("Texto a ser traduzido", "Digite o texto a ser traduzido:").strip()

        except (AttributeError, TypeError):

            # Usuário cancelou a janela ou ocorreu um erro
            textoOriginal = "Criado por"

        # Dando um feedback para o usuário
        if textoOriginal == "Criado por" or textoOriginal == "sair":

            # Exibe uma mensagem na tela com um botão "OK"
            messagebox.showinfo("Criado por:", "Este programa foi feito por:\nLord Dralkhy - Magno da Silva Gomes")
            textoOriginal = "sair"

        # Se o usuário não pediu para fechar o programa, ou digitou algo, faça a tradução
        if not textoOriginal == "sair" and not textoOriginal == "":

            # Exibe uma mensagem na tela com um botão "OK"
            messagebox.showinfo("Texto traduzido", traduzirFrase(textoOriginal, idiomaDestino))

    # Dando um feedback para o usuário
    print('Feito')

if __name__ == "__main__":
    main()