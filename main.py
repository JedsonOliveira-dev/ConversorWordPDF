import PySimpleGUI as sg
from os import path, listdir
from docx2pdf import convert
from tkinter import filedialog

layout = [
    [sg.Text('Converta Word para pdf')],
    [sg.Button('Converter')]
]


def diretorio():
    try:
        pasta = filedialog.askdirectory()
        caminho_pasta = path.join(pasta)
        lista_diretorios = listdir(pasta)
        return caminho_pasta, lista_diretorios
    except FileNotFoundError:
        return 1, 1


def gera_pdf():
    caminho_pasta_, lista_diretorios_ = diretorio()
    if lista_diretorios_ and caminho_pasta_:
        try:
            for item in range(len(lista_diretorios_)):
                arquivo = lista_diretorios_[item]
                try:
                    convert(caminho_pasta_ + "\\" + arquivo)
                except AssertionError:
                    pass
            sg.popup("Pronto", f"Conversão concluida.")
        except TypeError:
            pass
    else:
        sg.popup("Aviso", "Não há arquivos para conversão na pasta.", text_color='red', background_color='white', font='arial 10')


janela = sg.Window('Conversor', layout)

while True:
    event, values = janela.read()

    if event == sg.WIN_CLOSED:
        break

    gera_pdf()
