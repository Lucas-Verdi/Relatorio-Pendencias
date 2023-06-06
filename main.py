import time
import tkinter as tk
from tkinter import *
from tkinter import filedialog
import xlwings
import pyautogui
import win32com.client as win32
from pyautogui import sleep
import sys
from threading import Thread

class Th(Thread):

    def __init__(self, num):
        Thread.__init__(self)
        self.num = num


    def run(self):
        # Criando janela para selecionar o arquivo
        root = tk.Tk()
        root.withdraw()
        arquivo = filedialog.askopenfilename()

        # Abrindo a planilha selecionada
        pastadetrabalho = xlwings.Book(arquivo)

        # Abre o Excel em tela cheia
        excel_window = pyautogui.getWindowsWithTitle("Excel")[0]
        excel_window.maximize()

        #Selecionando a planilha
        planilha = pastadetrabalho.sheets["posicoes pendencias vendas comp"]

        #Deletando colunas desnecessárias
        planilha.range('D:F').delete()
        planilha.range('E:E').delete()
        planilha.range('F:H').delete()
        planilha.range('G:H').delete()
        planilha.range('H:I').delete()
        planilha.range('J:W').delete()
        planilha.range('K:U').delete()
        planilha.range('M:T').delete()


        #Movendo as colunas para as posições corretas

        planilha.range('A:L').select()
        pyautogui.moveTo(400, 0)
        pyautogui.click()
        pyautogui.hotkey('ctrl', 'x')
        pyautogui.press('right', presses=12)
        pyautogui.hotkey('ctrl', 'v')

        sleep(0.1)

        planilha.range('O:O').select()
        pyautogui.hotkey('ctrl', 'x')
        pyautogui.press('left', presses=14)
        pyautogui.hotkey('ctrl', 'v')

        sleep(0.1)


        planilha.range('V:V').select()
        pyautogui.hotkey('ctrl', 'x')
        pyautogui.press('left', presses=20)
        pyautogui.hotkey('ctrl', 'v')

        sleep(0.1)


        planilha.range('S:S').select()
        pyautogui.hotkey('ctrl', 'x')
        pyautogui.press('left', presses=16)
        pyautogui.hotkey('ctrl', 'v')

        sleep(0.1)


        planilha.range('R:R').select()
        pyautogui.hotkey('ctrl', 'x')
        pyautogui.press('left', presses=14)
        pyautogui.hotkey('ctrl', 'v')

        sleep(0.1)


        planilha.range('N:N').select()
        pyautogui.hotkey('ctrl', 'x')
        pyautogui.press('left', presses=9)
        pyautogui.hotkey('ctrl', 'v')

        sleep(0.1)


        planilha.range('T:U').select()
        pyautogui.hotkey('ctrl', 'x')
        pyautogui.press('left', presses=14)
        pyautogui.hotkey('ctrl', 'v')

        sleep(0.1)


        planilha.range('Q:Q').select()
        pyautogui.hotkey('ctrl', 'x')
        pyautogui.press('left', presses=9)
        pyautogui.hotkey('ctrl', 'v')

        sleep(0.1)


        planilha.range('M:M').select()
        pyautogui.hotkey('ctrl', 'x')
        pyautogui.press('left', presses=4)
        pyautogui.hotkey('ctrl', 'v')

        sleep(0.1)


        planilha.range('P:P').select()
        pyautogui.hotkey('ctrl', 'x')
        pyautogui.press('left', presses=6)
        pyautogui.hotkey('ctrl', 'v')

        sleep(0.1)


        planilha.range('W:X').select()
        pyautogui.hotkey('ctrl', 'x')
        pyautogui.press('left', presses=12)
        pyautogui.hotkey('ctrl', 'v')


def start():
    a = Th(1)
    a.start()

#Interface
janela = Tk()
janela.title('Pendências')
Label1 = Label(janela, text='Insira a pasta de trabalho:')
Label1.grid(column=0, row=0, padx=10, pady=10)
Botao1 = Button(janela, text='Inserir')
Botao1.bind("<Button>",  lambda e: start())
Botao1.grid(column=0, row=1, padx=10, pady=10)
janela.mainloop()




