#pip install pyperclip
#pip install pyautogui
#pip install pandas
#pip install numpy
#pip install openpyxl
#pip install win32

import time

import pandas as pd
import pyautogui
import pyperclip

pyautogui.PAUSE = 1

time.sleep(0)
pyautogui.position()
print(pyautogui.position())

# Passo 1 - Entrar no Sistema ( No Caso Entrar no Linak )

#pyautogui.hotkey("ctrl", "t")
#pyautogui.write("https://drive.google.com/drive/folders/149xknr9JvrlEnhNWO49zPcw0PW5icxga")
#pyautogui.press("enter")

pyautogui.press("win")
pyautogui.write("chrome")
pyautogui.press("enter")
pyperclip.copy("https://drive.google.com/drive/folders/149xknr9JvrlEnhNWO49zPcw0PW5icxga")
pyautogui.hotkey("ctrl", "v")
pyautogui.press("enter")

time.sleep(10)

# Passo 2 - Navegar até o Local do Arquivo ( Entrar na Pasta Exportar )

pyautogui.click(x=96, y=279, clicks=2)
time.sleep(8)

# Passo 3 - Fazer Dawload do Arquivo

pyautogui.click(x=96, y=279)
time.sleep(8)

# Passo 4 - Calcular Indicadores

tabela = pd.read_excel(r"C:/Users/Hertz Informatica/Downloads/Vendas - Dez.xlsx")
print(tabela)

Faturamento = tabela["Valor Final"].sum()
Quantidade = tabela["Quantidade"].sum()

# Passo 5 - Entrar no Email

pyautogui.hotkey("ctrl", "t")
pyperclip.copy("https://mail.google.com/mail/u/1/#inbox")
pyautogui.hotkey("ctrl", "v")
pyautogui.press("enter")
time.sleep(8)

# Passo 6 - Enviar por Email o Relatorio

pyautogui.click(x=38, y=200)
pyautogui.write("email@gmail.com")
pyautogui.press("tab") #Seleciona Email
pyautogui.press("tab") #Pula pro Campo de Assunto
pyperclip.copy("Relatorio de Vendas") #Escrever o Assunto
pyautogui.hotkey("ctrl", "v") #Escrever o Assunto
pyautogui.press("tab") #Pular pro Corpo do Email

texto = f"""
Prezados, Bom Dia

O faturamento de Ontem foi de: R$ {Faturamento:,.2f}
A Quantidade de Vendas de Ontem foi de: {Quantidade:}

Att;
Eduardo Sousa

"""

pyperclip.copy(texto)
pyautogui.hotkey("ctrl", "v")

#Clicar Botão de Envio

pyautogui.hotkey("ctrl", "enter")
