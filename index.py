import pyautogui
import pyperclip 
import time
import pandas as pd

pyautogui.alert("Código irá começar!")
pyautogui.PAUSE = 1

pyautogui.press("win")
pyautogui.write("edge")
pyautogui.press("enter")
pyperclip.copy("https://drive.google.com/drive/folders/149xknr9JvrlEnhNWO49zPcw0PW5icxga")
pyautogui.hotkey("ctrl", "v")
pyautogui.press("enter")
time.sleep(3)

pyautogui.click(x=487, y=451, clicks=4)
time.sleep(1)
pyautogui.click(x=495, y=584)
time.sleep(1)
pyautogui.click(x=1663, y=198)
time.sleep(1)

pyautogui.click(x=1363, y=659)
time.sleep(4)

tabela = pd.read_excel(r"C:\Users\fabio\Downloads\Vendas - Dez.xlsx")

ftm = tabela["Valor Final"].sum()
qtd = tabela["Quantidade"].sum()

pyautogui.hotkey("ctrl", "t")

pyperclip.copy("https://outlook.live.com/mail/0/junkemail")
pyautogui.hotkey("ctrl", "v")
pyautogui.press("enter")
time.sleep(3)

pyautogui.click(x=237, y=181)
time.sleep(1)
pyautogui.click(x=463, y=249)
pyautogui.write("fabiodanille@hotmail.com")
time.sleep(1)
pyautogui.press("tab")
pyautogui.press("tab")

time.sleep(1)
pyperclip.copy("Relatório de Vendas")
pyautogui.hotkey("ctrl", "v")

pyautogui.press("tab")

texto = f"""Prezados,

Segue relatório de vendas.
Faturamento: R${ftm:,.2f}
Quantidade de produtos vendidos: {qtd:,}

Fico à disposição,
AT.TE.,
Fabio Danille
"""
time.sleep(1)
pyperclip.copy(texto)
pyautogui.hotkey("ctrl","v")

pyautogui.hotkey("ctrl", "enter")
pyautogui.alert("Código finalizado!")

