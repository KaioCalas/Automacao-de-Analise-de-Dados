import pyautogui
import pyperclip
import time

pyautogui.PAUSE = 3


#ABRIR O NAVEGADOR
pyautogui.press('winleft')
pyautogui.write('firefox')
pyautogui.press('enter')

#ENTRAR NO SITE

pyautogui.hotkey('ctrl', 't')
link = 'https://drive.google.com/drive/folders/149xknr9JvrlEnhNWO49zPcw0PW5icxga'
pyperclip.copy(link)
pyautogui.hotkey('ctrl', 'v')
pyautogui.press('enter')

#BAIXAR O ARQUIVO EXPORTAR
time.sleep(5)
pyautogui.click(x=471, y=287, clicks=2)
pyautogui.click(x=471, y=287)
pyautogui.click(x=1161, y=203)
pyautogui.click(x=952, y=612)
pyautogui.click(x=432, y=477)
pyautogui.press('enter')

#Calcular o faturamento e quantidade
import pandas as pd
tabela = pd.read_excel(r'/home/caio/Downloads/Vendas - Dez.xlsx')
faturamento = tabela['Valor Final'].sum()
quantidade = tabela['Quantidade'].sum()

#Enviar o Email

link2 = 'https://mail.google.com/mail/u/0/'
pyperclip.copy(link2)
pyautogui.hotkey('ctrl', 't')
pyautogui.hotkey('ctrl', 'v')
pyautogui.press('enter')
time.sleep(5)
pyautogui.click(x=84, y=226)
time.sleep(3)

pyautogui.write('caladkain@gmail.com')
pyautogui.press('tab')
pyautogui.press('tab')

assunto = 'Relatório de Vendas'
pyperclip.copy(assunto)
pyautogui.hotkey('ctrl', 'v')
pyautogui.press('tab')

texto = f'''
Prezados, Bom dia!

O faturamento foi de R${faturamento:,.2f}.
A quantidade de produtos foi de R${quantidade:,.2f}.

Abs
Kaio Costa.
'''
pyautogui.write(texto)
pyautogui.hotkey('ctrl', 'enter')

