
import openpyxl 
import pyperclip 
import pyautogui
from time import sleep 
#entrar na planilha 
workbook = openpyxl.load_workbook("produtos_ficticios.xlsx")
sheet_produtos = workbook ['produtos']
 #copirar informaçoes de um campo e coolar no campo correspondente 
for linha in sheet_produtos.inter_rows(min_row=2):
 nome_produto = linha[0].value 
 pyperclip.copy(nome_produto)
 pyautogui.click(1518,305,duration=1)
 pyautogui.hotkey('crtl','v')
 descrição = linha[1].value
 pyperclip.copy(descricao)
 pyautogui.click(1472,413,duration =1)
sleep(3)
codigo_produto = linha[2].value 
#precisa pegar coordenadas do local de digitação e acrescentar pyautogui.click
pyautogui.click(0000,0000, duration=1)
pyautogui.hotkey #para barra de digitação
pyautogui.click(0000,0000, duration=1)
sleep(3)
 
 #####segundapagina######
 
 
material = linha[11].value 
fabricante = linha[12].value  
pais_origem = linha[13].value
observaçoes = linha[14].value 
codigo_de_barras = linha[15].value 
localização_armazem = linha[16].value

 
 
 
 
 