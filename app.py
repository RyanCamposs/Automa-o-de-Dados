import pyautogui
import openpyxl
import pyperclip
from time import sleep

#entrar na planilha
workbook = openpyxl.load_workbook('produtos_ficticios.xlsx')
sheet_produtos = workbook['Produtos']

#para cada linha na pagina de produtos, começar a partir da linha 2
for linha in sheet_produtos.iter_rows(min_row=2):
   nome_produto = linha[0].value
   pyperclip.copy(nome_produto)
   pyautogui.click(1253, 342, duration=1)
   pyautogui.hotkey('ctrl', 'v')
   
   descricao = linha[1].value
   pyperclip.copy(descricao)
   pyautogui.click(1140, 462, duration=1)
   pyautogui.hotkey('ctrl', 'v')
   
   categoria = linha[2].value
   pyperclip.copy(categoria)
   pyautogui.click(1133, 569, duration=1)
   pyautogui.hotkey('ctrl', 'v')
   
   codigo_produto = linha[3].value
   pyperclip.copy(codigo_produto)
   pyautogui.click(1143, 651, duration=1)
   pyautogui.hotkey('ctrl', 'v')
   
   peso = linha[4].value
   pyperclip.copy(peso)
   pyautogui.click(1106, 737, duration=1)
   pyautogui.hotkey('ctrl', 'v')
   
   dimensoes = linha[5].value
   pyperclip.copy(dimensoes)
   pyautogui.click(1118, 822, duration=1)
   pyautogui.hotkey('ctrl', 'v')
   
   pyautogui.click(1119,881,duration=1)
   sleep(3)
   
   
   preco = linha[6].value
   pyperclip.copy(preco)
   pyautogui.click(1136,383, duration=1)
   pyautogui.hotkey('ctrl', 'v')
   
   quantidade_em_estoque = linha[7].value
   pyperclip.copy(quantidade_em_estoque)
   pyautogui.click(1106,475, duration=1)
   pyautogui.hotkey('ctrl', 'v')
   
   data_validade = linha[8].value
   pyperclip.copy(data_validade)
   pyautogui.click(1107,563, duration=1)
   pyautogui.hotkey('ctrl', 'v')
   
   cor = linha[9].value
   pyperclip.copy(cor)
   pyautogui.click(1113,647, duration=1)
   pyautogui.hotkey('ctrl', 'v')
   
   
   
   tamanho = linha[10].value
   sleep(1)
   pyautogui.click(1133,738, duration=1)
   if tamanho == 'Pequeno':
      pyautogui.click(1123,770,duration=1)
   elif tamanho == "Médio":
      pyautogui.click(1110, 790, duration=1)
   else: 
      pyautogui.click(1111, 810, duration=1)
      
   material = linha[11].value
   pyperclip.copy(material)
   pyautogui.click(1118,817, duration=1)
   pyautogui.hotkey('ctrl', 'v')
      
   pyautogui.click(1120,870,duration=1)
   sleep(3)
   
   fabricante = linha[12].value
   pyperclip.copy(fabricante)
   pyautogui.click(1137,411, duration=1)
   pyautogui.hotkey('ctrl', 'v')
   
   pais_de_origem = linha[13].value
   pyperclip.copy(pais_de_origem)
   pyautogui.click(1118,491, duration=1)
   pyautogui.hotkey('ctrl', 'v')
   
   observacao = linha[14].value
   pyperclip.copy(observacao)
   pyautogui.click(1108,602, duration=1)
   pyautogui.hotkey('ctrl', 'v')
   
   codigo_de_barra = linha[15].value
   pyperclip.copy(codigo_de_barra)
   pyautogui.click(1109,714, duration=1)
   pyautogui.hotkey('ctrl', 'v')
   
   localizacao_no_armazem = linha[16].value
   pyperclip.copy(localizacao_no_armazem)
   pyautogui.click(1110,801, duration=1)
   pyautogui.hotkey('ctrl', 'v')
  
  # Concluir o envio
   pyautogui.click(1115,860,duration=1)
  # Confirmar envio
   pyautogui.click(1616,168,duration=1)
  # Adicionar o proximo produto
   pyautogui.click(1434,615,duration=1)   

