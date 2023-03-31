import pyautogui
import time
import webbrowser
import pandas as pd

#váriaveis fixas
destinatario = " " #insira o email de destino
assunto_desc = "Relatório de Vendas de Ontem"

# Entrar no sistema da empresa
pyautogui.alert("Vai começar, NÃO TOQUE EM NADA!!! Aperte OK para continuar")
webbrowser.open(" ")#Pasta no Drive com a planilha desejada
time.sleep(5)

# Navegar até o local
pyautogui.doubleClick(368,268, duration=2)
time.sleep(5)
pyautogui.click(361,335, duration=2, button='right')

# Fazer donwload do relatório
pyautogui.click(473,641, duration= 2)
time.sleep(10)
# Lê a planilha do excel
df = pd.read_excel(r" ") #inserir o caminho da planilha dentro de seu computador
print(df)

# Calcular os indicadores
faturamento = df['Valor Final'].sum()
print(faturamento)
produtos = df['Quantidade'].sum()
print(produtos)

# Enviar o email para outro email
webbrowser.open(destinatario)
time.sleep(5)
pyautogui.click(73,167, duration=2) #novo email
pyautogui.click(881,241, duration=2) #clica no "para"
time.sleep(3)

#Digitando as informações do email
pyautogui.write(' ' ) #insira seu email aqui
time.sleep(3)
pyautogui.press('enter')#completa o email
pyautogui.press('tab') #vai para assunto
time.sleep(3)
pyautogui.write(assunto_desc)
pyautogui.press('tab') #descrição
time.sleep(3)
texto = f"""

Prezados, bom dia

O faturamento de ontem foi de: R${faturamento:,.2f}
A quantidade de produtos foi de: {produtos:,}

Abçs
Guilherme"""

#Envia o email
pyautogui.write(texto)
pyautogui.press('tab')
time.sleep(3)
pyautogui.press('enter')

pyautogui.alert("Processo Finalizado.")

