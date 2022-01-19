# Passo 1 - importanto as bibliotecas
import pandas as pd
import pyautogui as pg
import pyperclip as pc
import time
import openpyxl
# Passo 2 - Cadastrando o Funcionário
nome = "Kaique Denobi Felipe"
tel = "**********"
sal = "**********"

# Passo 4 - Iniciando o programa
print("Bem vindo!")
print("Vamos dar inicio ao calculo de faturamentos.")
print("Iniciando...")
pg.alert("O Programa está começando, não mexa no computador até o final")
pg.PAUSE = 1
# Passo 5 - Baixando a PLanilha
time.sleep(15)
pg.hotkey("winleft" , "1")
pg.hotkey("Ctrl" , "t")
Tabela = "https://docs.google.com/spreadsheets/d/1l_8dBM8BMrd9vC0nqmZ0qG4k3lwLbtRb/edit?usp=sharing&ouid=113271375905089042953&rtpof=true&sd=true"
pc.copy(Tabela)
pg.hotkey("Ctrl" , "v")
pg.press("Enter")
time.sleep(20)
pg.click(x=84, y=162)
pg.click(x=327, y=431)
pg.click(x=538, y=436)
time.sleep(10)
# Passo 6 - Lendo os dados da planilha
pg.hotkey("winleft" , "9")
excel = pd.read_excel(r'C:\Users\Maq2\Downloads\FATURAMENTO KAIQUE.xlsx')
faturamento = excel['Faturamento'].sum()
fat_loja_mensal = excel['Fat_Mensal_Loja'].sum()
percentual = (faturamento/fat_loja_mensal)*100

# Passo 7 enviando o e-mail
pg.hotkey("winleft" , "1")
pg.hotkey("Ctrl" , "Shift" , "n")
link_email = 'https://outlook.live.com/owa/'
pc.copy(link_email)
pg.hotkey("Ctrl" , "v")
pg.press("Enter")
time.sleep(20)
pg.click(x=1179, y=128)
time.sleep(20)
email = '**************'
pc.copy(email)
pg.hotkey("Ctrl" , "v")
pg.press("Enter")
senha = '**********'
pc.copy(senha)
pg.hotkey("Ctrl" , "v")
pg.press("Enter")
pg.press("Enter")
time.sleep(15)
pg.click(x=123, y=193)
pc.copy(email)
pg.hotkey("Ctrl" , "v")
pg.press("Tab")
assunto = "Email relatório Kaique"
pc.copy(assunto)
pg.hotkey("Ctrl" , "v")
Corpo_email = f""""Olá Bom dia,
Segue abaixo o relatório Anual do vendedor {nome},
Os dados de contato do mesmo : E-mail {email}, Telefone {tel}.
O salario do Funcionário {nome} é de R$ {sal}.
Seu rendimento anual é de R$ {faturamento}, o faturamento anual da loja é de R$ {fat_loja_mensal}
O funcionário {nome} vendeu uma porcentagem de {percentual} eem relação ao rendimento anual da loja!

Att
Kaique Denobi Felipe"""
pg.press("Tab")


pc.copy(Corpo_email)
pg.hotkey("Ctrl" , "v")
pg.hotkey("Ctrl" , "Enter")
time.sleep(20)
pg.click(x=331, y=355)
time.sleep(20)

pg.hotkey("Alt" , "F4")
pg.alert("FIM")
