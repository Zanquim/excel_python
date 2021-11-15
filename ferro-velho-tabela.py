from os import read
from openpyxl import Workbook, load_workbook
import pandas as pd
from bs4 import BeautifulSoup
import requests
import email.message
import smtplib
import time
import pyperclip
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
import urllib
import pyautogui
import pandas as pd 
from matplotlib import pyplot as plt
import tabula

ferro_velho = load_workbook('Entrada 2.xlsx')

aba_ativa = ferro_velho.active

#def inserir_na_celula():
    #for célula1 in range(0, 4):
        #letra1 = str(input(f'Insira uma letra: ').upper())
        #numero1 = str(input(f'Insira um número: ').upper())
        #aba_ativa[letra1 + numero1] = str(input(f'Insira um valor: '))
        #ferro_velho.save('Entrada 2.xlsx')

#inserir_na_celula()


ferro = pd.read_excel('Entrada 2.xlsx')
ferro = ferro.set_index('Material')
ferro1 = str(ferro)
print(ferro1)


navegador = webdriver.Chrome()
navegador.get('https://web.whatsapp.com/')

while len(navegador.find_elements_by_id('side')) < 1:
  time.sleep(1)

numw =  [5511988016162]

for enviar in numw:
  texto = urllib.parse.quote(ferro1)
  
  link = f'https://web.whatsapp.com/send?phone={enviar}&text={texto}'
  navegador.get(link)

  while len(navegador.find_elements_by_id('side')) < 1:
    time.sleep(1)
  time.sleep(10)
  
  pyautogui.moveTo(1015, 669)
  pyautogui.click()
  time.sleep(10)