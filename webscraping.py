from selenium.webdriver.chrome.service import Service
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
import pandas as pd

service = Service(executable_path="chromedriver.exe")
navegador = webdriver.Chrome(service=service)

# PEGAR COTAÇÃO DO DÓLAR

navegador.get('https://www.google.com')

navegador.find_element('xpath', '/html/body/div[1]/div[3]/form/div[1]/div[1]/div[1]/div/div[2]/input').send_keys('cotação dolar')

navegador.find_element('xpath', '/html/body/div[1]/div[3]/form/div[1]/div[1]/div[1]/div/div[2]/input').send_keys(Keys.ENTER)

cotacao_dolar = navegador.find_element('xpath', '//*[@id="knowledge-currency__updatable-data-column"]/div[1]/div[2]/span[1]').get_attribute('data-value')
#cotacao_dolar = navegador.find_element('xpath', '//*[@id="knowledge-currency__updatable-data-column"]/div[1]/div[2]/span[1]').text

print(f' A cotação do dolár é R$ {cotacao_dolar}')

# PEGAR COTAÇÃO DO EURO

navegador.get('https://www.google.com')

navegador.find_element('xpath', '/html/body/div[1]/div[3]/form/div[1]/div[1]/div[1]/div/div[2]/input').send_keys('cotação euro')

navegador.find_element('xpath', '/html/body/div[1]/div[3]/form/div[1]/div[1]/div[1]/div/div[2]/input').send_keys(Keys.ENTER)

cotacao_euro = navegador.find_element('xpath', '//*[@id="knowledge-currency__updatable-data-column"]/div[1]/div[2]/span[1]').get_attribute('data-value')
#cotacao_dolar = navegador.find_element('xpath', '//*[@id="knowledge-currency__updatable-data-column"]/div[1]/div[2]/span[1]').text

print(f' A cotação do euro é R$ {cotacao_euro}')

# PEGAR COTAÇÃO DO OURO

navegador.get('https://www.melhorcambio.com/ouro-hoje')

cotação_ouro = navegador.find_element('xpath', '//*[@id="comercial"]').get_attribute('value').replace(',', '.')

print(f' A cotação do ouro é R$ {cotação_ouro}')

# IMPORTAR DADOS DA PLANILHA

tabela = pd.read_excel('Produtos.xlsx')
print(tabela)

# RECALCULAR A COTAÇÃO E ATUALIZAR A PLANILHA

tabela.loc[tabela['Moeda'] == 'Dólar', 'Cotação'] = float(cotacao_dolar)
tabela.loc[tabela['Moeda'] == 'Euro', 'Cotação'] = float(cotacao_euro)
tabela.loc[tabela['Moeda'] == 'Ouro', 'Cotação'] = float(cotação_ouro)

# preço de compra = cotação * preço original
tabela['Preço de Compra'] = tabela['Cotação'] * tabela['Preço Original']
# preço de venda = preço de compra * margem
tabela['Preço de Venda'] = tabela['Preço de Compra'] * tabela['Margem']

tabela.to_excel('Produtos.xlsx', index=False)

