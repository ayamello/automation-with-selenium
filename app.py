from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
import pandas as pd

# criar navegador
navegador = webdriver.Edge()

# entrar no google
navegador.get('https://www.google.com.br/')

# localizar input para escrever no google e escrever cotação dólar
navegador.find_element(By.XPATH, 
            '/html/body/div[1]/div[3]/form/div[1]/div[1]/div[1]/div/div[2]/input').send_keys('cotação dólar') # XPATH = caminho para o elemento de um site
# pressionar enter
navegador.find_element(By.XPATH, 
            '/html/body/div[1]/div[3]/form/div[1]/div[1]/div[1]/div/div[2]/input').send_keys(Keys.ENTER)

# pegar cotação do dólar
cotacao_dolar = navegador.find_element(By.XPATH, 
            '//*[@id="knowledge-currency__updatable-data-column"]/div[1]/div[2]/span[1]').get_attribute('data-value')


# pegar cotação do euro
navegador.get('https://www.google.com.br/')

navegador.find_element(By.XPATH, 
            '/html/body/div[1]/div[3]/form/div[1]/div[1]/div[1]/div/div[2]/input').send_keys('cotação euro') # XPATH = caminho para o elemento de um site

navegador.find_element(By.XPATH, 
            '/html/body/div[1]/div[3]/form/div[1]/div[1]/div[1]/div/div[2]/input').send_keys(Keys.ENTER)

cotacao_euro = navegador.find_element(By.XPATH, 
            '//*[@id="knowledge-currency__updatable-data-column"]/div[1]/div[2]/span[1]').get_attribute('data-value')


# pegar cotação do ouro
navegador.get('https://www.melhorcambio.com/ouro-hoje')

cotacao_ouro = navegador.find_element(By.XPATH, '//*[@id="comercial"]').get_attribute('value')
cotacao_ouro = cotacao_ouro.replace(',','.')

navegador.quit() #fechar navegador


# importar base e atualizar as cotações
table = pd.read_excel('Produtos.xlsx')


# atualizar cotações na base de dados
table.loc[table['Moeda'] == 'Dólar', 'Cotação'] = float(cotacao_dolar) # .loc[linha, coluna]
table.loc[table['Moeda'] == 'Euro', 'Cotação'] = float(cotacao_euro)
table.loc[table['Moeda'] == 'Ouro', 'Cotação'] = float(cotacao_ouro)

print(table)

# atualizar colunas de preços
table['Preço de Compra'] = table['Preço Original'] * table['Cotação']

table['Preço de Venda'] = table['Preço de Compra'] * table['Margem']


# exportar tabela atualizada
table.to_excel('new_table_Produtos.xlsx', index=False)