from selenium import webdriver
from selenium.webdriver.common.by import By
import openpyxl

# Acessar o site
driver = webdriver.Chrome()
driver.get('https://www.kabum.com.br/computadores/pc/pc-gamer')

# Extrair todos os títulos
# //tag[@atributo='valor']
titulos = driver.find_elements(By.XPATH, "//span[@class='sc-d79c9c3f-0 nlmfp sc-cdc9b13f-16 eHyEuD nameCard']")

# Extrair todos os preços
precos = driver.find_elements(By.XPATH, "//span[@class='sc-620f2d27-2 bMHwXA priceCard']")

# Criando a planilha
workbook = openpyxl.Workbook()

# Criando a página "Produtos"
workbook.create_sheet('Produtos')

# Selecionando a página "Produtos"
sheet_produtos = workbook['Produtos']
sheet_produtos['A1'].value = 'Produto'
sheet_produtos['B1'].value = 'Preço'

# Inserir os títulos e preços na planilha
# zip() - para/interrompe a iteração caso um dos elementos (titulos ou precos) acabe.
for titulo, preco in zip(titulos, precos):
    sheet_produtos.append([titulo.text, preco.text])

workbook.save('produtos.xlsx')
