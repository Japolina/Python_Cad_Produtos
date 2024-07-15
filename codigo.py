from selenium import webdriver
from selenium.webdriver.common.by import By
import openpyxl

# Acessar site de sua preferência com a página que deseja fazer os registros na planilha, meu exemplo foi com site Kabum
driver = webdriver.Chrome()
driver.get('https://www.kabum.com.br/computadores/pc/pc-gamer')

# Extrair todos os titulos em inspecionar do site
titulos = driver.find_elements(By.XPATH, "//span[@class='sc-d79c9c3f-0 nlmfp sc-9d1f1537-16 fQnige nameCard']") # //nome-da-tag[@nome-da-classe='nome']

# Extrair todos os preços em inspecionar do site
precos = driver.find_elements(By.XPATH, "//span[@class='sc-b1f5eb03-2 iaiQNF priceCard']")

# Criando planilha
workbook = openpyxl.Workbook()
# Criando página 'produtos'
workbook.create_sheet('produtos')
# Selecionando a página produtos
sheet_produtos = workbook['produtos']
sheet_produtos['A1'].value = 'Produto'
sheet_produtos['B1'].value = 'Preço'

# Inserir os titulos e preços na planilha
for titulo, preco in zip(titulos, precos): # Se o produto estiver esgotado, o ZIP irá reconhecer o valor nulo e não irá inserir na planilha
    sheet_produtos.append([titulo.text, preco.text])

workbook.save('produtos.xlsx')