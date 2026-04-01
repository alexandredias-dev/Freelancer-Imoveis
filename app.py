
from selenium import webdriver
from selenium.webdriver.common.by import By
from datetime import datetime
import openpyxl

#1- Entrar no site https://www.imoveismartinelli.com.br/pesquisa-de-imoveis/?locacao_venda=V&id_cidade%5B%5D=41&finalidade=&dormitorio=&garagem=&vmi=&vma=&ordem=4


driver = webdriver.Chrome()
driver.get("https://www.imoveismartinelli.com.br/pesquisa-de-imoveis/?locacao_venda=V&id_cidade%5B%5D=41&finalidade=&dormitorio=&garagem=&vmi=&vma=&ordem=4")

#2- Inserir o preço, link da casa e data dentro da planilha
precos = driver.find_elements(By.XPATH,"//div[@class='card-valores']/div")

links = driver.find_elements(By.XPATH,"//a[@class='carousel-cell is-selected']")

workbook = openpyxl.load_workbook('imoveis.xlsx')
pagina_imoveis = workbook['precos']

for preco, link in zip(precos, links):
    preco_texto = preco.text
    link_url = link.get_attribute('href')
    data_atual = datetime.now().strftime("%d-%m-%Y")
    pagina_imoveis.append([preco_texto, link_url, data_atual])

workbook.save('imoveis.xlsx')