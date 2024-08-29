from selenium import webdriver
from selenium.webdriver.common.by import By
from datetime import datetime
import openpyxl

driver = webdriver.Chrome()
driver.get('https://www.imoveismartinelli.com.br/pesquisa-de-imoveis/?locacao_venda=V&id_cidade%5B%5D=129&finalidade=&dormitorio=&garagem=&vmi=&vma=&ordem=4')

# Seleciona o container que envolve o preço
precos_container = driver.find_elements(By.XPATH, "//div[@class='card-valores']")
links = driver.find_elements(By.XPATH, "//a[@class='carousel-cell is-selected']")

workbook = openpyxl.load_workbook('imoveis.xlsx') #Nome da planilha
pagina_imoveis = workbook['precos'] #Nome da folha da planilha

for preco_container, link in zip(precos_container, links):
    preco_promocional = ""
    preco_normal = ""
    
    try:
        # Captura o preço promocional
        preco_promocional_element = preco_container.find_element(By.XPATH, ".//div[not(small)]")
        preco_promocional = preco_promocional_element.text.strip()
    except:
        pass

    if not preco_promocional:
        try:
            # Captura o preço riscado, se não houver preço promocional
            preco_normal_element = preco_container.find_element(By.XPATH, ".//div[small]")
            preco_normal = preco_normal_element.text.strip().split()[-1]
        except:
            pass

    # Formata o preço final a ser inserido na planilha
    preco_final = preco_promocional if preco_promocional else preco_normal

    link_pronto = link.get_attribute('href')
    data_atual = datetime.now().strftime('%d/%m/%Y')
    pagina_imoveis.append([preco_final, link_pronto, data_atual])

    workbook.save('imoveis.xlsx')
