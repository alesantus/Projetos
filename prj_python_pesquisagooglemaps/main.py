# -*- coding = utf-8 -*-
#Importando bibliotecas
import os 
import sys
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
import time
import keyboard
import pandas as pd

#Cria uma função responsável por apagar o arquivo com base no caminho informado, pode ser reutilizado
def removearquivo(caminhoarquivo):
    if os.path.exists(caminhoarquivo):
        os.remove(caminhoarquivo)
        print(f"Arquivo '{caminhoarquivo}' apagado com sucesso")
    else:
        print(f"Arquivo '{caminhoarquivo}' não existe.")


# Inicia algumas variaveis que serão utilizadas no código
# Certifique-se de ter o driver do Chrome instalado e configurado no PATH do sistema.

driverpath = r".\Drive\chromedriver.exe"
caminhoxlsx = r".\Documentos\dados_empresas.xlsx"
site = 'https://www.google.com/maps'
nomes, contatos, localidades = [],[],[] 

print("Iniciando rotina")

#Chama a função criada anteriormente
removearquivo(caminhoarquivo=caminhoxlsx)

# Cria uma instância do navegador Chrome
service = Service(driverpath)
driver = webdriver.Chrome(service=service)

# Abre o Google Chrome no google maps
driver.get(site)
driver.maximize_window()

# Localiza a barra de pesquisa e procura por contabilidade
barrapesquisa = driver.find_element(By.XPATH, r'//*[@id="searchboxinput"]')
barrapesquisa.send_keys("Contabilidade")
time.sleep(2)
keyboard.press('enter')
time.sleep(5)

#Gera uma lista com as empresas visiveis em tela retornadas pelo google maps
try:
    empresas = driver.find_elements(By.CSS_SELECTOR, r'a[class="hfpxzc"]')
    print('Elemento localizado - Resultados google maps')
except:
    print('Elemento não localizado - Resultados google maps')
    print('Finalizando rotina - Erro: Elementos dos resultados do google não encontrados')
    sys.exit()

    
#Itera sobre a lista de empresas encontradas
for empresa in empresas: 
       
    empresa.click()
    time.sleep(3)
    
    #Inclui o nome da empresa na lista de nomes
    nomes.append(empresa.get_attribute('aria-label'))
    
    #Procura n° telefone, caso encontre faz a inclusão na lista contatos
    try:
        telefone = driver.find_element(By.CSS_SELECTOR, r"button[aria-label~='Telefone:']")
        contatos.append(telefone.get_attribute('outerText'))
    except:
        print(str(nomes[-1]) + ': elemento não localizado - Telefone')
        contatos.append("-")
        
    #Procura o endereço, caso encontre faz a inclusão na lista localidades
    try:
        endereco = driver.find_element(By.CSS_SELECTOR, r'button[aria-label~="Endereço:"]')
        localidades.append(endereco.get_attribute('outerText'))
    except:
        print(str(nomes[-1]) + ': elemento não localizado - Endereço')
        localidades.append("-")
    
print("Fim da iteração dos resultados encontrados no google")

#Iniciar dicionário com os valores capturados
dadosempresas = {'nome_empresa': nomes,
                    'contato_telefone': contatos,
                    'endereço_comercial': localidades} 

#Gerar dataFrame com base no dicionário anterior
df_dadosempresas = pd.DataFrame(data=dadosempresas)  

#Escrever dataFrame em uma planilha excel, criando assim um arquivo xlsx
writer = pd.ExcelWriter(caminhoxlsx)
df_dadosempresas.to_excel(writer, sheet_name='dados_empresas', index=False)
writer.close()     

driver.quit()
print("Rotina finalizada com sucesso")
sys.exit()