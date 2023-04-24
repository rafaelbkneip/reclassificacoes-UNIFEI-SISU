import xlsxwriter  
from selenium import webdriver
import selenium.webdriver.support.expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait
from time import sleep
from datetime import date
from PyPDF2 import PdfReader, PdfWriter
import pdfplumber

from selenium.webdriver.common.keys import Keys

options = Options()
options.add_experimental_option("detach", True)
navegador = webdriver.Chrome(ChromeDriverManager().install(), options=options)
navegador2 = webdriver.Chrome(ChromeDriverManager().install(), options=options)

#Processos seletivos da UNIFAL
navegador.get("https://prg.unifei.edu.br/cops/sisu-2023-campus-itajuba-chamadas-da-lista-de-espera-convocados/")

sleep(10)

chamadas = navegador.find_elements(By.TAG_NAME, 'a')

print(chamadas)
print(len(chamadas))

nome = []
cota_inscricao = []
pontuacao = []
cota_convocacao = []
cursos=[]
situacao =[]

links= []

for p in range(len(chamadas)):
    try:
       print(chamadas[p].text)
       if(chamadas[p].text.split(" ")[-1] == '(convocados)'):
        links.append(chamadas[p].get_attribute('href'))
    except:
        print('Não é válido')

print(links, len(links))


for k in range(len(links)):
    navegador2.get(links[k])

    sleep(10)

    navegador2.find_element(By.XPATH, '//*[@id="download"]').click()

    

    nome_documento = (navegador2.find_element(By.XPATH, '//*[@id="header"]').get_attribute('data-name'))

    print(nome_documento)

    



    
    sleep(20)
    
    readpdf = PdfReader("C:\\Users\\rafae\\Downloads\\" + nome_documento)
        

    print(len(readpdf.pages))

    for j in range(len(readpdf.pages)):

        text = readpdf.pages[j].extract_text()

        listas = text.split("\n")

        

        for i in range(4, len(listas)):

            nome_aux=""

            cursos.append
            split = (listas[i].split(" "))

            if(split[0] != 'NOME' and split[0]  != 'CONVOCAÇÃONOTACOTA' and split[0]  != 'INSCRIÇÃO'and split[0]  != 'Itajubá'):
                

                cursos.append(listas[4])
                cota_convocacao.append(split[-1])
                pontuacao.append(split[-2])
                cota_inscricao.append(split[-3])
                situacao.append(split[-4])

                print(split[0])
                for r in range(len(split)-4):
                    nome_aux = nome_aux + split[r] + " "

                nome.append(nome_aux)


    print(nome, len(nome))
    print(cursos, len(cursos))
    print(cota_convocacao, len(cota_convocacao))
    print(pontuacao, len(pontuacao))
    print(cota_inscricao, len(cota_inscricao))
    print(situacao, len(situacao)) 


workbook = xlsxwriter.Workbook('UNIFEI.xlsx')
sheet = workbook.add_worksheet()   


for i in range(len(nome)):

    sheet.write(i+1, 0, nome[i])
    sheet.write(i+1, 1, cursos[i])
    sheet.write(i+1, 2, cota_convocacao[i])
    sheet.write(i+1, 3, pontuacao[i])
    sheet.write(i+1, 4, cota_inscricao[i])
    sheet.write(i+1, 5, situacao[i]) 
    

workbook.close()