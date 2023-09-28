# pip install selenium
# pip install webdriver_manager
# pip install pywin32
from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service
from time import sleep
from bs4 import BeautifulSoup
import win32com.client as win32

bim = int(input('Digite o bimestre que deseja ver as notas: '))

while bim > 5 or bim < 1:
    print('Digite um bimestre qvue exista!')
    int(input('Digite o bimestre que deseja ver as notas: '))

email = input('Digite o email que você que que envie as notas: ')

servico = Service(ChromeDriverManager().install())
navegador = webdriver.Chrome(service=servico)

# acessar o site
navegador.get("https://suap.ifsp.edu.br/accounts/login/?next=/")
sleep(1)

# logar em uma conta
navegador.find_element('xpath', '//*[@id="id_username"]').send_keys('CV3015602')
sleep(2)
navegador.find_element('xpath', '//*[@id="id_password"]').send_keys('senha')
sleep(2)
navegador.find_element('xpath', '/html/body/div[1]/main/div[1]/form/div[5]/input').click()
sleep(3)

# acessar o boletim
navegador.get('https://suap.ifsp.edu.br/edu/aluno/CV3015602/?tab=boletim')
sleep(5)

# acessa o detalhar por matéria
# Admininstração de Banco de Dados
navegador.get('https://suap.ifsp.edu.br/edu/detalhar_matricula_diario_boletim/205255/6628139/?_popup=1')
sleep(2)
adb = navegador.find_element('xpath', '//*[@id="content"]/div[4]/div/table[' + str(bim) + ']')
htmlContent = adb.get_attribute('outerHTML')
adb = BeautifulSoup(htmlContent, 'html.parser')
sleep(2)


# Biologia
navegador.get('https://suap.ifsp.edu.br/edu/detalhar_matricula_diario_boletim/205255/6627953/')
bio = navegador.find_element('xpath', '//*[@id="content"]/div[4]/div/table[' + str(bim) + ']')
htmlContent = bio.get_attribute('outerHTML')
bio = BeautifulSoup(htmlContent, 'html.parser')
sleep(2)

# Filosofia
navegador.get('https://suap.ifsp.edu.br/edu/detalhar_matricula_diario_boletim/205255/6628325/?_popup=1')
filo = navegador.find_element('xpath', '//*[@id="content"]/div[4]/div/table[' + str(bim) + ']')
htmlContent = filo.get_attribute('outerHTML')
filo = BeautifulSoup(htmlContent, 'html.parser')
sleep(2)

# Física
navegador.get('https://suap.ifsp.edu.br/edu/detalhar_matricula_diario_boletim/205255/6627984/')
fis = navegador.find_element('xpath', '//*[@id="content"]/div[4]/div/table[' + str(bim) + ']')
htmlContent = fis.get_attribute('outerHTML')
fis = BeautifulSoup(htmlContent, 'html.parser')
sleep(2)

# Geografia
navegador.get('https://suap.ifsp.edu.br/edu/detalhar_matricula_diario_boletim/205255/6628077/?_popup=1')
geo = navegador.find_element('xpath', '//*[@id="content"]/div[4]/div/table[' + str(bim) + ']')
htmlContent = geo.get_attribute('outerHTML')
geo = BeautifulSoup(htmlContent, 'html.parser')
sleep(2)

# História
navegador.get('https://suap.ifsp.edu.br/edu/detalhar_matricula_diario_boletim/205255/6628046/?_popup=1')
hist = navegador.find_element('xpath', '//*[@id="content"]/div[4]/div/table[' + str(bim) + ']')
htmlContent = hist.get_attribute('outerHTML')
hist = BeautifulSoup(htmlContent, 'html.parser')
sleep(2)

# Inglês
navegador.get('https://suap.ifsp.edu.br/edu/detalhar_matricula_diario_boletim/205255/6628108/?_popup=1')
ing = navegador.find_element('xpath', '//*[@id="content"]/div[4]/div/table[' + str(bim) + ']')
htmlContent = ing.get_attribute('outerHTML')
ing = BeautifulSoup(htmlContent, 'html.parser')
sleep(2)

# Linguagem de Programação Visual
navegador.get('https://suap.ifsp.edu.br/edu/detalhar_matricula_diario_boletim/205255/6628263/?_popup=1')
lpv = navegador.find_element('xpath', '//*[@id="content"]/div[4]/div/table[' + str(bim) + ']')
htmlContent = lpv.get_attribute('outerHTML')
lpv = BeautifulSoup(htmlContent, 'html.parser')
sleep(2)

# Lingua Portuguesa e Literatura
navegador.get('https://suap.ifsp.edu.br/edu/detalhar_matricula_diario_boletim/205255/6628170/?_popup=1')
pt = navegador.find_element('xpath', '//*[@id="content"]/div[4]/div/table[' + str(bim) + ']')
htmlContent = pt.get_attribute('outerHTML')
pt = BeautifulSoup(htmlContent, 'html.parser')
sleep(2)

# Matemática
navegador.get('https://suap.ifsp.edu.br/edu/detalhar_matricula_diario_boletim/205255/6628201/')
mat = navegador.find_element('xpath', '//*[@id="content"]/div[4]/div/table[' + str(bim) + ']')
htmlContent = mat.get_attribute('outerHTML')
mat = BeautifulSoup(htmlContent, 'html.parser')
sleep(2)

# Programação de Dispositivos Moveis
navegador.get('https://suap.ifsp.edu.br/edu/detalhar_matricula_diario_boletim/205255/6628232/')
pdm = navegador.find_element('xpath', '//*[@id="content"]/div[4]/div/table[' + str(bim) + ']')
htmlContent = pdm.get_attribute('outerHTML')
pdm = BeautifulSoup(htmlContent, 'html.parser')
sleep(2)

# Projeto Integrador
navegador.get('https://suap.ifsp.edu.br/edu/detalhar_matricula_diario_boletim/205255/6628294/?_popup=1')
tcc = navegador.find_element('xpath', '//*[@id="content"]/div[4]/div/table[' + str(bim) + ']')
htmlContent = tcc.get_attribute('outerHTML')
tcc = BeautifulSoup(htmlContent, 'html.parser')
sleep(2)

# Química
navegador.get('https://suap.ifsp.edu.br/edu/detalhar_matricula_diario_boletim/205255/6628015/?_popup=1')
qui = navegador.find_element('xpath', '//*[@id="content"]/div[4]/div/table[' + str(bim) + ']')
htmlContent = qui.get_attribute('outerHTML')
qui = BeautifulSoup(htmlContent, 'html.parser')
sleep(2)

# Sociologia
navegador.get('https://suap.ifsp.edu.br/edu/detalhar_matricula_diario_boletim/205255/6628356/')
soc = navegador.find_element('xpath', '//*[@id="content"]/div[4]/div/table[' + str(bim) + ']')
htmlContent = soc.get_attribute('outerHTML')
soc = BeautifulSoup(htmlContent, 'html.parser')
sleep(2)

# enviar esse relatório por email
outlook = win32.Dispatch('outlook.application')
mail = outlook.CreateItem(0)
mail.To = email
mail.Subject = 'Relatório de Notas do 3º Bimestre'
mail.HTMLBody = f'''
<p>Prezados,</p>

<p>Segue o Relaório de Notas do 3º Bimestre.</p>

<p>Admininstração de Banco de Dados:</p>
{adb}

<p>Biologia:</p>
{bio}

<p>Filosofia:</p>
{filo}

<p>Física:</p>
{fis}

<p>Geografia:</p>
{geo}

<p>História:</p>
{hist}

<p>Inglês:</p>
{ing}

<p>Linguagem de Programação Visual:</p>
{lpv}

<p>Lingua Portuguesa e Literatura:</p>
{pt}

<p>Matemática:</p>
{mat}

<p>Programação de Dispositivos Moveis:</p>
{pdm}

<p>Projeto Integrador:</p>
{tcc}

<p>Química:</p>
{qui}

<p>Sociologia:</p>
{soc}

<p>Qualquer coisa, estou a disposição.</p>

<p>att.,</p>
<p>Jesbica</p>
'''

try:
    mail.Send()
    print('email enviado')
except:
    print('erro')
