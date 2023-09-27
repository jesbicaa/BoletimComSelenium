from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service
from time import sleep
from bs4 import BeautifulSoup
import win32com.client as win32

servico = Service(ChromeDriverManager().install())
navegador = webdriver.Chrome(service=servico)

# acessar o site
navegador.get("https://suap.ifsp.edu.br/accounts/login/?next=/")
sleep(1)

# logar em uma conta
navegador.find_element('xpath', '//*[@id="id_username"]').send_keys('seu_prontuario')
sleep(2)
navegador.find_element('xpath', '//*[@id="id_password"]').send_keys('sua_senha')
sleep(2)
navegador.find_element('xpath', '/html/body/div[1]/main/div[1]/form/div[5]/input').click()
sleep(3)

# acessar o boletim
navegador.get('https://suap.ifsp.edu.br/edu/aluno/seu_prontuario/?tab=boletim')
sleep(3)

# acessa o detalhar por matéria
# Admininstração de Banco de Dados
navegador.get('https://suap.ifsp.edu.br/edu/detalhar_matricula_diario_boletim/205255/6628139/?_popup=1')
sleep(2)
adb3bi = navegador.find_element('xpath', '//*[@id="content"]/div[4]/div/table[3]')
htmlContent = adb3bi.get_attribute('outerHTML')
adb3bi = BeautifulSoup(htmlContent, 'html.parser')
sleep(2)


# Biologia
navegador.get('https://suap.ifsp.edu.br/edu/detalhar_matricula_diario_boletim/205255/6627953/')
bio3bi = navegador.find_element('xpath', '//*[@id="content"]/div[4]/div/table[3]')
htmlContent = bio3bi.get_attribute('outerHTML')
bio3bi = BeautifulSoup(htmlContent, 'html.parser')
sleep(2)

# Filosofia
navegador.get('https://suap.ifsp.edu.br/edu/detalhar_matricula_diario_boletim/205255/6628325/?_popup=1')
filo3bi = navegador.find_element('xpath', '//*[@id="content"]/div[4]/div/table[3]')
htmlContent = filo3bi.get_attribute('outerHTML')
filo3bi = BeautifulSoup(htmlContent, 'html.parser')
sleep(2)

# Física
navegador.get('https://suap.ifsp.edu.br/edu/detalhar_matricula_diario_boletim/205255/6627984/')
fis3bi = navegador.find_element('xpath', '//*[@id="content"]/div[4]/div/table[3]')
htmlContent = fis3bi.get_attribute('outerHTML')
fis3bi = BeautifulSoup(htmlContent, 'html.parser')
sleep(2)

# Geografia
navegador.get('https://suap.ifsp.edu.br/edu/detalhar_matricula_diario_boletim/205255/6628077/?_popup=1')
geo3bi = navegador.find_element('xpath', '//*[@id="content"]/div[4]/div/table[3]')
htmlContent = geo3bi.get_attribute('outerHTML')
geo3bi = BeautifulSoup(htmlContent, 'html.parser')
sleep(2)

# História
navegador.get('https://suap.ifsp.edu.br/edu/detalhar_matricula_diario_boletim/205255/6628046/?_popup=1')
hist3bi = navegador.find_element('xpath', '//*[@id="content"]/div[4]/div/table[3]')
htmlContent = hist3bi.get_attribute('outerHTML')
hist3bi = BeautifulSoup(htmlContent, 'html.parser')
sleep(2)

# Inglês
navegador.get('https://suap.ifsp.edu.br/edu/detalhar_matricula_diario_boletim/205255/6628108/?_popup=1')
ing3bi = navegador.find_element('xpath', '//*[@id="content"]/div[4]/div/table[3]')
htmlContent = ing3bi.get_attribute('outerHTML')
ing3bi = BeautifulSoup(htmlContent, 'html.parser')
sleep(2)

# Linguagem de Programação Visual
navegador.get('https://suap.ifsp.edu.br/edu/detalhar_matricula_diario_boletim/205255/6628263/?_popup=1')
lpv3bi = navegador.find_element('xpath', '//*[@id="content"]/div[4]/div/table[3]')
htmlContent = lpv3bi.get_attribute('outerHTML')
lpv3bi = BeautifulSoup(htmlContent, 'html.parser')
sleep(2)

# Lingua Portuguesa e Literatura
navegador.get('https://suap.ifsp.edu.br/edu/detalhar_matricula_diario_boletim/205255/6628170/?_popup=1')
pt3bi = navegador.find_element('xpath', '//*[@id="content"]/div[4]/div/table[3]')
htmlContent = pt3bi.get_attribute('outerHTML')
pt3bi = BeautifulSoup(htmlContent, 'html.parser')
sleep(2)

# Matemática
navegador.get('https://suap.ifsp.edu.br/edu/detalhar_matricula_diario_boletim/205255/6628201/')
mat3bi = navegador.find_element('xpath', '//*[@id="content"]/div[4]/div/table[3]')
htmlContent = mat3bi.get_attribute('outerHTML')
mat3bi = BeautifulSoup(htmlContent, 'html.parser')
sleep(2)

# Programação de Dispositivos Moveis
navegador.get('https://suap.ifsp.edu.br/edu/detalhar_matricula_diario_boletim/205255/6628232/')
pdm3bi = navegador.find_element('xpath', '//*[@id="content"]/div[4]/div/table[3]')
htmlContent = pdm3bi.get_attribute('outerHTML')
pdm3bi = BeautifulSoup(htmlContent, 'html.parser')
sleep(2)

# Projeto Integrador
navegador.get('https://suap.ifsp.edu.br/edu/detalhar_matricula_diario_boletim/205255/6628294/?_popup=1')
tcc3bi = navegador.find_element('xpath', '//*[@id="content"]/div[4]/div/table[3]')
htmlContent = tcc3bi.get_attribute('outerHTML')
tcc3bi = BeautifulSoup(htmlContent, 'html.parser')
sleep(2)

# Química
navegador.get('https://suap.ifsp.edu.br/edu/detalhar_matricula_diario_boletim/205255/6628015/?_popup=1')
qui3bi = navegador.find_element('xpath', '//*[@id="content"]/div[4]/div/table[3]')
htmlContent = qui3bi.get_attribute('outerHTML')
qui3bi = BeautifulSoup(htmlContent, 'html.parser')
sleep(2)

# Sociologia
navegador.get('https://suap.ifsp.edu.br/edu/detalhar_matricula_diario_boletim/205255/6628356/')
soc3bi = navegador.find_element('xpath', '//*[@id="content"]/div[4]/div/table[3]')
htmlContent = soc3bi.get_attribute('outerHTML')
soc3bi = BeautifulSoup(htmlContent, 'html.parser')
sleep(2)

# enviar esse relatório por email
outlook = win32.Dispatch('outlook.application')
mail = outlook.CreateItem(0)
mail.To = 'terceirotcc2023@gmail.com'
mail.Subject = 'Relatório de Notas do 3º Bimestre'
mail.HTMLBody = f'''
<p>Prezados,</p>

<p>Segue o Relaório de Notas do 3º Bimestre.</p>

<p>Admininstração de Banco de Dados:</p>
{adb3bi}

<p>Biologia:</p>
{bio3bi}

<p>Filosofia:</p>
{filo3bi}

<p>Física:</p>
{fis3bi}

<p>Geografia:</p>
{geo3bi}

<p>História:</p>
{hist3bi}

<p>Inglês:</p>
{ing3bi}

<p>Linguagem de Programação Visual:</p>
{lpv3bi}

<p>Lingua Portuguesa e Literatura:</p>
{pt3bi}

<p>Matemática:</p>
{mat3bi}

<p>Programação de Dispositivos Moveis:</p>
{pdm3bi}

<p>Projeto Integrador:</p>
{tcc3bi}

<p>Química:</p>
{qui3bi}

<p>Sociologia:</p>
{soc3bi}

<p>Qualquer coisa, estou a disposição.</p>

<p>att.,</p>
<p>Jesbica</p>
'''

try:
    mail.Send()
    print('email enviado')
except:
    print('erro')
