import shutil
import warnings
import pyautogui
import time
import pyperclip
import pandas as pd
import os
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service 
from pathlib import Path
from selenium.webdriver.chrome.options import Options
from datetime import date, timedelta


# Definindo usuário
user = os.getlogin()

# Definindo datas de produção
data = date.today()
indice_da_semana = data.weekday()

if indice_da_semana == 0:
    data_nova = data - timedelta(days=3)
else:
    data_nova = data - timedelta(days=1)

ano = data_nova.year
mes = data_nova.month
dia = data.day
dia_novo = data_nova.day

# Configurando o WebDriver do Chrome

service = Service(ChromeDriverManager().install())
navegador = webdriver.Chrome(service=service)

download_path = rf"C:\Users\{user}\Downloads"

navegador.execute_cdp_cmd(
    "Page.setDownloadBehavior",
    {
        "behavior": "allow",
        "downloadPath": download_path
    }
)
      
def Import_Diluição_Crystal():

    # Indicando essa janela como original
    original_window = navegador.current_window_handle

    # Abrindo Reports POMS
    navegador.get('http://brsprpt001.global.baxter.com/reportsystem/Reporte.aspx')
    wait = WebDriverWait(navegador, 20)

    # Clica em Consultar Adição de Matéria Prima
    wait.until(EC.element_to_be_clickable((By.XPATH, '/html/body/form/table/tbody/tr[3]/td[1]/div/div/div[10]/table[1]/tbody/tr[2]/td[4]/a'))).click()

    # View Report
    wait.until(EC.element_to_be_clickable((By.XPATH, '/html/body/form/table/tbody/tr[3]/td[2]/table/tbody/tr/td[2]/input[1]'))).click()

    # Aguarda a nova aba ou janela aparecer e muda para ela
    wait.until(EC.number_of_windows_to_be(2))

    for window_handle in navegador.window_handles:
        if window_handle != original_window:
            navegador.switch_to.window(window_handle)
            break

    # Campo DATA INICIAL - Preenchendo Campos 
    wait.until(EC.element_to_be_clickable((By.XPATH, '/html/body/table/tbody/tr/td/table/tbody/tr[2]/td/div/form/div/div/fieldset[1]/table/tbody/tr/td/table/tbody/tr[2]/td/table/tbody/tr/td[1]/table/tbody/tr/td/table/tbody/tr[2]/td[1]/input'))).send_keys(mes)
    wait.until(EC.element_to_be_clickable((By.XPATH, '/html/body/table/tbody/tr/td/table/tbody/tr[2]/td/div/form/div/div/fieldset[1]/table/tbody/tr/td/table/tbody/tr[2]/td/table/tbody/tr/td[1]/table/tbody/tr/td/table/tbody/tr[2]/td[1]/input'))).send_keys('/')
    wait.until(EC.element_to_be_clickable((By.XPATH, '/html/body/table/tbody/tr/td/table/tbody/tr[2]/td/div/form/div/div/fieldset[1]/table/tbody/tr/td/table/tbody/tr[2]/td/table/tbody/tr/td[1]/table/tbody/tr/td/table/tbody/tr[2]/td[1]/input'))).send_keys(dia)
    wait.until(EC.element_to_be_clickable((By.XPATH, '/html/body/table/tbody/tr/td/table/tbody/tr[2]/td/div/form/div/div/fieldset[1]/table/tbody/tr/td/table/tbody/tr[2]/td/table/tbody/tr/td[1]/table/tbody/tr/td/table/tbody/tr[2]/td[1]/input'))).send_keys('/')
    wait.until(EC.element_to_be_clickable((By.XPATH, '/html/body/table/tbody/tr/td/table/tbody/tr[2]/td/div/form/div/div/fieldset[1]/table/tbody/tr/td/table/tbody/tr[2]/td/table/tbody/tr/td[1]/table/tbody/tr/td/table/tbody/tr[2]/td[1]/input'))).send_keys(ano)

    # Campo DATA FINAL - Preenchendo Campos
    wait.until(EC.element_to_be_clickable((By.XPATH, '/html/body/table/tbody/tr/td/table/tbody/tr[2]/td/div/form/div/div/fieldset[2]/table/tbody/tr/td/table/tbody/tr[2]/td/table/tbody/tr/td[1]/table/tbody/tr/td/table/tbody/tr[2]/td[1]/input'))).send_keys(mes)
    wait.until(EC.element_to_be_clickable((By.XPATH, '/html/body/table/tbody/tr/td/table/tbody/tr[2]/td/div/form/div/div/fieldset[2]/table/tbody/tr/td/table/tbody/tr[2]/td/table/tbody/tr/td[1]/table/tbody/tr/td/table/tbody/tr[2]/td[1]/input'))).send_keys('/')
    wait.until(EC.element_to_be_clickable((By.XPATH, '/html/body/table/tbody/tr/td/table/tbody/tr[2]/td/div/form/div/div/fieldset[2]/table/tbody/tr/td/table/tbody/tr[2]/td/table/tbody/tr/td[1]/table/tbody/tr/td/table/tbody/tr[2]/td[1]/input'))).send_keys(dia_novo)
    wait.until(EC.element_to_be_clickable((By.XPATH, '/html/body/table/tbody/tr/td/table/tbody/tr[2]/td/div/form/div/div/fieldset[2]/table/tbody/tr/td/table/tbody/tr[2]/td/table/tbody/tr/td[1]/table/tbody/tr/td/table/tbody/tr[2]/td[1]/input'))).send_keys('/')
    wait.until(EC.element_to_be_clickable((By.XPATH, '/html/body/table/tbody/tr/td/table/tbody/tr[2]/td/div/form/div/div/fieldset[2]/table/tbody/tr/td/table/tbody/tr[2]/td/table/tbody/tr/td[1]/table/tbody/tr/td/table/tbody/tr[2]/td[1]/input'))).send_keys(ano)

    # Botão 'OK'
    navegador.find_element('xpath', '//*[@id="CrystalReportViewer1_submitButton"]').click()
    time.sleep(3)

    # Botão de extrair
    wait.until(EC.element_to_be_clickable((By.XPATH, '/html/body/form/div[3]/div/div[1]/table/tbody/tr/td[1]/table/tbody/tr/td[1]/table/tbody/tr/td[2]/table/tbody/tr/td[1]/table/tbody/tr/td[3]/table/tbody/tr/td/div/img'))).click()
    time.sleep(1)

    # Abre lista de formatos
    wait.until(EC.element_to_be_clickable((By.XPATH, '/html/body/table[3]/tbody/tr/td/table/tbody/tr[2]/td/table/tbody/tr[1]/td/div/table/tbody/tr/td/div/table/tbody/tr/td[1]/table/tbody/tr/td[2]/div'))).click()
    time.sleep(1)

    # Seleciona o formato
    wait.until(EC.element_to_be_clickable((By.XPATH, '/html/body/table[4]/tbody/tr/td/table/tbody/tr[5]/td[2]/span'))).click()
    time.sleep(1)

    # Antes do download, verifica se ja existe um arquivo com o memso nome, e caso já exista, exclui o antigo.
    nome_arq = Path(r'C:\Users\{}\Downloads\CrystalReportViewer1.xlsx'.format(user))

    if os.path.exists(nome_arq):

        os.remove(r'C:\Users\{}\Downloads\CrystalReportViewer1.xlsx'.format(user))

    # Confirma a Exportação do Download
    wait.until(EC.element_to_be_clickable((By.XPATH, '/html/body/table[3]/tbody/tr/td/table/tbody/tr[2]/td/table/tbody/tr[2]/td/table/tbody/tr/td/table/tbody/tr/td[2]/nobr/a'))).click()
    time.sleep(5)

    # Trocando o nome do arquivo
    nome_atual = os.path.join(r'C:\Users\{}\Downloads'.format(user), 'CrystalReportViewer1.xlsx')
    nom_nov = r'C:\Users\{}\Downloads\Report Diluição Crystal.xlsx'.format(user)
    os.rename(nome_atual, nom_nov)
    
    # Movendo o arquivo
    destino_pasta = r"C:\Users\{}\OneDrive - Baxter\Apontamento\Fechamentos\2026\Carga Soluções\01 - CAPD\Report POMS - Diluição".format(user)
    destino_final = os.path.join(destino_pasta, "Report Diluição Crystal.xlsx")
    # Se o arquivo já existir no destino, remove
    if os.path.exists(destino_final):
        os.remove(destino_final)

    shutil.move(nom_nov, destino_final)

def Import_Diluição_POMSNET():

    # Abrindo Reports POMS
    navegador.execute_script("window.open('');") 
    navegador.switch_to.window(navegador.window_handles[2]) 
    navegador.get(f'https://pomsbr.aws.baxter.com/poms/DesktopDefault.aspx?ReturnUrl=%2fPOMS%2fWelcome.aspx')
    wait = WebDriverWait(navegador, 20)

    # Login POMSNET
    wait.until(EC.element_to_be_clickable((By.XPATH, '/html/body/form/div[3]/div/div[2]/div[3]/div/div[2]/div[1]/div/div[2]/input'))).send_keys("SOARESM4")
    wait.until(EC.element_to_be_clickable((By.XPATH, '/html/body/form/div[3]/div/div[2]/div[3]/div/div[2]/div[2]/div/div[2]/input'))).send_keys("Mss1003*BfG!zXw")
    navegador.switch_to.active_element.send_keys(Keys.ENTER)

    # Menu Hamburguer
    wait.until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[2]/div[1]/div/ul/li/a'))).click()

    # Relatórios
    wait.until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[2]/div[1]/div/div[4]/div/ul[108]/li[3]/div/a/img'))).click()

    # Diluição
    wait.until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[2]/div[1]/div/div[4]/div/ul[98]/li[11]/div/a/img'))).click()

    # Custos
    wait.until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[2]/div[1]/div/div[4]/div/ul[90]/li[2]/div/a/img'))).click()

    # Adição de Matéria Prima
    wait.until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[2]/div[1]/div/div[4]/div/ul[77]/li[2]/div/a/img'))).click()

    # Campo DATA INICIAL - Preenchendo Campos 
    wait.until(EC.element_to_be_clickable((By.XPATH, '/html/body/form/div[11]/div/div/div[2]/div[2]/div[1]/div/div/span/input'))).send_keys(dia)
    wait.until(EC.element_to_be_clickable((By.XPATH, '/html/body/form/div[11]/div/div/div[2]/div[2]/div[1]/div/div/span/input'))).send_keys('/')
    wait.until(EC.element_to_be_clickable((By.XPATH, '/html/body/form/div[11]/div/div/div[2]/div[2]/div[1]/div/div/span/input'))).send_keys(mes)
    wait.until(EC.element_to_be_clickable((By.XPATH, '/html/body/form/div[11]/div/div/div[2]/div[2]/div[1]/div/div/span/input'))).send_keys('/')
    wait.until(EC.element_to_be_clickable((By.XPATH, '/html/body/form/div[11]/div/div/div[2]/div[2]/div[1]/div/div/span/input'))).send_keys(ano)
    wait.until(EC.element_to_be_clickable((By.XPATH, '/html/body/form/div[11]/div/div/div[2]/div[2]/div[1]/div/div/span/input'))).send_keys(" ")
    wait.until(EC.element_to_be_clickable((By.XPATH, '/html/body/form/div[11]/div/div/div[2]/div[2]/div[1]/div/div/span/input'))).send_keys("00:00:00")

    # Campo DATA FINAL - Preenchendo Campos
    wait.until(EC.element_to_be_clickable((By.XPATH, '/html/body/form/div[11]/div/div/div[2]/div[2]/div[2]/div/div/span/input'))).send_keys(dia_novo)
    wait.until(EC.element_to_be_clickable((By.XPATH, '/html/body/form/div[11]/div/div/div[2]/div[2]/div[2]/div/div/span/input'))).send_keys('/')
    wait.until(EC.element_to_be_clickable((By.XPATH, '/html/body/form/div[11]/div/div/div[2]/div[2]/div[2]/div/div/span/input'))).send_keys(mes)
    wait.until(EC.element_to_be_clickable((By.XPATH, '/html/body/form/div[11]/div/div/div[2]/div[2]/div[2]/div/div/span/input'))).send_keys('/')
    wait.until(EC.element_to_be_clickable((By.XPATH, '/html/body/form/div[11]/div/div/div[2]/div[2]/div[2]/div/div/span/input'))).send_keys(ano)
    wait.until(EC.element_to_be_clickable((By.XPATH, '/html/body/form/div[11]/div/div/div[2]/div[2]/div[2]/div/div/span/input'))).send_keys(" ")
    wait.until(EC.element_to_be_clickable((By.XPATH, '/html/body/form/div[11]/div/div/div[2]/div[2]/div[2]/div/div/span/input'))).send_keys("00:00:00")

    # Botão 'OK'
    wait.until(EC.element_to_be_clickable((By.XPATH, '/html/body/form/div[11]/div/div/div[2]/div[1]/div[2]/button[1]'))).click()
    time.sleep(5)

    # Botão de exportar
    wait.until(EC.element_to_be_clickable((By.XPATH, '/html/body/form/div[10]/div/div[2]/div[1]/ul[1]/li[11]/a'))).click()

    # Seleciona o formato Excel
    wait.until(EC.element_to_be_clickable((By.XPATH, '/html/body/form/div[10]/div/div[2]/div[1]/ul[1]/li[11]/div/div/div/ul/li[4]/a'))).click()
    time.sleep(5)

    download_path2 = rf"C:\Users\{user}\Downloads"
    padrao_nome = "RelatóriodeAdiçãodeMatériaPrima.xlsx"

    arquivos = [
        os.path.join(download_path2, f)
        for f in os.listdir(download_path2)
        if f.startswith("RelatóriodeAdiçãodeMatériaPrima") and f.endswith(".xlsx")
    ]

    if not arquivos:
        raise Exception("Arquivo de relatório POMSNet não encontrado na pasta de downloads.")
    
    arquivo_mais_recente = max(arquivos, key=os.path.getctime)

    novo_nome2 = os.path.join(download_path2, "Report Diluição POMSNET.xlsx")

    if os.path.exists(novo_nome2):
        os.remove(novo_nome2)

    os.rename(arquivo_mais_recente, novo_nome2)

    destino_pasta2 = r"C:\Users\{}\OneDrive - Baxter\Apontamento\Fechamentos\2026\Carga Soluções\01 - CAPD\Report POMS - Diluição".format(user)
    destino_final2 = os.path.join(destino_pasta2, "Report Diluição POMSNET.xlsx")

    if os.path.exists(destino_final2):
        os.remove(destino_final2)

    os.rename(novo_nome2, destino_final2)

Import_Diluição_Crystal()
Import_Diluição_POMSNET()
# By Matheus Siqueira
