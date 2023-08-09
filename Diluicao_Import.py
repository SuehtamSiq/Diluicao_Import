from time import sleep 
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service as ChromeService
import os
from pathlib import Path
import shutil


def Import_Diluição():

    # Definindo usuário
    user = os.getlogin()


    # Browser
    navegador = webdriver.Chrome(service=ChromeService('chromedriver.exe'))

    # Indicando essa janela como original
    original_window = navegador.current_window_handle

    # Abrindo Reports POMS
    navegador.get('http://brsprpt001.global.baxter.com/reportsystem/Reporte.aspx')
    wait = WebDriverWait(navegador, 20)

    # Clica em Consultar Adição de Matéria Prima
    navegador.find_element('xpath', '//*[@id="TreeView1t61"]').click()

    # View Report
    navegador.find_element('xpath', '//*[@id="ContentPlaceHolder1_btnImprimirLogo"]').click()

    # Aguarda a nova aba ou janela aparecer e muda para ela
    wait.until(EC.number_of_windows_to_be(2))

    for window_handle in navegador.window_handles:
        if window_handle != original_window:
            navegador.switch_to.window(window_handle)
            break

    # Campo DATA INICIAL - Preenchendo Campos 
    wait.until(EC.element_to_be_clickable((By.XPATH, '/html/body/table/tbody/tr/td/table/tbody/tr[2]/td/div/form/div/div/fieldset[1]/table/tbody/tr/td/table/tbody/tr[2]/td/table/tbody/tr/td[1]/table/tbody/tr/td/table/tbody/tr[2]/td[1]/input'))).send_keys('08/')
    wait.until(EC.element_to_be_clickable((By.XPATH, '/html/body/table/tbody/tr/td/table/tbody/tr[2]/td/div/form/div/div/fieldset[1]/table/tbody/tr/td/table/tbody/tr[2]/td/table/tbody/tr/td[1]/table/tbody/tr/td/table/tbody/tr[2]/td[1]/input'))).send_keys('08/')
    wait.until(EC.element_to_be_clickable((By.XPATH, '/html/body/table/tbody/tr/td/table/tbody/tr[2]/td/div/form/div/div/fieldset[1]/table/tbody/tr/td/table/tbody/tr[2]/td/table/tbody/tr/td[1]/table/tbody/tr/td/table/tbody/tr[2]/td[1]/input'))).send_keys('2023')

    # Campo DATA FINAL
    navegador.find_element('xpath', '//*[@id="CrystalReportViewer1_p1DiscreteValue"]').click()

    # Campo DATA FINAL - Preenchendo Campos
    wait.until(EC.element_to_be_clickable((By.XPATH, '/html/body/table/tbody/tr/td/table/tbody/tr[2]/td/div/form/div/div/fieldset[2]/table/tbody/tr/td/table/tbody/tr[2]/td/table/tbody/tr/td[1]/table/tbody/tr/td/table/tbody/tr[2]/td[1]/input'))).send_keys('08/')
    wait.until(EC.element_to_be_clickable((By.XPATH, '/html/body/table/tbody/tr/td/table/tbody/tr[2]/td/div/form/div/div/fieldset[2]/table/tbody/tr/td/table/tbody/tr[2]/td/table/tbody/tr/td[1]/table/tbody/tr/td/table/tbody/tr[2]/td[1]/input'))).send_keys('08/')
    wait.until(EC.element_to_be_clickable((By.XPATH, '/html/body/table/tbody/tr/td/table/tbody/tr[2]/td/div/form/div/div/fieldset[2]/table/tbody/tr/td/table/tbody/tr[2]/td/table/tbody/tr/td[1]/table/tbody/tr/td/table/tbody/tr[2]/td[1]/input'))).send_keys('2023')

    # Botão 'OK'
    navegador.find_element('xpath', '//*[@id="CrystalReportViewer1_submitButton"]').click()
    sleep(3)

    # Botão de extrair
    wait.until(EC.element_to_be_clickable((By.XPATH, '/html/body/form/div[3]/div/div[1]/table/tbody/tr/td[1]/table/tbody/tr/td[1]/table/tbody/tr/td[2]/table/tbody/tr/td[1]/table/tbody/tr/td[3]/table/tbody/tr/td/div/img'))).click()
    sleep(1)

    # Abre lista de formatos
    wait.until(EC.element_to_be_clickable((By.XPATH, '/html/body/table[3]/tbody/tr/td/table/tbody/tr[2]/td/table/tbody/tr[1]/td/div/table/tbody/tr/td/div/table/tbody/tr/td[1]/table/tbody/tr/td[2]/div'))).click()
    sleep(1)

    # Seleciona o formato
    wait.until(EC.element_to_be_clickable((By.XPATH, '/html/body/table[4]/tbody/tr/td/table/tbody/tr[5]/td[2]/span'))).click()
    sleep(1)

    # Antes do download, verifica se ja existe um arquivo com o memso nome, e caso já exista, exclui o antigo.
    nome_arq = Path(r'C:\Users\{}\Downloads\CrystalReportViewer1.xlsx'.format(user))

    try:
        nome_arq.unlink()
    except OSError as e:
        print(f'\Error:{ e.strerror}')

    # Confirma a Exportação do Download
    wait.until(EC.element_to_be_clickable((By.XPATH, '/html/body/table[3]/tbody/tr/td/table/tbody/tr[2]/td/table/tbody/tr[2]/td/table/tbody/tr/td/table/tbody/tr/td[2]/nobr/a'))).click()
    sleep(5)

    #Trocando o nome do arquivo
    nome_atual = os.path.join(r"C:\Users\{}\Downloads".format(user), "CrystalReportViewer1.xlsx")
    nome_novo = os.path.join(r"C:\Users\{}\Downloads".format(user), "Diluição - 08.08.23.xlsx")

    nome_consolidado = shutil.move(nome_atual, nome_novo)

    print("Arquivo alterado para: ",nome_consolidado)

    navegador.quit()


def Tratamento_Dados():

    plan = pd.read_excel(r"C:\Users\{}\Downloads\Diluição - 08.08.23.xlsx".format(user))

    # Filtar linhas onde 'Lote Produto' não contém 'PR'
    plan = plan[plan['Lote Produto'].str.contains('PR', na=False) 
                & plan['Status'].str.contains('IN PROCESS', case=False, na=False)]

    # Resetar índices após a filtragem
    plan.reset_index(drop=True, inplace=True)

    plan.to_excel(r"C:\Users\{}\Downloads\Diluição - 08.08.23.xlsx".format(user), index=False)
