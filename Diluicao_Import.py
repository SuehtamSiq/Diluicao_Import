from time import sleep 
import pandas as pd
from selenium import webdriver
import os
from pathlib import Path
import shutil
import datetime
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
import openpyxl
import warnings



# Definindo usuário
user = os.getlogin()

# Definindo a Data de produção
data = datetime.date.today()

indice_da_semana = data.weekday()

ano = data.year
mes = data.month
dia = data.day

if indice_da_semana == 0:
    dia_novo = dia - 3
else:
    dia_novo = dia - 1

def Import_Diluição():
    
    # Browser
    navegador = webdriver.Chrome()

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
    wait.until(EC.element_to_be_clickable((By.XPATH, '/html/body/table/tbody/tr/td/table/tbody/tr[2]/td/div/form/div/div/fieldset[1]/table/tbody/tr/td/table/tbody/tr[2]/td/table/tbody/tr/td[1]/table/tbody/tr/td/table/tbody/tr[2]/td[1]/input'))).send_keys(mes)
    wait.until(EC.element_to_be_clickable((By.XPATH, '/html/body/table/tbody/tr/td/table/tbody/tr[2]/td/div/form/div/div/fieldset[1]/table/tbody/tr/td/table/tbody/tr[2]/td/table/tbody/tr/td[1]/table/tbody/tr/td/table/tbody/tr[2]/td[1]/input'))).send_keys('/')
    wait.until(EC.element_to_be_clickable((By.XPATH, '/html/body/table/tbody/tr/td/table/tbody/tr[2]/td/div/form/div/div/fieldset[1]/table/tbody/tr/td/table/tbody/tr[2]/td/table/tbody/tr/td[1]/table/tbody/tr/td/table/tbody/tr[2]/td[1]/input'))).send_keys(dia_novo)
    wait.until(EC.element_to_be_clickable((By.XPATH, '/html/body/table/tbody/tr/td/table/tbody/tr[2]/td/div/form/div/div/fieldset[1]/table/tbody/tr/td/table/tbody/tr[2]/td/table/tbody/tr/td[1]/table/tbody/tr/td/table/tbody/tr[2]/td[1]/input'))).send_keys('/')
    wait.until(EC.element_to_be_clickable((By.XPATH, '/html/body/table/tbody/tr/td/table/tbody/tr[2]/td/div/form/div/div/fieldset[1]/table/tbody/tr/td/table/tbody/tr[2]/td/table/tbody/tr/td[1]/table/tbody/tr/td/table/tbody/tr[2]/td[1]/input'))).send_keys(ano)
    
    # Campo DATA FINAL - Preenchendo Campos
    if indice_da_semana == 0:
        wait.until(EC.element_to_be_clickable((By.XPATH, '/html/body/table/tbody/tr/td/table/tbody/tr[2]/td/div/form/div/div/fieldset[2]/table/tbody/tr/td/table/tbody/tr[2]/td/table/tbody/tr/td[1]/table/tbody/tr/td/table/tbody/tr[2]/td[1]/input'))).send_keys(mes)
        wait.until(EC.element_to_be_clickable((By.XPATH, '/html/body/table/tbody/tr/td/table/tbody/tr[2]/td/div/form/div/div/fieldset[2]/table/tbody/tr/td/table/tbody/tr[2]/td/table/tbody/tr/td[1]/table/tbody/tr/td/table/tbody/tr[2]/td[1]/input'))).send_keys('/')
        wait.until(EC.element_to_be_clickable((By.XPATH, '/html/body/table/tbody/tr/td/table/tbody/tr[2]/td/div/form/div/div/fieldset[2]/table/tbody/tr/td/table/tbody/tr[2]/td/table/tbody/tr/td[1]/table/tbody/tr/td/table/tbody/tr[2]/td[1]/input'))).send_keys(dia)
        wait.until(EC.element_to_be_clickable((By.XPATH, '/html/body/table/tbody/tr/td/table/tbody/tr[2]/td/div/form/div/div/fieldset[2]/table/tbody/tr/td/table/tbody/tr[2]/td/table/tbody/tr/td[1]/table/tbody/tr/td/table/tbody/tr[2]/td[1]/input'))).send_keys('/')
        wait.until(EC.element_to_be_clickable((By.XPATH, '/html/body/table/tbody/tr/td/table/tbody/tr[2]/td/div/form/div/div/fieldset[2]/table/tbody/tr/td/table/tbody/tr[2]/td/table/tbody/tr/td[1]/table/tbody/tr/td/table/tbody/tr[2]/td[1]/input'))).send_keys(ano)

    else:
        wait.until(EC.element_to_be_clickable((By.XPATH, '/html/body/table/tbody/tr/td/table/tbody/tr[2]/td/div/form/div/div/fieldset[2]/table/tbody/tr/td/table/tbody/tr[2]/td/table/tbody/tr/td[1]/table/tbody/tr/td/table/tbody/tr[2]/td[1]/input'))).send_keys(mes)
        wait.until(EC.element_to_be_clickable((By.XPATH, '/html/body/table/tbody/tr/td/table/tbody/tr[2]/td/div/form/div/div/fieldset[2]/table/tbody/tr/td/table/tbody/tr[2]/td/table/tbody/tr/td[1]/table/tbody/tr/td/table/tbody/tr[2]/td[1]/input'))).send_keys('/')
        wait.until(EC.element_to_be_clickable((By.XPATH, '/html/body/table/tbody/tr/td/table/tbody/tr[2]/td/div/form/div/div/fieldset[2]/table/tbody/tr/td/table/tbody/tr[2]/td/table/tbody/tr/td[1]/table/tbody/tr/td/table/tbody/tr[2]/td[1]/input'))).send_keys(dia_novo)
        wait.until(EC.element_to_be_clickable((By.XPATH, '/html/body/table/tbody/tr/td/table/tbody/tr[2]/td/div/form/div/div/fieldset[2]/table/tbody/tr/td/table/tbody/tr[2]/td/table/tbody/tr/td[1]/table/tbody/tr/td/table/tbody/tr[2]/td[1]/input'))).send_keys('/')
        wait.until(EC.element_to_be_clickable((By.XPATH, '/html/body/table/tbody/tr/td/table/tbody/tr[2]/td/div/form/div/div/fieldset[2]/table/tbody/tr/td/table/tbody/tr[2]/td/table/tbody/tr/td[1]/table/tbody/tr/td/table/tbody/tr[2]/td[1]/input'))).send_keys(ano)
    
    
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

    if os.path.exists(nome_arq):
        
        os.remove(r'C:\Users\{}\Downloads\CrystalReportViewer1.xlsx'.format(user))
        

    # Confirma a Exportação do Download
    wait.until(EC.element_to_be_clickable((By.XPATH, '/html/body/table[3]/tbody/tr/td/table/tbody/tr[2]/td/table/tbody/tr[2]/td/table/tbody/tr/td/table/tbody/tr/td[2]/nobr/a'))).click()
    sleep(5)

    #Trocando o nome do arquivo
    nome_atual = os.path.join(r'C:\Users\{}\Downloads'.format(user), 'CrystalReportViewer1.xlsx')
    nome_novo = os.path.join(r'C:\Users\{}\Downloads'.format(user), 'Diluição - {}.{}.{}.xlsx'.format(dia, mes, ano))

    nome_consolidado = shutil.move(nome_atual, nome_novo)

    print('Arquivo alterado para: ',nome_consolidado)

    navegador.quit()


def Tratamento_Dados():
    
    with warnings.catch_warnings(record=True):
        warnings.simplefilter('always')
        plan = pd.read_excel(r'C:\Users\{}\Downloads\Diluição - {}.{}.{}.xlsx'.format(user, dia, mes, ano), engine='openpyxl')

    # Filtar linhas onde 'Lote Produto' não contém 'PR'
    plan = plan[plan['Lote Produto'].str.contains('PR', na=False) 
                & plan['Status'].str.contains('IN PROCESS', case=False, na=False)]

    # Resetar índices após a filtragem
    plan.reset_index(drop=True, inplace=True)

    plan.to_excel(r'C:\Users\{}\Downloads\Diluição - {}.{}.{}.xlsx'.format(user, dia, mes, ano), engine='openpyxl', index=False)
    
    caminho_arquivo = r'C:\Users\{}\Downloads\Diluição - {}.{}.{}.xlsx'.format(user, dia, mes, ano)
    
    os.startfile(caminho_arquivo)
    
Import_Diluição() 
Tratamento_Dados()


# By Matheus Siqueira
