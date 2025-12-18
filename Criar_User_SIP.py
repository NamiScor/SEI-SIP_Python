import time
import pandas as pd 
import pyautogui as py
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from IPython.display import display

driver = webdriver.Edge()
wait = WebDriverWait(driver, 8)
driver.maximize_window()  
driver.get('https://sip-treinamento.treslagoas.ms.gov.br/sip/login.php?sigla_orgao_sistema=PMTL&sigla_sistema=SIP')
time.sleep(2)

# processo de login
driver.find_element(By.ID, "txtUsuario").send_keys("yasmim.silva")
driver.find_element(By.ID, "pwdSenha").send_keys("PMTL@2020")
driver.find_element(By.ID, "selOrgao").click()
time.sleep(1)
py.press('down')
driver.find_element(By.ID, "sbmAcessar").click()
time.sleep(2)

wait.until(EC.element_to_be_clickable((By.ID, "linkMenu51"))).click()
wait.until(EC.element_to_be_clickable((By.ID, "linkMenu53"))).click()

dbu = pd.read_excel("Servidores Outubro.xlsx", dtype={"Matrícula": str})

# Pesquisa pela sigla
for linha in dbu.index:
    nome = str(dbu.loc[linha, "Nome"]).strip().title()
    div = str(dbu.loc[linha, "Nome Divisão"]).strip()
    cpf = str(dbu.loc[linha, "CPF"]).strip()

     # condição da célula estar vazia na planilha
    if pd.isna(nome) or str(nome).strip() == "":
        print('automação encerrada')
        break
    
    nome_sip = wait.until(EC.element_to_be_clickable((By.ID, "txtNomeRegistroCivilUsuario")))
    nome_sip.click()
    nome_sip.clear()
    nome_sip.send_keys(nome)
    py.press('enter')
    time.sleep(1)

    # Condicional do usuário NÃO EXISTIR no SIP
    try:
        wait.until(EC.presence_of_element_located((By.XPATH, "//table[@class='infraTable']//tr[td]")))
        display (f"Usuário {nome} EXISTENTE")
        time.sleep(2)
        continue

    except TimeoutException:
        # criação de um novo usuário no SIP
        print(f'Usuário {nome} NÃO EXISTE, prosseguindo com o cadastro...')    
        driver.find_element(By.ID, "btnNovo").click()
        time.sleep(1)

        driver.find_element(By.ID, 'selOrgao').click()
        py.press('down', presses=2)
        py.press('enter')
        time.sleep(1)

        driver.find_element(By.ID, 'txtSigla').click()
        sigla = '.'.join([nome.split()[0], nome.split()[-1]]).lower()
        driver.find_element(By.ID, 'txtSigla').send_keys(sigla)
        time.sleep(1)

        driver.find_element(By.ID,'txtNome').click()
        driver.find_element(By.ID,'txtNome').send_keys(nome.title())
        time.sleep(1)

        driver.find_element(By.ID,'txtCpf').click()
        driver.find_element(By.ID,'txtCpf').send_keys(cpf)
        time.sleep(1)

        driver.find_element(By.NAME,'sbmCadastrarUsuario').click()
        print(f'Usuário {nome} CRIADO, possui a silga: {sigla}')
        time.sleep(1)

print('Operação finalizada...')
driver.quit()