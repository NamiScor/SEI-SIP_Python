# <...EXPLICAÇÃO DO CÓDIGO...>
# PARTE 2
# Automação referente as permissões no SIP para a liberação do usuário no SEI
import time
import pandas as pd 
import pyautogui as py
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.support.ui import Select
from IPython.display import display

driver = webdriver.Edge()
driver.maximize_window()  
driver.get('https://sip')
time.sleep(2)

# processo de login
driver.find_element(By.ID, "txtUsuario").send_keys("  ")
driver.find_element(By.ID, "pwdSenha").send_keys("  ")
driver.find_element(By.ID, "selOrgao").click()
time.sleep(1)
py.press('down')
driver.find_element(By.ID, "sbmAcessar").click()
time.sleep(2)

wait = WebDriverWait(driver, 10)
dbu = pd.read_excel("  .xlsx", dtype={"Matrícula": str})

# Navegação até a tela permissão de usuários, IMPORTANTE deixar os times para processar cada comando e conseguir executá-los
driver.find_element(By.ID, "linkMenu24").click()
time.sleep(1)
driver.find_element(By.ID, "linkMenu27").click()
time.sleep(1)
driver.find_element(By.ID, "selSistema").click()
py.press('down')
py.press('enter')

# Seleciona o Órgão da Unidade
org_unidade = wait.until(EC.presence_of_element_located((By.ID, "selOrgaoUnidade")))
Select(org_unidade).select_by_visible_text("PMTL")   

# Seleciona o Órgão do Usuário
org_usuario = wait.until(EC.presence_of_element_located((By.ID, "selOrgaoUsuario")))
Select(org_usuario).select_by_visible_text("PMTL")
   
# Pesquisa pela sigla
for linha in dbu.index:
    nome = str(dbu.loc[linha, "Nome"]).strip().title()
    div = str(dbu.loc[linha, "Nome Divisão"]).strip()
    cpf = str(dbu.loc[linha, "CPF"]).strip()

    # condição da célula estar vazia na planilha
    if pd.isna(nome) or str(nome).strip() == "":
        print('automação ENCERRADA')
        break

    campo = wait.until(EC.element_to_be_clickable((By.ID, "txtUsuario")))
    campo.click()
    campo.clear()
    campo.send_keys(nome)
    time.sleep(1)
    py.press('down')
    py.press('enter')
    time.sleep(1)

    try:
        # Tenta localizar a tabela de permissões (se houver resultados)
        wait.until(EC.presence_of_element_located((By.XPATH, "//table[@class='infraTable']//tr[td]")))
        display (f"Usuário {nome} já possui permissões na {div}")
        driver.find_element(By.ID, "txtUsuario").click()
        driver.find_element(By.ID, "txtUsuario").clear()
        continue 

    except TimeoutException:
        # Se não encontrou nenhuma linha na tela de pesquisa com o texto: "Nenhum registro encontrado"
        driver.find_element(By.ID, "btnNova").click()
        time.sleep(1)

        driver.find_element(By.ID, "selUnidade").send_keys(div)
        py.press('enter')
        time.sleep(2)

        driver.find_element(By.ID, "selPerfil").click()
        py.press('down', presses=4)
        py.press('enter')
        time.sleep(1)

        driver.find_element(By.NAME, "sbmCadastrarPermissao").click()
        print(f'{nome} foi permissionada: {div}')
        time.sleep(1)
print('Operação FINALIZADA...')
driver.quit()


# COMPLETO