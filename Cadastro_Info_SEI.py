# <...EXPLICAÇÃO DO CÓDIGO...>
# PARTE 3 (FINAL)
# Automação para cadastro de informações pessoais do usuário no SEI, com permissão já consedida no SIP
# Após ser criado ou não os usuários (ETAPA 1), e permissionados para aparecer seu registro no SEI através da permissão obrigatória do SIP para o SEI (ETAPA 2),
# assim, será preciso preencher os dados dessess para ficar registrado em seu cadastro no sistema e ser dispobilizado tal para realizar alguma operação dentro do sistema.
import time
import pandas as pd 
import pyautogui as py
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
driver = webdriver.Edge()
driver.maximize_window()    
driver.get('https://sei')
time.sleep(2)

# processo de login
driver.find_element(By.ID, "txtUsuario").send_keys("  ")
driver.find_element(By.ID, "pwdSenha").send_keys("  ")
driver.find_element(By.ID, "selOrgao").click()
time.sleep(1)
py.press('down')
driver.find_element(By.ID, "sbmAcessar").click()
time.sleep(2)

# Navegação até a tela de cadastro de usuários, IMPORTANTE deixar os times para processar cada comando e conseguir executá-los
driver.find_element(By.ID, "linkMenu1").click() 
time.sleep(1)
driver.find_element(By.ID, "linkMenu126").click()
time.sleep(1)
driver.find_element(By.ID, "linkMenu127").click()
time.sleep(1)
driver.find_element(By.ID, 'selOrgao').click()
driver.find_element(By.XPATH, '//*[@id="selOrgao"]/option[4]').click()
time.sleep(1)

wait=webdriver.support.ui.WebDriverWait(driver, 10)
db = pd.read_excel("  .xlsx", dtype={"Matrícula": str})

for linha in db.index:  
    camp_name = db.loc[linha, "Nome"]
    camp_sexo = str(db.loc[linha, "Sexo"]).strip().upper()  
    camp_cpf = db.loc[linha, "CPF"]
    camp_matric = db.loc[linha, "Matrícula"]

    # condição da célula estar vazia
    if pd.isna(camp_name) or str(camp_name).strip() == "":
        print('automação ENCERRADA')
        driver.quit()
        exit()  

    #Pesquisa do usuário
    driver.find_element(By.ID, "txtNomeUsuario").clear()
    driver.find_element(By.ID, "txtNomeUsuario").send_keys(camp_name)
    driver.find_element(By.ID, "btnPesquisar").click()
    time.sleep(1)

    # Condicional caso o usuário não for encontrado durante a pesquisa 
    try:
        # se o usuário existe
        btn_edit_info = WebDriverWait(driver, 3).until(EC.element_to_be_clickable((By.XPATH,"//img[@alt='Alterar Usuário']/parent::a")))
        btn_edit_info.click()
    except TimeoutException:
        # se não existe
        print(f"Usuário {camp_name} não encontrado.") 
        driver.find_element(By.ID, "txtNomeUsuario").clear()
        continue

    btn_open_window = wait.until(EC.element_to_be_clickable((By.ID, "imgAlterarContato")))
    btn_open_window.click()
    time.sleep(1)

    # Abaixo, tudo corresponde a troca de contexto para o modal de edição aberto, para o selenium identificar o modal
    # Espera o iframe do modal aparecer
    iframe = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.CSS_SELECTOR, "iframe[name='modal-frame']")))

    # Troca o foco para dentro do iframe
    driver.switch_to.frame(iframe)

    # Agora você pode acessar os campos do CPF e Matrícula para preenche-los
    cpf_input = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "txtCpf")))
    cpf_input.clear()
    cpf_input.send_keys(camp_cpf)
    time.sleep(1)
    matricula_input = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "txtMatricula")))
    matricula_input.clear()
    matricula_input.send_keys(str(camp_matric))
    time.sleep(1)
    btn_topo = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID, "btnInfraTopo")))
    btn_topo.click()
    time.sleep(1)

    # Para o cadastro é preciso definir o gênero do usuário, conforme os dados dos usuários na base
    if camp_sexo == "F":
        label_fem = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.CSS_SELECTOR, "label[for='optFeminino']")))
        label_fem.click()
        time.sleep(1)
    elif camp_sexo == "M":
        label_mas = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.CSS_SELECTOR, "label[for='optMasculino']")))
        label_mas.click()
        time.sleep(1)
    else:
        print('Erro de identificação')
        exit()

    # Botão "Salvar" dentro do modal
    btn_salvar = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.NAME, "sbmAlterarContato")))
    btn_salvar.click()
    time.sleep(1)

    # Botão "salvar" para fechar a parte de edição do usuário e dar continuidade ao loop e volta para janela de pesquisa inicial
    driver.switch_to.default_content() 
    
    # Botão "Salvar" da tela principal de edição do usuário, FECHA para ir para a próxima linha do loop
    btn_closed = wait.until(EC.element_to_be_clickable((By.NAME, "sbmAlterarUsuario")))
    btn_closed.click()

    print(f'Usuário {camp_name} preenchido!')
    
print(f'Operação FINALIZADA...')
driver.quit()

# COMPLETO