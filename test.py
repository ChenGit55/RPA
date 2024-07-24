import time
from selenium import webdriver
from selenium.webdriver.common.by import By
import pyautogui
import pandas as pd


empresas_arquivo = [["06/2024", 391, "Empresa teste 1", "123123123123123"],["03/2024", 292, "Empresa teste 2", "098098098098098"], ["05/2024", 148, "Empresa teste 3", "045045045000134"],]

#abrindo arquivo
empresas = pd.DataFrame(empresas_arquivo, columns=["Competência",'Código ER',"Empresa","CNPJ-CPF"] )

def open_browser(url_start_path):
    driver = webdriver.Chrome()
    pyautogui.shortcut('winleft','up')
    driver.get(url_start_path)
    time.sleep(2)

    return driver
################################################################################################################################
# Tela 1

try:
    driver = open_browser("https://www.dominioweb.com.br/")
    tela_login = pyautogui.locateCenterOnScreen('tela_login.png')
    #credenciais
    user = 'usuarioteste@teste.com'
    password = 'Uteste@123!'

    #cria uma variável com os campos da tela de login
    user_field = driver.find_element(By.XPATH, "/html/body/app-root/app-login/div/div[1]/fieldset/div/div/section/form/label[1]/span[2]/input")
    password_field = driver.find_element(By.XPATH, "/html/body/app-root/app-login/div/div[1]/fieldset/div/div/section/form/label[2]/span[2]/input")
    enter_btn = driver.find_element(By.XPATH, "/html/body/app-root/app-login/div/div[1]/fieldset/div/div/section/form/div/button")

    user_field.send_keys(user) #preenche o campo de usuário
    password_field.send_keys(password)#preenche o campo de senha
    enter_btn.click() #clica no botão de entrar
    time.sleep(2)

except:
    print(f"Error ao carregar")
    exit()

# Login concluído e abrindo o programa Domínio no computador
pyautogui.press('winleft')
pyautogui.write('Domínio')
pyautogui.press('enter')

time.sleep(2)
################################################################################################################################
# Tela 2
################################################################################################################################
# Tela 3

try:
    tela_dominio_web= pyautogui.locateCenterOnScreen('tela_dominio_web.png')
    icone_escrita_fiscal= pyautogui.locateCenterOnScreen('escrita_fiscal.png')
    pyautogui.doubleClick(icone_escrita_fiscal)
    time.sleep(2)

except:
    print(f"Error ao carregar")

 ################################################################################################################################
# Tela 4

user_manager = 'Gerente'
password_manager = 'teste@123'

try:
    tela_login_escrita_fiscal =  pyautogui.locateCenterOnScreen('login_escrita_fiscal.png')
    user_manager_input = pyautogui.locateCenterOnScreen('user_manager_input.png')
    pyautogui.click(user_manager_input)
    pyautogui.write(user_manager)
    pyautogui.press('tab')
    pyautogui.write(password_manager)
    time.sleep(2)

except:
    print(f"Error ao carregar")
################################################################################################################################
# Tela 5
################################################################################################################################
# Tela 6


try:
    tela_escrita_fiscal =  pyautogui.locateCenterOnScreen('scrita_fiscal.png')
    for indice,linha in empresas.iterrows():
        cod_er = linha['Código ER']
        pyautogui.press('f8')

        #seleciona a opcao Código
        cod_radio = pyautogui.locateCenterOnScreen('cod_radio.png')
        pyautogui.click(cod_radio)

        #selecionar e ativar a emrpesa
        cod_input = pyautogui.locateCenterOnScreen('cod_input.png')
        pyautogui.click(cod_input)
        pyautogui.write(cod_er)
        pyautogui.shortcut('alt','a')
        pyautogui.press('enter')
        time.sleep(2)

except:
    print(f"Error ao carregar")
################################################################################################################################
# Tela 7
# Após ativar a empresa, no menu superior, selecionar a opção Relatórios e selecionar as demais opções conforme a imagem

#seleciona submenus de relatorio
pyautogui.shortcut('alt','r')
pyautogui.press('n')
pyautogui.press('f')
pyautogui.press('d')
pyautogui.press('m')
time.sleep(2)
################################################################################################################################
 # Tela 8


impostos_nao_calculados = pyautogui.locateCenterOnScreen('impostos_nao_calculados.png')
try:

    comp_filed = pyautogui.locateCenterOnScreen('competencia.png')
    for indice,linha in empresas.iterrows():
        comp = linha['Competência']
        cod_er = linha['Código ER']
        empresa = linha['Empresa']

        #seleciona e troca a competencia
        pyautogui.click(comp_filed)
        pyautogui.write(comp)

        #caminho para salvar arquivo
        file = f'{cod_er}.RFB'
        save_path = f'M:\DCTF\{file}'

        #seleciona o campo para digitar o caminho e digita o caminho
        path_filed = pyautogui.locateCenterOnScreen('caminho.png')
        pyautogui.write(save_path)

        ################################################################################################################################
        # Tela 09

        #exporta
        pyautogui.shortcut('alt','x')
        pyautogui.press('enter')

        if impostos_nao_calculados:
            nome_arquivo = f'{empresa}_erro.txt'
            pyautogui.press('enter')

            with open(nome_arquivo, 'w') as arquivo:
                arquivo.write(empresa)

except:
    print(f"Error ao carregar")

################################################################################################################################
# Tela 10

# Abrir o programa DCTF
pyautogui.press('winleft')
pyautogui.write('DCTF')
pyautogui.press('enter')
################################################################################################################################
# Tela 11

linha_DCTF = pyautogui.locateCenterOnScreen('linha_DCTF.png')
campo_nome_arquivo = pyautogui.locateCenterOnScreen('campo_nome_do_arquivo.png')
try:
    tela_DCTF = pyautogui.locateCenterOnScreen('tela_DCTF.png')
    for indice, linha in empresas:
        cod_er = linha['Código ER']
        file = f'{cod_er}.RFB'

        #selecionar importar
        pyautogui.shortcut('ctrl','m')
        time.sleep(1)
        #clica na linha DCTF
        pyautogui.doubleClick(linha_DCTF)
        #seleciona o campo Nome do Arquivo
        pyautogui.click(campo_nome_arquivo)
        #escreve o nome do arquivo e aperta ok
        pyautogui.write(file)
        pyautogui.shortcut('alt','o')
        pyautogui.press('enter')
        time.sleep(5)
        #aperta ok no caso de sucesso
        pyautogui.press('enter')
        pyautogui.shortcut('alt','c')
        pyautogui.press('enter')
        pyautogui.shortcut('ctrl','a')
        pyautogui.shortcut('alt','o')

    time.sleep(5)

except:
    print(f"Error ao carregar")