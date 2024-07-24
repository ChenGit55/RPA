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

driver = open_browser("https://www.google.com.br/")

try:
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

# except pyautogui.ImageNotFoundException:
#     print('TELA DE LOGIN NAO CARREGADA')
#     exit()

# Login concluído e abrindo o programa Domínio no computador
pyautogui.press('winleft')
pyautogui.write('Domínio')
pyautogui.press('enter')

time.sleep(2)
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
# Janela de Login
# Utilize as credenciais:

try:
    user_manager = 'Gerente'
    password_manager = 'teste@123'
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
# Login realizado e programa Domínio dentro da área fiscal

tela_escrita_fiscal =  pyautogui.locateCenterOnScreen('scrita_fiscal.png')
time.sleep(2)
################################################################################################################################
# Tela 6
# Apertar f8 para abrir a janela de seleção de empresa
# Escolher a opção código
# Digitar o código da empresa com base na planilha
# Ativar a empresa


#abrir menu
try:
    tela_escrita_fiscal =  pyautogui.locateCenterOnScreen('scrita_fiscal.png')
    for indice,linha in empresas.iterrows():
        pyautogui.press('f8')

        #seleciona a opcao Código
        cod_radio = pyautogui.locateCenterOnScreen('cod_radio.png')
        pyautogui.click(cod_radio)

        #selecionar e ativar a emrpesa
        cod_input = pyautogui.locateCenterOnScreen('cod_input.png')
        pyautogui.click(cod_input)
        pyautogui.write(linha['Código ER'])
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
# Nessa tela, será feita a exportação do relatório referente à empresa ativada anteriormente
# Alterar a competência conforme a planilha, alterar o caminho de onde será salvo e selecionar a opção Exportar
# Caminho que deverá ser escrito (altere a palavra Teste para o código da empresa na planilha):
# 	- M:\DCTF\Teste.RFB

#seleciona e troca a competencia

impostos_nao_calculados = pyautogui.locateCenterOnScreen('impostos_nao_calculados.png')
try:
    comp_filed = pyautogui.locateCenterOnScreen('competencia.png')
    for indice,linha in empresas.iterrows():
        comp = linha['Competência']
        cod_er = linha['Código ER']
        empresa = linha['Empresa']

        pyautogui.click(comp_filed)
        pyautogui.write(comp)

        #caminho para salvar arquivo
        file = f'{cod_er}.RFB'
        save_path = f'M:\DCTF\{file}'

        #seleciona o campo para digitar o caminho e digita o caminho
        path_filed = pyautogui.locateCenterOnScreen('caminho.png')
        pyautogui.write(save_path)

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
# No menu superior, selecionar a opção Declaração e depois a opção Importar

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
