from selenium import webdriver
from selenium.webdriver.common.by import By
import time
import pyautogui
import pandas as pd

empresas_arquivo = [["06/2024", 391, "Empresa teste 1", "123123123123123"],["03/2024", 292, "Empresa teste 2", "098098098098098"], ["05/2024", 148, "Empresa teste 3", "045045045000134"],]

#abrindo arquivo
empresas = pd.DataFrame(empresas_arquivo, columns=["Competência",'Código ER',"Empresa","CNPJ-CPF"] )
cod_er = empresas['Código ER']

#itera sobre as empresas
for indice, linha in empresas.iterrows():
    comp = linha['Competência']
    cod_er = linha['Código ER']
    ################################################################################################################################
    # Tela 1
    # Abrir o navegador, entrar no site https://www.dominioweb.com.br/ e realizar o login
    # Use as credenciais:
    # 	- user: usuarioteste@teste.com
    # 	- password: Uteste@123!

    # Configuração do driver
    driver = webdriver.Chrome()

    # URL da página a ser aberta
    url = "https://www.dominioweb.com.br/"
    driver.get(url)

    time.sleep(1)

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
    ################################################################################################################################
    # Tela 2
    # Login concluído e abrindo o programa Domínio no computador
    time.sleep(2)
    ################################################################################################################################
    # Tela 3
    # Menu de opções do Domínio
    # Selecionar a opção de Escrita Fiscal

    #cria uma váriavel pra selecionar a escrita fiscal e clica
    escrita_fiscal = driver.find_element(By.XPATH, "caminho para o elemento")
    escrita_fiscal.click()
    ################################################################################################################################
    # Tela 4
    # Janela de Login
    # Utilize as credenciais:
    # 	- user: Gerente
    # 	- password: teste@123

    #credenciais pra um novo login
    user_manager = 'Gerente'
    password_manager = 'teste@123'

    #cria uma variável com os novos campos de login
    user_manager_field = driver.find_element(By.XPATH, "camnho para o elemento nome de usuário")
    password_manager_field = driver.find_element(By.XPATH, "camnho para o elemento senha")
    enter_manager_btn = driver.find_element(By.XPATH, "/camnho para o elemento OK(botao)")

    # #preenche os campos e clica
    user_manager_field.send_keys(user)
    password_manager_field.send_keys(password)
    enter_manager_btn.click()

    time.sleep(2)
    ################################################################################################################################
    # Tela 5
    # Login realizado e programa Domínio dentro da área fiscal
    ################################################################################################################################
    # Tela 6
    # Apertar f8 para abrir a janela de seleção de empresa
    # Escolher a opção código
    # Digitar o código da empresa com base na planilha
    # Ativar a empresa

    #abrir menu
    pyautogui.press('f8')

    #seleciona a opcao Código
    cod_radio = user_manager_field = driver.find_element(By.XPATH, "camnho para o elemento radio(Código)")
    cod_radio.click()

    #selecionar e ativar a emrpesa
    cod_input_box = driver.find_element(By.XPATH, "camnho para o elemento input")
    cod_input_box.send_keys(cod_er)
    pyautogui.shortcut('alt','a')
    pyautogui.press('enter')

    time.sleep(2)
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
    competencia_field = driver.find_element(By.XPATH, "camnho para o elemento competencia")
    competencia_field.send_keys(comp)

    #caminho para salvar arquivo
    file = f'{cod_er}.RFB'
    save_path = f'M:\DCTF\{file}'

    #seleciona o campo para digitar o caminho e digita o caminho
    pyautogui.shortcut('alt','c')
    pyautogui.write(save_path)
    pyautogui.press('enter')

    #exporta
    pyautogui.shortcut('alt','x')
    pyautogui.press('enter')
    ################################################################################################################################
    # Tela 9
    # Exportação realizada com sucesso

    #clicar em ok depois de exportado com sucesso
    pyautogui.press('enter')
    ################################################################################################################################
    # Tela 10
    # Abrir o programa DCTF

    pyautogui.press('winleft')
    pyautogui.write('DCTF')
    pyautogui.press('enter')
    ################################################################################################################################
    # Tela 11
    # No menu superior, selecionar a opção Declaração e depois a opção Importar

    #selecionar importar
    pyautogui.shortcut('ctrl','m')
    ################################################################################################################################
    # Tela 12
    # Selecionar a opção DCTF utilizando doubleclick, escrever o nome do arquivo que foi salvo na caixa Nome do Arquivo e selecionar o botão ok

    #selecionar DCFT baseado na posição que esta na tela
    dcft_x = 200
    dcft_y = 200
    pyautogui.doubleClick(dcft_x,dcft_y)

    #selecionar o campo nome do arquivo
    file_name_x = 200
    file_name_y = 200
    pyautogui.doubleClick(file_name_x,file_name_y)
    pyautogui.write(file)

    pyautogui.shortcut('alt','o')
    pyautogui.press('enter')
    ################################################################################################################################
    # Tela 13
    # Importação realizada com sucesso
    # Clicar em ok para fechar a janela
    # O programa voltará para a tela 12, selecionar o botão Cancelar para fechar a janela

    pyautogui.press('enter')
    ################################################################################################################################
    # Tela 14
    # Para abrir essa janela, o programa estará na tela 10, selecione a segunda opção em baixo do menu Declaração
    # Selecionar a opção OK para importar o arquivo para a empresa selecionada
    pyautogui.shortcut('ctrl','a')
    pyautogui.shortcut('alt','o')

    time.sleep(5)