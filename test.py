from gettext import find
import time
from selenium import webdriver
from selenium.webdriver.common.by import By
import pyautogui as pyg
import pandas as pd
import os

base_path = os.path.dirname(__file__)
pyg.PAUSE = 1

# criando data frame com uma lista de empresas
empresas_arquivo = [
    ["06/2024", 391, "Empresa teste 1", "123123123123123"],
    ["03/2024", 292, "Empresa teste 2", "098098098098098"],
    ["05/2024", 148, "Empresa teste 3", "045045045000134"],
]
empresas = pd.DataFrame(
    empresas_arquivo, columns=["Competência", "Código ER", "Empresa", "CNPJ-CPF"]
)


def open_browser(url_start_path):
    driver = webdriver.Chrome()
    pyg.shortcut("winleft", "up")
    driver.get(url_start_path)

    return driver


# funçaao que retorna as cordenadas da imagem
def find_image(file_name):
    base_path = os.path.dirname(__file__)
    full_path = os.path.join(base_path, file_name)
    location = pyg.locateCenterOnScreen(full_path, confidence=0.7)
    return location


################################################################################################################################
# Tela 1

driver = open_browser("https://www.dominioweb.com.br/")
time.sleep(2)
try:
    tela_login = find_image("tela_login.png")
    user_field = find_image("user_login.png")
    # credenciais
    user = "usuarioteste@teste.com"
    password = "Uteste@123!"

    # cria uma variável com os campos da tela de login
    pyg.click(user_field)
    pyg.write(user)
    pyg.press("tab")
    pyg.write(password)
    pyg.press("tab")
    pyg.press("tab")
    pyg.press("enter")
    time.sleep(2)

except:
    print(f"Error ao carregar tela de login")
    exit()

################################################################################################################################
# Tela 2
time.sleep(1)
################################################################################################################################
# Tela 3

try:
    tela_dominio_web = find_image("tela_dominio_web.png")
    icone_escrita_fiscal = find_image("icone_escrita_fiscal.png")
    pyg.doubleClick(icone_escrita_fiscal)
    time.sleep(1)

except:
    print(f"Error ao carregar tela do domínio web")

################################################################################################################################
# Tela 4

user_manager = "Gerente"
password_manager = "teste@123"

try:
    tela_login_escrita_fiscal = find_image("login_escrita_fiscal.png")
    user_manager_input = find_image("user_manager_input.png")
    pyg.click(user_manager_input)
    pyg.write(user_manager)
    pyg.press("tab")
    pyg.write(password_manager)
    time.sleep(2)

except:
    print(f"Error ao carregar dominio web login")
################################################################################################################################
# Tela 5
################################################################################################################################
# Tela 6

try:
    for indice, linha in empresas.iterrows():
        cod_er = linha["Código ER"]
        pyg.press("f8")

        # seleciona a opcao Código
        cod_radio = find_image("cod_radio.png")
        pyg.click(cod_radio)

        # selecionar e ativar a emrpesa
        cod_input = find_image("cod_input.png")
        pyg.click(cod_input)
        pyg.write(cod_er)
        pyg.shortcut("alt", "a")
        pyg.press("enter")
        time.sleep(2)

except:
    print(f"Error ao carregar")
################################################################################################################################
# Tela 7

# seleciona submenus de relatorio
pyg.shortcut("alt", "r")
pyg.press("n")
pyg.press("f")
pyg.press("d")
pyg.press("m")
time.sleep(2)
################################################################################################################################
# Tela 8


try:
    comp_field = find_image("competencia.png")
    for indice, linha in empresas.iterrows():
        comp = linha["Competência"]
        cod_er = linha["Código ER"]
        empresa = linha["Empresa"]

        # caminho para salvar arquivo
        file = f"{cod_er}.RFB"
        save_path = f"M:\DCTF\{file}"

        # seleciona e troca a competencia
        pyg.click(comp_field)
        pyg.write(comp)

        # seleciona o campo para digitar o caminho e digita o caminho
        path_filed = find_image("caminho.png")
        pyg.click(path_filed)
        pyg.write(save_path)

        ################################################################################################################################
        # Tela 09

        # exporta
        pyg.shortcut("alt", "x")
        pyg.press("enter")

        impostos_nao_calculados = find_image("impostos_nao_calculados.png")
        if impostos_nao_calculados:
            nome_arquivo = f"{empresa}_erro.txt"
            pyg.press("enter")

            with open(nome_arquivo, "w") as arquivo:
                arquivo.write(empresa)

except:
    print(f"Error ao carregar")

################################################################################################################################
# Tela 10

# Abrir o programa DCTF
pyg.press("winleft")
pyg.write("DCTF")
pyg.press("enter")

################################################################################################################################
# Tela 11

try:
    tela_DCTF = find_image("tela_DCTF.png")
    linha_DCTF = find_image("linha_DCTF.png")
    campo_nome_arquivo = find_image("campo_nome_do_arquivo.png")

    for indice, linha in empresas:
        cod_er = linha["Código ER"]
        file = f"{cod_er}.RFB"

        # selecionar importar
        pyg.shortcut("ctrl", "m")
        time.sleep(1)
        # clica na linha DCTF
        pyg.doubleClick(linha_DCTF)
        # seleciona o campo Nome do Arquivo
        pyg.click(campo_nome_arquivo)
        # escreve o nome do arquivo e aperta ok
        pyg.write(file)
        pyg.shortcut("alt", "o")
        pyg.press("enter")
        time.sleep(2)
        # aperta ok no caso de sucesso
        pyg.press("enter")
        pyg.shortcut("alt", "c")
        pyg.press("enter")
        pyg.shortcut("ctrl", "a")
        pyg.shortcut("alt", "o")

    time.sleep(2)

except:
    print(f"Error ao carregar")

driver.quit()
