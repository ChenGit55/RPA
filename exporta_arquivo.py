from selenium import webdriver
import time
import pyautogui as pyg
import pandas as pd
import os

# criando data frame com uma lista de empresas
empresas_arquivo = [
    ["06/2024", 391, "Empresa teste 1", "123123123123123"],
    ["03/2024", 292, "Empresa teste 2", "098098098098098"],
    ["05/2024", 148, "Empresa teste 3", "045045045000134"],
]
empresas = pd.DataFrame(
    empresas_arquivo, columns=["Competência", "Código ER", "Empresa", "CNPJ-CPF"]
)


# configura o navegador
def open_browser(url_start_path):
    driver = webdriver.Chrome()
    pyg.shortcut("winleft", "up")
    driver.get(url_start_path)

    return driver


# funçaao que retorna as cordenadas da imagem
def find_image(file_name, max_attempts=5, delay=0.5):
    base_path = os.path.dirname(__file__)  # os.getcdw()
    full_path = os.path.join(base_path, file_name)
    location = None
    for _ in range(max_attempts):
        location = pyg.locateCenterOnScreen(full_path, confidence=0.7)
        if location is None:
            time.sleep(delay)
            break
    return location


driver = open_browser("https://www.dominioweb.com.br/")
time.sleep(2)

try:
    user_field = find_image("user_login.png")
    # credenciais
    user = "usuarioteste@teste.com"
    password = "Uteste@123!"

    # preenche o campo de email e senha
    pyg.click(user_field)
    pyg.write(user)
    pyg.press("tab")
    pyg.write(password)
    pyg.press("tab")
    pyg.press("tab")
    pyg.press("enter")

except:
    print(f"Error ao carregar tela de login")
    exit()

# selecionar escrita fiscal
time.sleep(2)
try:
    icone_escrita_fiscal = find_image("icone_escrita_fiscal.png")
    pyg.doubleClick(icone_escrita_fiscal)

except:
    print(f"Error ao carregar tela do painel domínio web")
    exit()


# escrita fiscal login
time.sleep(2)
try:
    user_manager = "Gerente"
    password_manager = "teste@123"
    tela_login_escrita_fiscal = find_image("login_escrita_fiscal.png")
    user_manager_input = find_image("user_manager_input.png")
    pyg.click(user_manager_input)
    pyg.write(user_manager)
    pyg.press("tab")
    pyg.write(password_manager)


except:
    print(f"Error ao carregar tela de login do programa escrita fiscal")
    exit()

time.sleep(2)
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


except:
    print(f"Error ao carregar tela de troca de empresas")
    exit()

# seleciona submenus de relatorio
pyg.shortcut("alt", "r")
pyg.press("n")
pyg.press("f")
pyg.press("d")
pyg.press("m")

# tela pra exportar o arquivo
time.sleep(2)
try:
    comp_field = find_image("competencia.png")
    path_filed = find_image("caminho.png")
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
        pyg.click(path_filed)
        pyg.write(save_path)

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
    print(f"Error ao carregar exportação DCTF Mensal")
    exit()

# Abrir o programa DCTF
pyg.press("winleft")
pyg.write("DCTF")
pyg.press("enter")

time.sleep(2)
try:
    tela_DCTF = find_image("tela_DCTF.png")
    linha_DCTF = find_image("linha_DCTF.png")
    campo_nome_arquivo = find_image("campo_nome_do_arquivo.png")

    for indice, linha in empresas:
        cod_er = linha["Código ER"]
        file = f"{cod_er}.RFB"

        # selecionar importar
        pyg.shortcut("ctrl", "m")
        # clica na linha DCTF
        pyg.doubleClick(linha_DCTF)
        # seleciona o campo Nome do Arquivo
        pyg.click(campo_nome_arquivo)
        # escreve o nome do arquivo e aperta ok
        pyg.write(file)
        pyg.shortcut("alt", "o")
        pyg.press("enter")

        # aperta ok no caso de sucesso
        pyg.press("enter")
        pyg.shortcut("alt", "c")
        pyg.press("enter")
        pyg.shortcut("ctrl", "a")
        pyg.shortcut("alt", "o")


except:
    print(f"Error ao carregar DCTF Mensal")
    exit()

driver.quit()
