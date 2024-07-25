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


def click_on_image(file_name, max_attempts=5, delay=0.5, double=False):
    base_path = os.path.dirname(__file__)  # os.getcdw()
    full_path = os.path.join(base_path, file_name)
    location = None
    for _ in range(max_attempts):
        time.sleep(delay)
        location = pyg.locateCenterOnScreen(full_path, confidence=0.7)
        if location is not None:
            if double:
                pyg.doubleClick(location)
            else:
                pyg.click(location)
            return location


driver = open_browser("https://www.dominioweb.com.br/")
time.sleep(2)

try:
    # user_field = find_image("user_login.png")
    # credenciais
    user = "usuarioteste@teste.com"
    password = "Uteste@123!"

    # preenche o campo de email e senha
    click_on_image("user_login.png")
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
    click_on_image("icone_escrita_fiscal.png")

except:
    print(f"Error ao carregar tela do painel domínio web")
    exit()

# escrita fiscal login
time.sleep(2)
try:
    user_manager = "Gerente"
    password_manager = "teste@123"
    click_on_image("user_manager_input.png")
    pyg.write(user_manager)
    pyg.press("tab")
    pyg.write(password_manager)

except:
    print(f"Error ao carregar tela de login do programa escrita fiscal")
    exit()

time.sleep(2)
for indice, linha in empresas.iterrows():
    competencia = linha["Competência"]
    cod_er = linha["Código ER"]
    empresa = linha["Empresa"]

    # caminho para salvar arquivo
    file = f"{cod_er}.RFB"
    save_path = f"M:\DCTF\{file}"

    # altera a empresa
    try:
        pyg.press("f8")
        # selecionar e ativar a emrpesa
        click_on_image("cod_input.png")
        pyg.write(cod_er)
        pyg.shortcut("alt", "a")
        pyg.press("enter")
    except:
        print(f"Error ao carregar tela de troca de empresas")
        exit()

    # abre o menu DCTF mensal caso nao esteja aberto
    pyg.shortcut("alt", "r")
    pyg.press("n")
    pyg.press("f")
    pyg.press("d")
    pyg.press("m")

    # tela pra exportar o arquivo
    time.sleep(2)
    try:
        # seleciona e troca a competencia
        click_on_image("competencia.png")
        pyg.write(competencia)

        # seleciona o campo para digitar o caminho e digita o caminho
        click_on_image("caminho.png")
        pyg.write(save_path)

        # exporta
        pyg.shortcut("alt", "x")
        pyg.press("enter")

        # em caso de erro salva um arquivo de texto com o nome da empresa
        try:
            click_on_image("impostos_nao_calculados.png")
            nome_arquivo = f"local/ficticio/{empresa}_erro.txt"
            with open(nome_arquivo, "w") as arquivo:
                arquivo.write(empresa)

        except:
            pass

        pyg.press("enter")
        pyg.shortcut("alt", "f4")
    except:
        print(f"Error ao importar DCTF Mensal")
        exit()

# Abrir o programa DCTF
pyg.press("winleft")
pyg.write("DCTF")
pyg.press("enter")

time.sleep(2)
try:
    for indice, linha in empresas:
        cod_er = linha["Código ER"]
        file = f"{cod_er}.RFB"

        # selecionar importar
        pyg.shortcut("ctrl", "m")
        # clica na linha DCTF
        click_on_image("linha_DCTF.png", double=True)
        # seleciona o campo Nome do Arquivo
        click_on_image("campo_nome_do_arquivo.png")
        # escreve o nome do arquivo e aperta ok
        pyg.write(file)
        pyg.shortcut("alt", "o")
        pyg.press("enter")
        # aperta ok e cancel pra fechar a janela
        pyg.press("enter")
        pyg.shortcut("alt", "c")
        pyg.press("enter")
        # importa para emresa selecionada
        pyg.shortcut("ctrl", "a")
        pyg.shortcut("alt", "o")

except:
    print(f"Error ao carregar DCTF Mensal")
    exit()

driver.quit()
