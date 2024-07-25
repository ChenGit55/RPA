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


class AutomatedTask:

    def __init__(self):
        self.base_path = os.path.dirname(__file__)

    def find_image(self, file_name, max_attempts=5, delay=0.5):
        try:
            full_path = os.path.join(self.base_path, file_name)
            location = None
            for _ in range(max_attempts):
                time.sleep(delay)
                location = pyg.locateCenterOnScreen(full_path, confidence=0.7)
                return location
        except:
            print(f"Erro ao carregar tela da imagem {file_name}")
            exit()

    def click_img_location(self, file_name, max_attempts=5, delay=0.5, double=False):
        location = self.find_image(file_name, max_attempts, delay)
        if double:
            pyg.doubleClick(location)
        else:
            pyg.click(location)

    def write_img_location(
        self, file_name, text="", max_attempts=5, delay=0.5, double=False
    ):
        self.click_img_location(file_name, max_attempts, delay, double)
        pyg.write(text)


driver = open_browser("https://www.dominioweb.com.br/")
task = AutomatedTask()
time.sleep(2)

# credenciais
user = "usuarioteste@teste.com"
password = "Uteste@123!"

# preenche o campo de email e senha
task.write_img_location("user_login.png", user)
pyg.press("tab")
pyg.write(password)
pyg.press("tab")
pyg.press("tab")
pyg.press("enter")

# selecionar escrita fiscal
time.sleep(2)
task.click_img_location("icone_escrita_fiscal.png")

# escrita fiscal login
time.sleep(2)
user_manager = "Gerente"
password_manager = "teste@123"
task.write_img_location("user_manager_input.png", user_manager)
pyg.press("tab")
pyg.write(password_manager)

time.sleep(2)
for indice, linha in empresas.iterrows():
    competencia = linha["Competência"]
    cod_er = linha["Código ER"]
    empresa = linha["Empresa"]

    # caminho para salvar arquivo
    file = f"{cod_er}.RFB"
    save_path = f"M:\DCTF\{file}"

    # altera a empresa
    pyg.press("f8")
    # selecionar e ativar a emrpesa
    task.write_img_location("cod_input.png", cod_er)
    pyg.shortcut("alt", "a")
    pyg.press("enter")

    # abre o menu DCTF mensal caso nao esteja aberto
    pyg.shortcut("alt", "r")
    pyg.press("n")
    pyg.press("f")
    pyg.press("d")
    pyg.press("m")

    # tela pra exportar o arquivo
    time.sleep(2)
    # seleciona e troca a competencia
    task.write_img_location("competencia.png", competencia)

    # seleciona o campo para digitar o caminho e digita o caminho
    task.write_img_location("caminho.png", save_path)

    # exporta
    pyg.shortcut("alt", "x")
    pyg.press("enter")

    # em caso de erro salva um arquivo de texto com o nome da empresa
    try:
        task.find_image("impostos_nao_calculados.png")
        nome_arquivo = f"local/ficticio/{empresa}_erro.txt"
        with open(nome_arquivo, "w") as arquivo:
            arquivo.write(empresa)
    except:
        pass

    pyg.press("enter")
    pyg.shortcut("alt", "f4")

# Abrir o programa DCTF
pyg.press("winleft")
pyg.write("DCTF")
pyg.press("enter")

time.sleep(2)
for indice, linha in empresas:
    cod_er = linha["Código ER"]
    file = f"{cod_er}.RFB"

    # selecionar importar
    pyg.shortcut("ctrl", "m")
    # clica na linha DCTF
    task.click_img_location("linha_DCTF.png", double=True)
    # seleciona o campo Nome do Arquivo e escreve o caminho
    task.write_img_location("campo_nome_do_arquivo.png", file)
    pyg.shortcut("alt", "o")
    pyg.press("enter")
    # aperta ok e cancel pra fechar a janela
    pyg.press("enter")
    pyg.shortcut("alt", "c")
    pyg.press("enter")
    # importa para emresa selecionada
    pyg.shortcut("ctrl", "a")
    pyg.shortcut("alt", "o")


driver.quit()
