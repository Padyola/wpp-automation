import time
import pandas as pd
import random
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.service import Service


# Caminho para o driver do Chrome
CHROMEDRIVER_PATH = "chromedriver.exe"

# Caminho para a planilha Excel
EXCEL_FILE = "dados.xlsx"

# Tempo de espera para carregar elementos
WAIT_TIME = 5 

# Configurações do Selenium
options = webdriver.ChromeOptions()
options.add_argument("--user-data-dir=./chrome_profile")  # Mantém sessão do usuário
options.add_argument("--disable-gpu")  # Desativa aceleração de GPU
options.add_argument("--no-sandbox")  # Evita problemas de permissão
options.add_argument("--remote-debugging-port=9222")  # Debugger remoto
options.add_argument("--start-maximized")  # Maximiza a janela

# Inicializa o WebDriver
service = Service(CHROMEDRIVER_PATH)
driver = webdriver.Chrome(service=service, options=options)

# Lendo a planilha Excel
df = pd.read_excel(EXCEL_FILE, dtype=str)  # Lê como string para evitar erros
df = df.fillna("")  # Substitui valores nulos por string vazia
print(df)

# Itera sobre as linhas da planilha
for index, row in df.iterrows():
    nome_empresa = str(row.get("NOME DA EMPRESA", "")).strip()
    numero = str(row.get("NUMERO", "")).strip()

    if not numero.isdigit():
        print(f"[ERRO] Número inválido para {nome_empresa}: {numero}")
        continue  # Pula para o próximo

    MENSAGEM = f"Olá, somos da {nome_empresa}! Gostaríamos de conversar sobre uma possível parceria. Quando podemos falar?"

    # Formata o link do WhatsApp
    link_whatsapp = f"https://web.whatsapp.com/send?phone={numero}&text={MENSAGEM}"
    driver.get(link_whatsapp)
    
    # Aguarda carregamento do WhatsApp Web
    time.sleep(WAIT_TIME)
    input_box = driver.switch_to.active_element
    input_box.send_keys(Keys.ENTER)
    print(f"[SUCESSO] Mensagem enviada para {nome_empresa} ({numero})")
    
    time.sleep(random.randint(2, 4))


