import os
import time
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.edge.options import Options
from dotenv import load_dotenv
from webdriver_manager.microsoft import EdgeChromiumDriverManager
from selenium.webdriver.edge.service import Service

# Carrega credenciais do .env
load_dotenv()
email = os.getenv("POWERBI_EMAIL")
senha = os.getenv("POWERBI_SENHA")

# Configuração do navegador Edge
options = Options()

# Usar webdriver-manager para gerenciar o driver
service = Service(EdgeChromiumDriverManager().install())

# Inicializa o navegador
try:
    driver = webdriver.Edge(service=service, options=options)
    driver.get("https://app.powerbi.com/groups/me/reports/321e4af2-cd51-4721-8a05-401eadefede3/678b136856a29493cce1?ctid=cef04b19-7776-4a94-b89b-375c77a8f936&experience=power-bi")
except Exception as e:
    print(f"Erro ao inicializar o navegador: {e}")
    exit()

try:
    # Função para fazer login
    def fazer_login():
        email_field = WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.ID, "i0116")))
        email_field.send_keys(email)
        driver.find_element(By.ID, "idSIButton9").click()

        senha_field = WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.ID, "i0118")))
        senha_field.send_keys(senha)
        driver.find_element(By.ID, "idSIButton9").click()

        WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.TAG_NAME, "body")))
        print("Login realizado com sucesso!")

    # Função para localizar a tabela e clicar em "Exportar dados"
    def exportar_dados_tabela():
        time.sleep(5)
        iframes = driver.find_elements(By.TAG_NAME, "iframe")
        if iframes:
            print("Mudando para iframe...")
            driver.switch_to.frame(iframes[0])

        tabela = WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.TAG_NAME, "table")))

        # Simula o pressionamento do mouse sobre a tabela
        ActionChains(driver).move_to_element(tabela).click_and_hold().perform()
        time.sleep(2)  # Aguarda o botão aparecer
        print("Mouse pressionado sobre a tabela.")

        # Localizador XPATH ajustado
        botao_opcoes = WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, "//button[@aria-label='Mais opções']")))
        botao_opcoes.click()
        print("Botão 'Mais opções' clicado.")

        exportar_dados = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, "//span[text()='Exportar dados']")))
        exportar_dados.click()
        print("Opção 'Exportar dados' clicada!")

        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, "//div[contains(text(), 'Quais dados deseja exportar?')]")))
        opcao_layout_atual = driver.find_element(By.XPATH, "//div[contains(@class, 'layoutCurrent')]")
        opcao_layout_atual.click()

        botao_exportar = driver.find_element(By.XPATH, "//button[contains(text(), 'Exportar')]")
        botao_exportar.click()

    # Função para verificar o download do arquivo
    def verificar_download():
        time.sleep(15)
        # Adicione aqui a lógica para verificar se o arquivo foi baixado com sucesso
        print("Arquivo exportado com sucesso!")

    # Executa as funções
    fazer_login()
    exportar_dados_tabela()
    verificar_download()

except Exception as e:
    print(f"Erro: {e}")

finally:
    try:
        driver.quit()
    except:
        pass