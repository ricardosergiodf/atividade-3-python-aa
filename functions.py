import logging
import datetime
import os
import sys
import traceback
from botcity.web import WebBot, Browser, By
import pandas as pd
from botcity.web.browsers.chrome import default_options
from botcity.maestro import *

def setup_logging():
    log_path = "C:/Users/ricar/Desktop/-/Compass/atividades-praticas-compass/Sprint-4/ativ-pratica-3-python-aa/resources/logfiles"
    # Verifica se a pasta "logfiles" existe, se não, cria-a
    if not os.path.exists(log_path):
        os.makedirs(log_path)
    data_atual = datetime.datetime.now().strftime("%d-%m-%Y_%H-%M-%S")
    nome_arquivo_log = f"{log_path}/logfile-{data_atual}.txt"

    logging.basicConfig(
        filename=nome_arquivo_log,
        level=logging.INFO,
        format="(%(asctime)s) - %(levelname)s - %(message)s",
        datefmt="%d/%m/%Y %H:%M:%S"
    )


def bot_driver_setup():
    web_bot = WebBot()
    def_options = default_options(
        headless = web_bot.headless,
    )

    web_bot.options = def_options
    web_bot.browser = Browser.CHROME
    web_bot.driver_path = r"C:\Users\ricar\Downloads\chromedriver-win64\chromedriver-win64\chromedriver.exe"
    return web_bot


def error_exception():
    exc_type, exc_value, exc_tb = sys.exc_info()  # Captura a exceção atual

    if exc_type:  # Verifica se há um erro capturado
        tb = traceback.extract_tb(exc_tb)  # Obtém detalhes do traceback
        last_trace = tb[-1]  # Última linha do traceback (onde o erro ocorreu)

        line_number = last_trace.lineno  # Número da linha do erro
        error_type = exc_type.__name__  # Tipo do erro (ex: ValueError, KeyError)
        error_message = str(exc_value)  # Mensagem do erro

        logging.error(f"Erro do tipo {error_type} na linha {line_number}: {error_message}.")
    else:
        logging.error("Erro não identificado.")

    return None


def browse_url(web_bot, url):
    try:
        logging.info(f"Abrindo Browser na URL: {url}")
        return web_bot.browse(url)
    except Exception:
        error_exception()
        return False


def browse_close(web_bot):
    try:
        logging.info("Fechando Browser.")
        return web_bot.stop_browser()
    except Exception:
        error_exception()
        return False


def download_excel(web_bot, url):
    try:
        browse_url(web_bot, url)
        download_excel_btn = web_bot.find_element("/html[1]/body[1]/app-root[1]/div[2]/app-rpa1[1]/div[1]/div[1]/div[6]/a[1]", By.XPATH)
        download_excel_btn.click()
        logging.info("Clicado no botao de Download Excel, fazendo o download do arquivo.")
        web_bot.wait(1000)
        return True
    except Exception:
        error_exception()
        return False


def read_excel():
    try:
        logging.info("Lendo o arquivo Excel.")
        challenge_spreadsheet = os.path.abspath("./challenge.xlsx")
        data = pd.read_excel(challenge_spreadsheet)
        if data.empty:
            raise ValueError("O arquivo Excel está vazio.")
        else:
            return data  # retorna os dados do arquivo excel
    except Exception:
        error_exception()
        return False
    

def start_challenge(web_bot):
    try:
        start_btn = web_bot.find_element("/html[1]/body[1]/app-root[1]/div[2]/app-rpa1[1]/div[1]/div[1]/div[6]/button[1]", By.XPATH)
        start_btn.click()
        logging.info("Capturado e clicado no botao de Start.")
        return True
    except Exception:
        error_exception()
        return False


def fill_fields(web_bot, data):
    try:
        logging.info(f"Capturando elementos, preenchendo nos campos e clicando no botao de Submit.")
        # loop para pegar todas as infos da planilha excel e dar o submit com todas as infos de cada linha
        for index, row in data.iterrows():
            first_name = web_bot.find_element('//input[@ ng-reflect-name="labelFirstName"]', By.XPATH)
            last_name = web_bot.find_element('//input[@ ng-reflect-name="labelLastName"]', By.XPATH)
            company_name = web_bot.find_element('//input[@ ng-reflect-name="labelCompanyName"]', By.XPATH)
            role_in_company = web_bot.find_element('//input[@ ng-reflect-name="labelRole"]', By.XPATH)
            address = web_bot.find_element('//input[@ ng-reflect-name="labelAddress"]', By.XPATH)
            email = web_bot.find_element('//input[@ ng-reflect-name="labelEmail"]', By.XPATH)
            phone_number = web_bot.find_element('//input[@ ng-reflect-name="labelPhone"]', By.XPATH)
            submit_btn = web_bot.find_element('input[value="Submit"]', By.CSS_SELECTOR)

            first_name.send_keys(row["First Name"])
            last_name.send_keys(row["Last Name "])
            company_name.send_keys(row["Company Name"])
            role_in_company.send_keys(row["Role in Company"])
            address.send_keys(row["Address"])
            email.send_keys(row["Email"])
            phone_number.send_keys(row["Phone Number"])
            submit_btn.click()
    
        return True
    except Exception:
        error_exception()
        return False
    

def success_message(web_bot):
    try:
        success_message = web_bot.find_element("message2", By.CLASS_NAME).text
        logging.info(f"Mensagem de sucesso: {success_message}")
        return True
    except Exception:
        error_exception()
        return False