from selenium import webdriver
from selenium.webdriver.edge.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time

def getRelatorio(password, email_user, data_inicial, data_final):

    prestadores =  {
        "Neomater": 94,
        "Neotin": 122,
        "Pediatrico": 73
    }

    edge_options = Options()
    edge_options.add_argument('--headless')
    edge_options.add_argument('--disable-gpu')
    edge_options.add_argument('--no-sandbox')
    edge_options.add_argument('--disable-dev-shm-usage')
    edge_options.add_argument('--window-size=1920,1080')

    navegador = webdriver.Edge(options=edge_options)

    navegador.get("https://cisbaf.hygiahub.com.br/login")

    wait = WebDriverWait(navegador, 10)
    login = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="loginform"]/form/div[4]/div/button')))
    email = navegador.find_element(By.NAME, "_username")
    senha = navegador.find_element(By.NAME, "_password")

    email.send_keys(email_user)
    senha.send_keys(password)
    login.click()
    time.sleep(3)

    for prestador, id_prestador in prestadores.items():
        URL = f"https://cisbaf.hygiahub.com.br/relatorio_financeiro/html/?modelo=ANALITICO&status[]=AGENDADO&status[]=REALIZADO&fornecedor[]={id_prestador}&visualizacoes[]=MUNICIPIO&dt_inicial={data_inicial}&dt_final={data_final}&fundo_municipio=%23457b9d&fundo_fornecedor=%23118ab2&fundo_profissional=%233a0ca3&fundo_unidade=%239b5de5&fundo_usuario=%23a8dadc&fundo_totais=%23457b9d&tipo=html"
        navegador.get(URL)
        time.sleep(2)

        navegador.save_screenshot(f'pagina-{prestador}.png')
    navegador.quit()
