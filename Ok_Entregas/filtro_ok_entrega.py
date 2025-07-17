import schedule
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
import time
import pyautogui
from cred import OK_ENTREGA_URL, OK_ENTREGA_SENHA, OK_ENTREGA_USUARIO


def filtro_ok_entrega():
    options = Options()
    for arg in [
        "--start_maximized",
        "--disable-extensions",
        "--allow-running-insecure-content",
        f"--unsafely-treat-insecure-origin-as-secure={OK_ENTREGA_URL}",
        "--disable-gpu",
        "--no-sandbox",
        "--disable-dev-shm-usage"
    ]: options.add_argument(arg)

    options.add_experimental_option("prefs", {
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "safebrowsing.enabled": False,
    })

    service = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service, options=options)

    try:
        driver.get(OK_ENTREGA_URL)
        driver.maximize_window()

        email_button = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.NAME, "camp_identificacao"))
        )

        email_button.send_keys(OK_ENTREGA_USUARIO)

        senha = driver.find_element(By.XPATH, '//input[@placeholder="Informe sua senha ..."]')

        senha.send_keys(OK_ENTREGA_SENHA)

        botao_entrar = driver.find_element(By.CLASS_NAME, "loginButton")

        botao_entrar.click()

        print("Entrando no sistema... Aguarde.")
        time.sleep(5)

        hambuguer_menu = WebDriverWait(driver, 10).until((
            EC.presence_of_element_located((By.CLASS_NAME, "hamburger"))
        ))

        hambuguer_menu.click()

        consulta_button = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, '//a[span[contains(text(), "Consulta")]]'))
        )

        consulta_button.click()
        time.sleep(1)

        exportacao_excel = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, '//a[text()="Exportação Excel"]'))
        )

        exportacao_excel.click()
        time.sleep(3)

        close_btn = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.ID, "icon_close"))
        )
        close_btn.click()

        WebDriverWait(driver, 10).until(
            EC.invisibility_of_element_located((By.ID, "alertModalAviso"))
        )

        time.sleep(5)

        filtro_btn = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "a.nav-link.bell-link"))
        )

        filtro_btn.click()

        intermediario_link = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, "//a[text()='Intermediário']"))
        )

        intermediario_link.click()

        timeout = 10
        start_time = time.time()
        selecionar_btn = None

        while time.time() - start_time < timeout:
            selecionar_btn = pyautogui.locateOnScreen("selecionar.png", confidence=0.8)
            if selecionar_btn:
                pyautogui.click(selecionar_btn)
                print("Cliquei no botão Selecionar.")
                break
            time.sleep(0.5)

        if not selecionar_btn:
            print("Botão Selecionar não encontrado na tela.")

        opcao = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, "//li[@class='option' and text()='Emissão NF']")))
        opcao.click()


        # campo_data = WebDriverWait(driver, 10).until(
        #     EC.element_to_be_clickable((By.XPATH, '//input[@placeholder="Selecione a data de entrega"]'))
        # )
        #
        # campo_data.clear()
        # campo_data.send_keys("01/01/2025")



    except Exception as e:
        print(f"Erro durante a execução: {e}")
    finally:
        time.sleep(3)
        driver.quit()


filtro_ok_entrega()
schedule.every().day.at("07:30").do(filtro_ok_entrega)

while True:
    schedule.run_pending()
    print("Wait to filter! \U0001F600")
    time.sleep(1)
