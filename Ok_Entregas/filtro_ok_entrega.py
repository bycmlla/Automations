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

        time.sleep(5)

        ok_btn = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, '//button[normalize-space(text())="Ok"]'))
        )
        ok_btn.click()

        WebDriverWait(driver, 10).until(
            EC.invisibility_of_element_located((By.ID, "alertModalAviso"))
        )

        filtro_avancado = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, '//h4[normalize-space(text())="Filtro Avançado"]'))
        )
        filtro_avancado.click()

        time.sleep(3)

        try:
            driver.execute_script(
                "window.scrollBy(0, 800)")

            dropdown_emissao = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((
                    By.XPATH,
                    '//h5[normalize-space(text())="Emissão:"]/ancestor::div[contains(@class, "input-daterange")]//button[contains(@class, "multiselect")]'
                ))
            )

            driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", dropdown_emissao)

            WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH,
                                            '//h5[normalize-space(text())="Emissão:"]/ancestor::div[contains(@class, "input-daterange")]//button[contains(@class, "multiselect")]'
                                            ))
            ).click()
            time.sleep(0.5)

            actions = ActionChains(driver)
            actions.move_by_offset(0, 0).perform()

            emissao_lista = pyautogui.locateOnScreen('ul.png', confidence=0.8)
            if emissao_lista:
                x, y, w, h = emissao_lista
                regiao_filtro_area = (x, y, w, h)

                pos = pyautogui.locateCenterOnScreen('data_especifica.png', confidence=0.8, region=regiao_filtro_area)
                if pos:
                    print("Clicando manualmente em:", pos)
                    pyautogui.click(pos)
                else:
                    print("Não encontrado")
            else:
                print("Área não localizada")


        except Exception as e:
            print(f"Erro durante a execução: {e}")

        campo_data = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, '//input[@placeholder="Selecione a data de entrega"]'))
        )

        campo_data.clear()
        campo_data.send_keys("01/01/2025")

        actions = ActionChains(driver)
        actions.move_by_offset(0, 0).click().perform()

        time.sleep(1)

        btn_filtrar = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((
                By.XPATH,
                '//div[contains(@class, "col-md-12") and .//select[@id="limitar-consulta"]]//button[@id="bt-filtrar"]'
            ))
        )
        btn_filtrar.click()

    except Exception as e:
        print(f"Erro durante a execução: {e}")
    finally:
        time.sleep(3)
        driver.quit()


schedule.every().day.at("07:30").do(filtro_ok_entrega)

while True:
    schedule.run_pending()
    print("Wait to filter! \U0001F600")
    time.sleep(1)
