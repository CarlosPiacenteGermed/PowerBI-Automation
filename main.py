from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from dotenv import load_dotenv
import time
import os
import pandas as pd


navegador = webdriver.Chrome()
    # Pressiona ESC
actions = ActionChains(navegador)
navegador.get("https://app.powerbi.com/groups/me/reports/ea2cf9b1-6ee5-4bfe-bc3f-a03da4128881/7383820f33dae729a192?ctid=e6e393d7-971b-4b04-bce1-ba3967f60dd4")
time.sleep(2)

size = WebDriverWait(navegador, 120).until(
    EC.element_to_be_clickable((By.CSS_SELECTOR, '[data-testid="app-bar-view-menu-btn"]'))
)
size.click()


fit_to_width = WebDriverWait(navegador, 30).until(
    EC.element_to_be_clickable((By.CSS_SELECTOR, '[data-testid="fit-to-width-btn"]'))
)
fit_to_width.click()
actions.send_keys(Keys.ESCAPE).perform()
time.sleep(1) # Aguarda um pouco para garantir que o ESC teve efeito

load_dotenv()
bookmark_names = os.getenv("BOOKMARKS").split(",")
username = os.getenv("USERNAME")
download_folder = f"C:\\Users\\{username}\\Downloads"  # ajuste conforme seu sistema
original_filename = "data.xlsx"



for name in bookmark_names:
    bookmark = WebDriverWait(navegador, 120).until(
        EC.element_to_be_clickable((By.ID, "bookmarkButton")),
    )
    # Clica no dropdown
    bookmark.click()
    material = WebDriverWait(navegador, 120).until(
        EC.element_to_be_clickable((By.XPATH, f'//div[@role="listitem" and @aria-label="{name.strip()}"]'))
    )
    material.click()

    actions.send_keys(Keys.ESCAPE).perform()
    time.sleep(1) # Aguarda um pouco para garantir que o ESC teve efeito
    # Captura todos os visualWrapper (retorna uma lista)
    visual_wrappers = WebDriverWait(navegador, 120).until(
        EC.presence_of_all_elements_located((By.CLASS_NAME, "visualWrapper"))
    )

    # Move o cursor até o sexto visualWrapper (índice 5)
    # actions.move_to_element(visual_wrappers[4]).perform()

    import pyautogui

    # Move o mouse para uma posição aproximada (x, y)
    pyautogui.moveTo(500, 900) 

    menu = WebDriverWait(navegador, 30).until(
        EC.element_to_be_clickable((By.CLASS_NAME, "vcMenuBtn")),
    )
    menu.click()
    export_menu = WebDriverWait(navegador, 30).until(
        EC.element_to_be_clickable((By.CSS_SELECTOR, '[data-testid="pbimenu-item.Exportar dados"]'))
    )
    export_menu.click()

    # Clica na sessão "Dados resumidos"
    dados_resumidos = WebDriverWait(navegador, 30).until(
        EC.element_to_be_clickable((By.XPATH, '//section[.//span[text()="Dados resumidos"]]'))
    )
    dados_resumidos.click()

    # Clica no botão "Exportar"
    botao_exportar = WebDriverWait(navegador, 30).until(
        EC.element_to_be_clickable((By.CSS_SELECTOR, '[data-testid="export-btn"]'))
    )
    botao_exportar.click()
    # Move o mouse para uma posição aproximada (x, y)
    pyautogui.moveTo(500, 500) 
    
    time.sleep(5)  # ajuste conforme necessário

    # Renomeia o arquivo
    new_filename = f"{name.strip()}.xlsx"
    original_path = os.path.join(download_folder, original_filename)
    new_path = os.path.join(download_folder, new_filename)

    if os.path.exists(original_path):
        if os.path.exists(new_path):
            os.remove(new_path)  # Remove o arquivo antigo se já existir
        os.rename(original_path, new_path)
        print(f"Arquivo renomeado para: {new_filename}")
    else:
        print(f"Arquivo {original_filename} não encontrado.")
navegador.quit()
import tratamento
tratamento.main()
time.sleep(20)  # Aguarda um pouco para garantir que o ESC teve efeito