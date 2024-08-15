from selenium import webdriver
from selenium.webdriver.chrome.service import Service as ChromeService
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import pandas as pd
import time
from datetime import datetime

driver = webdriver.Chrome(service=ChromeService(ChromeDriverManager().install()))

driver.get('https://web.whatsapp.com/')
input('Por favor, escaneie o código QR e pressione Enter após o login...')

nome_grupo = 'MOVIMENTE-SE GN'
data_inicial = datetime.strptime('13/08/2024', '%d/%m/%Y')

def localizar_mensagens_no_grupo(palavras_chave):
    global nome_grupo, data_inicial  

    search_box = WebDriverWait(driver, 20).until(
        EC.presence_of_element_located((By.XPATH, '//div[@contenteditable="true"][@data-tab="3"]'))
    )
    search_box.clear()
    search_box.send_keys(nome_grupo + Keys.ENTER)

    WebDriverWait(driver, 20).until(
        EC.presence_of_element_located((By.XPATH, '//div[@role="gridcell"]'))
    )

    last_height = driver.execute_script("return document.querySelector('div[role=\"grid\"]').scrollHeight")
    scroll_attempts = 0
    max_attempts = 100  
    scroll_wait_time = 50

    while scroll_attempts < max_attempts:
        driver.execute_script("document.querySelector('div[role=\"grid\"]').scrollTo(0, 0);")
        time.sleep(scroll_wait_time)  
        new_height = driver.execute_script("return document.querySelector('div[role=\"grid\"]').scrollHeight")
        if new_height == last_height:
            break
        last_height = new_height
        scroll_attempts += 1

    mensagens_encontradas = []
    all_messages = driver.find_elements(By.XPATH, '//div[contains(@class, "message-in") or contains(@class, "message-out")]')

    print(f"Total de mensagens encontradas: {len(all_messages)}")

    for message in reversed(all_messages):
        try:
            texto_mensagem = ""
            spans = message.find_elements(By.XPATH, './/span[@dir="ltr"]')
            for span in spans:
                texto_mensagem += span.text + " "

            imgs = message.find_elements(By.XPATH, './/img[@alt]')
            for img in imgs:
                texto_mensagem += img.get_attribute("alt") + " "

            texto_mensagem = texto_mensagem.strip()
            print(f"Mensagem encontrada: {texto_mensagem}")
        except Exception as e:
            print(f"Erro ao obter texto da mensagem: {e}")
            continue

        for palavra_chave in palavras_chave:
            if palavra_chave in texto_mensagem:
                try:
                    message_data = message.find_element(By.XPATH, './/div[contains(@class, "copyable-text")]')
                    message_time = message_data.get_attribute('data-pre-plain-text')
                    message_date = datetime.strptime(message_time.split(']')[0][1:], '%H:%M, %m/%d/%Y')
                    
                    if message_date >= data_inicial:
                        sender_name = message_time.split('] ')[1].split(':')[0].strip()
                        print(f"Mensagem encontrada com palavra-chave: {texto_mensagem}, Data: {message_date}, Contato: {sender_name}")

                        mensagem_encontrada = {
                            'Data de Envio': message_date.strftime('%d/%m/%Y, %H:%M'),
                            'Mensagem': texto_mensagem,
                            'Contato': sender_name
                        }

                        mensagens_encontradas.append(mensagem_encontrada)
                except Exception as e:
                    print(f"Erro ao obter metadados da mensagem: {e}")
                    continue

    return mensagens_encontradas

palavras_chave = [
    '#Movimente-sePago', '#movimente-sePago', '#movimente-sepago', '#Movimente-sepago', '#MOVIMENTE-SEPAGO', '#MOVIMENTE-SEpago',
    '#movimenteseday', '#Movimenteseday', '#MovimenteseDay', '#movimenteseDay', '#MOVIMENTESEDAY', '#MovimenteSEday',
    '#movimentese day', '#Movimentese day', '#Movimentese Day', '#movimentese Day', '#MOVIMENTESE DAY', '#MovimenteSE day', 
    '#Movimenteseday', '#movimenteseday', '#MovimenteseDay', '#movimenteseDay', '#MOVIMENTESEDAY', '#MovimenteSEday', 
    '#Movimentese day', '#movimentese day', '#Movimentese Day', '#movimentese Day', '#MOVIMENTESE DAY', '#MovimenteSE day',
    '#movimenteseDay', '#MovimenteseDay', '#movimenteseday', '#Movimenteseday', '#MOVIMENTESEDAY', '#MovimenteSEday',
    '#movimentese Day', '#Movimentese Day', '#movimentese day', '#Movimentese day', '#MOVIMENTESE DAY', '#MovimenteSE day',
    '#MovimenteseDay', '#movimenteseDay', '#Movimenteseday', '#movimenteseday', '#MOVIMENTESEDAY', '#MovimenteSEday',
    '#Movimentese Day', '#movimentese Day', '#Movimentese day', '#movimentese day', '#MOVIMENTESE DAY', '#MovimenteSE day',
    '#Movimente-seday', '#movimente-seday', '#Movimente-Seday', '#movimente-seday', '#MOVIMENTE-SEDAY', '#Movimente-SEday',
    '#Movimente-se day', '#movimente-se day', '#Movimente-Se day', '#movimente-se day', '#MOVIMENTE-SE DAY', '#Movimente-SE day'
]

mensagens_encontradas = localizar_mensagens_no_grupo(palavras_chave)

def salvar_em_planilha(dados, nome_arquivo):
    df = pd.DataFrame(dados)
    df.to_excel(nome_arquivo, index=False)
    print(f'Dados salvos em {nome_arquivo}')

if mensagens_encontradas:
    nome_arquivo = 'mensagens_encontradas_13_08.xlsx'
    salvar_em_planilha(mensagens_encontradas, nome_arquivo)
else:
    print(f'Nenhuma mensagem com as palavras-chave "{palavras_chave}" encontrada no grupo "{nome_grupo}"')