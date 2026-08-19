import pyautogui
import time

# Abrir navegador Edge
pyautogui.press("win")
pyautogui.write("Edge")
pyautogui.press("enter")
time.sleep(5)

# Acesso LugaRH
pyautogui.write("https://empresa.lugarh.com.br/documentos?Tab=Active&Order=newer&Filter=&PageSize=10&PageNumber=1&QueryTerms=")
pyautogui.press("enter")
time.sleep(14)

# Pesquisa e Seleção do Candidato
pyautogui.click(x=419, y=207)
pyautogui.write("Camily Vitoria")
pyautogui.press("enter")
time.sleep(5)
pyautogui.click(x=422, y=741)

# INICIAR PREENCHIMENTO DOS DADOS DE CONTRATAÇÃO