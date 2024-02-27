import openpyxl
from urllib.parse import quote
import webbrowser
from time import sleep
import pyautogui

webbrowser.open("https://web.whatsapp.com/")
sleep(30)

pyautogui.hotkey("ctrl","w")
sleep(2)

workbook = openpyxl.load_workbook("Dados.xlsx")
pagina_clientes = workbook["Plan1"]

for linha in pagina_clientes.iter_rows(min_row=1):
    nome = linha[0].value
    telefone = linha[1].value
    vencimento = linha[2].value
    vencimento = vencimento.strftime("%d/%m/%Y")

    mensagem = f"Olá, {nome}, seu boleto vence no dia {vencimento}. (mensagem enviada com python e precisava de numeros reais)"
    link_mensagem_whatsapp = f"https://web.whatsapp.com/send?phone={telefone}&text={quote(mensagem)}"

    webbrowser.open(link_mensagem_whatsapp)
    sleep(15)
    try:
        seta = pyautogui.locateCenterOnScreen("seta.png")
        sleep(4)

        pyautogui.click(seta[0],seta[1])
        sleep(4)

        pyautogui.hotkey("ctrl","w")
        sleep(4)
    except:
        print(f"não foi possivel enviar a mensagem para {nome}")
        with open("erros.csv", "a", newline="utf-8") as arquivo:
            arquivo.write(f"{nome},{telefone}")