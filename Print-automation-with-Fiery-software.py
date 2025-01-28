import win32api
import os
import time
import pyautogui

#

def verificar(a):
    try:
        b = pyautogui.locateCenterOnScreen(a)
    except:
        return False
    else:
        return True
def imprimir(a,b,c):
    pyautogui.hotkey('ctrl', 'p')
    pyautogui.click(x=864, y=297)
    pyautogui.click(x=864, y=320+b)
    pyautogui.press("tab")
    pyautogui.press("enter")
    pyautogui.click(x=1182, y=334)
    pyautogui.click(x=1225, y=520+c)
    pyautogui.press("enter")
    pyautogui.press("tab")
    pyautogui.press("tab")
    pyautogui.press("tab")
    pyautogui.press("tab")
    pyautogui.press("tab")
    pyautogui.write(a)  
    pyautogui.press("tab")
    pyautogui.press("tab")
    pyautogui.press("tab")
    pyautogui.press("tab")
    pyautogui.press("tab")
    pyautogui.press("tab")
    pyautogui.press("tab")
    pyautogui.press("tab")  
    pyautogui.press("tab")
    pyautogui.press("tab")
    pyautogui.press("tab")
    pyautogui.press("tab")
    pyautogui.press("tab")
    pyautogui.press("tab")
    pyautogui.press("tab")
    pyautogui.press("tab")
    pyautogui.press("enter")
try:
    anexo = 0
    arquivo = int(input("Escolha 1 arquivo tendo em vista as opções:\n  1 - 3 anos\n  2 - 6º ano \n  3 - 7º ano\n  4 - 8º ano \n  5 - 9º ano \n  6 - Alfabetização 1\n  7 - Alfabetização 2 \n  8 - Anexo de Alfabetização 1\n  9 - Anexo de Alfabetização 2\n\nEscolha: "))
    copias = int(input("Insira o número de cópias para cada volume: "))
    maq = int(input("Escolha uma MAQ: "))
    if arquivo <= 0 or arquivo > 9:
        raise ValueError("Arquivo escolhido não é válido, rode o programa novamente e insira um valor válido")
    elif copias < 1:
        raise ValueError("Número de cópias inválido, rode o programa novamente e insira um valor válido")
    elif maq < 6 or maq > 10:
        raise ValueError("Máquins inserida não é válida, verifique opções válidas e rode o programa novamente")
    copias = str(copias)
    maq = 18*(maq-6)
    if arquivo == 1:
        caminho = r"C:\Users\Usuário\Desktop\Teste 1"
    elif arquivo == 2:  
        caminho = r"C:\Users\Usuário\Desktop\2025 impressões\6º Ano"
    elif arquivo == 3:  
        caminho = r"C:\Users\Usuário\Desktop\2025 impressões\7º Ano"
    elif arquivo == 4:  
        caminho = r"C:\Users\Usuário\Desktop\2025 impressões\8º Ano"
    elif arquivo == 5:  
        caminho = r"C:\Users\Usuário\Desktop\2025 impressões\9º Ano"
    elif arquivo == 6:  
        caminho = r"C:\Users\Usuário\Desktop\2025 impressões\Alfabetização 1"
    elif arquivo == 7:  
        caminho = r"C:\Users\Usuário\Desktop\2025 impressões\Alfabetização 2"
    elif arquivo == 8:  
        caminho = r"C:\Users\Usuário\Desktop\2025 impressões\Alfabetização 1\Anexo"
        anexo -= 18
    elif arquivo == 9:  
        caminho = r"C:\Users\Usuário\Desktop\2025 impressões\Alfabetização 2\Anexo"
        anexo -= 18
    lista_arquivos = os.listdir(caminho)
    pyautogui.PAUSE = 0.8
    for arquivo in lista_arquivos:
            
        win32api.ShellExecute(0, "open", arquivo, None, caminho, 0)
        time.sleep(5)
        imprimir(copias,maq,anexo)
        time.sleep(10)
        j = False
        while j:
            j = verificar("progresso.png")
        pyautogui.hotkey("ctrl","w")
    pyautogui.hotkey("ctrl","q")
except Exception as erro:
    print(erro)

