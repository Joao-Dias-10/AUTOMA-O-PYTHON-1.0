import pyautogui
import time
import shutil
import MySQLdb
import glob
import pandas as pd
import win32com.client as win32
from selenium.webdriver.common.by import By
from datetime import datetime

def enviar_email():

    # criar a integração com o outlook
    outlook = win32.Dispatch('outlook.application')

    # criar um email
    email = outlook.CreateItem(0)

    # configurar as informações do seu e-mail
    email.To = "pessoateste1@eu.inf.br; pessoateste2@eu.inf.br"
    email.Subject = "Informação sobre atualização da base..."
    email.HTMLBody = f"""
    <p>Olá, Base atualizada com Sucesso</p>
    """

    # Anexar arquivo
    #  anexo = "C://Users/Downloads/arquivo.xlsx"
    #  email.Attachments.Add(anexo)

    email.Send()
    print("Email Enviado")

def algo_deu_errado_email():

    # criar a integração com o outlook
    outlook = win32.Dispatch('outlook.application')

    # criar um email
    email = outlook.CreateItem(0)

    # configurar as informações do seu e-mail
    email.To = "pessoateste1@eu.inf.br; pessoateste2@eu.inf.br"
    email.Subject = "Informação sobre atualização da base..."
    email.HTMLBody = f"""
    <p>Olá, Houve algum problema na atualização da base!</p>

    """

    #Anexar arquivo
    #anexo = "C://User/Downloads/arquivo.xlsx"
    #email.Attachments.Add(anexo)

    email.Send()
    print("Email Enviador")

def jogar_no_banco():
    con = MySQLdb.connect(host='',user='',password='',database='')        
    tabela = pd.read_csv(r'C:/Users/User/Documents/BASE.csv',encoding='ISO-8859-1')
    cursorr = con.cursor()
    arquivos = glob.glob('C:/Users/User/Documents/BASE/*.csv')
    for i in arquivos:
        tabela = pd.read_csv(rf'{i}', encoding='ISO-8859-1',sep = ";")
        print(tabela)
        con.commit()
        for i in tabela.fillna(" ").itertuples():
            #colunas da tabela do banco depois ass do excel
            cursorr.execute('insert into banco(coluna1,coluna2,coluna3,coluna4,) values(%s,%s,%s,%s)',
                            (i.coluna1,i.coluna2,i.coluna3,i.coluna4,))
        
        con.commit()

def baixar_base_executar():

    try:
        pyautogui.PAUSE= 0.5

        #PRESSIONAR APP EXCUTAVEL EM ARQUIVOS
        pyautogui.press('winleft')
        pyautogui.doubleClick(590, 327) 
        pyautogui.hotkey('win', 'Up')
        time.sleep(2)
        pyautogui.click(916,66) 
        pyautogui.write('EXECUTADOR')
        pyautogui.press('Enter')
        time.sleep(2)
        
        #Executar
        pyautogui.press('Tab')
        pyautogui.press('Tab')
        pyautogui.press('Enter')
        time.sleep(20)
        
        #Login e senha
        pyautogui.write('12356')
        pyautogui.press('Tab')
        pyautogui.press('Tab')
        time.sleep(1)
        pyautogui.write('88933570')
        time.sleep(1)
        pyautogui.hotkey('Ctrl', 'Enter')
        time.sleep(5)
        pyautogui.click(756, 444)   
        time.sleep(30)

        #dowloads_arquivos
        pyautogui.doubleClick(324, 526)
        time.sleep(5)

        #dowload
        pyautogui.click(1204, 142)
        time.sleep(1)
        pyautogui.press('Tab')
        pyautogui.press('Tab')
        pyautogui.press('Enter')
        time.sleep(15)
        pyautogui.click(742, 430)#ok
        pyautogui.click(1342,8)#fechar

        #logout
        pyautogui.click(450, 615)#logout
        pyautogui.click(703, 443)#sim
        time.sleep(1)
        pyautogui.click(710, 493)#cancelar
        time.sleep(1)
        pyautogui.click(1342,8)# fechar pasta
        time.sleep(2)

        #  mover base baixada
        Base_baixada = r'C:\Users\User\Documents\base.csv'
        # movendo e renomeando
        Base_movida = r'C:\Users\User\Documents\BASESS\base_excel.csv'
        shutil.move(Base_baixada, Base_movida)
        time.sleep(10)


        jogar_no_banco()

        enviar_email()

       


    except:
        algo_deu_errado_email()

# CHAMANDO O QUE RODA TUDO
baixar_base_executar()







