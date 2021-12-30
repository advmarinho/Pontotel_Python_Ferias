from selenium import webdriver
from time import sleep
import pyautogui as abrirsite
import getpass
from openpyxl import load_workbook
import pandas as pd

# pyinstaller --onefile --noconsole .\nome.py
# pyinstaller --onefile --console .\appSiteE1.py
# cxfreeze appSite.py --target-dir dist
# python -m venv folhasSrcSraSrv criando ambiente virtual
# deactivate
# \Scripts\Activate.ps1 ativando ambiente
#Atualizando python -m pip install --upgrade pip  e verificar pip list


print('Siga as instruções  ')
print('Software by ADS - Anderson Marinho  \n\n')
abrirFerias0 = input(r'Digite Caminho da SRR férias: ') + str('\\SRR.xlsx')
abrirFerias0 = pd.read_excel(abrirFerias0, sheet_name='SRR', index_col=None, header=None)
print(abrirFerias0.head())
print('-->| ',str(len(abrirFerias0)) + str(' - ') + str('Funcionário(s)'))


abrirFerias = input(r'Digite Caminho da SRR férias: ') + str('\\SRR.xlsx')
nome_arquivo = abrirFerias
print(nome_arquivo)
print("\n\n\t |   --> Robô ENCONTROU esses dados para incluir Férias <--                   \n")

# int(input("Quantas férias devem ser lançadas: "))
planilha_aberta = load_workbook(filename=nome_arquivo)
sheet_selecionada = planilha_aberta['SRR']
linhaWS = input("A partir da linha tal: ")
qtdlenWS = len(abrirFerias0)
qtdlinhasSite = int(qtdlenWS) - int(linhaWS) + 1
print('Serão enviadas essa(s) quantidade de linhas ', qtdlinhasSite)
print('\nRecomenda-se 30 segundos para cada linha')
segundosUsar = int(input("Quantos segundos devem ser usados no preenchimento das férias no site: "))

login = input("Digite o login: ")
senha = getpass.getpass("Digite a senha: ")

chrome_driver = webdriver.Chrome(input(r'Digite Caminho do CHROME férias: ') + str('\\chromedriver.exe'))

newSite = "https://gestao.pontotel.com.br/#/cognito/login"
chrome_driver.get(newSite)
chrome_driver.maximize_window()
sleep(4)
texto = chrome_driver.find_element_by_xpath('//*[@id="email"]')
sleep(0.5)
abrirsite.write(str(login))
sleep(0.5)
abrirsite.press('enter')
sleep(0.5)
texto = chrome_driver.find_element_by_xpath('//*[@id="view-space"]/div/div/div[1]/form/div[2]/input')
sleep(0.5)
abrirsite.write(senha)
sleep(0.5)
abrirsite.press('enter')
sleep(4)
abrirsite.press('esc', 3)
sleep(0.5)



contador1 = 1
while(contador1 < int(qtdlinhasSite) + 1):
    for linha in range(int(linhaWS), len(sheet_selecionada['C']) + 1):
        matricula = sheet_selecionada['C%s' % linha].value
        nomeFunc = sheet_selecionada['D%s' % linha].value
        inicioFer = sheet_selecionada['F%s' % linha].value.strftime('%d/%m/%Y')
        fimFer = sheet_selecionada['G%s' % linha].value.strftime('%d/%m/%Y')
        print((str(matricula) + str(' - ') + str(nomeFunc) + str(' - ') + str(inicioFer) + str(' - ') + str(fimFer)))

        matricula = str(matricula)
        nomeFunc = str(nomeFunc)
        inicioFer = str(inicioFer)
        fimFer = str(fimFer)
        
        
        chrome_driver.get('https://gestao.pontotel.com.br/#/employee/list')
        sleep(2)
        chrome_driver.find_element_by_xpath('//*[@id="view-space"]/div/div[2]/div[1]/div[1]/div[3]/div[1]/input').send_keys(matricula)
        sleep(2)
        chrome_driver.find_element_by_xpath('//*[@id="employee-list-row-0"]/div[1]/span/span').click()
        sleep(2)
        chrome_driver.find_element_by_xpath('//*[@id="view-space"]/div/div[2]/div[1]/div[3]/div[2]/div/div[1]/div/div[10]/div[5]').click()
        sleep(2)
        contador1 = contador1 + 1
        print(contador1)
        def cronometro(i, f, p):
            print(f'Cronometro de {i} até {f} de {p} em {p}')
            sleep(2)
            if i < f:
                cont = i
                while cont <= f:
                    print(f'{cont} ', end='-', flush=True)
                    sleep(0.5)
                    cont += p
                print(' Outro funcionário')
            else:
                cont = i
                while cont >= f:
                    print(f'{cont} ', end='', flush=True)
                    sleep(0.5)
                    cont -= p
                print(' Outro funcionário')
        cronometro((segundosUsar), 0, 1)
        
        abrirsite.press('enter')
        print('Seguindo para novo funcionário ')
        sleep(2)

        
else:
    print("\n\n\t |   --> Fim da execução <--                   \n")
    chrome_driver.quit()

input('Pressione enter para sair ')
