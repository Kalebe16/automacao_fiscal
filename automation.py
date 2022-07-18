
from time import sleep
from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager 
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys 
from selenium.webdriver.common.action_chains import ActionChains 
import pandas as pd
from pyautogui import alert, confirm



# Cria um popup para confirmar se o usuario deseja iniciar a automação
popup_inicial = confirm(text='INICIAR AUTOMAÇÃO?', title='AUTOMAÇÃO FISCAL', buttons=['INICIAR', 'CANCELAR'])


# Se a resposta para pergunta_inicial for clicar no botão iniciar, inicie a automação
if popup_inicial == "INICIAR":

    alert(text='EVITE DE MOVIMENTAR O MOUSE E TECLADO DURANTE A AUTOMAÇÃO', title='ALERTA', button='OK')
    


    # inicializando e instanciando o navegador
    servico = Service(ChromeDriverManager().install())
    navegador = webdriver.Chrome(service=servico)
    action = ActionChains(navegador)

    # acessando os dados da minha tabela
    tabela = pd.read_excel("emitir.xlsx", dtype=str)
    print(tabela)


    # percorrendo cada linha da minha tabela
    for i, cpf in enumerate(tabela["CPF"]):  
        # nome_variavel["nome_coluna"]
        # enumerate me retorna o numero da linha dentro da variavel i, ou seja a posição da linha atual do laço
        # tabela.loc["linha", "coluna"], localizar a posiçao exata de um item na tabela passando a coordenada de linha e coluna
        email = tabela.loc[i, "Email"]
        descricao = tabela.loc[i, "Descrição"]
        valor = tabela.loc[i, "Valor"]


        # acessando o link do formulario
        url = "https://forms.gle/8jayUa4QqEsG347DA"
        navegador.get(url)

        # preenche o campo email
        navegador.find_element(By.XPATH, '//*[@id="mG61Hd"]/div[2]/div/div[2]/div[1]/div/div[1]/div[2]/div[1]/div/div[1]/input').send_keys(email)

        # preenche o cpf
        navegador.find_element(By.XPATH, '//*[@id="mG61Hd"]/div[2]/div/div[2]/div[2]/div/div/div[2]/div/div[1]/div/div[1]/input').send_keys(cpf)
        

        # preenche a descrição
        navegador.find_element(By.XPATH, '//*[@id="mG61Hd"]/div[2]/div/div[2]/div[3]/div/div/div[2]/div/div[1]/div[2]/textarea').send_keys(descricao)


        # preenche o tipo de serviço prestado
        sleep(0.5)
        navegador.find_element(By.XPATH, '//*[@id="mG61Hd"]/div[2]/div/div[2]/div[4]/div/div/div[2]/div/div[1]/div[2]').click()
        

        # selecioa o item atividades fisicas no campo de seleção
        for i in range(8):
            sleep(0.5)
            action.key_down(Keys.ARROW_DOWN).perform()

        
        action.key_down(Keys.ENTER).perform()

        
        # preenche o valor
        sleep(1)
        navegador.find_element(By.XPATH, '//*[@id="mG61Hd"]/div[2]/div/div[2]/div[5]/div/div/div[2]/div/div[1]/div/div[1]/input').send_keys(str(valor))

        # clica no botão enviar
        navegador.find_element(By.XPATH, '//*[@id="mG61Hd"]/div[2]/div/div[3]/div[1]/div[1]/div/span/span').click()



    # isso faz com que o selenium não feche o navegador no final do programa
    popup_final = alert(text='CLIQUE EM "ENCERRAR" PARA FINALIZAR O PROGRAMA', title='AUTOMAÇÃO FISCAL', button='ENCERRAR')

    navegador.close()


# Se a resposta para o popup_inicial for fechar a janela ou apertar em 'cancelar', não inicie o programa
else:
    None