import numpy.random.common
import numpy.random.bounded_integers
import numpy.random.entropy
import pandas as pd
from selenium import webdriver
import time
import tkinter as tk
from tkinter import filedialog

root = tk.Tk()
root.withdraw()
file = filedialog.askopenfilename(initialdir='C:/Users/%s')
arq = pd.read_excel('Teste de automacao.xlsx')

browser = webdriver.Chrome()
browser.get('https://docs.google.com/forms/d/e/1FAIpQLSeGkKoum1HEonSxWeZP8r2PDHBncBVxRn61O4x4SwgTILDWtQ/viewform')

i = 0
for i in range(i, len(arq)):
    
    #inserindo o nome
    user_input = browser.find_element_by_xpath('/html/body/div/div[2]/form/div/div/div[2]/div[1]/div/div[2]/div/div[1]/div/div[1]/input')
    user_input.send_keys("Rachid Hassan Rafael Elihimas")
    time.sleep(0.1)
    
    #inserindo o produto
    if arq['Item de entrega'][i] == "Produto A":
        user_input = browser.find_element_by_xpath('/html/body/div/div[2]/form/div/div/div[2]/div[2]/div/div[2]/div/span/div/div[1]/label/div/div[2]/div/span')
        user_input.click()
    elif arq['Item de entrega'][i] == "Produto B":
        user_input = browser.find_element_by_xpath('/html/body/div/div[2]/form/div/div/div[2]/div[2]/div/div[2]/div/span/div/div[2]/label/div/div[2]/div/span')
        user_input.click()
    elif arq['Item de entrega'][i] == "Produto C":
        user_input = browser.find_element_by_xpath('/html/body/div/div[2]/form/div/div/div[2]/div[2]/div/div[2]/div/span/div/div[3]/label/div/div[2]/div/span')
        user_input.click()
    elif arq['Item de entrega'][i] == "Produto D":
        user_input = browser.find_element_by_xpath('/html/body/div/div[2]/form/div/div/div[2]/div[2]/div/div[2]/div/span/div/div[4]/label/div/div[2]/div/span')
        user_input.click()  
    time.sleep(0.1)
    
    #inserindo o CEP
    user_input = browser.find_element_by_xpath('/html/body/div/div[2]/form/div/div/div[2]/div[3]/div/div[2]/div/div[1]/div/div[1]/input')
    user_input.send_keys(str(arq['CEP'][i]))
    time.sleep(0.1)
    
    #inserindo o Endereço
    user_input = browser.find_element_by_xpath('/html/body/div/div[2]/form/div/div/div[2]/div[4]/div/div[2]/div/div[1]/div/div[1]/input')
    user_input.send_keys(str(arq['Endereço'][i]))
    time.sleep(0.1)
    
    #inserindo o Complemento
    if str(arq['Complemento'][i]) == False:
         user_input = browser.find_element_by_xpath('/html/body/div/div[2]/form/div/div/div[2]/div[5]/div/div[2]/div/div[1]/div/div[1]/input')
         user_input.send_keys(str(arq['Complemento'][i]))
    elif str(arq['Complemento'][i]) == True:
         user_input = browser.find_element_by_xpath('/html/body/div/div[2]/form/div/div/div[2]/div[5]/div/div[2]/div/div[1]/div/div[1]/input')
         user_input.send_keys("Sem Complemento")
    time.sleep(0.1)
    
    #enviando
    login_button = browser.find_element_by_xpath('/html/body/div/div[2]/form/div/div/div[3]/div[1]/div/div/span/span')
    login_button.click()
    time.sleep(0.1)

    #outra resposta
    user_input = browser.find_element_by_xpath('/html/body/div[1]/div[2]/div[1]/div/div[4]/a')
    user_input.click()
    time.sleep(0.1)

    i += 1
   
browser.close()
root.mainloop()

print('Ação finalizada')