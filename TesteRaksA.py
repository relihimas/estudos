import numpy.random.common
import numpy.random.bounded_integers
import numpy.random.entropy
import requests
import time
import pandas as pd
import tkinter as tk
from tkinter import filedialog

# Início, solicitação e leitura de arquivo
root = tk.Tk()
root.withdraw()
root.filename = filedialog.askopenfilename(initialdir='C:/Users/%s', title='Selecione o arquivo', filetypes=(
('Arquivos do Excel (*.xlsx)', "*.xlsx"), ("Todos os arquivos", "*.*")))

arq = pd.read_excel(root.filename)

# URL do desafio (POST do formulário)
url = 'https://docs.google.com/forms/d/e/1FAIpQLSeGkKoum1HEonSxWeZP8r2PDHBncBVxRn61O4x4SwgTILDWtQ/formResponse'

# Tempo entre tentativas caso erro conexão
time_delay = 0.5

# Iniciando request session
s = requests.Session()
s.get(url)

# hard code
nome = 'Rachid Hassan Rafael Elihimas'

# contador e loop para inserção de dados
i = 0
for i in range(i, len(arq)):
    if (str(arq['Complemento'][i]) == 'nan'):
        complemento = 'Sem complemento'
    else:
        complemento = str(arq['Complemento'][i])

        # Inclusão dos dados - formato JSON
        form_data = {
            'entry.1349055322': nome,
            'entry.1587979055': str(arq['Item de entrega'][i]),
            'entry.702002479': str(arq['CEP'][i]),
            'entry.2052823336': str(arq['Endereço'][i]),
            'entry.268273082': complemento,
            'fvv': 1,
            'draftResponse': [],
            'pageHistory': 0,
            'fbzx': -2716156778472223046
        }
# Loop para conexão
    while True:
        try:
            r = s.post(url, data = form_data)
            print(i)
        except:
            print('Erro de conexão, tentando novamente...')
            time.sleep(time_delay)
            continue
        break

print('Concluído')
