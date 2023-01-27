# Automação de Sistemas e Processos com Python
automação web.

Desafio:
Todos os dias, o nosso sistema atualiza as vendas do dia anterior. O seu trabalho diário, como analista, é enviar um e-mail para a diretoria, assim que começar a trabalhar, com o faturamento e a quantidade de produtos vendidos no dia anterior

E-mail da diretoria: seugmail+diretoria@gmail.com
Local onde o sistema disponibiliza as vendas do dia anterior: https://drive.google.com/drive/folders/149xknr9JvrlEnhNWO49zPcw0PW5icxga?usp=sharing

Para resolver isso, vamos usar o pyautogui, uma biblioteca de automação de comandos do mouse e do teclado

Referência do pyautogui: https://pyautogui.readthedocs.io/en/latest/quickstart.html

!pip install pyautogui
Requirement already satisfied: pyautogui in c:\users\fabio\anaconda3\lib\site-packages (0.9.53)
Requirement already satisfied: pyscreeze>=0.1.21 in c:\users\fabio\anaconda3\lib\site-packages (from pyautogui) (0.1.28)
Requirement already satisfied: PyTweening>=1.0.1 in c:\users\fabio\anaconda3\lib\site-packages (from pyautogui) (1.0.4)
Requirement already satisfied: pygetwindow>=0.0.5 in c:\users\fabio\anaconda3\lib\site-packages (from pyautogui) (0.0.9)
Requirement already satisfied: pymsgbox in c:\users\fabio\anaconda3\lib\site-packages (from pyautogui) (1.0.9)
Requirement already satisfied: mouseinfo in c:\users\fabio\anaconda3\lib\site-packages (from pyautogui) (0.1.3)
Requirement already satisfied: pyrect in c:\users\fabio\anaconda3\lib\site-packages (from pygetwindow>=0.0.5->pyautogui) (0.2.0)
Requirement already satisfied: pyperclip in c:\users\fabio\anaconda3\lib\site-packages (from mouseinfo->pyautogui) (1.8.2)
import pyautogui
import pyperclip
import time
​
pyautogui.PAUSE = 1
​
# pyautogui.click -> clicar
# pyautogui.press -> apertar 1 tecla
# pyautogui.hotkey -> conjunto de teclas
# pyautogui.write -> escreve um texto
​
# Passo 1: Entrar no sistema da empresa (no nosso caso é o link do drive)
pyautogui.hotkey("ctrl", "t")
pyperclip.copy("https://drive.google.com/drive/folders/149xknr9JvrlEnhNWO49zPcw0PW5icxga?usp=sharing")
pyautogui.hotkey("ctrl", "v")
pyautogui.press("enter")
​
time.sleep(5)
​
# Passo 2: Navegar no sistema e encontrar a base de vendas (entrar na pasta exportar)
pyautogui.click(x=357, y=263, clicks=2)
time.sleep(2)
​
# Passo 3: Fazer o download da base de vendas
pyautogui.click(x=378, y=370) # clicar no arquivo
pyautogui.click(x=1156, y=160) # clicar nos 3 pontinhos
pyautogui.click(x=967, y=586) # clicar no fazer download
time.sleep(5) # esperar o download acabar
Vamos agora ler o arquivo baixado para pegar os indicadores
Faturamento
Quantidade de Produtos
# Passo 4: Importar a base de vendas pro Python
import pandas as pd
​
tabela = pd.read_excel(r"C:\Users\joaol\Downloads\Vendas - Dez.xlsx")
display(tabela)
---------------------------------------------------------------------------
FileNotFoundError                         Traceback (most recent call last)
~\AppData\Local\Temp\ipykernel_4212\514751453.py in <module>
      2 import pandas as pd
      3 
----> 4 tabela = pd.read_excel(r"C:\Users\joaol\Downloads\Vendas - Dez.xlsx")
      5 display(tabela)

~\anaconda3\lib\site-packages\pandas\util\_decorators.py in wrapper(*args, **kwargs)
    309                     stacklevel=stacklevel,
    310                 )
--> 311             return func(*args, **kwargs)
    312 
    313         return wrapper

~\anaconda3\lib\site-packages\pandas\io\excel\_base.py in read_excel(io, sheet_name, header, names, index_col, usecols, squeeze, dtype, engine, converters, true_values, false_values, skiprows, nrows, na_values, keep_default_na, na_filter, verbose, parse_dates, date_parser, thousands, decimal, comment, skipfooter, convert_float, mangle_dupe_cols, storage_options)
    455     if not isinstance(io, ExcelFile):
    456         should_close = True
--> 457         io = ExcelFile(io, storage_options=storage_options, engine=engine)
    458     elif engine and engine != io.engine:
    459         raise ValueError(

~\anaconda3\lib\site-packages\pandas\io\excel\_base.py in __init__(self, path_or_buffer, engine, storage_options)
   1374                 ext = "xls"
   1375             else:
-> 1376                 ext = inspect_excel_format(
   1377                     content_or_path=path_or_buffer, storage_options=storage_options
   1378                 )

~\anaconda3\lib\site-packages\pandas\io\excel\_base.py in inspect_excel_format(content_or_path, storage_options)
   1248         content_or_path = BytesIO(content_or_path)
   1249 
-> 1250     with get_handle(
   1251         content_or_path, "rb", storage_options=storage_options, is_text=False
   1252     ) as handle:

~\anaconda3\lib\site-packages\pandas\io\common.py in get_handle(path_or_buf, mode, encoding, compression, memory_map, is_text, errors, storage_options)
    793         else:
    794             # Binary mode
--> 795             handle = open(handle, ioargs.mode)
    796         handles.append(handle)
    797 

FileNotFoundError: [Errno 2] No such file or directory: 'C:\\Users\\joaol\\Downloads\\Vendas - Dez.xlsx'

# Passo 5: Calcular os indicadores da empresa
faturamento = tabela["Valor Final"].sum()
print(faturamento)
quantidade = tabela["Quantidade"].sum()
print(quantidade)
Vamos agora enviar um e-mail pelo gmail
# Passo 6: Enviar um e-mail para a diretoria com os indicadores de venda
​
# abrir aba
pyautogui.hotkey("ctrl", "t")
​
# entrar no link do email - https://mail.google.com/mail/u/0/#inbox
pyperclip.copy("https://mail.google.com/mail/u/0/#inbox")
pyautogui.hotkey("ctrl", "v")
pyautogui.press("enter")
time.sleep(5)
​
# clicar no botão escrever
pyautogui.click(x=240, y=415)
​
# preencher as informações do e-mail
pyautogui.write("pythonimpressionador@gmail.com")
pyautogui.press("tab") # selecionar o email
​
pyautogui.press("tab") # pular para o campo de assunto
pyperclip.copy("Relatório de Vendas")
pyautogui.hotkey("ctrl", "v")
​
pyautogui.press("tab") # pular para o campo de corpo do email
​
texto = f"""
Prezados,
​
Segue relatório de vendas.
Faturamento: R${faturamento:,.2f}
Quantidade de produtos vendidos: {quantidade:,}
​
Qualquer dúvida estou à disposição.
Att.,
Lira do Python
"""
​
# formatação dos números (moeda, dinheiro)
​
pyperclip.copy(texto)
pyautogui.hotkey("ctrl", "v")
​
# enviar o e-mail
pyautogui.hotkey("ctrl", "enter")
Use esse código para descobrir qual a posição de um item que queira clicar
Lembre-se: a posição na sua tela é diferente da posição na minha tela
time.sleep(5)
pyautogui.position()
​
