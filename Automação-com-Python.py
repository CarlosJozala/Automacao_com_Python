#!/usr/bin/env python
# coding: utf-8

# # Automação de Sistemas e Processos com Python
# 
# ### Desafio:
# 
# Todos os dias, o nosso sistema atualiza as vendas do dia anterior.
# O seu trabalho diário, como analista, é enviar um e-mail para a diretoria, assim que começar a trabalhar, com o faturamento e a quantidade de produtos vendidos no dia anterior
# 
# E-mail da diretoria: seugmail+diretoria@gmail.com<br>
# Local onde o sistema disponibiliza as vendas do dia anterior: https://drive.google.com/drive/folders/149xknr9JvrlEnhNWO49zPcw0PW5icxga?usp=sharing
# 
# Para resolver isso, vamos usar o pyautogui, uma biblioteca de automação de comandos do mouse e do teclado

# In[93]:


# !pip install pyautogui


# In[94]:


import pyautogui
import pyperclip
import time

#Entrar no Sistema

pyautogui.PAUSE = 1

pyautogui.hotkey("ctrl", "t")
pyperclip.copy("https://drive.google.com/drive/folders/149xknr9JvrlEnhNWO49zPcw0PW5icxga?usp=sharing")
pyautogui.hotkey("ctrl", "v")
pyautogui.press("enter")

#Carregamento do site

time.sleep(5)

# Definindo paramêtros
time.sleep(3)
pyautogui.click(x=1237, y=313)
pyautogui.click(x=411, y=335, clicks=2)
time.sleep(2)

# Posição para Downloads
pyautogui.click(x=572, y=416)
pyautogui.click(x=1679, y=205)
pyautogui.click(x=1503, y=661)


# # Vamos agora ler o arquivo baixado para pegar os indicadores
# 
# - Faturamento
# - Quantidade de Produtos

# In[95]:


import pandas as pd

tabela = pd.read_excel(r"C:\Users\dujoz\Downloads\Vendas - Dez.xlsx")
display(tabela)


# In[96]:


quantidade = tabela["Quantidade"].sum()

faturamento = tabela["Valor Final"].sum()

print (quantidade)
print(faturamento)


# ### Vamos agora enviar um e-mail pelo gmail

# In[97]:


#envio por e-mail

pyautogui.hotkey("ctrl", "t")
pyperclip.copy("mail.google.com/mail/u/0/#inbox")
pyautogui.hotkey("ctrl", "v")
pyautogui.press("enter")
time.sleep(4)

#criar e-mail - botão de +
pyautogui.click(x=84, y=235)
time.sleep(2)

#Escrever destinatário
pyautogui.write("dujozala@gmail.com")
pyautogui.press("tab") #Selecionar e-mail
pyautogui.press("tab") #passar pro assunto

#Escrever assunto
pyperclip.copy("Relatório de Vendas")
pyautogui.hotkey("ctrl", "v")
pyautogui.press("tab") # Passar para o corpo de e-mail

#Escrever o corpo de e-mail
texto = f"""
Prezados, bom dia

O faturamento de ontem foi de: R$ {faturamento:,.2f}
A quantidade de produtos foi de: R$ {quantidade:,}

Abraços,

Carlos Jozala

"""

pyperclip.copy(texto)
pyautogui.hotkey("ctrl", "v")

#enviar o e-mail

pyautogui.hotkey("ctrl", "enter")


# #### Use esse código para descobrir qual a posição de um item que queira clicar
# 
# - Lembre-se: a posição na sua tela é diferente da posição na minha tela

# 

# 

# In[98]:


time.sleep(3)
pyautogui.position()


# In[99]:


pyautogui.click(x=269, y=338, clicks=2)

