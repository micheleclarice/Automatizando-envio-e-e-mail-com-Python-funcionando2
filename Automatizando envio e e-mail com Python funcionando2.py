#!/usr/bin/env python
# coding: utf-8

# In[57]:


get_ipython().system('pip install pyautogui')
get_ipython().system('pip install pyperclip')


# In[ ]:





# In[58]:


import pyautogui


# In[59]:


import pyperclip
import time


# In[60]:


pyautogui.PAUSE = 1


# In[61]:


import pyautogui
import pyperclip
import time

pyautogui.PAUSE = 1

# Passo 1: Entrar no sistema (no nosso caso, entrar no link)
pyautogui.hotkey("ctrl", "t")
pyperclip.copy("https://drive.google.com/drive/folders/149xknr9JvrlEnhNWO49zPcw0PW5icxga")
pyautogui.hotkey("ctrl", "v")
pyautogui.press("enter")

time.sleep(3)

# Passo 2: Navegar até o local do relatório (entrar na pasta Exportar)
pyautogui.click(x=438, y=382,clicks=2)
time.sleep(2)

# Passo 3: Fazer o download do relatório
pyautogui.click(x=427, y=575)
pyautogui.click(x=1667, y=248)
pyautogui.click(x=1407, y=729)
time.sleep(5)


# In[62]:


##Vamos agora ler o arquivo baixado para pegar os indicadores
#Faturamento
#Quantidade de Produtos


# In[63]:


import pandas as pd


# In[64]:


# Passo 4: Calcular os indicadores


# In[65]:


tabela = pd.read_excel(r"C:\Users\miche\Downloads\Vendas - Dez.xlsx")


# In[66]:


display(tabela)
faturamento = tabela["Valor Final"].sum()
quantidade = tabela["Quantidade"].sum()


# In[67]:


#Vamos agora enviar um e-mail pelo Hotmail


# In[68]:


# Passo 5: Entrar no email
pyautogui.hotkey("ctrl", "t")
pyperclip.copy("https://mail.google.com/mail/u/0/#inbox")
pyautogui.hotkey("ctrl", "v")
pyautogui.press("enter")
time.sleep(5)

# Passo 6: Enviar por e-mail o resultado
pyautogui.click(x=97, y=256)
time.sleep(7)

pyautogui.write("micheleclarice@hotmail.com")
pyautogui.press("tab") # seleciona o email
pyautogui.press("tab") # pula pro campo de assunto
pyperclip.copy("Relatório de Vendas")
pyautogui.hotkey("ctrl", "v") # escrever o assunto
pyautogui.press("tab") #pular pro corpo do email

texto = f"""
Prezados, bom dia

O faturamento de ontem foi de: R${faturamento:,.2f}
A quantidade de produtos foi de: {quantidade:,}

Abs
Michele Santana"""

pyperclip.copy(texto)
pyautogui.hotkey("ctrl", "v")
time.sleep(2)
pyautogui.hotkey("ctrl", "enter")


# In[69]:


# clicar no botão enviar

# apertar Ctrl Enter
pyautogui.hotkey("ctrl", "enter")


# In[ ]:





# In[70]:


##Use esse código para descobrir qual a posição de um item que queira clicar
##Lembre-se: a posição na sua tela é diferente da posição na minha tela


# In[71]:


time.sleep(5)
pyautogui.position()


# In[ ]:





# In[ ]:




