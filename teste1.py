import pyautogui
from time import sleep

# INICIANDO O PROGRAMA
pyautogui.alert("O CÓDIGO SERÁ EXECUTADO. NÃO MEXA NO COMPUTADOR")


#                                       COPIANDO OS DADOS DA PRIMEIRA ABA DA PLANILHA


# Abrir o arquivo EXCEL
sleep(1)
pyautogui.press(['winleft'])
sleep(1)
pyautogui.write('TESTE123') # <<<<<<<<<<<< COLOCAR O NOME DO ARQUIVO EXCEL
sleep(1)
pyautogui.press(['enter'])

# Localizar a primeira celula
sleep(5)
pyautogui.keyDown('ctrl')
pyautogui.press(['g'])
pyautogui.keyUp('ctrl')
pyautogui.write('B2')  # <<<<<<<<<<<< ALTERAR AQUI O CAMPO INICIAL DA TABELA
pyautogui.press(['enter'])

# Selecionar toda a tabela
pyautogui.keyDown('ctrl')
pyautogui.keyDown('shift')
pyautogui.press(['space'])
pyautogui.keyUp('ctrl')
pyautogui.keyUp('shift')

# Copiar os dados
sleep(3)
pyautogui.hotkey('ctrl', 'c')

# Abrir o arquivo POWER POINT
pyautogui.press(['winleft'])
sleep(1)
pyautogui.write('POWERTESTE.') # <<<<<<<<<<<< COLOCAR O NOME DO ARQUIVO PPT
sleep(1)
pyautogui.press(['enter'])

# Selecionar a tabela para ser atualizada no Power Point
sleep(3)
pyautogui.moveTo(665,357)
pyautogui.mouseDown()
pyautogui.moveTo(950,808)
sleep(0.5)
pyautogui.mouseUp()

# Colando os dados
pyautogui.hotkey('ctrl', 'v')


#                                       COPIANDO OS DADOS DA SEGUNDA ABA DA PLANILHA


# Voltando para o Excel
sleep(1)
pyautogui.keyDown('alt')
pyautogui.press(['tab'])
pyautogui.keyUp('alt')

# Mudando para a próxima aba
sleep(1)
pyautogui.keyDown('ctrl')
pyautogui.press(['pgdn'])
pyautogui.keyUp('ctrl')

# Localizar a primeira celula
sleep(5)
pyautogui.keyDown('ctrl')
pyautogui.press(['g'])
pyautogui.keyUp('ctrl')
pyautogui.write('B2')  # <<<<<<<<<<<< ALTERAR AQUI O CAMPO INICIAL DA TABELA
pyautogui.press(['enter'])

# Selecionar toda a tabela
pyautogui.keyDown('ctrl')
pyautogui.keyDown('shift')
pyautogui.press(['space'])
pyautogui.keyUp('ctrl')
pyautogui.keyUp('shift')

# Copiar os dados
sleep(3)
pyautogui.hotkey('ctrl', 'c')

# Voltando para o Power Point
sleep(1)
pyautogui.keyDown('alt')
pyautogui.press(['tab'])
pyautogui.keyUp('alt')

# Mudando para o próximo slide
sleep(1)
pyautogui.keyDown('ctrl')
pyautogui.press(['pgdn'])
pyautogui.keyUp('ctrl')

# Selecionar a tabela para ser atualizada no Power Point
sleep(3)
pyautogui.moveTo(665,357)
pyautogui.mouseDown()
pyautogui.moveTo(1079,947)
sleep(0.5)
pyautogui.mouseUp()

# Colando os dados
pyautogui.hotkey('ctrl', 'v')





# FINALIZANDO O PROGRAMA
sleep(3)
pyautogui.alert("O código foi finalizado. Você já pode utilizar o computador!")

