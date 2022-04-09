import pyautogui

screenWidth, screenHeight = pyautogui.size() # Get the size of the primary monitor.
print("TAMANO DA TELA", screenWidth, screenHeight)


currentMouseX, currentMouseY = pyautogui.position() # Get the XY position of the mouse.
print(f"X = {currentMouseX}    Y = {currentMouseY}")


# POSIÇÃO INICIAL X E Y NO POWER POINT (SLIDE 1)
# X = 665    Y = 357

# POSIÇÃO FINAL X E Y NO POWER POINT (SLIDE 1)
# X = 950    Y = 808


# POSIÇÃO INICIAL X E Y NO POWER POINT (SLIDE 1)
# X = 665    Y = 357

# POSIÇÃO FINAL X E Y NO POWER POINT (SLIDE 1)
# X = 1079    Y = 947