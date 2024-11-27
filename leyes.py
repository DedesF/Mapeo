import pyautogui
import time


# Realizar un clic izquierdo en la posición actual
pyautogui.click()

# Mover el mouse a otra posición y realizar un clic derecho
pyautogui.moveTo(600, 400, duration=1)
pyautogui.click(button='right')

# Esperar 2 segundos antes de continuar
time.sleep(2)

# Mover el mouse en un bucle
for i in range(5):  # 5 movimientos
    pyautogui.moveRel(50, 50, duration=0.5)  # Mover el mouse en diagonal
    time.sleep(0.5)  # Pausa entre movimientos

print("Movimientos del mouse completados")

# Presionar Tab
#pyautogui.press("tab")

# Escribir texto
pyautogui.write("Hola Mundo!", interval=0.1)

# Esperar 2 segundos
time.sleep(2)

# Presionar Enter
pyautogui.press("enter")

pyautogui.write("Leyes", interval=0.1)
