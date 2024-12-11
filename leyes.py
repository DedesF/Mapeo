import pyautogui
import pandas as pd
import time
from pywinauto import Application
import sys
import os

# Determinar la ruta al archivo leyes.csv
def obtener_ruta_archivo(nombre_archivo):
    # Verificar si estamos en un ejecutable
    if getattr(sys, 'frozen', False):
        # Ejecutándose desde un ejecutable creado con PyInstaller
        base_path = sys._MEIPASS
    else:
        # Ejecutándose como un script normal
        base_path = os.path.abspath(".")
    
    return os.path.join(base_path, nombre_archivo)

# Ruta al archivo leyes.csv
ruta_leyes = obtener_ruta_archivo('leyes.csv')
# Intentar cargar el archivo
try:
    df_leyes = pd.read_csv(ruta_leyes, encoding='UTF-8', header=0)
    print("Archivo leyes.csv cargado exitosamente.")
    print(df_leyes.head())
except FileNotFoundError:    
    print(f"Ruta calculada para leyes.csv: {ruta_leyes}")

#Conecto con la aplicación de acQuire a través de pywinauto
app = Application(backend='win32').connect(title_re=".*acQuire 4.*")

# Instancio la ventana principal del programa
window = app.top_window()

# Mazimizo la ventana de acQuire
window.maximize()

# Obtener el tamaño actual de la pantalla para no tener si el programa se ejecuta en una pantalla con diferente resolución
screen_width, screen_height = pyautogui.size()

# Calcular las coordenadas absolutas
xmin_rel = 0.64
ymin_rel = 0.88

xmin = int(screen_width * xmin_rel)
ymin = int(screen_height * ymin_rel)

# Mover el mouse y realizar un clic
pyautogui.moveTo(xmin, ymin, duration=0.5)  # duration controla la velocidad del movimiento
pyautogui.click()

def ingresar_datos(df):
    for index, row in df.iterrows():
        print(f"Procesando fila {index}: {row}")
        print(f"Desde: {row['Desde']}, Hasta: {row['Hasta']}, % Sulf: {row['% Sulf']}")
        screen_width, screen_height = pyautogui.size()

        # Calcular las coordenadas absolutas
        x_rel = 0.48
        y_rel = 0.36

        x = int(screen_width * x_rel)
        y = int(screen_height * y_rel)

        # Mover el mouse y realizar un clic
        pyautogui.moveTo(x, y, duration=0.5)  # duration controla la velocidad del movimiento
        pyautogui.doubleClick()

        pyautogui.write(f"{row['Desde']}")        
        pyautogui.press('tab')

        pyautogui.write(f"{row['Hasta']}")
        pyautogui.press('tab', presses=4)        
        
        pyautogui.write(f"{row['% Sulf']}")
        pyautogui.press('tab')

        pyautogui.write(f"{row['Py']}")
        pyautogui.press('tab')

        pyautogui.write(f"{row['Cpy']}")
        pyautogui.press('tab')

        pyautogui.write(f"{row['Cc']}")
        pyautogui.press('tab', presses=2)

        pyautogui.write(f"{row['Bo']}")
        pyautogui.press('tab')

        pyautogui.write(f"{row['Mo']}")
        
        pyautogui.press('f9')

        # Pausa para asegurar que el programa externo procese la entrada
        time.sleep(1.5)

ingresar_datos(df_leyes)