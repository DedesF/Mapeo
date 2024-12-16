import pyautogui
import pandas as pd
import time
from pywinauto import Application
import sys
import os


#Funciones
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

def ingresar_leyes(df):
    # Obtener el tamaño actual de la pantalla para no tener si el programa se ejecuta en una pantalla con diferente resolución
    screen_width, screen_height = pyautogui.size()
    # Calcular las coordenadas absolutas para seleccionar la celda 'Desde':mineralizacion y comenzar con la secuencia
    x_rel = 0.48
    y_rel = 0.36
    x = int(screen_width * x_rel)
    y = int(screen_height * y_rel)
    # Mover el mouse y realizar un clic
    pyautogui.moveTo(x, y)  # duration controla la velocidad del movimiento
    pyautogui.doubleClick()

    # Ingreso las leyes
    for index, row in df.iterrows():
        #print(f"Procesando fila {index}: {row}")
        #print(f"Desde: {row['Desde']}, Hasta: {row['Hasta']}, % Sulf: {row['% Sulf']}")
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
        #time.sleep(1.5)

def ingresar_vetillas(df):
    # Obtener el tamaño actual de la pantalla para no tener si el programa se ejecuta en una pantalla con diferente resolución
    screen_width, screen_height = pyautogui.size()
    # Calcular las coordenadas absolutas para seleccionar la celda 'Desde':mineralizacion y comenzar con la secuencia
    x_rel = 0.48
    y_rel = 0.34
    x = int(screen_width * x_rel)
    y = int(screen_height * y_rel)
    # Mover el mouse y realizar un clic
    pyautogui.moveTo(x, y)  # si agregas 'duration=' controla la velocidad del movimiento
    pyautogui.doubleClick()

    for index, row in df.iterrows():
        pyautogui.write(f"{row['Desde']}")        
        pyautogui.press('tab')
        pyautogui.write(f"{row['Hasta']}")
        pyautogui.press('tab', presses=4)
        pyautogui.write(f"{row['T4']}")
        pyautogui.press('tab')
        pyautogui.write(f"{row['A']}")
        pyautogui.press('tab')
        pyautogui.write(f"{row['B']}")
        pyautogui.press('tab')
        pyautogui.write(f"{row['ED']}")
        pyautogui.press('tab')
        pyautogui.write(f"{row['D']}")
        pyautogui.press('tab')
        if row['Comentarios'] == '0':
            pyautogui.press('f9')            
        else:
            pyautogui.write(f"{row['Comentarios']}")
            pyautogui.press('f9')



# definir las rutas a los archivos que se usaran
ruta_leyes = obtener_ruta_archivo('leyes.csv')
ruta_vetillas = obtener_ruta_archivo('vetillas.csv')

# Intentar cargar los archivos
try:
    df_vetillas = pd.read_csv(ruta_vetillas, encoding='UTF-8', sep=',', header=0)
    print("Archivo vetillas.csv cargado exitosamente.")
    print(df_vetillas.head())    
except FileNotFoundError:
    print(f"Ruta calculada para leyes.csv: {df_vetillas}")

try:
    df_leyes = pd.read_csv(ruta_leyes, encoding='UTF-8', sep=',', header=0)
    print("Archivo leyes.csv cargado exitosamente.")
    print(df_leyes.head())    
except FileNotFoundError:
    print(f"Ruta calculada para vetillas.csv: {df_leyes}")

#Conecto con la aplicación de acQuire a través de pywinauto
app = Application(backend='win32').connect(title_re=".*acQuire 4.*")
# Instancio la ventana principal del programa
window = app.top_window()
# Maximizo la ventana de acQuire
window.maximize()

#La ventana collar debe  estar seleccionada
pyautogui.press('f8') #---> selecciono hoja litología
pyautogui.press('f8') #---> selecciono hoja volumen de alteracion
pyautogui.press('f8') #---> selecciono hoja alteracion
pyautogui.press('f8') #---> selecciono hoja estructuras

#-------------------------------------------------------------

pyautogui.press('f8') #---> selecciono hoja mineralizacion
#ingresar_leyes(df_leyes)

#-------------------------------------------------------------

pyautogui.press('f8') #---> selecciono hoja minzone
pyautogui.press('f8') #---> selecciono hoja piso/techo

#-------------------------------------------------------------

pyautogui.press('f8') #---> selecciono hoja vetillas
ingresar_vetillas(df_vetillas)

#-------------------------------------------------------------

pyautogui.press('f8') #---> selecciono hoja unidad geologica