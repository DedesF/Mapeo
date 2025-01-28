import pyautogui
import pandas as pd
from pywinauto import Application
import time
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

def ingresar_litologia(df):
    # Obtener el tamaño actual de la pantalla para no tener si el programa se ejecuta en una pantalla con diferente resolución
    screen_width, screen_height = pyautogui.size()
    # Calcular las coordenadas absolutas para seleccionar la celda 'Desde':mineralizacion y comenzar con la secuencia
    x_rel = 0.46
    y_rel = 0.36
    x = int(screen_width * x_rel)
    y = int(screen_height * y_rel)
    # Mover el mouse y realizar un clic
    pyautogui.moveTo(x, y)  # duration controla la velocidad del movimiento
    pyautogui.doubleClick()

    # Ingreso las litologia
    for index, row in df.iterrows():
        #print(f"Procesando fila {index}: {row}")
        #print(f"Desde: {row['Desde']}, Hasta: {row['Hasta']}, % Sulf: {row['% Sulf']}")
        pyautogui.write(f"{row['Desde']}")        
        pyautogui.press('tab')
        pyautogui.write(f"{row['Hasta']}")
        pyautogui.press('tab')
        pyautogui.press('esc')
        time.sleep(0.2)
        pyautogui.write(f"{row['Litologia']}")
        time.sleep(0.2)
        if pd.isna(row['Comentarios']) or row['Comentarios'] == '':
            pyautogui.press('f9')            
        else:
            pyautogui.press('tab', presses=4)            
            time.sleep(0.2)
            pyautogui.write(f"{row['Comentarios']}")
            pyautogui.press('f9')
        # Pausa para asegurar que el programa externo procese la entrada
        #time.sleep(1.5)

def ingresar_volalt(df):
    # Obtener el tamaño actual de la pantalla para no tener si el programa se ejecuta en una pantalla con diferente resolución
    screen_width, screen_height = pyautogui.size()
    # Calcular las coordenadas absolutas para seleccionar la celda 'Desde':mineralizacion y comenzar con la secuencia
    x_rel = 0.46
    y_rel = 0.36
    x = int(screen_width * x_rel)
    y = int(screen_height * y_rel)
    # Mover el mouse y realizar un clic
    pyautogui.moveTo(x, y)  # duration controla la velocidad del movimiento
    pyautogui.doubleClick()

    # Ingreso el volumen de alteracion
    for index, row in df.iterrows():
        #print(f"Procesando fila {index}: {row}")
        #print(f"Desde: {row['Desde']}, Hasta: {row['Hasta']}, % Sulf: {row['% Sulf']}")
        #Defino las columnas que se deben sumar
        columnas =['K (Biot)','CS','QS','SI','K (Feld)','ARC','Epid','Cl','NR', 'K (Hist)']

        pyautogui.write(f"{row['Desde']}")        
        pyautogui.press('tab')
        pyautogui.write(f"{row['Hasta']}")
        pyautogui.press('tab', presses=4)
        pyautogui.write(f"{row['K (Biot)']}")
        pyautogui.press('tab')
        pyautogui.write(f"{row['CS']}")
        pyautogui.press('tab')
        pyautogui.write(f"{row['QS']}")
        pyautogui.press('tab')
        pyautogui.write(f"{row['SI']}")
        pyautogui.press('tab')
        pyautogui.write(f"{row['K (Feld)']}")
        pyautogui.press('tab')
        pyautogui.write(f"{row['ARC']}")
        pyautogui.press('tab')
        pyautogui.write(f"{row['Epid']}")
        pyautogui.press('tab')
        pyautogui.write(f"{row['Cl']}")
        pyautogui.press('tab')
        pyautogui.write(f"{row['NR']}")
        pyautogui.press('tab')
        pyautogui.write(f"{row['K (Hist)']}")        
        if pd.isna(row['Comentarios']) or row['Comentarios'] == '':
            pyautogui.press('f9')            
        else:
            pyautogui.press('tab')
            pyautogui.write(f"{row['Comentarios']}")
            pyautogui.press('f9')        
        # Sumar los valores de las columnas con alteración para validarla en acquire cuando la suma es igual a 0
        if pd.to_numeric(row[columnas], errors='coerce').fillna(0).sum() == 0:
            pyautogui.press('tab')
            pyautogui.press('enter')
        # Pausa para asegurar que el programa externo procese la entrada
        #time.sleep(1.5)

def ingresar_alteracion(df):
    # Obtener el tamaño actual de la pantalla para no tener si el programa se ejecuta en una pantalla con diferente resolución
    screen_width, screen_height = pyautogui.size()
    # Calcular las coordenadas absolutas para seleccionar la celda 'Desde':mineralizacion y comenzar con la secuencia
    x_rel = 0.47
    y_rel = 0.36
    x = int(screen_width * x_rel)
    y = int(screen_height * y_rel)
    # Mover el mouse y realizar un clic
    pyautogui.moveTo(x, y)  # duration controla la velocidad del movimiento
    pyautogui.doubleClick()

    # Ingreso las litologia
    for index, row in df.iterrows():
        #print(f"Procesando fila {index}: {row}")
        #print(f"Desde: {row['Desde']}, Hasta: {row['Hasta']}, % Sulf: {row['% Sulf']}")
        pyautogui.write(f"{row['Desde']}")        
        pyautogui.press('tab')
        pyautogui.write(f"{row['Hasta']}")
        pyautogui.press('tab', presses=4)
        if pd.isna(row['Comentarios']) or row['Comentarios'] == '':
            pyautogui.press('tab')            
        else:            
            pyautogui.write(f"{row['Comentarios']}")
            pyautogui.press('tab')
        pyautogui.press('esc')
        pyautogui.write(f"{row['Tipo Alt']}")
        pyautogui.press('f9')       
        # Pausa para asegurar que el programa externo procese la entrada
        #time.sleep(1.5)'''

def ingresar_estructuras(df):
    # Obtener el tamaño actual de la pantalla para no tener si el programa se ejecuta en una pantalla con diferente resolución
    screen_width, screen_height = pyautogui.size()
    # Calcular las coordenadas absolutas para seleccionar la celda 'Desde':mineralizacion y comenzar con la secuencia
    x_rel = 0.47
    y_rel = 0.35
    x = int(screen_width * x_rel)
    y = int(screen_height * y_rel)
    # Mover el mouse y realizar un clic
    pyautogui.moveTo(x, y)  # duration controla la velocidad del movimiento
    pyautogui.doubleClick()

    # Ingreso las estructuras
    for index, row in df.iterrows():
        #print(f"Procesando fila {index}: {row}")
        #print(f"Desde: {row['Desde']}, Hasta: {row['Hasta']}, % Sulf: {row['% Sulf']}")
        pyautogui.write(f"{row['Desde']}")        
        pyautogui.press('tab')
        pyautogui.write(f"{row['Hasta']}")
        pyautogui.press('tab', presses=4)
        pyautogui.write(f"{row['Relleno de Est']}")
        if pd.isna(row['Angulo']) or row['Angulo'] == '':
            pyautogui.press('tab', presses=2)            
        else:
            pyautogui.press('tab')
            pyautogui.write(f"{row['Angulo']}")
            pyautogui.press('tab')
        if pd.isna(row['Comentarios']) or row['Comentarios'] == '':
            pyautogui.press('f9')            
        else:            
            pyautogui.write(f"{row['Comentarios']}")
            pyautogui.press('f9')
        # Pausa para asegurar que el programa externo procese la entrada
        #time.sleep(1.5)'''

def ingresar_mineralizacion(df):
    # Obtener el tamaño actual de la pantalla para no tener si el programa se ejecuta en una pantalla con diferente resolución
    screen_width, screen_height = pyautogui.size()
    # Calcular las coordenadas absolutas para seleccionar la celda 'Desde':mineralizacion y comenzar con la secuencia
    x_rel = 0.49
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
        #time.sleep(1.5)'''

def ingresar_minzone(df):
    # Obtener el tamaño actual de la pantalla para no tener si el programa se ejecuta en una pantalla con diferente resolución
    screen_width, screen_height = pyautogui.size()
    # Calcular las coordenadas absolutas para seleccionar la celda 'Desde':mineralizacion y comenzar con la secuencia
    x_rel = 0.46
    y_rel = 0.37
    x = int(screen_width * x_rel)
    y = int(screen_height * y_rel)
    # Mover el mouse y realizar un clic
    pyautogui.moveTo(x, y)  # duration controla la velocidad del movimiento
    pyautogui.doubleClick()

    # Ingreso las litologia
    for index, row in df.iterrows():
        #print(f"Procesando fila {index}: {row}")
        #print(f"Desde: {row['Desde']}, Hasta: {row['Hasta']}, % Sulf: {row['% Sulf']}")
        pyautogui.write(f"{row['Desde']}")        
        pyautogui.press('tab')
        pyautogui.write(f"{row['Hasta']}")
        pyautogui.press('tab')
        pyautogui.press('esc')
        pyautogui.write(f"{row['MinZone']}")        
        if pd.isna(row['Comentarios']) or row['Comentarios'] == '':
            pyautogui.press('f9')            
        else:
            pyautogui.press('tab', presses=4)
            pyautogui.write(f"{row['Comentarios']}")
            pyautogui.press('f9')
        # Pausa para asegurar que el programa externo procese la entrada
        #time.sleep(1.5)'''

def ingresar_pisotecho(df):
    # Obtener el tamaño actual de la pantalla para no tener si el programa se ejecuta en una pantalla con diferente resolución
    screen_width, screen_height = pyautogui.size()
    # Calcular las coordenadas absolutas para seleccionar la celda 'Desde':mineralizacion y comenzar con la secuencia
    x_rel = 0.46
    y_rel = 0.37
    x = int(screen_width * x_rel)
    y = int(screen_height * y_rel)
    # Mover el mouse y realizar un clic
    pyautogui.moveTo(x, y)  # duration controla la velocidad del movimiento
    pyautogui.doubleClick()

    # Ingreso las litologia
    for index, row in df.iterrows():
        #print(f"Procesando fila {index}: {row}")
        #print(f"Desde: {row['Desde']}, Hasta: {row['Hasta']}, % Sulf: {row['% Sulf']}")
        pyautogui.write(f"{row['Profundidad']}")        
        pyautogui.press('tab')
        pyautogui.press('esc')
        pyautogui.write(f"{row['Piso Techo']}")
        if pd.isna(row['Comentarios']) or row['Comentarios'] == '':
            pyautogui.press('f9')
        else:
            pyautogui.press('tab', presses=4)
            pyautogui.write(f"{row['Comentarios']}")
            pyautogui.press('f9')
        # Pausa para asegurar que el programa externo procese la entrada
        #time.sleep(1.5)'''

def ingresar_vetillas(df):
    # Obtener el tamaño actual de la pantalla para no tener si el programa se ejecuta en una pantalla con diferente resolución
    screen_width, screen_height = pyautogui.size()
    # Calcular las coordenadas absolutas para seleccionar la celda 'Desde':mineralizacion y comenzar con la secuencia
    x_rel = 0.47
    y_rel = 0.36
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
        if pd.isna(row['Comentarios']) or row['Comentarios'] == '':
            pyautogui.press('f9')            
        else:
            pyautogui.write(f"{row['Comentarios']}")
            pyautogui.press('f9')

def ingresar_ug(df):
    # Obtener el tamaño actual de la pantalla para no tener si el programa se ejecuta en una pantalla con diferente resolución
    screen_width, screen_height = pyautogui.size()
    # Calcular las coordenadas absolutas para seleccionar la celda 'Desde':mineralizacion y comenzar con la secuencia
    x_rel = 0.46
    y_rel = 0.36
    x = int(screen_width * x_rel)
    y = int(screen_height * y_rel)
    # Mover el mouse y realizar un clic
    pyautogui.moveTo(x, y)  # duration controla la velocidad del movimiento
    pyautogui.doubleClick()

    # Ingreso las litologia
    for index, row in df.iterrows():
        #print(f"Procesando fila {index}: {row}")
        #print(f"Desde: {row['Desde']}, Hasta: {row['Hasta']}, % Sulf: {row['% Sulf']}")
        pyautogui.write(f"{row['Desde']}")
        pyautogui.press('tab')
        pyautogui.write(f"{row['Hasta']}")
        pyautogui.press('tab')
        pyautogui.press('esc')
        time.sleep(0.3)
        pyautogui.write(f"{row['UG']}")
        time.sleep(0.3)
        if pd.isna(row['Comentarios']) or row['Comentarios'] == '':
            pyautogui.press('f9')            
        else:
            pyautogui.press('tab', presses=4)
            pyautogui.write(f"{row['Comentarios']}")
            pyautogui.press('f9')
        # Pausa para asegurar que el programa externo procese la entrada
        #time.sleep(1.5)

# definir las rutas a los archivos que se usaran
ruta_lit = obtener_ruta_archivo('litologia.csv')
ruta_volalt = obtener_ruta_archivo('volalt.csv')
ruta_alteracion = obtener_ruta_archivo('alteracion.csv')
ruta_estructuras = obtener_ruta_archivo('estructuras.csv')
ruta_mineralizacion = obtener_ruta_archivo('mineralizacion.csv')
ruta_minzone = obtener_ruta_archivo('minzone.csv')
ruta_pisotecho = obtener_ruta_archivo('pisotecho.csv')
ruta_vetillas = obtener_ruta_archivo('vetillas.csv')
ruta_UG = obtener_ruta_archivo('ug.csv')


# Intentar cargar los archivos
try:
    df_litologia = pd.read_csv(ruta_lit, encoding='UTF-8', sep=',', header=0)
    print("Archivo litologia.csv cargado exitosamente.")
    print(df_litologia.head())    
except FileNotFoundError:
    print(f"Ruta calculada para litologia.csv: {ruta_lit}")

try:
    df_volalt = pd.read_csv(ruta_volalt, encoding='UTF-8', sep=',', header=0)
    print("Archivo volalt.csv cargado exitosamente.")
    print(df_volalt.head())    
except FileNotFoundError:
    print(f"Ruta calculada para el volalt.csv: {ruta_volalt}")

try:
    df_alteracion = pd.read_csv(ruta_alteracion, encoding='UTF-8', sep=',', header=0)
    print("Archivo alteracion.csv cargado exitosamente.")
    print(df_alteracion.head())    
except FileNotFoundError:
    print(f"Ruta calculada para alteracion.csv: {ruta_alteracion}")

try:
    df_estructuras = pd.read_csv(ruta_estructuras, encoding='UTF-8', sep=',', header=0)
    print("Archivo estructuras.csv cargado exitosamente.")
    print(df_estructuras.head())    
except FileNotFoundError:
    print(f"Ruta calculada para estructuras.csv: {ruta_estructuras}")

try:
    df_mineralizacion = pd.read_csv(ruta_mineralizacion, encoding='UTF-8', sep=',', header=0)
    print("Archivo mineralizacion.csv cargado exitosamente.")
    print(df_mineralizacion.head())    
except FileNotFoundError:
    print(f"Ruta calculada para mineralizacion.csv: {ruta_mineralizacion}")

try:
    df_minzone = pd.read_csv(ruta_minzone, encoding='UTF-8', sep=',', header=0)
    print("Archivo minzone.csv cargado exitosamente.")
    print(df_minzone.head())    
except FileNotFoundError:
    print(f"Ruta calculada para minzone.csv: {ruta_minzone}")

try:
    df_pisotecho = pd.read_csv(ruta_pisotecho, encoding='UTF-8', sep=',', header=0)
    print("Archivo pisotecho.csv cargado exitosamente.")
    print(df_pisotecho.head())    
except FileNotFoundError:
    print(f"Ruta calculada para pisotecho.csv: {ruta_pisotecho}")

try:
    df_vetillas = pd.read_csv(ruta_vetillas, encoding='UTF-8', sep=',', header=0)
    print("Archivo vetillas.csv cargado exitosamente.")
    print(df_vetillas.head())    
except FileNotFoundError:
    print(f"Ruta calculada para vetillas.csv: {df_vetillas}")

try:
    df_UG = pd.read_csv(ruta_UG, encoding='UTF-8', sep=',', header=0)
    print("Archivo ug.csv cargado exitosamente.")
    print(df_UG.head())    
except FileNotFoundError:
    print(f"Ruta calculada para ug.csv: {ruta_UG}")

#Conecto con la aplicación de acQuire a través de pywinauto
app = Application(backend='win32').connect(title_re=".*acQuire 4.*")
# Instancio la ventana principal del programa
window = app.top_window()
# Maximizo la ventana de acQuire
window.maximize()
#-------------------------------------------------------------
#La ventana collar debe  estar seleccionada
pyautogui.press('f8') #---> selecciono hoja litología
ingresar_litologia(df_litologia)
#-------------------------------------------------------------
pyautogui.press('f8') #---> selecciono hoja volumen de alteracion
ingresar_volalt(df_volalt)
#-------------------------------------------------------------
pyautogui.press('f8') #---> selecciono hoja alteracion
ingresar_alteracion(df_alteracion)
#-------------------------------------------------------------
pyautogui.press('f8') #---> selecciono hoja estructuras
ingresar_estructuras(df_estructuras)
#-------------------------------------------------------------
pyautogui.press('f8') #---> selecciono hoja mineralizacion
ingresar_mineralizacion(df_mineralizacion)
#-------------------------------------------------------------
pyautogui.press('f8') #---> selecciono hoja minzone
ingresar_minzone(df_minzone)
#-------------------------------------------------------------
pyautogui.press('f8') #---> selecciono hoja piso/techo
ingresar_pisotecho(df_pisotecho)
#-------------------------------------------------------------
pyautogui.press('f8') #---> selecciono hoja vetillas
ingresar_vetillas(df_vetillas)
#-------------------------------------------------------------
pyautogui.press('f8') #---> selecciono hoja unidad geologica
ingresar_ug(df_UG)