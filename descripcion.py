import os
from pywinauto import Desktop

# Obtener el escritorio del usuario
desktop = os.path.join(os.path.join(os.environ['USERPROFILE']), 'Desktop')

# Ruta del archivo de texto
output_file = os.path.join(desktop, "ventanas_abiertas.txt")

# Abrir archivo en modo escritura
with open(output_file, "w", encoding="utf-8") as file:
    file.write("Ventanas abiertas en el sistema:\n")
    file.write("=" * 50 + "\n")
    
    # Iterar sobre todas las ventanas visibles
    for window in Desktop(backend="win32").windows():
        try:
            # Obtener información de la ventana
            title = window.window_text()
            class_name = window.class_name()
            pid = window.process_id()

            # Escribir información en el archivo
            file.write(f"Título: {title}\n")
            file.write(f"Clase: {class_name}\n")
            file.write(f"PID: {pid}\n")
            file.write("-" * 50 + "\n")
        except Exception as e:
            # Manejar errores, si alguna ventana no puede ser accedida
            file.write(f"Error al acceder a una ventana: {str(e)}\n")
            file.write("-" * 50 + "\n")

print(f"Archivo creado exitosamente en: {output_file}")