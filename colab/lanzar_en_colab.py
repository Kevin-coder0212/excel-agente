import os
import time

# 1. Asegúrate de poner el nombre exacto de tu archivo aquí
zip_path = '/content/excel agente.zip'
extract_dir = '/content/mi_agente'

# 2. Descomprimir el archivo zip
print("📦 Descomprimiendo el archivo zip...")
!unzip -o -q "$zip_path" -d "$extract_dir"

# 3. Buscar automáticamente la carpeta que contiene app.py
ruta_app = None
for root, dirs, files in os.walk(extract_dir):
    if 'app.py' in files:
        ruta_app = root
        break

if ruta_app:
    os.chdir(ruta_app)
    print(f"✅ Carpeta de la app encontrada en: {ruta_app}")

    # 4. Instalar dependencias si existe requirements.txt
    if os.path.exists("requirements.txt"):
        print("⚙️ Instalando dependencias...")
        !pip install -q -r requirements.txt

    # 5. Obtener la IP pública limpia (contraseña del túnel)
    print("\n⚠️ IMPORTANTE: Copia este número de IP (te lo pedirá como contraseña en el navegador):")
    !curl -s ipv4.icanhazip.com

    # 6. Instalar LocalTunnel
    print("\n🚇 Instalando túnel...")
    !npm install -g localtunnel > /dev/null 2>&1

    # 7. Arrancar la aplicación en segundo plano
    print("\n🚀 Iniciando tu agente Excel...")
    get_ipython().system_raw('nohup python app.py > flask.log 2>&1 &')
    time.sleep(3)

    # 8. Generar el enlace público
    print("\n🌐 HAZ CLIC EN EL ENLACE QUE TERMINA EN .loca.lt ABAJO 🌐")
    !lt --port 5000

else:
    print("❌ Error: No se encontró app.py dentro del zip. Verifica la estructura.")
