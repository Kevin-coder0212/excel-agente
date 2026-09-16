import os
import uuid
import pandas as pd
from flask import Flask, render_template, request, send_file
from werkzeug.utils import secure_filename
import motor_excel_sin_ia as motor

app = Flask(__name__)
UPLOAD_FOLDER = 'uploads'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

# Límite de tamaño de archivo: 15 MB. Evita que alguien tumbe el servidor
# subiendo un archivo gigante.
app.config['MAX_CONTENT_LENGTH'] = 15 * 1024 * 1024


@app.route('/')
def index():
    return render_template('index.html')


@app.route('/download-template')
def download_template():
    file_path = "plantilla_ejemplo.xlsx"
    if not os.path.exists(file_path):
        df_sample = pd.DataFrame({
            "ID": [1, 2, 3],
            "Producto": ["Laptop", "Mouse", "Teclado"],
            "Ventas": [1200, 45, 85]
        })
        df_sample.to_excel(file_path, index=False)
    return send_file(file_path, as_attachment=True)


@app.route('/download-guide')
def download_guide():
    file_path = "ghostly_guia.pdf"
    if os.path.exists(file_path):
        return send_file(file_path, as_attachment=True)
    else:
        return "El archivo PDF aún no se ha subido.", 404


@app.route('/upload', methods=['POST'])
def upload_file():
    if 'file' not in request.files:
        return "No se encontró ningún archivo", 400

    file = request.files['file']
    funcion = request.form.get('funcion')          # <-- viene del <select> del menú desplegable
    hoja = request.form.get('hoja', 'Ventas')       # <-- nombre de la hoja a usar

    if file.filename == '':
        return "Archivo no seleccionado", 400
    if not funcion:
        return "No seleccionaste ninguna función", 400
    if not file.filename.lower().endswith('.xlsx'):
        return "Solo se aceptan archivos .xlsx (Excel moderno). Si tu archivo es .xls, ábrelo en Excel y guárdalo como .xlsx primero.", 400

    # Nombre único por subida, para que dos personas usando el sitio al
    # mismo tiempo no se pisen entre sí con archivos del mismo nombre.
    nombre_seguro = secure_filename(file.filename)
    nombre_unico = f"{uuid.uuid4().hex}_{nombre_seguro}"
    filepath = os.path.join(UPLOAD_FOLDER, nombre_unico)
    file.save(filepath)

    try:
        archivo_listo = motor.procesar_archivo(filepath, funcion, hoja, request.form)
        return send_file(archivo_listo, as_attachment=True, download_name=nombre_seguro)
    except Exception as e:
        return f"Error procesando el archivo: {str(e)}", 500
    finally:
        # Limpieza: borramos el archivo del servidor después de responder,
        # para no ir llenando el disco con cada subida.
        try:
            if os.path.exists(filepath):
                os.remove(filepath)
        except OSError:
            pass


if __name__ == '__main__':
    # debug=True solo para cuando lo corres tú mismo en tu laptop/Colab.
    # En Render, gunicorn arranca la app sin pasar por aquí, así que
    # este modo debug nunca se activa en producción.
    app.run(debug=True)
