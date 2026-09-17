# Lanzador en Google Colab (uso temporal / pruebas)

Este script permite correr el sitio (`app.py` + `motor_excel_sin_ia.py`) desde
un notebook de Google Colab, expuesto al público con un túnel temporal de
[Localtunnel](https://github.com/localtunnel/localtunnel), sin necesidad de
tener nada corriendo en tu laptop.

> **Importante:** el sitio ya tiene un despliegue **permanente** en Render:
> **https://excel-agente.onrender.com** — ese es el que se debe compartir y
> usar normalmente. Este método de Colab es útil solo para:
> - Probar cambios rápidamente antes de subirlos a Render.
> - Tener un respaldo si Render llegara a estar caído.
> - Correr una prueba puntual sin afectar el sitio en producción.

## Cómo usarlo

1. Comprime la carpeta del proyecto (`app.py`, `motor_excel_sin_ia.py`,
   `templates/`, `requirements.txt`, etc.) en un archivo `.zip`. El nombre
   del zip no importa, pero si le pones otro nombre distinto a
   `excel agente.zip`, actualiza la variable `zip_path` al inicio del script.
2. Abre un notebook nuevo en [Google Colab](https://colab.research.google.com).
3. En la primera celda, sube el `.zip` (ícono de carpeta a la izquierda →
   subir archivo, o arrastrarlo directo al panel de archivos).
4. Pega el contenido de `lanzar_en_colab.py` en una celda y ejecútala.
5. El script va a:
   - Descomprimir el zip y encontrar solo la carpeta que tiene `app.py`
     (sin importar cómo se llame la carpeta por dentro).
   - Instalar las dependencias de `requirements.txt`.
   - Mostrar una IP — **cópiala**, es la "contraseña" que Localtunnel te va
     a pedir la primera vez que abras el link en el navegador.
   - Instalar Localtunnel y arrancar `app.py` en segundo plano.
   - Imprimir un link público terminado en `.loca.lt` — ese es el que abres.

## Notas

- La sesión de Colab se desconecta sola tras un rato de inactividad o si
  cierras la pestaña — cuando eso pase, el link deja de funcionar y hay que
  volver a correr todo el proceso desde el paso 3.
- Si algo falla al iniciar Flask, revisa el archivo `flask.log` que se genera
  en la misma carpeta (`!cat flask.log` en una celda nueva de Colab) para ver
  el error real.
- Cada vez que subas una versión nueva del `.zip` (por ejemplo tras agregar
  una función nueva al motor), tienes que volver a correr el script completo
  desde el paso 3 — no hace falta reiniciar todo el notebook, solo repetir la
  subida y ejecución de la celda.
