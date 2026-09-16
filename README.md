# Ghostly Analyst V5 — Motor Excel sin IA

Sitio web que procesa archivos Excel (`.xlsx`) aplicando 18 funciones comunes
(resúmenes, conteos, márgenes, limpieza de datos, gráficos, etc.), todo con
fórmulas reales de Excel — **sin usar ningún modelo de IA**. El usuario elige
la función y las columnas desde un formulario simple; no hace falta escribir
instrucciones en lenguaje natural.

## Cómo funciona

1. El usuario sube un archivo `.xlsx` en el sitio.
2. El navegador lee las hojas del archivo al instante (con SheetJS) y llena
   solo el menú de "Hoja a usar" — sin escribir nada a mano.
3. El usuario elige qué función quiere (del menú desplegable) y llena las
   columnas que esa función necesita.
4. El servidor (Flask) recibe esos datos y llama al motor
   (`motor_excel_sin_ia.py`), que procesa el archivo con `pandas` y
   `openpyxl` y agrega una hoja nueva con el resultado (fórmulas reales,
   no valores fijos).
5. El usuario descarga el Excel ya procesado.

## Estructura del proyecto

```
excel agente/
├── app.py                    # Servidor Flask: rutas, subida de archivos
├── motor_excel_sin_ia.py     # El motor: las 18 funciones + el despachador
├── templates/index.html      # Formulario web (diseño + JS del menú dinámico)
├── requirements.txt          # Dependencias de Python
├── Procfile                  # Comando de arranque para Render
├── .gitignore                # Archivos que no se suben al repositorio
├── ghostly_guia.pdf          # Guía descargable (pendiente de actualizar)
└── uploads/                  # Carpeta temporal, se limpia sola tras cada uso
```

## Las 18 funciones disponibles

| Función | Qué hace |
|---|---|
| Resumen por categoría | Suma valores agrupados por categoría (fórmula `SUMIF`) |
| Búsqueda de valores | Busca valores y trae un resultado asociado (dentro de la misma hoja o cruzando dos hojas) |
| Eliminar duplicados | Quita filas repetidas según las columnas que elijas (con respaldo automático) |
| Formato condicional | Pinta celdas de verde/rojo según umbrales |
| Crear gráfico | Inserta un gráfico de barras, líneas o pastel sobre datos ya resumidos |
| Conteo de valores | Cuenta ocurrencias por valor (fórmula `COUNTIF`) |
| Limpieza de texto | Quita espacios sobrantes al inicio/final de cada celda (con respaldo automático) |
| Limpieza profunda | Además, unifica espacios dobles en medio y puede capitalizar el texto |
| Cálculo de FCR | Fórmula de conversión alimenticia (alimento / biomasa), agrupado por lote |
| Filtrar y ordenar | Filtra por un valor y ordena por otra columna |
| Consolidar varias hojas | Une dos o más hojas en una sola, marcando el origen de cada fila |
| Cálculo de margen | Margen $ y % entre costo y venta |
| Alertar vencimientos | Pinta en rojo (vencido) o amarillo (por vencer) según fechas |
| Operación entre columnas | Suma, resta, multiplica o divide dos columnas fila por fila |
| Promedio por categoría | Promedio agrupado (fórmula `AVERAGEIF`) |
| Aplicar descuento | Calcula precio final con % fijo o variable por fila |
| Máximo / Mínimo por categoría | Valores extremos agrupados (`MAXIFS`/`MINIFS`) |
| Porcentaje del total | Qué % representa cada categoría del total general |

> **Nota:** las hojas de resultado se crean con fórmulas de Excel reales
> (no números fijos). Al abrir el archivo en Excel/Google Sheets, verás el
> cálculo automáticamente; si lo lees por código sin abrirlo antes, las
> celdas de fórmula no traen el valor cacheado.

## Correr localmente

```bash
pip install -r requirements.txt
python app.py
```
Abre `http://localhost:5000` en tu navegador.

## Desplegar en Render

1. Sube esta carpeta a un repositorio de GitHub.
2. En [render.com](https://render.com) → "New +" → "Web Service" → conecta el repositorio.
3. Build Command: `pip install -r requirements.txt`
4. Start Command: se detecta solo por el `Procfile` (`gunicorn app:app --bind 0.0.0.0:$PORT`)
5. Plan Free funciona para pruebas — se "duerme" tras inactividad y tarda unos
   segundos en despertar en la siguiente visita.

## Limitaciones conocidas

- Solo acepta `.xlsx` (Excel moderno). Los `.xls` viejos deben convertirse primero.
- Límite de 15 MB por archivo subido.
- Sin autenticación: cualquiera con el link puede usar el sitio.
- No probado aún con archivos de miles de filas (sí con datasets pequeños/medianos).
- La guía en PDF (`ghostly_guia.pdf`) puede tener contenido desactualizado del
  flujo anterior basado en IA; pendiente de revisión.

## Historial

Este proyecto reemplaza una versión anterior (`agente_excel_v7_final.py`) que
interpretaba instrucciones en lenguaje natural usando un modelo de IA local
(Ollama). Esa dependencia se quitó por completo: ahora el usuario elige la
función directamente de un menú, y el motor ejecuta la lógica sin ningún
modelo de por medio.
