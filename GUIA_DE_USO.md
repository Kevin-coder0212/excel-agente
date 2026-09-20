# Guía de uso — Ghostly Analyst V5

Sitio web: **https://excel-agente.onrender.com**

Esta guía explica cómo usar el sitio paso a paso, y qué hace cada una de las
18 funciones disponibles. No necesitas saber programar ni escribir ninguna
instrucción especial — todo se hace eligiendo opciones de un menú.

> ⏳ **Nota:** el sitio corre en un plan gratis. Si nadie lo visita por un
> rato, "se duerme" y la primera visita después de eso puede tardar unos
> 20-30 segundos en cargar. Es normal, no está roto.

---

## 1. Cómo usar el sitio (paso a paso)

1. **Entra al sitio** y sube tu archivo Excel (`.xlsx`) arrastrándolo al
   recuadro punteado, o haciendo clic para buscarlo en tu computadora.
   - Solo se aceptan archivos `.xlsx` (Excel moderno). Si el tuyo es `.xls`
     (formato viejo), ábrelo en Excel y guárdalo como `.xlsx` primero
     (Archivo → Guardar como → Libro de Excel).
2. **Elige la hoja** que quieres usar — el menú se llena solo con los
   nombres reales de las hojas de tu archivo, no hace falta escribirlo.
3. **Elige qué quieres hacer** del menú "¿Qué quieres hacer con tus datos?"
   (ver la lista completa de funciones más abajo).
4. **Llena los campos que aparezcan** — cada función pide solo la
   información que necesita (por ejemplo, qué columna usar).
   - Puedes escribir el **nombre** de la columna tal como aparece en tu
     Excel (ej. `Ventas`) o el **número** de columna (ej. `5` si es la
     quinta columna) — ambos funcionan igual.
5. Dale clic a **"Procesar Datos"** y espera unos segundos.
6. El sitio te va a **descargar automáticamente** tu mismo Excel, pero con
   una hoja nueva agregada con el resultado.

### Sobre los resultados

Los resultados se calculan con **fórmulas reales de Excel** (como
`SUMIF`, `COUNTIF`, `PROMEDIO`), no con números fijos. Esto quiere decir:

- Al abrir el archivo descargado en Excel o Google Sheets, vas a ver el
  resultado calculado automáticamente.
- Si cambias un dato en la hoja original después, el resultado de la hoja
  nueva se va a actualizar solo (porque es una fórmula, no un número
  pegado).
- Si abres el archivo con algún programa que no calcule fórmulas
  automáticamente, es posible que veas la fórmula en vez del número —
  simplemente ábrelo en Excel una vez y se va a calcular solo.

---

## 2. Las 18 funciones, explicadas

### 📊 Resumen por categoría
Suma los valores de una columna, agrupados por categoría. Ejemplo: total
de ventas por producto.
- **Pide:** columna a agrupar, columna con los valores a sumar.

### 🔍 Búsqueda de valores
Busca uno o varios valores en una columna y trae el dato de otra columna.
Puede buscar dentro de la misma hoja, o cruzar información entre dos hojas
distintas del mismo archivo.
- **Pide:** columna donde buscar, columna a traer como resultado, los
  valores a buscar (separados por coma), y opcionalmente el nombre de otra
  hoja si el resultado está en otro lado.

### 🗑️ Eliminar duplicados
Quita filas repetidas, comparando las columnas que elijas. Antes de borrar
nada, guarda automáticamente una copia de seguridad del archivo original.
- **Pide:** las columnas a comparar (separadas por coma).
- ⚠️ Compara el texto exacto — si dos filas dicen "Mouse" y "Mouse " (con
  un espacio de más), no las va a considerar iguales. Usa primero
  "Limpieza de texto" o "Limpieza profunda" si sospechas que tus datos
  tienen espacios de más.

### 🎨 Formato condicional
Pinta de verde las celdas que superen un valor alto, y de rojo las que
estén por debajo de un valor bajo.
- **Pide:** la columna con los valores, el umbral alto y el umbral bajo.

### 📈 Crear gráfico
Inserta un gráfico (de barras, líneas o pastel) sobre datos que ya estén
resumidos en una hoja (por ejemplo, después de usar "Resumen por
categoría").
- **Pide:** el nombre de la hoja con los datos ya resumidos, y el tipo de
  gráfico.

### 🔢 Conteo de valores
Cuenta cuántas veces aparece cada valor distinto en una columna.
- **Pide:** la columna a contar.

### 🧹 Limpieza de texto
Quita los espacios sobrantes al inicio y al final de cada celda de texto.
Guarda una copia de seguridad antes de modificar nada.
- **Pide:** las columnas a limpiar (separadas por coma).

### 🧼 Limpieza profunda
Como la anterior, pero más agresiva: además de quitar espacios de los
bordes, une los espacios dobles/triples que estén en medio del texto, y
opcionalmente pone Mayúscula Inicial En Cada Palabra (útil para que
"maria lopez" y "Maria Lopez" queden exactamente iguales).
- **Pide:** las columnas a limpiar, y si quieres capitalizar (sí/no).

### 🦐 Cálculo de FCR
Calcula la conversión alimenticia (alimento total ÷ biomasa ganada) por
lote — típico en acuacultura/avicultura, pero puedes reusar las columnas
para cualquier cálculo con la misma lógica (algo repartido entre grupos).
- **Pide:** columna de alimento, columna de biomasa, columna de lote.

### 🔀 Filtrar y ordenar
Filtra las filas que cumplan un valor específico, y las ordena por otra
columna (de mayor a menor o al revés).
- **Pide:** columna y valor del filtro (opcional), columna para ordenar,
  y el orden (ascendente/descendente).

### 🔗 Consolidar varias hojas
Une dos o más hojas del mismo archivo en una sola, marcando de qué hoja
vino cada fila.
- **Pide:** los nombres de las hojas a unir, separados por coma.

### 💰 Cálculo de margen
Calcula el margen en dinero y en porcentaje entre una columna de costo y
una de venta.
- **Pide:** columna de costo, columna de venta.

### ⏰ Alertar vencimientos
Pinta de rojo las fechas que ya pasaron, y de amarillo las que están
próximas a vencer (dentro de los días que elijas).
- **Pide:** columna con las fechas, días de anticipación para la alerta.

### ➕ Operación entre columnas
Suma, resta, multiplica o divide dos columnas, fila por fila.
- **Pide:** las dos columnas, y el tipo de operación.

### 📐 Promedio por categoría
Calcula el promedio de una columna, agrupado por categoría.
- **Pide:** columna a agrupar, columna con los valores.

### 🏷️ Aplicar descuento
Calcula el precio final después de un descuento. Puedes usar un porcentaje
fijo para todas las filas, o un porcentaje distinto por fila (si tienes
esa columna en tu Excel).
- **Pide:** columna de precio, y el porcentaje (fijo o por columna).

### ⬆️⬇️ Máximo / Mínimo por categoría
Encuentra el valor más alto y/o más bajo de cada categoría.
- **Pide:** columna a agrupar, columna con los valores, y si quieres
  máximo, mínimo, o ambos.

### 🥧 Porcentaje del total
Calcula qué porcentaje representa cada categoría sobre el total general.
- **Pide:** columna a agrupar, columna con los valores.

---

## 3. Preguntas frecuentes

**¿Necesito que mi Excel tenga los encabezados en la primera fila?**
No necesariamente — el sitio detecta solo en qué fila están los
encabezados, aunque tengas títulos o filas vacías arriba.

**¿Puedo procesar la misma función varias veces sobre el mismo archivo?**
Sí. Cada vez que procesas algo, se agrega una hoja nueva con el resultado
(por ejemplo `ResumenAuto`, `ConteoAuto`) sin borrar las anteriores. Si
vuelves a correr la misma función, esa hoja se reemplaza con el resultado
más reciente.

**Me salió un mensaje de error, ¿qué hago?**
El sitio ahora muestra el motivo real del error (por ejemplo, "falta la
columna X" o "la hoja Y no existe"). Revisa que hayas llenado todos los
campos obligatorios y que el nombre de la hoja/columna sea el correcto.

**¿Mis archivos quedan guardados en el servidor?**
No. El archivo se borra automáticamente del servidor justo después de
procesarlo — no se queda almacenado en ningún lado.

**¿Hay límite de tamaño de archivo?**
Sí, 15 MB por archivo.
