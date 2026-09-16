"""
MOTOR EXCEL — SIN IA (basado en agente_excel_v7_final.py)
=========================================================
Es el mismo motor que ya te funcionaba, pero sin el intérprete de
lenguaje natural (Ollama/Qwen2.5). Aquí tú llamas directamente a la
función que necesitas (o usas ejecutar_funcion con el nombre como
texto) y le pasas los números de columna — nada de IA, nada de
requests a un modelo local.

Sigue siendo GENÉRICO: no está pensado para un Excel específico.
Cualquier archivo .xlsx con encabezados (en cualquier fila, se
detectan solos) puede usar cualquiera de las 17 funciones.

- openpyxl puro (rápido con pocas columnas): Resumen, Conteo, FCR,
  Margen, Formato condicional, Gráfico, Vencimientos, Operación,
  Promedio, Descuento, Máximo/Mínimo, Porcentaje del total.
- pandas + openpyxl (rápido con la fila completa / muchas filas):
  Filtrar/ordenar, Duplicados, Búsqueda, Consolidar, Limpieza.

Las fórmulas reales de Excel (SUMIF/COUNTIF/etc.) se siguen viendo
al hacer clic en la celda — pandas solo acelera el cálculo por
dentro, openpyxl sigue siendo quien escribe lo que tú ves.
"""

import re
import difflib
import unicodedata
from datetime import datetime, timedelta

import pandas as pd
from openpyxl import load_workbook
from openpyxl.chart import BarChart, LineChart, PieChart, Reference
from openpyxl.styles import PatternFill
from openpyxl.utils import get_column_letter

# =====================================================================
# MODO DETALLADO — interruptor aparte del motor
# =====================================================================
# Cuando está en True, cada función imprime sus micro-pasos internos
# (qué columna lee, qué compara, qué encuentra, qué escribe) en vez de
# solo el resumen final. Actívalo/desactívalo con:
#   activar_modo_detallado()
#   desactivar_modo_detallado()
MODO_DETALLADO = False


def activar_modo_detallado():
    global MODO_DETALLADO
    MODO_DETALLADO = True
    print("🔍 Modo detallado ACTIVADO — ahora verás cada paso interno del motor.")


def desactivar_modo_detallado():
    global MODO_DETALLADO
    MODO_DETALLADO = False
    print("🔍 Modo detallado DESACTIVADO — vuelves al resumen normal.")


def _debug(mensaje):
    """Imprime solo si el modo detallado está activo. No afecta nada
    del comportamiento normal cuando está apagado."""
    if MODO_DETALLADO:
        print(f"   🔎 {mensaje}")


# =====================================================================
# UTILIDADES BASE
# =====================================================================

def normalizar(texto):
    texto = str(texto).lower()
    texto = "".join(c for c in unicodedata.normalize("NFD", texto) if unicodedata.category(c) != "Mn")
    return texto


def _detectar_fila_encabezado(ws, max_filas_revisar=15):
    """Busca la fila con más celdas de texto no vacías en las primeras filas —
    esa es la fila de encabezados real, sin asumir que es la fila 1."""
    mejor_fila, mejor_cantidad = 1, -1
    limite = min(max_filas_revisar, ws.max_row)
    for fila in range(1, limite + 1):
        cantidad = sum(
            1 for col in range(1, ws.max_column + 1)
            if ws.cell(row=fila, column=col).value not in (None, "")
        )
        if cantidad > mejor_cantidad:
            mejor_cantidad = cantidad
            mejor_fila = fila
    return mejor_fila


def leer_columnas(archivo, hoja="Ventas"):
    wb = load_workbook(archivo)
    if hoja not in wb.sheetnames:
        raise ValueError(f"La hoja '{hoja}' no existe. Hojas disponibles: {wb.sheetnames}")
    ws = wb[hoja]
    fila_enc = _detectar_fila_encabezado(ws)
    columnas = {}
    for col in range(1, ws.max_column + 1):
        nombre = ws.cell(row=fila_enc, column=col).value
        if nombre:
            columnas[col] = nombre
    return columnas


def _fila_encabezado_de(archivo, hoja):
    wb = load_workbook(archivo)
    ws = wb[hoja]
    return _detectar_fila_encabezado(ws)


def _leer_df(archivo, hoja):
    """Lee la hoja completa como DataFrame de pandas — usado solo por
    las funciones que necesitan la fila completa (filtrar, duplicados,
    búsqueda, consolidar, limpieza)."""
    wb = load_workbook(archivo, read_only=True)
    ws = wb[hoja]
    fila_enc = _detectar_fila_encabezado(ws)
    wb.close()
    df = pd.read_excel(archivo, sheet_name=hoja, header=fila_enc - 1)
    return df, fila_enc


def _validar_columna_idx(df, columna, nombre_param="columna"):
    if not isinstance(columna, int) or columna < 1 or columna > len(df.columns):
        raise ValueError(f"{nombre_param} inválida: {columna}. El archivo tiene {len(df.columns)} columnas.")


def _validar_hoja(wb, hoja):
    if hoja not in wb.sheetnames:
        raise ValueError(f"La hoja '{hoja}' no existe. Hojas disponibles: {wb.sheetnames}")


def _validar_columna(ws, columna, nombre_param="columna"):
    if not isinstance(columna, int) or columna < 1 or columna > ws.max_column:
        raise ValueError(f"{nombre_param} inválida: {columna}. El archivo tiene {ws.max_column} columnas.")


def _hoja_salida(wb, nombre):
    if nombre in wb.sheetnames:
        del wb[nombre]
    return wb.create_sheet(nombre)


def crear_respaldo(archivo):
    """Copia el archivo con marca de tiempo antes de una operación que
    modifica datos permanentemente (Duplicados, Limpieza). Se guarda en
    una carpeta 'respaldos_excel' junto al archivo original."""
    import shutil
    import os

    carpeta = os.path.join(os.path.dirname(os.path.abspath(archivo)) or ".", "respaldos_excel")
    os.makedirs(carpeta, exist_ok=True)

    nombre_base = os.path.splitext(os.path.basename(archivo))[0]
    marca_tiempo = datetime.now().strftime("%Y-%m-%d_%H%M%S")
    ruta_respaldo = os.path.join(carpeta, f"{nombre_base}_BACKUP_{marca_tiempo}.xlsx")

    shutil.copy(archivo, ruta_respaldo)
    print(f"🛡️ Respaldo creado antes de modificar datos: {ruta_respaldo}")
    return ruta_respaldo


def _a_numero(valor, default=0):
    if valor is None:
        return default
    if isinstance(valor, (int, float)):
        return valor
    try:
        limpio = re.sub(r"[^\d.\-]", "", str(valor).replace(",", ""))
        return float(limpio) if limpio else default
    except (ValueError, TypeError):
        return default


def _a_fecha(valor):
    if isinstance(valor, datetime):
        return valor
    if valor is None:
        return None
    for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y", "%m/%d/%Y"):
        try:
            return datetime.strptime(str(valor).strip(), fmt)
        except (ValueError, TypeError):
            continue
    return None


# =====================================================================
# LISTA DE FUNCIONES Y SUS PARÁMETROS (sin IA — solo referencia/validación)
# =====================================================================

FUNCIONES_DISPONIBLES = [
    "resumen", "busqueda", "duplicados", "formato", "grafico",
    "conteo", "limpieza", "limpieza_profunda", "fcr", "filtrar", "consolidar", "margen", "vencimientos",
    "operacion", "promedio", "descuento", "maximo_minimo", "porcentaje_total"
]

PARAMS_POR_FUNCION = {
    "resumen": ["columna_agrupar", "columna_valores"],
    "busqueda": ["columna_buscar", "columna_resultado", "valores_a_buscar", "hoja_resultado"],
    "duplicados": ["columnas_verificar"],
    "formato": ["columna_valores", "umbral_alto", "umbral_bajo"],
    "grafico": ["hoja_datos", "tipo_grafico"],
    "conteo": ["columna_contar"],
    "limpieza": ["columnas_limpiar"],
    "limpieza_profunda": ["columnas_limpiar", "capitalizar"],
    "fcr": ["columna_alimento", "columna_biomasa", "columna_lote"],
    "filtrar": ["columna_filtro", "valor_filtro", "columna_ordenar", "orden"],
    "consolidar": ["hojas_a_unir"],
    "margen": ["columna_costo", "columna_venta"],
    "vencimientos": ["columna_fecha", "dias_alerta"],
    "operacion": ["columna_a", "columna_b", "tipo_operacion"],
    "promedio": ["columna_agrupar", "columna_valores"],
    "descuento": ["columna_precio", "columna_porcentaje", "porcentaje_fijo"],
    "maximo_minimo": ["columna_agrupar", "columna_valores", "tipo"],
    "porcentaje_total": ["columna_agrupar", "columna_valores"],
}


def podar_parametros(funcion, parametros):
    permitidos = PARAMS_POR_FUNCION.get(funcion, [])
    return {k: v for k, v in parametros.items() if k in permitidos}


# =====================================================================
# DESPACHADOR — SIN IA
# =====================================================================
# Antes esto lo llenaba un modelo de lenguaje (Ollama) interpretando
# una instrucción en español. Ahora TÚ le dices directamente qué
# función quieres y con qué columnas, por ejemplo:
#
#   ejecutar_funcion("ventas.xlsx", "resumen",
#                     hoja="Ventas",
#                     columna_agrupar=2, columna_valores=5)
#
# Sigue siendo genérico para cualquier archivo: solo necesita que le
# digas la hoja y los números de columna (o el nombre de la columna,
# tal cual está escrito en el Excel — lo reconoce con coincidencia
# aproximada, útil si te equivocas en un acento o una letra).

def ejecutar_funcion(archivo, funcion, hoja="Ventas", **params):
    """Punto de entrada único para las 17 funciones, sin IA.

    archivo   -> ruta al .xlsx
    funcion   -> nombre de la función ("resumen", "busqueda", etc.)
    hoja      -> hoja principal a usar (por defecto "Ventas")
    **params  -> los parámetros propios de esa función (columna_agrupar=2, etc.)
                 Puedes pasar el número de columna O el nombre de la columna
                 como texto; ambos se resuelven automáticamente.
    """
    funcion_norm = normalizar(funcion)
    funcion_final, mejor_score = None, 0
    for opcion in FUNCIONES_DISPONIBLES:
        score = difflib.SequenceMatcher(None, funcion_norm, opcion).ratio()
        if score > mejor_score:
            mejor_score, funcion_final = score, opcion
    if mejor_score < 0.4:
        print(f"❌ No reconozco la función '{funcion}'. Opciones válidas: {FUNCIONES_DISPONIBLES}")
        return

    columnas_map = leer_columnas(archivo, hoja)
    advertencias = []

    def limpiar(v):
        """Si v es el nombre de una columna (texto), lo convierte al número
        de columna correspondiente comparando contra los encabezados reales
        del archivo. Si ya es un número, lo deja igual."""
        if isinstance(v, list):
            v = v[0] if v else None
        if v is None:
            return None
        if isinstance(v, str):
            v_norm = normalizar(v)
            mejor_num, mejor_sim = None, 0
            for num, nombre in columnas_map.items():
                sim = difflib.SequenceMatcher(None, v_norm, normalizar(nombre)).ratio()
                if sim > mejor_sim:
                    mejor_sim, mejor_num = sim, num
            if mejor_sim >= 0.4:
                if mejor_sim < 0.7:
                    advertencias.append(
                        f"⚠️ Interpreté '{v}' como la columna '{columnas_map.get(mejor_num)}' con {round(mejor_sim*100)}% de similitud — verifica que sea correcto.")
                return mejor_num
            try:
                return int(v)
            except (ValueError, TypeError):
                return v
        return v

    def param(nombre, default=None, requerido=False):
        valor = params.get(nombre, default)
        valor = limpiar(valor) if not isinstance(valor, list) else [limpiar(x) if isinstance(x, str) and x not in params.get("_no_columna", []) else x for x in valor]
        if valor is None and requerido:
            raise KeyError(f"Falta el parámetro obligatorio '{nombre}' para la función '{funcion_final}'.")
        return valor

    def _nombre_col(num):
        return columnas_map.get(num, f"columna {num}") if isinstance(num, int) else num

    resumen_params = []
    for k, v in params.items():
        if isinstance(v, int) and v in columnas_map:
            resumen_params.append(f"{k} = '{columnas_map.get(v)}'")
        else:
            resumen_params.append(f"{k} = {v}")
    print(f"⚙️ Ejecutando función '{funcion_final}' con {', '.join(resumen_params) if resumen_params else 'sin parámetros adicionales'}.")
    if mejor_score < 0.7:
        print(f"⚠️ El nombre de función que diste se parece un {round(mejor_score*100)}% a '{funcion_final}'. Si no era lo que querías, revisa el nombre.")

    try:
        if funcion_final == "resumen":
            crear_resumen(archivo, param("columna_agrupar", requerido=True), param("columna_valores", requerido=True), hoja, "ResumenAuto")
        elif funcion_final == "busqueda":
            hoja_res = params.get("hoja_resultado")
            crear_busqueda(archivo, hoja, param("columna_buscar", requerido=True), param("columna_resultado", requerido=True),
                            params.get("valores_a_buscar"), "BusquedaAuto", hoja_res)
        elif funcion_final == "duplicados":
            crear_respaldo(archivo)
            eliminar_duplicados(archivo, hoja, param("columnas_verificar"))
        elif funcion_final == "formato":
            aplicar_formato_condicional(archivo, hoja, param("columna_valores", requerido=True),
                                         params.get("umbral_alto", 9999999), params.get("umbral_bajo", -9999999))
        elif funcion_final == "grafico":
            crear_grafico(archivo, params.get("hoja_datos"), params.get("tipo_grafico", "barras"), "Gráfico automático")
        elif funcion_final == "conteo":
            crear_conteo(archivo, hoja, param("columna_contar", requerido=True), "ConteoAuto")
        elif funcion_final == "limpieza":
            crear_respaldo(archivo)
            limpiar_texto(archivo, hoja, param("columnas_limpiar"))
        elif funcion_final == "limpieza_profunda":
            crear_respaldo(archivo)
            texto_cap = str(params.get("capitalizar", "")).strip().lower()
            capitalizar = texto_cap in ("si", "sí", "true", "1", "yes")
            limpieza_profunda(archivo, hoja, param("columnas_limpiar"), capitalizar)
        elif funcion_final == "fcr":
            calcular_fcr(archivo, hoja, param("columna_alimento", requerido=True), param("columna_biomasa", requerido=True),
                         param("columna_lote", requerido=True), "FCR_Auto")
        elif funcion_final == "filtrar":
            filtrar_ordenar(archivo, hoja, params.get("columna_filtro"), params.get("valor_filtro"),
                             param("columna_ordenar", requerido=True), params.get("orden", "desc"), "FiltradoAuto")
        elif funcion_final == "consolidar":
            consolidar_hojas(archivo, params.get("hojas_a_unir"), "ConsolidadoAuto")
        elif funcion_final == "margen":
            calcular_margen(archivo, hoja, param("columna_costo", requerido=True), param("columna_venta", requerido=True), "MargenAuto")
        elif funcion_final == "vencimientos":
            dias = params.get("dias_alerta")
            if dias is None:
                dias = 30
                advertencias.append("ℹ️ No especificaste 'dias_alerta' — usé 30 días por defecto.")
            alertar_vencimientos(archivo, hoja, param("columna_fecha", requerido=True), dias)
        elif funcion_final == "operacion":
            calcular_operacion(archivo, hoja, param("columna_a", requerido=True), param("columna_b", requerido=True),
                                params.get("tipo_operacion", "suma"), "OperacionAuto")
        elif funcion_final == "promedio":
            calcular_promedio(archivo, hoja, param("columna_agrupar", requerido=True), param("columna_valores", requerido=True), "PromedioAuto")
        elif funcion_final == "descuento":
            col_pct = param("columna_porcentaje")
            pct_fijo = params.get("porcentaje_fijo")
            if col_pct is None and pct_fijo is None:
                pct_fijo = 10
                advertencias.append("ℹ️ No especificaste el % de descuento — usé 10% por defecto.")
            calcular_descuento(archivo, hoja, param("columna_precio", requerido=True), col_pct, pct_fijo, "DescuentoAuto")
        elif funcion_final == "maximo_minimo":
            calcular_maximo_minimo(archivo, hoja, param("columna_agrupar", requerido=True), param("columna_valores", requerido=True),
                                    params.get("tipo", "ambos"), "MaxMinAuto")
        elif funcion_final == "porcentaje_total":
            calcular_porcentaje_total(archivo, hoja, param("columna_agrupar", requerido=True), param("columna_valores", requerido=True), "PorcentajeAuto")
        else:
            print(f"Función '{funcion}' no reconocida.")
    except Exception as e:
        print(f"❌ Error al ejecutar la función: {e}")
        print("   Esto es un error técnico (columna/hoja inválida, dato faltante) — revisa el mensaje de arriba.")
        raise

    if advertencias:
        print("\n📋 Puntos a verificar (no son errores, pero conviene que los revises):")
        for a in advertencias:
            print("   " + a)
    else:
        print("\n✅ Sin ambigüedades detectadas.")


# =====================================================================
# PUENTE CON LA WEB (Flask) — SIN IA
# =====================================================================
# El formulario del sitio manda campos de texto (todo lo que viene de
# un <form> HTML es texto). Esta función traduce esos campos de texto
# a los tipos correctos (números, listas) y llama a ejecutar_funcion.
#
# Se usa así desde app.py:
#   motor.procesar_archivo(filepath, funcion, hoja, request.form)

CAMPOS_LISTA = {"columnas_verificar", "columnas_limpiar", "hojas_a_unir", "valores_a_buscar"}
CAMPOS_NUMERICOS = {"umbral_alto", "umbral_bajo", "porcentaje_fijo", "dias_alerta"}


def procesar_archivo(archivo, funcion, hoja, form):
    """Convierte los campos de texto del formulario web en los parámetros
    que necesita ejecutar_funcion, y ejecuta la función elegida.

    archivo -> ruta al .xlsx ya guardado en el servidor
    funcion -> valor del <select> del formulario (ej. "resumen")
    hoja    -> nombre de la hoja escrito/elegido en el formulario
    form    -> el objeto de campos del formulario (request.form en Flask,
               o cualquier diccionario con .get())
    """
    campos_de_esta_funcion = PARAMS_POR_FUNCION.get(funcion, [])
    params = {}

    for campo in campos_de_esta_funcion:
        valor_crudo = form.get(campo)
        if valor_crudo is None or str(valor_crudo).strip() == "":
            continue

        if campo in CAMPOS_LISTA:
            # Ej: "1,3,4" -> [1, 3, 4]  |  "Ventas,Compras" -> ["Ventas", "Compras"]
            partes = [v.strip() for v in str(valor_crudo).split(",") if v.strip()]
            valor = [int(v) if v.isdigit() else v for v in partes]
        elif campo in CAMPOS_NUMERICOS:
            try:
                valor = float(valor_crudo)
                if valor.is_integer():
                    valor = int(valor)
            except (ValueError, TypeError):
                valor = valor_crudo
        else:
            # Columnas: puede venir como número ("3") o nombre ("Producto") —
            # ejecutar_funcion ya sabe resolver ambos casos.
            try:
                valor = int(valor_crudo)
            except (ValueError, TypeError):
                valor = valor_crudo

        params[campo] = valor

    ejecutar_funcion(archivo, funcion, hoja=hoja, **params)
    return archivo


# =====================================================================
# FUNCIONES DEL MOTOR (1/17 y 2/17) — con detección automática de encabezado
# =====================================================================

def crear_resumen(archivo, columna_agrupar, columna_valores, hoja="Ventas", nombre_hoja_salida="ResumenAuto"):
    _debug(f"Abriendo archivo '{archivo}', hoja '{hoja}'...")
    wb = load_workbook(archivo)
    _validar_hoja(wb, hoja)
    ws = wb[hoja]
    _validar_columna(ws, columna_agrupar, "columna_agrupar")
    _validar_columna(ws, columna_valores, "columna_valores")
    fila_enc = _detectar_fila_encabezado(ws)
    _debug(f"Encabezados detectados en fila {fila_enc}")

    col_agrupar_letra = get_column_letter(columna_agrupar)
    col_valores_letra = get_column_letter(columna_valores)
    ultima_fila = ws.max_row
    nombre_col_agrupar = ws.cell(row=fila_enc, column=columna_agrupar).value
    nombre_col_valores = ws.cell(row=fila_enc, column=columna_valores).value
    _debug(f"Leyendo columna '{nombre_col_agrupar}' (columna {columna_agrupar}) para agrupar...")
    _debug(f"Leyendo columna '{nombre_col_valores}' (columna {columna_valores}) para sumar...")

    categorias, vistos = [], set()
    for fila in range(fila_enc + 1, ultima_fila + 1):
        clave = ws.cell(row=fila, column=columna_agrupar).value
        if clave is not None and clave not in vistos:
            vistos.add(clave)
            categorias.append(clave)
    _debug(f"Categorías únicas encontradas: {categorias}")

    salida = _hoja_salida(wb, nombre_hoja_salida)
    salida.append(["Categoría", "Total (fórmula)"])
    for i, categoria in enumerate(categorias, start=2):
        salida.cell(row=i, column=1, value=categoria)
        formula = (f"=SUMIF('{hoja}'!{col_agrupar_letra}{fila_enc+1}:{col_agrupar_letra}{ultima_fila},"
                   f"A{i},'{hoja}'!{col_valores_letra}{fila_enc+1}:{col_valores_letra}{ultima_fila})")
        salida.cell(row=i, column=2, value=formula)
        _debug(f"Fila {i}: '{categoria}' → escribiendo fórmula {formula}")

    _debug(f"Guardando archivo con la hoja '{nombre_hoja_salida}'...")
    wb.save(archivo)
    print(f"Resumen creado con fórmulas SUMIF en hoja '{nombre_hoja_salida}' ({len(categorias)} categorías). [Encabezado detectado en fila {fila_enc}]")
    return categorias


def crear_busqueda(archivo, hoja, columna_buscar, columna_resultado, valores_a_buscar, nombre_hoja_salida="BusquedaAuto", hoja_resultado=None):
    """Busca valores en 'hoja' (columna_buscar) y trae el dato de columna_resultado.
    Si hoja_resultado se especifica y es distinta de 'hoja', cruza entre dos
    hojas distintas (tipo BUSCARV real) usando columna_buscar como llave
    común en ambas hojas."""
    if not valores_a_buscar:
        raise ValueError("valores_a_buscar está vacío.")

    _debug(f"Iniciando búsqueda en hoja '{hoja}', columna {columna_buscar}...")
    df_buscar, _ = _leer_df(archivo, hoja)
    _validar_columna_idx(df_buscar, columna_buscar, "columna_buscar")

    valores_norm = [str(v).strip().lower() for v in valores_a_buscar]
    col_buscar_nombre = df_buscar.columns[columna_buscar - 1]
    _debug(f"Leyendo columna '{col_buscar_nombre}' (columna {columna_buscar})...")
    _debug(f"Comparando cada fila con: {valores_a_buscar}")

    if hoja_resultado and hoja_resultado != hoja:
        # --- Modo BUSCARV real: cruza entre dos hojas usando columna_buscar como llave ---
        _debug(f"Modo cruce entre hojas detectado: '{hoja}' → '{hoja_resultado}'")
        df_result, _ = _leer_df(archivo, hoja_resultado)
        _validar_columna_idx(df_result, columna_resultado, "columna_resultado")
        col_llave_result = df_result.columns[columna_buscar - 1] if columna_buscar <= len(df_result.columns) else df_result.columns[0]
        col_resultado_nombre = df_result.columns[columna_resultado - 1]
        _debug(f"En la hoja '{hoja_resultado}', usando '{col_llave_result}' como llave y trayendo '{col_resultado_nombre}'")

        mask = df_buscar[col_buscar_nombre].astype(str).str.strip().str.lower().isin(valores_norm)
        llaves = df_buscar.loc[mask, col_buscar_nombre]
        _debug(f"Encontradas {len(llaves)} filas que coinciden en '{hoja}'")

        cruce = df_result.set_index(df_result[col_llave_result].astype(str).str.strip().str.lower())[col_resultado_nombre]
        resultados = []
        for llave in llaves:
            llave_norm = str(llave).strip().lower()
            valor_encontrado = cruce.get(llave_norm, "⚠️ No encontrado en la otra hoja")
            _debug(f"'{llave}' → buscando en '{hoja_resultado}'... encontrado: {valor_encontrado}")
            resultados.append([llave, valor_encontrado])

        _debug(f"Escribiendo {len(resultados)} filas en hoja '{nombre_hoja_salida}'...")
        wb = load_workbook(archivo)
        salida = _hoja_salida(wb, nombre_hoja_salida)
        salida.append(["Coincidencia", f"Resultado (desde '{hoja_resultado}')"])
        for fila in resultados:
            salida.append(fila)
        wb.save(archivo)
        _debug(f"Archivo guardado.")
        print(f"[pandas] Búsqueda cruzada entre hojas '{hoja}' y '{hoja_resultado}': {len(resultados)} coincidencias en hoja '{nombre_hoja_salida}'.")
        return resultados
    else:
        # --- Modo normal: búsqueda dentro de la misma hoja ---
        _validar_columna_idx(df_buscar, columna_resultado, "columna_resultado")
        col_resultado_nombre = df_buscar.columns[columna_resultado - 1]
        _debug(f"Trayendo resultado de la columna '{col_resultado_nombre}' (misma hoja)")
        mask = df_buscar[col_buscar_nombre].astype(str).str.strip().str.lower().isin(valores_norm)
        resultado_df = df_buscar.loc[mask, [col_buscar_nombre, col_resultado_nombre]]
        _debug(f"Encontradas {len(resultado_df)} filas coincidentes")

        _debug(f"Escribiendo resultados en hoja '{nombre_hoja_salida}'...")
        wb = load_workbook(archivo)
        salida = _hoja_salida(wb, nombre_hoja_salida)
        salida.append(["Coincidencia", "Resultado"])
        for _, fila in resultado_df.iterrows():
            salida.append([fila[col_buscar_nombre], fila[col_resultado_nombre]])
        wb.save(archivo)
        _debug(f"Archivo guardado.")
        print(f"[pandas] Búsqueda completada: {len(resultado_df)} coincidencias en hoja '{nombre_hoja_salida}'.")
        return resultado_df


# =====================================================================
# FUNCIONES DEL MOTOR (3/17 y 4/17)
# =====================================================================

def eliminar_duplicados(archivo, hoja, columnas_verificar):
    if not columnas_verificar:
        raise ValueError("columnas_verificar está vacío.")
    _debug(f"Leyendo la hoja '{hoja}' completa con pandas...")
    df, fila_enc = _leer_df(archivo, hoja)
    for col in columnas_verificar:
        _validar_columna_idx(df, col, "columnas_verificar")

    cols_nombres = [df.columns[c - 1] for c in columnas_verificar]
    _debug(f"Verificando duplicados comparando las columnas: {cols_nombres}")
    antes = len(df)
    _debug(f"Total de filas antes de limpiar: {antes}")
    df_sin_dup = df.drop_duplicates(subset=cols_nombres, keep="first")
    eliminadas = antes - len(df_sin_dup)
    _debug(f"Filas duplicadas detectadas y eliminadas: {eliminadas}")

    _debug(f"Guardando la hoja '{hoja}' actualizada, sin duplicados...")
    with pd.ExcelWriter(archivo, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
        df_sin_dup.to_excel(writer, sheet_name=hoja, index=False, startrow=fila_enc - 1)
    print(f"[pandas] Duplicados eliminados: {eliminadas} filas removidas.")
    return int(eliminadas)


def aplicar_formato_condicional(archivo, hoja, columna_valores, umbral_alto, umbral_bajo):
    wb = load_workbook(archivo)
    _validar_hoja(wb, hoja)
    ws = wb[hoja]
    _validar_columna(ws, columna_valores, "columna_valores")
    fila_enc = _detectar_fila_encabezado(ws)
    _debug(f"Umbrales: verde si >= {umbral_alto}, rojo si <= {umbral_bajo}")

    verde = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
    rojo = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")

    contador_alto = contador_bajo = 0
    for fila in range(fila_enc + 1, ws.max_row + 1):
        celda = ws.cell(row=fila, column=columna_valores)
        valor = _a_numero(celda.value, default=None)
        if valor is None:
            continue
        if valor >= umbral_alto:
            celda.fill = verde
            contador_alto += 1
            _debug(f"Fila {fila}: valor {valor} >= {umbral_alto} → pintado VERDE")
        elif valor <= umbral_bajo:
            celda.fill = rojo
            contador_bajo += 1
            _debug(f"Fila {fila}: valor {valor} <= {umbral_bajo} → pintado ROJO")

    _debug("Guardando archivo con los colores aplicados...")
    wb.save(archivo)
    print(f"Formato aplicado: {contador_alto} altos (verde), {contador_bajo} bajos (rojo).")
    return {"altos": contador_alto, "bajos": contador_bajo}


# =====================================================================
# FUNCIONES DEL MOTOR (5/17 y 6/17)
# =====================================================================

def crear_grafico(archivo, hoja_datos, tipo_grafico="barras", titulo="Gráfico automático"):
    wb = load_workbook(archivo)
    _validar_hoja(wb, hoja_datos)
    ws = wb[hoja_datos]
    if ws.max_row < 2:
        raise ValueError(f"La hoja '{hoja_datos}' no tiene suficientes datos para graficar.")
    _debug(f"Tomando datos de la hoja '{hoja_datos}' ({ws.max_row} filas)")
    _debug(f"Tipo de gráfico solicitado: {tipo_grafico}")

    chart = {"barras": BarChart, "lineas": LineChart, "pastel": PieChart}.get(tipo_grafico, BarChart)()
    chart.title = titulo
    if tipo_grafico != "pastel":
        chart.y_axis.title = ws.cell(row=1, column=2).value or "Valor"
        chart.x_axis.title = ws.cell(row=1, column=1).value or "Categoría"

    datos = Reference(ws, min_col=2, min_row=1, max_row=ws.max_row)
    categorias = Reference(ws, min_col=1, min_row=2, max_row=ws.max_row)
    chart.add_data(datos, titles_from_data=True)
    chart.set_categories(categorias)
    _debug("Insertando el gráfico en la hoja...")

    ws.add_chart(chart, f"{get_column_letter(ws.max_column + 2)}2")
    wb.save(archivo)
    print(f"Gráfico ({tipo_grafico}) '{titulo}' agregado a la hoja '{hoja_datos}'.")


def crear_conteo(archivo, hoja, columna_contar, nombre_hoja_salida="ConteoAuto"):
    _debug(f"Abriendo archivo, hoja '{hoja}'...")
    wb = load_workbook(archivo)
    _validar_hoja(wb, hoja)
    ws = wb[hoja]
    _validar_columna(ws, columna_contar, "columna_contar")
    fila_enc = _detectar_fila_encabezado(ws)

    col_letra = get_column_letter(columna_contar)
    ultima_fila = ws.max_row
    _debug(f"Contando valores únicos en columna {columna_contar}...")

    valores, vistos = [], set()
    for fila in range(fila_enc + 1, ultima_fila + 1):
        valor = ws.cell(row=fila, column=columna_contar).value
        if valor is not None and valor not in vistos:
            vistos.add(valor)
            valores.append(valor)
    _debug(f"Valores únicos encontrados: {valores}")

    salida = _hoja_salida(wb, nombre_hoja_salida)
    salida.append(["Valor", "Cantidad (fórmula)"])
    for i, valor in enumerate(valores, start=2):
        salida.cell(row=i, column=1, value=valor)
        formula = f"=COUNTIF('{hoja}'!{col_letra}{fila_enc+1}:{col_letra}{ultima_fila},A{i})"
        salida.cell(row=i, column=2, value=formula)
        _debug(f"'{valor}' → fórmula {formula}")

    wb.save(archivo)
    print(f"Conteo creado con fórmulas COUNTIF en hoja '{nombre_hoja_salida}' ({len(valores)} valores distintos). [Encabezado detectado en fila {fila_enc}]")
    return valores


# =====================================================================
# FUNCIONES DEL MOTOR (7/17 y 8/17)
# =====================================================================

def limpiar_texto(archivo, hoja, columnas_limpiar):
    if not columnas_limpiar:
        raise ValueError("columnas_limpiar está vacío.")
    df, fila_enc = _leer_df(archivo, hoja)
    for col in columnas_limpiar:
        _validar_columna_idx(df, col, "columnas_limpiar")

    celdas_modificadas = 0
    for col in columnas_limpiar:
        col_nombre = df.columns[col - 1]
        original = df[col_nombre].copy()
        df[col_nombre] = df[col_nombre].apply(lambda x: x.strip() if isinstance(x, str) else x)
        celdas_modificadas += int((original != df[col_nombre]).sum())

    with pd.ExcelWriter(archivo, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
        df.to_excel(writer, sheet_name=hoja, index=False, startrow=fila_enc - 1)
    print(f"[pandas] Limpieza completada: {celdas_modificadas} celdas modificadas.")
    return celdas_modificadas


def limpieza_profunda(archivo, hoja, columnas_limpiar, capitalizar=False):
    """Limpieza más agresiva que limpiar_texto: además de quitar espacios
    al inicio/final, también junta los espacios dobles/triples de en medio
    en uno solo, y opcionalmente pone Mayúscula Inicial En Cada Palabra
    (útil para unificar 'maria lopez' y 'Maria Lopez' como el mismo texto)."""
    if not columnas_limpiar:
        raise ValueError("columnas_limpiar está vacío.")
    df, fila_enc = _leer_df(archivo, hoja)
    for col in columnas_limpiar:
        _validar_columna_idx(df, col, "columnas_limpiar")

    def limpiar_valor(x):
        if not isinstance(x, str):
            return x
        texto = re.sub(r"\s+", " ", x).strip()
        if capitalizar:
            texto = texto.title()
        return texto

    celdas_modificadas = 0
    for col in columnas_limpiar:
        col_nombre = df.columns[col - 1]
        original = df[col_nombre].copy()
        df[col_nombre] = df[col_nombre].apply(limpiar_valor)
        celdas_modificadas += int((original != df[col_nombre]).sum())

    with pd.ExcelWriter(archivo, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
        df.to_excel(writer, sheet_name=hoja, index=False, startrow=fila_enc - 1)

    extra = " y texto capitalizado" if capitalizar else ""
    print(f"[pandas] Limpieza profunda completada: {celdas_modificadas} celdas modificadas (espacios internos unificados{extra}).")
    return celdas_modificadas


def calcular_fcr(archivo, hoja, columna_alimento, columna_biomasa, columna_lote, nombre_hoja_salida="FCR_Auto"):
    wb = load_workbook(archivo)
    _validar_hoja(wb, hoja)
    ws = wb[hoja]
    for col, nombre in [(columna_alimento, "columna_alimento"), (columna_biomasa, "columna_biomasa"),
                         (columna_lote, "columna_lote")]:
        _validar_columna(ws, col, nombre)
    fila_enc = _detectar_fila_encabezado(ws)

    col_alimento_letra = get_column_letter(columna_alimento)
    col_biomasa_letra = get_column_letter(columna_biomasa)
    col_lote_letra = get_column_letter(columna_lote)
    ultima_fila = ws.max_row

    lotes, vistos = [], set()
    for fila in range(fila_enc + 1, ultima_fila + 1):
        lote = ws.cell(row=fila, column=columna_lote).value
        if lote is not None and lote not in vistos:
            vistos.add(lote)
            lotes.append(lote)
    _debug(f"Lotes únicos encontrados: {lotes}")

    salida = _hoja_salida(wb, nombre_hoja_salida)
    salida.append(["Lote", "Alimento total (fórmula)", "Biomasa ganada (fórmula)", "FCR (fórmula)"])
    for i, lote in enumerate(lotes, start=2):
        salida.cell(row=i, column=1, value=lote)
        f_alimento = (f"=SUMIF('{hoja}'!{col_lote_letra}{fila_enc+1}:{col_lote_letra}{ultima_fila},"
                      f"A{i},'{hoja}'!{col_alimento_letra}{fila_enc+1}:{col_alimento_letra}{ultima_fila})")
        f_biomasa = (f"=SUMIF('{hoja}'!{col_lote_letra}{fila_enc+1}:{col_lote_letra}{ultima_fila},"
                     f"A{i},'{hoja}'!{col_biomasa_letra}{fila_enc+1}:{col_biomasa_letra}{ultima_fila})")
        salida.cell(row=i, column=2, value=f_alimento)
        salida.cell(row=i, column=3, value=f_biomasa)
        salida.cell(row=i, column=4, value=f"=B{i}/C{i}")
        _debug(f"Lote '{lote}': alimento total = {f_alimento}, FCR = alimento/biomasa")

    wb.save(archivo)
    print(f"FCR calculado con fórmulas SUMIF para {len(lotes)} lotes en hoja '{nombre_hoja_salida}'.")
    return lotes


# =====================================================================
# FUNCIONES DEL MOTOR (9/17 y 10/17)
# =====================================================================

def filtrar_ordenar(archivo, hoja, columna_filtro, valor_filtro, columna_ordenar, orden="desc", nombre_hoja_salida="FiltradoAuto"):
    _debug(f"Leyendo hoja '{hoja}' completa con pandas...")
    df, fila_enc = _leer_df(archivo, hoja)
    _validar_columna_idx(df, columna_ordenar, "columna_ordenar")

    df_trabajo = df.copy()
    if columna_filtro and valor_filtro:
        col_filtro_nombre = df.columns[columna_filtro - 1]
        _debug(f"Filtrando donde '{col_filtro_nombre}' = '{valor_filtro}'...")
        df_trabajo = df_trabajo[df_trabajo[col_filtro_nombre].astype(str).str.strip().str.lower() == str(valor_filtro).strip().lower()]
        _debug(f"Filas que pasaron el filtro: {len(df_trabajo)}")

    col_ordenar_nombre = df.columns[columna_ordenar - 1]
    ascending = (orden != "desc")
    _debug(f"Ordenando por '{col_ordenar_nombre}', {'ascendente' if ascending else 'descendente'}...")
    df_ordenado = df_trabajo.sort_values(col_ordenar_nombre, ascending=ascending, na_position="last")

    wb = load_workbook(archivo)
    if nombre_hoja_salida in wb.sheetnames:
        del wb[nombre_hoja_salida]
    wb.save(archivo)

    with pd.ExcelWriter(archivo, engine="openpyxl", mode="a") as writer:
        df_ordenado.to_excel(writer, sheet_name=nombre_hoja_salida, index=False)
    print(f"[pandas] Filtrado/ordenado: {len(df_ordenado)} filas en hoja '{nombre_hoja_salida}'.")
    return len(df_ordenado)


def consolidar_hojas(archivo, hojas_a_unir, nombre_hoja_salida="ConsolidadoAuto"):
    if not hojas_a_unir or len(hojas_a_unir) < 2:
        raise ValueError("Se necesitan al menos 2 hojas en hojas_a_unir.")
    _debug(f"Voy a unir estas hojas: {hojas_a_unir}")

    dfs = []
    for h in hojas_a_unir:
        _debug(f"Leyendo hoja '{h}'...")
        df_h, _ = _leer_df(archivo, h)
        df_h["Origen"] = h
        _debug(f"'{h}' tiene {len(df_h)} filas")
        dfs.append(df_h)
    df_consolidado = pd.concat(dfs, ignore_index=True)
    _debug(f"Total combinado: {len(df_consolidado)} filas")

    wb = load_workbook(archivo)
    if nombre_hoja_salida in wb.sheetnames:
        del wb[nombre_hoja_salida]
    wb.save(archivo)

    _debug(f"Guardando todo en la hoja '{nombre_hoja_salida}'...")
    with pd.ExcelWriter(archivo, engine="openpyxl", mode="a") as writer:
        df_consolidado.to_excel(writer, sheet_name=nombre_hoja_salida, index=False)
    print(f"[pandas] Consolidado: {len(df_consolidado)} filas de {len(hojas_a_unir)} hojas en '{nombre_hoja_salida}'.")
    return len(df_consolidado)


# =====================================================================
# FUNCIONES DEL MOTOR (11/17 y 12/17)
# =====================================================================

def calcular_margen(archivo, hoja, columna_costo, columna_venta, nombre_hoja_salida="MargenAuto"):
    wb = load_workbook(archivo)
    _validar_hoja(wb, hoja)
    ws = wb[hoja]
    _validar_columna(ws, columna_costo, "columna_costo")
    _validar_columna(ws, columna_venta, "columna_venta")
    fila_enc = _detectar_fila_encabezado(ws)
    _debug(f"Calculando margen: columna_costo={columna_costo}, columna_venta={columna_venta}")

    col_costo_letra = get_column_letter(columna_costo)
    col_venta_letra = get_column_letter(columna_venta)

    salida = _hoja_salida(wb, nombre_hoja_salida)
    salida.append(["Fila origen", "Costo (ref)", "Venta (ref)", "Margen $ (fórmula)", "Margen % (fórmula)"])

    fila_salida = 2
    filas_incluidas = 0
    for fila in range(fila_enc + 1, ws.max_row + 1):
        costo = _a_numero(ws.cell(row=fila, column=columna_costo).value)
        venta = _a_numero(ws.cell(row=fila, column=columna_venta).value)
        if costo == 0 and venta == 0:
            continue

        ref_costo = f"'{hoja}'!{col_costo_letra}{fila}"
        ref_venta = f"'{hoja}'!{col_venta_letra}{fila}"
        _debug(f"Fila {fila}: costo={costo}, venta={venta} → margen = venta-costo")

        salida.cell(row=fila_salida, column=1, value=fila)
        salida.cell(row=fila_salida, column=2, value=f"={ref_costo}")
        salida.cell(row=fila_salida, column=3, value=f"={ref_venta}")
        salida.cell(row=fila_salida, column=4, value=f"=C{fila_salida}-B{fila_salida}")
        salida.cell(row=fila_salida, column=5, value=f"=(C{fila_salida}-B{fila_salida})/C{fila_salida}")

        fila_salida += 1
        filas_incluidas += 1

    wb.save(archivo)
    print(f"Margen calculado con fórmulas para {filas_incluidas} filas en hoja '{nombre_hoja_salida}'.")
    return filas_incluidas


def alertar_vencimientos(archivo, hoja, columna_fecha, dias_alerta=30):
    wb = load_workbook(archivo)
    _validar_hoja(wb, hoja)
    ws = wb[hoja]
    _validar_columna(ws, columna_fecha, "columna_fecha")
    fila_enc = _detectar_fila_encabezado(ws)
    _debug(f"Revisando fechas en columna {columna_fecha}, alerta a {dias_alerta} días")

    rojo = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
    amarillo = PatternFill(start_color="FFEB9C", end_color="FFEB9C", fill_type="solid")
    hoy = datetime.now()
    limite = hoy + timedelta(days=dias_alerta)
    _debug(f"Hoy: {hoy.date()}, límite de alerta: {limite.date()}")

    vencidos = por_vencer = 0
    for fila in range(fila_enc + 1, ws.max_row + 1):
        celda = ws.cell(row=fila, column=columna_fecha)
        fecha = _a_fecha(celda.value)
        if fecha is None:
            continue
        if fecha < hoy:
            celda.fill = rojo
            vencidos += 1
            _debug(f"Fila {fila}: {fecha.date()} ya pasó → ROJO (vencido)")
        elif fecha <= limite:
            celda.fill = amarillo
            por_vencer += 1
            _debug(f"Fila {fila}: {fecha.date()} está próxima → AMARILLO")

    wb.save(archivo)
    print(f"Vencimientos: {vencidos} ya vencidos (rojo), {por_vencer} próximos a {dias_alerta} días (amarillo).")
    return {"vencidos": vencidos, "por_vencer": por_vencer}


# =====================================================================
# FUNCIONES DEL MOTOR (13/17 y 14/17)
# =====================================================================

_SIMBOLOS_OPERACION = {"suma": "+", "resta": "-", "multiplicacion": "*", "division": "/"}
_NOMBRES_OPERACION = {"suma": "Suma", "resta": "Resta", "multiplicacion": "Multiplicación", "division": "División"}


def calcular_operacion(archivo, hoja, columna_a, columna_b, tipo_operacion="suma", nombre_hoja_salida="OperacionAuto"):
    """Suma, resta, multiplica o divide dos columnas cualquiera, fila por fila,
    con fórmula real de Excel (ej. =B2+C2)."""
    if tipo_operacion not in _SIMBOLOS_OPERACION:
        tipo_operacion = "suma"
    simbolo = _SIMBOLOS_OPERACION[tipo_operacion]

    wb = load_workbook(archivo)
    _validar_hoja(wb, hoja)
    ws = wb[hoja]
    _validar_columna(ws, columna_a, "columna_a")
    _validar_columna(ws, columna_b, "columna_b")
    fila_enc = _detectar_fila_encabezado(ws)

    col_a_letra = get_column_letter(columna_a)
    col_b_letra = get_column_letter(columna_b)
    nombre_col_a = ws.cell(row=fila_enc, column=columna_a).value or f"Columna{columna_a}"
    nombre_col_b = ws.cell(row=fila_enc, column=columna_b).value or f"Columna{columna_b}"
    _debug(f"Operación: '{nombre_col_a}' {simbolo} '{nombre_col_b}', fila por fila")

    salida = _hoja_salida(wb, nombre_hoja_salida)
    salida.append([nombre_col_a, nombre_col_b, f"{_NOMBRES_OPERACION[tipo_operacion]} (fórmula)"])

    fila_salida = 2
    for fila in range(fila_enc + 1, ws.max_row + 1):
        val_a = ws.cell(row=fila, column=columna_a).value
        val_b = ws.cell(row=fila, column=columna_b).value
        if val_a is None and val_b is None:
            continue
        ref_a = f"'{hoja}'!{col_a_letra}{fila}"
        ref_b = f"'{hoja}'!{col_b_letra}{fila}"
        salida.cell(row=fila_salida, column=1, value=f"={ref_a}")
        salida.cell(row=fila_salida, column=2, value=f"={ref_b}")
        salida.cell(row=fila_salida, column=3, value=f"={ref_a}{simbolo}{ref_b}")
        _debug(f"Fila {fila}: {val_a} {simbolo} {val_b} → fórmula ={ref_a}{simbolo}{ref_b}")
        fila_salida += 1

    wb.save(archivo)
    print(f"{_NOMBRES_OPERACION[tipo_operacion]} calculada con fórmulas para {fila_salida-2} filas en hoja '{nombre_hoja_salida}'.")
    return fila_salida - 2


def calcular_promedio(archivo, hoja, columna_agrupar, columna_valores, nombre_hoja_salida="PromedioAuto"):
    """Calcula el promedio de una columna, agrupado por categoría,
    con fórmula real de Excel (AVERAGEIF)."""
    wb = load_workbook(archivo)
    _validar_hoja(wb, hoja)
    ws = wb[hoja]
    _validar_columna(ws, columna_agrupar, "columna_agrupar")
    _validar_columna(ws, columna_valores, "columna_valores")
    fila_enc = _detectar_fila_encabezado(ws)

    col_agrupar_letra = get_column_letter(columna_agrupar)
    col_valores_letra = get_column_letter(columna_valores)
    ultima_fila = ws.max_row

    categorias, vistos = [], set()
    for fila in range(fila_enc + 1, ultima_fila + 1):
        clave = ws.cell(row=fila, column=columna_agrupar).value
        if clave is not None and clave not in vistos:
            vistos.add(clave)
            categorias.append(clave)
    _debug(f"Categorías encontradas para promediar: {categorias}")

    salida = _hoja_salida(wb, nombre_hoja_salida)
    salida.append(["Categoría", "Promedio (fórmula)"])
    for i, categoria in enumerate(categorias, start=2):
        salida.cell(row=i, column=1, value=categoria)
        formula = (f"=AVERAGEIF('{hoja}'!{col_agrupar_letra}{fila_enc+1}:{col_agrupar_letra}{ultima_fila},"
                   f"A{i},'{hoja}'!{col_valores_letra}{fila_enc+1}:{col_valores_letra}{ultima_fila})")
        salida.cell(row=i, column=2, value=formula)
        _debug(f"'{categoria}' → fórmula {formula}")

    wb.save(archivo)
    print(f"Promedio calculado con fórmulas AVERAGEIF en hoja '{nombre_hoja_salida}' ({len(categorias)} categorías).")
    return categorias


# =====================================================================
# FUNCIONES DEL MOTOR (15/17 y 16/17)
# =====================================================================

def calcular_descuento(archivo, hoja, columna_precio, columna_porcentaje=None, porcentaje_fijo=None, nombre_hoja_salida="DescuentoAuto"):
    """Calcula el precio final aplicando un % de descuento, con fórmula real.
    Si columna_porcentaje se especifica, usa un % distinto por fila.
    Si no, usa porcentaje_fijo (mismo % para todas las filas)."""
    wb = load_workbook(archivo)
    _validar_hoja(wb, hoja)
    ws = wb[hoja]
    _validar_columna(ws, columna_precio, "columna_precio")
    fila_enc = _detectar_fila_encabezado(ws)

    col_precio_letra = get_column_letter(columna_precio)
    usa_columna_pct = columna_porcentaje is not None
    if usa_columna_pct:
        _validar_columna(ws, columna_porcentaje, "columna_porcentaje")
        col_pct_letra = get_column_letter(columna_porcentaje)
        _debug(f"Usando % variable por fila desde columna {columna_porcentaje}")
    else:
        pct_debug = porcentaje_fijo if porcentaje_fijo is not None else 10
        _debug(f"Usando % fijo para todas las filas: {pct_debug}%")

    salida = _hoja_salida(wb, nombre_hoja_salida)
    encabezados = ["Precio original (ref)", "% Descuento", "Descuento $ (fórmula)", "Precio final (fórmula)"]
    salida.append(encabezados)

    fila_salida = 2
    for fila in range(fila_enc + 1, ws.max_row + 1):
        precio = ws.cell(row=fila, column=columna_precio).value
        if precio is None:
            continue
        ref_precio = f"'{hoja}'!{col_precio_letra}{fila}"
        salida.cell(row=fila_salida, column=1, value=f"={ref_precio}")

        if usa_columna_pct:
            ref_pct = f"'{hoja}'!{col_pct_letra}{fila}"
            salida.cell(row=fila_salida, column=2, value=f"={ref_pct}")
            salida.cell(row=fila_salida, column=3, value=f"={ref_precio}*(B{fila_salida}/100)")
        else:
            pct = porcentaje_fijo if porcentaje_fijo is not None else 10
            salida.cell(row=fila_salida, column=2, value=pct)
            salida.cell(row=fila_salida, column=3, value=f"={ref_precio}*(B{fila_salida}/100)")

        salida.cell(row=fila_salida, column=4, value=f"=A{fila_salida}-C{fila_salida}")
        _debug(f"Fila {fila}: precio {precio} → precio final = precio - descuento")
        fila_salida += 1

    wb.save(archivo)
    print(f"Descuento calculado con fórmulas para {fila_salida-2} filas en hoja '{nombre_hoja_salida}'.")
    return fila_salida - 2


def calcular_maximo_minimo(archivo, hoja, columna_agrupar, columna_valores, tipo="ambos", nombre_hoja_salida="MaxMinAuto"):
    """Encuentra el valor máximo y/o mínimo por categoría, con fórmula real
    (MAXIFS/MINIFS)."""
    wb = load_workbook(archivo)
    _validar_hoja(wb, hoja)
    ws = wb[hoja]
    _validar_columna(ws, columna_agrupar, "columna_agrupar")
    _validar_columna(ws, columna_valores, "columna_valores")
    fila_enc = _detectar_fila_encabezado(ws)

    col_agrupar_letra = get_column_letter(columna_agrupar)
    col_valores_letra = get_column_letter(columna_valores)
    ultima_fila = ws.max_row

    categorias, vistos = [], set()
    for fila in range(fila_enc + 1, ultima_fila + 1):
        clave = ws.cell(row=fila, column=columna_agrupar).value
        if clave is not None and clave not in vistos:
            vistos.add(clave)
            categorias.append(clave)
    _debug(f"Categorías encontradas: {categorias} (tipo solicitado: {tipo})")

    salida = _hoja_salida(wb, nombre_hoja_salida)
    encabezados = ["Categoría"]
    if tipo in ("maximo", "ambos"):
        encabezados.append("Máximo (fórmula)")
    if tipo in ("minimo", "ambos"):
        encabezados.append("Mínimo (fórmula)")
    salida.append(encabezados)

    rango_agrupar = f"'{hoja}'!{col_agrupar_letra}{fila_enc+1}:{col_agrupar_letra}{ultima_fila}"
    rango_valores = f"'{hoja}'!{col_valores_letra}{fila_enc+1}:{col_valores_letra}{ultima_fila}"

    for i, categoria in enumerate(categorias, start=2):
        salida.cell(row=i, column=1, value=categoria)
        col_actual = 2
        if tipo in ("maximo", "ambos"):
            formula_max = f"=MAXIFS({rango_valores},{rango_agrupar},A{i})"
            salida.cell(row=i, column=col_actual, value=formula_max)
            _debug(f"'{categoria}': máximo → {formula_max}")
            col_actual += 1
        if tipo in ("minimo", "ambos"):
            formula_min = f"=MINIFS({rango_valores},{rango_agrupar},A{i})"
            salida.cell(row=i, column=col_actual, value=formula_min)
            _debug(f"'{categoria}': mínimo → {formula_min}")

    wb.save(archivo)
    print(f"Máximo/Mínimo calculado con fórmulas MAXIFS/MINIFS en hoja '{nombre_hoja_salida}' ({len(categorias)} categorías).")
    return categorias


# =====================================================================
# FUNCIÓN DEL MOTOR (17/17)
# =====================================================================

def calcular_porcentaje_total(archivo, hoja, columna_agrupar, columna_valores, nombre_hoja_salida="PorcentajeAuto"):
    """Calcula qué % representa cada categoría del total general,
    con fórmula real (SUMIF dividido entre SUM total)."""
    wb = load_workbook(archivo)
    _validar_hoja(wb, hoja)
    ws = wb[hoja]
    _validar_columna(ws, columna_agrupar, "columna_agrupar")
    _validar_columna(ws, columna_valores, "columna_valores")
    fila_enc = _detectar_fila_encabezado(ws)

    col_agrupar_letra = get_column_letter(columna_agrupar)
    col_valores_letra = get_column_letter(columna_valores)
    ultima_fila = ws.max_row

    categorias, vistos = [], set()
    for fila in range(fila_enc + 1, ultima_fila + 1):
        clave = ws.cell(row=fila, column=columna_agrupar).value
        if clave is not None and clave not in vistos:
            vistos.add(clave)
            categorias.append(clave)
    _debug(f"Categorías encontradas: {categorias}")
    _debug(f"El total general se calcula sumando toda la columna {columna_valores}")

    salida = _hoja_salida(wb, nombre_hoja_salida)
    salida.append(["Categoría", "Total (fórmula)", "% del total (fórmula)"])

    rango_agrupar = f"'{hoja}'!{col_agrupar_letra}{fila_enc+1}:{col_agrupar_letra}{ultima_fila}"
    rango_valores = f"'{hoja}'!{col_valores_letra}{fila_enc+1}:{col_valores_letra}{ultima_fila}"
    total_general = f"SUM({rango_valores})"

    for i, categoria in enumerate(categorias, start=2):
        salida.cell(row=i, column=1, value=categoria)
        formula_total = f"=SUMIF({rango_agrupar},A{i},{rango_valores})"
        salida.cell(row=i, column=2, value=formula_total)
        salida.cell(row=i, column=3, value=f"=B{i}/{total_general}*100")
        _debug(f"'{categoria}': total = {formula_total}, % = total/suma_general*100")

    wb.save(archivo)
    print(f"Porcentaje del total calculado con fórmulas en hoja '{nombre_hoja_salida}' ({len(categorias)} categorías).")
    return categorias


# =====================================================================
# EJEMPLO DE USO (sin IA) — bórralo o coméntalo si vas a importar este
# archivo como módulo desde otro script
# =====================================================================
if __name__ == "__main__":
    ARCHIVO = "mi_archivo.xlsx"   # <-- cambia esto por la ruta real de tu Excel
    HOJA = "Ventas"               # <-- cambia esto por el nombre real de tu hoja

    # Ejemplos — descomenta y ajusta las columnas a tu archivo real:

    # ejecutar_funcion(ARCHIVO, "resumen", hoja=HOJA, columna_agrupar=2, columna_valores=5)
    # ejecutar_funcion(ARCHIVO, "conteo", hoja=HOJA, columna_contar=3)
    # ejecutar_funcion(ARCHIVO, "margen", hoja=HOJA, columna_costo=4, columna_venta=5)
    # ejecutar_funcion(ARCHIVO, "vencimientos", hoja=HOJA, columna_fecha=6, dias_alerta=15)
    # ejecutar_funcion(ARCHIVO, "grafico", hoja_datos="ResumenAuto", tipo_grafico="barras")

    print("Motor cargado. Descomenta un ejemplo arriba o llama a ejecutar_funcion(...) con tus propios parámetros.")
    print(f"Funciones disponibles: {FUNCIONES_DISPONIBLES}")
