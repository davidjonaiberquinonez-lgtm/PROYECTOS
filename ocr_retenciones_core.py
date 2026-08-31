"""
ocr_retenciones_core.py — Inteligencia de extracción OCR de Comprobantes de
Retención de IVA venezolanos, sin ningún framework web encima.

Este módulo NO depende de Flask ni de nada relacionado con servir HTTP —
solo sabe convertir un PDF/imagen en los 6 campos extraídos. Lo usan dos
servidores distintos:
  - servidor_retenciones.py: nuestra interfaz interna (HTML + histórico
    en Excel local).
  - api_retenciones_ocr.py: API pública sin interfaz, para que otro
    equipo le construya su propia UI encima.

Campos que extrae, por cada página/imagen:
  - fecha
  - nro_comprobante
  - cliente (el Agente de Retención que emite el comprobante — la empresa
    que procesa este comprobante aparece como "Proveedor"/"Sujeto
    Retenido" en el documento, y NO es el dato que buscamos acá)
  - nro_factura        (todos los números de factura del comprobante,
                        separados por coma — se mantiene así por
                        compatibilidad con quien ya consume este campo)
  - rif_cliente (RIF de ese mismo Agente de Retención)
  - monto_retenido     (el total retenido, ya sumado)
  - detalle            LISTA con una fila por cada factura de la tabla
                        "Compras Internas e Importaciones" del comprobante
                        (antes solo se leía el resumen de arriba y se
                        perdía este desglose — reportado en vivo, 31/08).
                        Cada fila trae: numero_factura, numero_control,
                        numero_nota_debito, numero_nota_credito,
                        tipo_trans, documento_afectado,
                        total_compras_con_iva, total_compras_sin_iva,
                        base_imponible, porcentaje_alicuota, impuesto_iva,
                        iva_retenido — tal cual esa fila de la tabla.

Motor principal: DeepSeek (deepseek-v4-flash-vision-exp) — modelo de
visión experimental de DeepSeek, lanzado el 21/08/2026. OJO: comprime cada
imagen a ~800x800px / 384 tokens antes de leerla, así que en comprobantes
con letra chica puede rendir peor que Gemini. Si falla o el resultado no
sirve, respaldo automático a Gemini (gemini-3.6-flash), y si ese también
falla, a OpenRouter (dots-studio/dots-3-note-preview, gratis). El campo
"motor" en cada resultado indica cuál de los tres respondió.
"""

import base64
import json
import os
import re
import time

import pymupdf
import requests
from dotenv import load_dotenv

load_dotenv()

# ---------------------------------------------------------------------------
# Configuración
# ---------------------------------------------------------------------------

GOOGLE_BASE_URL = "https://generativelanguage.googleapis.com/v1beta/models"
MODELO_VISION = os.environ.get("MODELO_VISION_GOOGLE", "gemini-3.6-flash")

clave_google = os.environ.get("GOOGLE_API_KEY")
if not clave_google:
    raise SystemExit("Falta GOOGLE_API_KEY en el archivo .env")

OPENROUTER_URL = "https://openrouter.ai/api/v1/chat/completions"
MODELO_VISION_OPENROUTER = os.environ.get("MODELO_VISION_OPENROUTER", "dots-studio/dots-3-note-preview:free")
clave_openrouter = os.environ.get("OPENROUTER_API_KEY")

DEEPSEEK_URL = "https://api.deepseek.com/chat/completions"
MODELO_VISION_DEEPSEEK = os.environ.get("MODELO_VISION_DEEPSEEK", "deepseek-v4-flash-vision-exp")
clave_deepseek = os.environ.get("DEEPSEEK_API_KEY")

EXTENSIONES_IMAGEN = {"png", "jpg", "jpeg", "webp", "bmp"}
EXTENSIONES_PERMITIDAS = EXTENSIONES_IMAGEN | {"pdf"}


# ---------------------------------------------------------------------------
# PDF -> imágenes (una por página)
# ---------------------------------------------------------------------------

def pdf_a_imagenes(contenido_pdf, dpi=200):
    """Convierte cada página de un PDF a una imagen PNG (bytes)."""
    documento = pymupdf.open(stream=contenido_pdf, filetype="pdf")
    zoom = dpi / 72
    matriz = pymupdf.Matrix(zoom, zoom)
    imagenes = []
    try:
        for pagina in documento:
            pixmap = pagina.get_pixmap(matrix=matriz)
            imagenes.append(pixmap.tobytes("png"))
    finally:
        documento.close()
    return imagenes


# ---------------------------------------------------------------------------
# Extracción con el modelo de visión
# ---------------------------------------------------------------------------

class ErrorOCR(Exception):
    pass


PROMPT_SISTEMA = """Actúa como un sistema avanzado de OCR y auditoría contable especializado en documentos fiscales venezolanos (SENIAT).
Tu tarea es analizar la imagen proporcionada de un "Comprobante de Retención de I.V.A." y extraer estrictamente los siguientes 7 campos de datos.

CONTEXTO IMPORTANTE: quien procesa este comprobante es la empresa que aparece como "Proveedor" o "Sujeto Retenido" en el documento (normalmente CRIST MEDICALS, C.A. o similar) — esa empresa NO es el dato que buscamos, es la nuestra. Lo que necesitamos es identificar al AGENTE DE RETENCIÓN: la farmacia u otra empresa que EMITE el comprobante y nos retiene el impuesto — ese es nuestro "cliente" en este contexto.

REGLAS DE EXTRACCIÓN:
1. "fecha": Busca el campo "Fecha de emisión". Devuelve el valor en formato DD/MM/YYYY.
2. "nro_comprobante": Busca el "Nro de Comprobante" (suele ser un código numérico largo de 14 dígitos que empieza con el año, ej. 2026...).
3. "cliente": Busca el nombre o razón social del AGENTE DE RETENCIÓN — quien EMITE el comprobante (suele estar en el bloque "Identificación Agente de Retención" o "Nombre o Razón Social del Agente de Retención", normalmente arriba del todo). NO uses el nombre del "Proveedor" ni del "Sujeto Retenido" — esa es nuestra propia empresa, no el cliente.
4. "nro_factura": Ubica la tabla central de operaciones ("Compras Internas e Importaciones") y extrae TODOS los valores de la columna "Nro de Factura"/"Número de Factura" de esa tabla, separados por coma en el mismo orden en que aparecen (ej. "00154065, 00153866, 00155170") — el comprobante puede traer más de una factura.
5. "rif_cliente": Busca el RIF del AGENTE DE RETENCIÓN (formato J-XXXXXXXX-X, V-XXXXXXXX-X) — el mismo emisor identificado en el campo "cliente". NO uses el RIF que aparece junto a "Proveedor" o "Sujeto Retenido" — ese es el RIF de nuestra propia empresa, no el que necesitamos.
6. "monto_retenido": Busca el total de "Impuesto Retenido I.V.A." o "Monto Impuesto Retenido" del comprobante (el total ya sumado, no de una fila individual).
7. "detalle": la tabla "Compras Internas e Importaciones" completa, UNA fila del JSON por cada fila real de esa tabla (no te saltees ninguna, incluso si el comprobante trae solo una). Cada fila tiene EXACTAMENTE estas claves, tal cual estén escritas en esa columna de la tabla (no las inventes ni las calcules si no están):
   - "numero_factura": columna "Número Factura" de esa fila.
   - "numero_control": columna "Número Control".
   - "numero_nota_debito": columna "Número Nota Débito".
   - "numero_nota_credito": columna "Número Nota Crédito".
   - "tipo_trans": columna "Tipo Trans" (el código de tipo de transacción, ej. "01").
   - "documento_afectado": columna "Documento Afectado".
   - "total_compras_con_iva": columna "Total Compras Incluye Iva".
   - "total_compras_sin_iva": columna "Total Compras Sin Iva".
   - "base_imponible": columna "Base Imponible".
   - "porcentaje_alicuota": columna "% Alíc" (el porcentaje, ej. "16.00").
   - "impuesto_iva": columna "Impuesto IVA".
   - "iva_retenido": columna "IVA Retenido" de esa fila puntual (la suma de esta columna en todas las filas debería dar el "monto_retenido" del punto 6 — si no coincide, igual transcribe lo que está escrito, no lo fuerces).

IMPORTANTE SOBRE LA ESTRUCTURA: las 7 claves de arriba van SIEMPRE las 7,
en un ÚNICO objeto JSON — nunca cierres el objeto antes de las 7, y nunca
omitas una clave aunque no tengas el dato (usa null, o [] en el caso de
"detalle" — pero la clave tiene que estar). No entregues "detalle" como
un bloque aparte ni en un JSON distinto: es una clave más adentro del
mismo objeto que fecha/nro_comprobante/etc.

FORMATO DE SALIDA:
Devuelve ÚNICAMENTE un objeto JSON válido con la siguiente estructura, sin texto adicional, sin bloques de código markdown (```json) y sin explicaciones:
{
  "fecha": "DD/MM/YYYY",
  "nro_comprobante": "valor",
  "cliente": "valor",
  "nro_factura": "valor",
  "rif_cliente": "valor",
  "monto_retenido": "valor",
  "detalle": [
    {
      "numero_factura": "valor", "numero_control": "valor", "numero_nota_debito": "valor",
      "numero_nota_credito": "valor", "tipo_trans": "valor", "documento_afectado": "valor",
      "total_compras_con_iva": "valor", "total_compras_sin_iva": "valor", "base_imponible": "valor",
      "porcentaje_alicuota": "valor", "impuesto_iva": "valor", "iva_retenido": "valor"
    }
  ]
}
Si un dato es completamente ilegible o no existe en el comprobante, usa el valor null. Si la tabla de compras no tiene ninguna fila legible, "detalle" debe ser una lista vacía: []."""

CAMPOS_ESPERADOS = ("fecha", "nro_comprobante", "cliente", "nro_factura", "rif_cliente", "monto_retenido")

CAMPOS_DETALLE = (
    "numero_factura", "numero_control", "numero_nota_debito", "numero_nota_credito",
    "tipo_trans", "documento_afectado", "total_compras_con_iva", "total_compras_sin_iva",
    "base_imponible", "porcentaje_alicuota", "impuesto_iva", "iva_retenido",
)


def _extraer_json(texto):
    """El modelo debería responder solo JSON, pero por si acaso agrega
    texto alrededor, se toma el primer bloque {...} que aparezca."""
    inicio = texto.find("{")
    fin = texto.rfind("}")
    if inicio == -1 or fin == -1 or fin < inicio:
        raise ErrorOCR(f"La respuesta del modelo no contiene JSON: {texto!r}")
    return json.loads(texto[inicio:fin + 1])


def _normalizar_clave(clave):
    """Reduce una clave a solo sus letras/números en minúscula, para poder
    comparar "rif_cliente", "rifCliente", "rif cliente" y "RIF-Cliente"
    como si fueran la misma clave."""
    return re.sub(r"[^a-z0-9]", "", clave.lower())


def _mapear_campos(datos):
    """Arma el diccionario de campos esperados tolerando que el modelo no
    respete el nombre exacto en snake_case — el respaldo de OpenRouter en
    particular ha devuelto camelCase ("rifCliente") y hasta con espacio
    ("rif cliente") en vez de "rif_cliente". Sin esto, esas variaciones se
    pierden como null aunque el dato sí vino en la respuesta.

    Agrega además "detalle" (lista de filas de la tabla de compras) —
    aditivo, no rompe a quien ya lee los 6 campos planos de siempre (ver
    docstring del módulo, 31/08): "nro_factura"/"monto_retenido" siguen
    siendo el resumen de todo el comprobante, "detalle" es el desglose
    fila por fila."""
    normalizado = {_normalizar_clave(clave): valor for clave, valor in datos.items()}
    campos = {campo: normalizado.get(_normalizar_clave(campo)) for campo in CAMPOS_ESPERADOS}
    campos["detalle"] = _mapear_detalle(normalizado.get(_normalizar_clave("detalle")))
    return campos


def _mapear_detalle(filas_crudas):
    """Tolera lo mismo que _mapear_campos pero fila por fila — si el
    modelo no devolvió "detalle" o no es una lista, [] (nunca se inventa
    una fila que no vino)."""
    if not isinstance(filas_crudas, list):
        return []
    detalle = []
    for fila in filas_crudas:
        if not isinstance(fila, dict):
            continue
        fila_normalizada = {_normalizar_clave(clave): valor for clave, valor in fila.items()}
        detalle.append({campo: fila_normalizada.get(_normalizar_clave(campo)) for campo in CAMPOS_DETALLE})
    return detalle


# Esquema forzado para Gemini — DeepSeek rechaza esto ("This response_format
# type is unavailable now", confirmado en vivo el 31/08: solo acepta
# response_format json_object, que garantiza JSON válido pero NO garantiza
# qué claves trae — un campo puede faltar sin que eso sea un error de
# formato). Gemini sí soporta responseSchema con "required": fuerza a que
# las 7 claves de arriba y las 12 de cada fila de "detalle" estén SIEMPRE
# presentes (null/[] si no hay dato, pero nunca ausentes) — por eso Gemini
# es el respaldo que de verdad garantiza la estructura completa cuando
# DeepSeek (que solo puede pedírselo por instrucción, no forzarlo) se
# queda corto.
_ESQUEMA_ITEM_DETALLE = {
    "type": "OBJECT",
    "properties": {campo: {"type": "STRING", "nullable": True} for campo in CAMPOS_DETALLE},
    "required": list(CAMPOS_DETALLE),
}
ESQUEMA_RESPUESTA_GEMINI = {
    "type": "OBJECT",
    "properties": {
        **{campo: {"type": "STRING", "nullable": True} for campo in CAMPOS_ESPERADOS},
        "detalle": {"type": "ARRAY", "items": _ESQUEMA_ITEM_DETALLE},
    },
    "required": list(CAMPOS_ESPERADOS) + ["detalle"],
}


def consultar_vision_google(imagen_bytes):
    b64 = base64.b64encode(imagen_bytes).decode()
    url = f"{GOOGLE_BASE_URL}/{MODELO_VISION}:generateContent?key={clave_google}"
    payload = {
        "systemInstruction": {"parts": [{"text": PROMPT_SISTEMA}]},
        "contents": [{
            "parts": [
                {"text": "Extrae los 7 campos de esta imagen (incluida la tabla de detalle completa) siguiendo exactamente las reglas. Responde solo con el JSON, nada más."},
                {"inline_data": {"mime_type": "image/png", "data": b64}},
            ],
        }],
        "generationConfig": {
            "temperature": 0.0,
            "maxOutputTokens": 4096,
            "responseMimeType": "application/json",
            "responseSchema": ESQUEMA_RESPUESTA_GEMINI,
        },
    }

    texto = None
    ultimo_error = None
    # Solo 2 intentos: con el respaldo de OpenRouter ya no vale la pena
    # hacer esperar al que llama hasta 150s en backoff antes de rendirse
    # — mejor fallar rápido y pasar la página al respaldo.
    for intento in range(2):
        try:
            respuesta = requests.post(url, json=payload, timeout=60)
            respuesta.raise_for_status()
            cuerpo = respuesta.json()
            texto = cuerpo["candidates"][0]["content"]["parts"][0]["text"]
            ultimo_error = None
            break
        except Exception as error:
            ultimo_error = error
            codigo = error.response.status_code if isinstance(error, requests.HTTPError) and error.response is not None else None
            if codigo in (429, 503):
                time.sleep(5 * (intento + 1))
                continue
            break
    if ultimo_error is not None:
        raise ErrorOCR(f"Gemini falló: {ultimo_error}") from ultimo_error

    datos = _extraer_json(texto)
    campos = _mapear_campos(datos)
    return campos, texto


def consultar_vision_openrouter(imagen_bytes):
    if not clave_openrouter:
        raise ErrorOCR("No hay OPENROUTER_API_KEY configurada para el respaldo")

    b64 = base64.b64encode(imagen_bytes).decode()
    payload = {
        "model": MODELO_VISION_OPENROUTER,
        "temperature": 0.0,
        "max_tokens": 6000,
        "response_format": {"type": "json_object"},
        "messages": [
            {"role": "system", "content": PROMPT_SISTEMA},
            {"role": "user", "content": [
                {"type": "text", "text": "Extrae los 7 campos de esta imagen (incluida la tabla de detalle completa) siguiendo exactamente las reglas. Responde solo con el JSON, nada más. Usa exactamente estos nombres de clave, sin traducirlos ni cambiarles el formato: fecha, nro_comprobante, cliente, nro_factura, rif_cliente, monto_retenido, detalle (cada fila con numero_factura, numero_control, numero_nota_debito, numero_nota_credito, tipo_trans, documento_afectado, total_compras_con_iva, total_compras_sin_iva, base_imponible, porcentaje_alicuota, impuesto_iva, iva_retenido)."},
                {"type": "image_url", "image_url": {"url": f"data:image/png;base64,{b64}"}},
            ]},
        ],
    }
    try:
        respuesta = requests.post(
            OPENROUTER_URL, json=payload,
            headers={"Authorization": f"Bearer {clave_openrouter}"}, timeout=90,
        )
        respuesta.raise_for_status()
        cuerpo = respuesta.json()
        texto = cuerpo["choices"][0]["message"]["content"]
    except Exception as error:
        raise ErrorOCR(f"OpenRouter también falló: {error}") from error

    if not texto:
        raise ErrorOCR("OpenRouter no devolvió contenido (se quedó sin tokens razonando)")

    datos = _extraer_json(texto)
    campos = _mapear_campos(datos)
    return campos, texto


def consultar_vision_deepseek(imagen_bytes):
    if not clave_deepseek:
        raise ErrorOCR("No hay DEEPSEEK_API_KEY configurada")

    b64 = base64.b64encode(imagen_bytes).decode()
    payload = {
        "model": MODELO_VISION_DEEPSEEK,
        "temperature": 0.0,
        "max_tokens": 8192,  # deepseek-v4-flash-vision-exp razona antes de responder
        # (campo "reasoning_content" aparte de "content", confirmado en vivo el 31/08:
        # ~600 tokens de razonamiento incluso con una imagen en blanco) — con 2048/4096
        # el modelo podía gastar todo el tope pensando y nunca llegar a escribir el
        # JSON: consumía tokens de entrada/salida igual, pero "content" quedaba vacío.
        # Encima ahora la respuesta trae la tabla de detalle completa, no solo el resumen.
        "response_format": {"type": "json_object"},
        "messages": [
            {"role": "system", "content": PROMPT_SISTEMA},
            {"role": "user", "content": [
                {"type": "text", "text": "Extrae los 7 campos de esta imagen (incluida la tabla de detalle completa) siguiendo exactamente las reglas. Responde solo con el JSON, nada más."},
                {"type": "image_url", "image_url": {"url": f"data:image/png;base64,{b64}"}},
            ]},
        ],
    }
    texto = None
    ultimo_error = None
    # Mismo criterio de reintento que consultar_vision_google: 429 (límite
    # de peticiones) y 503 (saturado) sí valen la pena reintentar una vez
    # con espera — antes esto solo lo tenía Gemini, y un 429/503 de
    # DeepSeek se reportaba como "falló" sin darle ni un segundo intento
    # (reportado en vivo, 31/08: fallaba al leer facturas reales por esto).
    for intento in range(2):
        try:
            respuesta = requests.post(
                DEEPSEEK_URL, json=payload,
                headers={"Authorization": f"Bearer {clave_deepseek}"}, timeout=60,
            )
            respuesta.raise_for_status()
            cuerpo = respuesta.json()
            texto = cuerpo["choices"][0]["message"]["content"]
            ultimo_error = None
            break
        except Exception as error:
            ultimo_error = error
            codigo = error.response.status_code if isinstance(error, requests.HTTPError) and error.response is not None else None
            if codigo in (429, 503):
                time.sleep(5 * (intento + 1))
                continue
            break
    if ultimo_error is not None:
        raise ErrorOCR(f"DeepSeek falló: {ultimo_error}") from ultimo_error

    if not texto:
        raise ErrorOCR("DeepSeek no devolvió contenido (se quedó sin tokens razonando)")

    datos = _extraer_json(texto)
    campos = _mapear_campos(datos)
    return campos, texto


def consultar_vision(imagen_bytes):
    """DeepSeek primero (más barato); si falla o no está configurado, se
    reintenta con Gemini, y si ese también falla, con OpenRouter. `motor`
    en el resultado indica cuál de los tres respondió.

    TEMPORAL: DeepSeek recién estrenó visión (deepseek-v4-flash-vision-exp,
    21/08/2026) y comprime cada imagen a ~800x800px/384 tokens antes de
    leerla — en comprobantes con letra chica puede rendir peor que Gemini.
    Si la precisión no convence, volver a poner a Gemini primero."""
    try:
        campos, texto = consultar_vision_deepseek(imagen_bytes)
        return campos, texto, "DeepSeek"
    except ErrorOCR as error_deepseek:
        try:
            campos, texto = consultar_vision_google(imagen_bytes)
            return campos, texto, "Gemini (respaldo)"
        except ErrorOCR as error_gemini:
            try:
                campos, texto = consultar_vision_openrouter(imagen_bytes)
                return campos, texto, "OpenRouter (respaldo)"
            except ErrorOCR as error_openrouter:
                raise ErrorOCR(f"{error_deepseek} — {error_gemini} — {error_openrouter}") from error_openrouter


def procesar_imagen(imagen_bytes, nombre_archivo, numero_pagina):
    resultado_pagina = {"pagina": numero_pagina, "archivo": nombre_archivo, "ok": False}
    try:
        campos, texto_crudo, motor = consultar_vision(imagen_bytes)
    except ErrorOCR as error:
        resultado_pagina["error"] = str(error)
        return resultado_pagina

    resultado_pagina.update({"ok": True, "campos": campos, "texto_ocr": texto_crudo, "motor": motor})
    return resultado_pagina


def procesar_documento(contenido, nombre_archivo, extension):
    """Punto de entrada de alto nivel: recibe los bytes crudos de un
    archivo (PDF o imagen) y su extensión, y devuelve la lista de
    resultados (uno por página). No sabe nada de HTTP ni de Flask."""
    if extension not in EXTENSIONES_PERMITIDAS:
        raise ValueError(f"Extensión no soportada: .{extension}")

    if extension == "pdf":
        imagenes = pdf_a_imagenes(contenido)
        if not imagenes:
            raise ValueError("El PDF no tiene páginas")
    else:
        imagenes = [contenido]

    resultados = []
    for indice, imagen_bytes in enumerate(imagenes, start=1):
        if indice > 1:
            # Freno preventivo: la capa gratuita de Gemini limita a ~15
            # peticiones/minuto. Sin esta pausa, un documento de varias
            # páginas dispara todas casi de corrido y choca con ese
            # límite (429) aunque cada página individual responda rápido.
            time.sleep(4)
        nombre = f"{nombre_archivo}-pagina{indice}.png" if extension == "pdf" else nombre_archivo
        resultados.append(procesar_imagen(imagen_bytes, nombre, indice))
    return resultados
