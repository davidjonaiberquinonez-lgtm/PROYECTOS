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
  - nro_factura
  - rif_cliente (RIF de ese mismo Agente de Retención)
  - monto_retenido

Motor principal: Google Gemini (gemini-3.6-flash). Motor de respaldo:
OpenRouter (dots-studio/dots-3-note-preview, gratis) — se usa automático
si Gemini agota sus reintentos por 429/503. El campo "motor" en cada
resultado indica cuál de los dos respondió.
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
Tu tarea es analizar la imagen proporcionada de un "Comprobante de Retención de I.V.A." y extraer estrictamente los siguientes 6 campos de datos.

CONTEXTO IMPORTANTE: quien procesa este comprobante es la empresa que aparece como "Proveedor" o "Sujeto Retenido" en el documento (normalmente CRIST MEDICALS, C.A. o similar) — esa empresa NO es el dato que buscamos, es la nuestra. Lo que necesitamos es identificar al AGENTE DE RETENCIÓN: la farmacia u otra empresa que EMITE el comprobante y nos retiene el impuesto — ese es nuestro "cliente" en este contexto.

REGLAS DE EXTRACCIÓN:
1. "fecha": Busca el campo "Fecha de emisión". Devuelve el valor en formato DD/MM/YYYY.
2. "nro_comprobante": Busca el "Nro de Comprobante" (suele ser un código numérico largo de 14 dígitos que empieza con el año, ej. 2026...).
3. "cliente": Busca el nombre o razón social del AGENTE DE RETENCIÓN — quien EMITE el comprobante (suele estar en el bloque "Identificación Agente de Retención" o "Nombre o Razón Social del Agente de Retención", normalmente arriba del todo). NO uses el nombre del "Proveedor" ni del "Sujeto Retenido" — esa es nuestra propia empresa, no el cliente.
4. "nro_factura": Ubica la tabla central de operaciones y extrae el valor bajo la columna "Nro de Factura" o "Número de Factura".
5. "rif_cliente": Busca el RIF del AGENTE DE RETENCIÓN (formato J-XXXXXXXX-X, V-XXXXXXXX-X) — el mismo emisor identificado en el campo "cliente". NO uses el RIF que aparece junto a "Proveedor" o "Sujeto Retenido" — ese es el RIF de nuestra propia empresa, no el que necesitamos.
6. "monto_retenido": Busca en la tabla inferior la columna "Impuesto Retenido I.V.A." o "Monto Impuesto Retenido" y extrae la cantidad exacta.

FORMATO DE SALIDA:
Devuelve ÚNICAMENTE un objeto JSON válido con la siguiente estructura, sin texto adicional, sin bloques de código markdown (```json) y sin explicaciones:
{
  "fecha": "DD/MM/YYYY",
  "nro_comprobante": "valor",
  "cliente": "valor",
  "nro_factura": "valor",
  "rif_cliente": "valor",
  "monto_retenido": "valor"
}
Si un dato es completamente ilegible o no existe en el comprobante, usa el valor null."""

CAMPOS_ESPERADOS = ("fecha", "nro_comprobante", "cliente", "nro_factura", "rif_cliente", "monto_retenido")


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
    pierden como null aunque el dato sí vino en la respuesta."""
    normalizado = {_normalizar_clave(clave): valor for clave, valor in datos.items()}
    return {campo: normalizado.get(_normalizar_clave(campo)) for campo in CAMPOS_ESPERADOS}


def consultar_vision_google(imagen_bytes):
    b64 = base64.b64encode(imagen_bytes).decode()
    url = f"{GOOGLE_BASE_URL}/{MODELO_VISION}:generateContent?key={clave_google}"
    payload = {
        "systemInstruction": {"parts": [{"text": PROMPT_SISTEMA}]},
        "contents": [{
            "parts": [
                {"text": "Extrae los 5 campos de esta imagen siguiendo exactamente las reglas. Responde solo con el JSON, nada más."},
                {"inline_data": {"mime_type": "image/png", "data": b64}},
            ],
        }],
        "generationConfig": {
            "temperature": 0.0,
            "maxOutputTokens": 4096,
            "responseMimeType": "application/json",
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
                {"type": "text", "text": "Extrae los 5 campos de esta imagen siguiendo exactamente las reglas. Responde solo con el JSON, nada más. Usa exactamente estos nombres de clave, sin traducirlos ni cambiarles el formato: fecha, nro_comprobante, cliente, nro_factura, rif_cliente, monto_retenido."},
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


def consultar_vision(imagen_bytes):
    """Gemini primero (más preciso); si se satura (429/503) o falla por
    cualquier otra razón, se reintenta automáticamente con OpenRouter en
    vez de devolver un error. `motor` en el resultado indica cuál de los
    dos respondió."""
    try:
        campos, texto = consultar_vision_google(imagen_bytes)
        return campos, texto, "Gemini"
    except ErrorOCR as error_gemini:
        try:
            campos, texto = consultar_vision_openrouter(imagen_bytes)
            return campos, texto, "OpenRouter (respaldo)"
        except ErrorOCR as error_openrouter:
            raise ErrorOCR(f"{error_gemini} — {error_openrouter}") from error_openrouter


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
