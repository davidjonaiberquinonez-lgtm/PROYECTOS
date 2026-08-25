"""
ocr_ps_core.py — Inteligencia de extracción OCR para "OCR-PS" (base de
datos de Psicotrópicos), sin ningún framework web encima.

Mismo patrón que ocr_retenciones_core.py y ocr_abonos_core.py (Gemini +
respaldo OpenRouter, mismo manejo de PDF/imagen), pero para el caso más
simple de los tres: el gerente de compras pasa las facturas de productos
psicotrópicos y solo hace falta UN dato por factura — el número de
factura — para cruzarlo contra la base de datos de control de
psicotrópicos.

Campo que extrae, por cada página/imagen:
  - numero_factura   El número de factura, como entero (ej. 174857). Si
                      la factura no tiene un número identificable, null.

Motor principal: Google Gemini (gemini-3.6-flash). Motor de respaldo:
OpenRouter (dots-studio/dots-3-note-preview, gratis) — automático si
Gemini agota reintentos por 429/503. El campo "motor" del resultado
indica cuál de los dos respondió.
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
# Configuración (comparte credenciales con los otros módulos de OCR)
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


PROMPT_SISTEMA = """Actúa como un sistema de OCR ultra enfocado. Tu único trabajo es encontrar el NÚMERO DE FACTURA en la imagen de una factura de compra (de un proveedor de productos psicotrópicos/farmacéuticos) y devolverlo — nada más.

DÓNDE BUSCAR: el número de factura suele estar cerca de las palabras "Factura", "Factura N°", "Nro. Factura", "Invoice", "N° de Control" (si hay varios números parecidos, el de "N° de Control" NO es el que buscas — preferí el que esté etiquetado como "Factura"). Es casi siempre la referencia numérica más prominente arriba del documento, junto al logo o encabezado del proveedor.

REGLAS:
1. Devuelve el número de factura como un ENTERO, sin ceros a la izquierda, sin guiones, puntos ni espacios (ej. si en la factura dice "N° 00174857" o "174.857", el valor es 174857).
2. Si hay letras o un prefijo pegado al número (ej. "A-00174857" o "FAC-174857"), devuelve solo la parte numérica.
3. Si genuinamente no encontrás ningún número de factura en la imagen, usa null — no inventes ni adivines un número.

FORMATO DE SALIDA:
Devuelve ÚNICAMENTE un objeto JSON válido con esta forma exacta, sin texto adicional, sin bloques de código markdown (```json) y sin explicaciones:
{
  "numero_factura": 174857
}
Si no se encuentra el número, usa: {"numero_factura": null}"""

CAMPOS_ESPERADOS = ("numero_factura",)


def _extraer_json(texto):
    """El modelo debería responder solo JSON, pero por si acaso agrega
    texto alrededor, se toma el primer bloque {...} que aparezca."""
    inicio = texto.find("{")
    fin = texto.rfind("}")
    if inicio == -1 or fin == -1 or fin < inicio:
        raise ErrorOCR(f"La respuesta del modelo no contiene JSON: {texto!r}")
    return json.loads(texto[inicio:fin + 1])


def _normalizar_clave(clave):
    return re.sub(r"[^a-z0-9]", "", clave.lower())


def _a_numero_factura(valor):
    """Tolera que el modelo devuelva el número como int, como string
    ("174857", "N° 174857", "174.857") o con basura alrededor — se queda
    solo con los dígitos. Vacío/no numérico -> None, nunca un error que
    tumbe el resto de la extracción."""
    if valor is None:
        return None
    if isinstance(valor, (int, float)):
        return int(valor)
    solo_digitos = re.sub(r"\D", "", str(valor))
    return int(solo_digitos) if solo_digitos else None


def _mapear_campos(datos):
    """Arma el diccionario de campos esperados tolerando variaciones de
    nombre de clave (mismo criterio que los otros módulos de OCR)."""
    normalizado = {_normalizar_clave(clave): valor for clave, valor in datos.items()}
    crudo = normalizado.get(_normalizar_clave("numero_factura"))
    return {"numero_factura": _a_numero_factura(crudo)}


def consultar_vision_google(imagen_bytes):
    b64 = base64.b64encode(imagen_bytes).decode()
    url = f"{GOOGLE_BASE_URL}/{MODELO_VISION}:generateContent?key={clave_google}"
    payload = {
        "systemInstruction": {"parts": [{"text": PROMPT_SISTEMA}]},
        "contents": [{
            "parts": [
                {"text": "Extrae el número de factura de esta imagen siguiendo exactamente las reglas. Responde solo con el JSON, nada más."},
                {"inline_data": {"mime_type": "image/png", "data": b64}},
            ],
        }],
        "generationConfig": {
            "temperature": 0.0,
            "maxOutputTokens": 512,
            "responseMimeType": "application/json",
        },
    }

    texto = None
    ultimo_error = None
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
        "max_tokens": 512,
        "response_format": {"type": "json_object"},
        "messages": [
            {"role": "system", "content": PROMPT_SISTEMA},
            {"role": "user", "content": [
                {"type": "text", "text": (
                    "Extrae el número de factura de esta imagen siguiendo exactamente las reglas. Responde "
                    "solo con el JSON, nada más. Usa exactamente esta clave: numero_factura."
                )},
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
    cualquier otra razón, se reintenta automáticamente con OpenRouter."""
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
    resultados (uno por página)."""
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
            time.sleep(4)
        nombre = f"{nombre_archivo}-pagina{indice}.png" if extension == "pdf" else nombre_archivo
        resultados.append(procesar_imagen(imagen_bytes, nombre, indice))
    return resultados
