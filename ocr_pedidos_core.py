"""
ocr_pedidos_core.py — Inteligencia de extracción OCR de "Pedidos" (notas de
entrega manuscritas de un cliente/sucursal), sin ningún framework web
encima.

Mismo patrón que ocr_ps_core.py y ocr_abonos_core.py (DeepSeek + respaldo
Gemini + respaldo final OpenRouter, mismo manejo de PDF/imagen), pero para
un documento distinto: una nota manuscrita donde se listan uno o más
artículos pedidos, no un dato único por página. Por eso la forma de
"campos" acá NO es un diccionario plano como en los otros módulos — es:

  - sede       "Sucursal" | "Principal" | null — para qué sede es el
               pedido, según lo que esté escrito/marcado en la nota.
  - items      Lista de artículos pedidos, en el mismo orden en que
               aparecen escritos. Cada ítem:
                 - articulo    Nombre y descripción del artículo, tal cual
                               está escrito.
                 - cantidad    Cantidad solicitada, tal cual está escrita
                               (no se fuerza a número — la letra manuscrita
                               puede traer "x2", "2 cajas", etc.).
                 - pediatrico  true SOLO si al lado de ESE artículo
                               puntual la nota dice "pediátrico" (o una
                               abreviación clara) — false en cualquier
                               otro caso, nunca se asume.

Motor principal: DeepSeek (deepseek-v4-flash-vision-exp). Si falla,
respaldo automático a Gemini (gemini-3.6-flash), y si ese también falla, a
OpenRouter (dots-studio/dots-3-note-preview, gratis). El campo "motor" del
resultado indica cuál de los tres respondió.
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


PROMPT_SISTEMA = """Actúa como un sistema de OCR especializado en notas de pedido manuscritas de una empresa farmacéutica venezolana (Crist Medicals, C.A.).

CONTEXTO DEL DOCUMENTO: es una nota manuscrita donde un cliente o una sucursal anota a mano uno o más artículos que necesita pedir, generalmente en una lista/tabla informal (un renglón por artículo, con su cantidad).

Tu tarea es extraer EXACTAMENTE esta estructura:

1. "sede": indica si el pedido es para una SUCURSAL o para la SEDE/CASA PRINCIPAL — busca esa palabra escrita o marcada en algún lugar de la nota (encabezado, esquina, un círculo/check junto a una de las dos opciones). Devuelve exactamente el texto "Sucursal" o "Principal" según corresponda. Si la nota no lo indica en ningún lado, usa null — no lo asumas ni lo adivines.

2. "items": una lista con TODOS los artículos pedidos, en el mismo orden en que aparecen escritos en la nota — no te saltees ninguno, incluso si algún renglón tiene letra difícil. Cada ítem de la lista tiene:
   - "articulo": el nombre y la descripción del artículo, tal cual está escrito (ej. "Amoxicilina 500mg susp", "Diclofenac gel", "Suero fisiológico 250ml").
   - "cantidad": la cantidad solicitada de ese artículo puntual, tal cual está escrita — puede ser un número solo, o venir con una unidad ("2 cajas", "x3", "1 caja y media"). Transcribila tal cual, sin inventar una unidad que no esté escrita.
   - "pediatrico": true SOLO si justo al lado de la descripción de ESE artículo aparece escrita la palabra "pediátrico"/"pediatrico" o una abreviación clara ("ped.", "PED"). Si no aparece nada de eso junto a ese renglón puntual, el valor es false — nunca marques true por asociación con otro renglón ni por suposición.

REGLAS GENERALES:
- Si una palabra o parte de un artículo es genuinamente ilegible, transcribe lo que sí se distingue con claridad y no inventes el resto.
- Si la nota no tiene ningún artículo legible, "items" debe ser una lista vacía: [].
- No inventes artículos, cantidades ni la sede si no están escritos en la imagen.

FORMATO DE SALIDA:
Devuelve ÚNICAMENTE un JSON válido con esta forma exacta, sin texto antes ni después, sin bloques de código markdown (```json) y sin explicaciones:
{
  "sede": "Sucursal",
  "items": [
    {"articulo": "...", "cantidad": "...", "pediatrico": false}
  ]
}
Si no se puede determinar la sede, usa "sede": null. Si no hay artículos legibles, usa "items": []."""

CAMPOS_ESPERADOS = ("sede", "items")


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


def _a_booleano(valor):
    """Tolera que "pediatrico" venga como bool real, string ("true"/"si"),
    o cualquier otra cosa rara del respaldo de OpenRouter — ante cualquier
    duda, false (mismo criterio "no inventes" del resto del prompt: si no
    es un true inequívoco, no se marca pediátrico)."""
    if isinstance(valor, bool):
        return valor
    if valor is None:
        return False
    return str(valor).strip().lower() in ("true", "si", "sí", "1", "yes", "x")


def _texto_o_none(valor):
    if valor is None:
        return None
    texto = str(valor).strip()
    return texto or None


def _mapear_item(item):
    if not isinstance(item, dict):
        return None
    normalizado = {_normalizar_clave(clave): valor for clave, valor in item.items()}
    articulo = _texto_o_none(normalizado.get(_normalizar_clave("articulo")))
    if not articulo:
        return None
    return {
        "articulo": articulo,
        "cantidad": _texto_o_none(normalizado.get(_normalizar_clave("cantidad"))),
        "pediatrico": _a_booleano(normalizado.get(_normalizar_clave("pediatrico"))),
    }


def _mapear_campos(datos):
    """Arma {"sede": ..., "items": [...]} tolerando variaciones de nombre
    de clave — mismo criterio que los otros módulos de OCR, adaptado acá
    porque la forma no es un diccionario plano sino sede + lista."""
    normalizado = {_normalizar_clave(clave): valor for clave, valor in datos.items()}
    sede = _texto_o_none(normalizado.get(_normalizar_clave("sede")))
    items_crudos = normalizado.get(_normalizar_clave("items"))
    items = []
    if isinstance(items_crudos, list):
        for item_crudo in items_crudos:
            item = _mapear_item(item_crudo)
            if item is not None:
                items.append(item)
    return {"sede": sede, "items": items}


def consultar_vision_google(imagen_bytes):
    b64 = base64.b64encode(imagen_bytes).decode()
    url = f"{GOOGLE_BASE_URL}/{MODELO_VISION}:generateContent?key={clave_google}"
    payload = {
        "systemInstruction": {"parts": [{"text": PROMPT_SISTEMA}]},
        "contents": [{
            "parts": [
                {"text": "Extrae la sede y la lista completa de artículos pedidos de esta imagen siguiendo exactamente las reglas. Responde solo con el JSON, nada más."},
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
                {"type": "text", "text": (
                    "Extrae la sede y la lista completa de artículos pedidos de esta imagen siguiendo exactamente "
                    "las reglas. Responde solo con el JSON, nada más. Usa exactamente estas claves: sede, items "
                    "(cada ítem con articulo, cantidad, pediatrico)."
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


def consultar_vision_deepseek(imagen_bytes):
    if not clave_deepseek:
        raise ErrorOCR("No hay DEEPSEEK_API_KEY configurada")

    b64 = base64.b64encode(imagen_bytes).decode()
    payload = {
        "model": MODELO_VISION_DEEPSEEK,
        "temperature": 0.0,
        "max_tokens": 8192,  # deepseek-v4-flash-vision-exp razona antes de responder
        # (campo "reasoning_content" aparte de "content", confirmado en vivo el 31/08:
        # ~600 tokens de razonamiento incluso con una imagen en blanco) — con 2048 el
        # modelo podía gastar todo el tope pensando y nunca llegar a escribir el JSON:
        # consumía tokens de entrada/salida igual, pero "content" quedaba vacío.
        "response_format": {"type": "json_object"},
        "messages": [
            {"role": "system", "content": PROMPT_SISTEMA},
            {"role": "user", "content": [
                {"type": "text", "text": "Extrae la sede y la lista completa de artículos pedidos de esta imagen siguiendo exactamente las reglas. Responde solo con el JSON, nada más."},
                {"type": "image_url", "image_url": {"url": f"data:image/png;base64,{b64}"}},
            ]},
        ],
    }
    texto = None
    ultimo_error = None
    # Mismo criterio de reintento que consultar_vision_google: 429/503
    # valen la pena reintentar con espera — sin esto un límite de
    # peticiones de DeepSeek se reportaba como "falló" directo (reportado
    # en vivo, 31/08).
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
    en el resultado indica cuál de los tres respondió."""
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
            time.sleep(4)
        nombre = f"{nombre_archivo}-pagina{indice}.png" if extension == "pdf" else nombre_archivo
        resultados.append(procesar_imagen(imagen_bytes, nombre, indice))
    return resultados
