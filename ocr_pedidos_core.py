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

from bin.evaluar_legibilidad import evaluar_legibilidad

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

# NVIDIA NIM Vision (11/09, reemplaza a DeepSeek en la cadena de respaldo
# — DeepSeek se quedó sin saldo en la cuenta, "402 Insufficient Balance"
# confirmado en vivo, a pedido explícito del usuario). Mismas 5 llaves
# que ya usa ARA_PROYECT para esto mismo (ara_vision.py) — acá alcanza
# con un pool simple con rotación por fallo, sin la rotación de modelos
# que tiene ese lado (un solo modelo, el primero de su lista de
# prioridad comprobada).
NVIDIA_NIM_URL = "https://integrate.api.nvidia.com/v1/chat/completions"
MODELO_VISION_NVIDIA = os.environ.get("MODELO_VISION_NVIDIA", "meta/llama-3.2-11b-vision-instruct")
CLAVES_NVIDIA = [os.environ.get(f"NVIDIA_API_KEY_{i}") for i in range(1, 6)]
CLAVES_NVIDIA = [clave for clave in CLAVES_NVIDIA if clave]

DEEPSEEK_URL = "https://api.deepseek.com/chat/completions"
MODELO_VISION_DEEPSEEK = os.environ.get("MODELO_VISION_DEEPSEEK", "deepseek-v4-flash-vision-exp")
clave_deepseek = os.environ.get("DEEPSEEK_API_KEY")

# Motor principal (03/09): Qwen3-VL-30B-A3B servido local en LAN (vLLM,
# compatible OpenAI, sin autenticación) — sin costo por token ni
# rate-limit de nube, por eso va de primero. Si el servidor local no
# responde, cae a DeepSeek -> Gemini -> OpenRouter igual que antes.
QWEN_VL_BASE_URL = os.environ.get("QWEN_VL_BASE_URL", "http://192.168.4.4:8001/v1")
QWEN_VL_URL = f"{QWEN_VL_BASE_URL}/chat/completions"
MODELO_VISION_QWEN = os.environ.get("MODELO_VISION_QWEN", "Qwen2.5-VL-7B")

EXTENSIONES_IMAGEN = {"png", "jpg", "jpeg", "webp", "bmp"}
# Notas de voz (11/09, a pedido explícito del usuario): WhatsApp manda
# .opus/.ogg, iPhone suele grabar .m4a, además mp3/wav/webm por si el
# audio viene de otro lado — bin/transcribir_audio.py decodifica
# cualquiera de estos vía torchaudio.
EXTENSIONES_AUDIO = {"opus", "ogg", "m4a", "mp3", "wav", "webm", "aac", "flac"}
EXTENSIONES_PERMITIDAS = EXTENSIONES_IMAGEN | EXTENSIONES_AUDIO | {"pdf"}


# ---------------------------------------------------------------------------
# PDF -> imágenes (una por página)
# ---------------------------------------------------------------------------

def pdf_a_imagenes(contenido_pdf, dpi=300):
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


PROMPT_SISTEMA = """Actúa como un sistema de OCR especializado en pedidos de una empresa farmacéutica venezolana (Crist Medicals, C.A.). La imagen puede ser UNA DE DOS COSAS DISTINTAS — fijate cuál es antes de extraer nada:

CASO A — NOTA O LISTA DE PEDIDO (el caso más común): uno o más artículos que se necesitan pedir, en una lista/tabla con un renglón por artículo. Puede ser manuscrita (a mano) O impresa/tipeada (una tabla generada por computadora, con columnas prolijas). En una tabla IMPRESA es común que el CÓDIGO DE BARRAS sea una columna propia y bien separada (normalmente la primera columna de la fila, un número de 8 a 13 dígitos, antes de la descripción del artículo) — en ese caso ESE número de esa columna ES el código de barras de esa fila, transcribilo directo en "codigo_barra", no lo ignores ni lo confundas con parte de la descripción.

CASO B — FOTO DE UN PRODUCTO SUELTO (caja, empaque o etiqueta de UN solo artículo, sin ninguna lista escrita a mano): en vez de una nota, alguien fotografió el producto en sí porque no sabe el nombre exacto ni el código para pedirlo. En este caso genera un ÚNICO ítem en "items" con lo que se lea en el empaque — "cantidad" y "pediatrico" quedan en null/false (no hay forma de saberlos de una foto de producto), y el ejecutivo confirma después el código y nombre oficial buscándolo a mano.

Tu tarea es extraer EXACTAMENTE esta estructura, sea CASO A o CASO B:

1. "sede": indica si el pedido es para una SUCURSAL o para la SEDE/CASA PRINCIPAL — busca esa palabra escrita o marcada en algún lugar de la nota (encabezado, esquina, un círculo/check junto a una de las dos opciones). Devuelve exactamente el texto "Sucursal" o "Principal" según corresponda. Si la nota no lo indica en ningún lado (siempre el caso en CASO B), usa null — no lo asumas ni lo adivines.

2. "items": una lista con TODOS los artículos pedidos (CASO A) o el único producto fotografiado (CASO B), en el mismo orden en que aparecen escritos en la nota — no te saltees ninguno, incluso si algún renglón tiene letra difícil. Cada ítem de la lista tiene:
   - "articulo": el nombre y la descripción del artículo, tal cual está escrito o impreso (ej. "Amoxicilina 500mg susp", "Diclofenac gel", "Suero fisiológico 250ml"). En CASO B, usa el nombre comercial/genérico y presentación que se lean en el empaque.
   - "cantidad": la cantidad SOLICITADA de ese artículo puntual (cuánto pide quien hizo la nota), tal cual está escrita — puede ser un número solo, o venir con una unidad ("2 cajas", "x3", "1 caja y media"). Transcribila tal cual, sin inventar una unidad que no esté escrita. ⚠️ NUNCA pongas "1" como relleno/valor por defecto cuando en realidad no pudiste leer la cantidad de ESE renglón — un pedido real casi siempre tiene cantidades DISTINTAS entre artículos (14, 24, 6, 10, 30...), así que si tus 9 o 10 artículos te quedan todos en la misma cantidad, es señal de que no estás leyendo la columna/número real de cada fila, estás adivinando. Mirá con cuidado el número específico de CADA renglón (puede estar en una columna aparte, o pegado al final de la línea) antes de transcribirlo; si de verdad no hay ningún número legible para ese artículo puntual, usa null, NUNCA "1". En CASO B, "cantidad" es SIEMPRE null, sin excepción — un empaque casi siempre trae impreso cuánto CONTIENE la caja (ej. "10 comprimidos", "x30 tabletas", "contenido: 100 unidades"), pero ESO NO ES UNA CANTIDAD PEDIDA, es una característica del producto (va dentro de "articulo" si querés, nunca en "cantidad"). No confundas "cuánto trae la caja" con "cuántas cajas se piden" — en una foto de un solo producto nadie pidió una cantidad todavía.
   - "pediatrico": true SOLO si justo al lado de la descripción de ESE artículo aparece escrita la palabra "pediátrico"/"pediatrico" o una abreviación clara ("ped.", "PED"). Si no aparece nada de eso junto a ese renglón puntual, el valor es false — nunca marques true por asociación con otro renglón ni por suposición. En CASO B, siempre false.
   - "codigo_barra": el número del código de barras de ESE artículo específico, si aparece en la imagen. Dos formas en que puede aparecer:
     (a) TABLA IMPRESA/TIPEADA con una columna dedicada al código de barras (normalmente la primera columna de la fila, 8 a 13 dígitos, antes de la descripción) — ese número de esa columna ES el código de barras, transcribilo directo.
     (b) Foto de un empaque/etiqueta con líneas verticales de código de barras dibujadas — el número IMPRESO DEBAJO de esas líneas es el código de barras.
     ⚠️ NO es el código de barras: código de lote, registro sanitario, código interno tipo "CPE" o similar, fecha de vencimiento, ni el código de barras de OTRO producto que aparezca de fondo en la foto. ⚠️ CADA ARTÍCULO TIENE SU PROPIO CÓDIGO DE BARRAS — dos productos DISTINTOS (ej. un complejo vitamínico y una furosemida) NUNCA pueden compartir el mismo código de barras. Si estás por escribir el mismo número de código de barras en dos o más artículos distintos de la lista, DETENETE: eso significa que encontraste un solo número en algún lado de la imagen (una etiqueta de envío, un sello) y lo estás repitiendo por error — en ese caso usa null en todos esos renglones en vez de repetir el número. Si tenés dudas genuinas de si un número es el código de barras o es otra cosa, usa null — mejor vacío que un código equivocado, pero NO ignores una columna de código de barras clara y bien separada solo por precaución. Transcribí únicamente los dígitos, sin espacios. Si no hay ningún código de barras real en la imagen, usa null — NUNCA inventes un número ni reutilices otro código que hayas visto.

REGLAS GENERALES:
- Si una palabra o parte de un artículo es genuinamente ilegible, transcribe lo que sí se distingue con claridad y no inventes el resto.
- Si la nota no tiene ningún artículo legible, "items" debe ser una lista vacía: [].
- No inventes artículos, cantidades, código de barras ni la sede si no están escritos/impresos en la imagen.

FORMATO DE SALIDA:
Devuelve ÚNICAMENTE un JSON válido con esta forma exacta, sin texto antes ni después, sin bloques de código markdown (```json) y sin explicaciones:
{
  "sede": "Sucursal",
  "items": [
    {"articulo": "...", "cantidad": "...", "pediatrico": false, "codigo_barra": null}
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
        "codigo_barra": _texto_o_none(normalizado.get(_normalizar_clave("codigo_barra"))),
    }


def _anular_codigos_barra_duplicados(items):
    """Un código de barras real nunca se repite entre DOS artículos
    distintos — si el modelo devolvió el mismo código_barra en 2+ items,
    es señal de que encontró un solo número en algún lado de la imagen
    (una etiqueta de envío, un sello) y lo copió por error en todas las
    filas (bug real reportado en vivo, 10/09: 9 artículos distintos, los
    9 con el mismo código de barras). Mismo criterio que el chequeo de
    CSDAC.IVA duplicado en retenciones: mejor null que un dato que
    sabemos que está mal — se anula en TODAS las filas que compartían ese
    valor, no solo en la segunda en adelante."""
    conteo = {}
    for item in items:
        codigo = item.get("codigo_barra")
        if codigo:
            conteo[codigo] = conteo.get(codigo, 0) + 1
    for item in items:
        codigo = item.get("codigo_barra")
        if codigo and conteo.get(codigo, 0) > 1:
            item["codigo_barra"] = None
    return items


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
    items = _anular_codigos_barra_duplicados(items)
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
                    "las reglas (o el único producto si es una foto de un producto suelto, CASO B). Responde solo "
                    "con el JSON, nada más. Usa exactamente estas claves: sede, items "
                    "(cada ítem con articulo, cantidad, pediatrico, codigo_barra)."
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


def consultar_vision_nvidia_nim(imagen_bytes):
    """NVIDIA NIM Vision — respaldo si Qwen-VL local no responde (11/09,
    reemplaza a DeepSeek). Prueba cada key del pool en orden; si una
    devuelve 401/403/429/503 (o falla la conexión), pasa a la
    siguiente sin gastar más tiempo en esa key."""
    if not CLAVES_NVIDIA:
        raise ErrorOCR("No hay ninguna NVIDIA_API_KEY_1..5 configurada")
    b64 = base64.b64encode(imagen_bytes).decode()
    payload = {
        "model": MODELO_VISION_NVIDIA,
        "temperature": 0.0,
        "max_tokens": 4096,
        "messages": [
            {"role": "system", "content": PROMPT_SISTEMA},
            {"role": "user", "content": [
                {"type": "text", "text": "Extrae la sede y la lista completa de artículos pedidos de esta imagen siguiendo exactamente las reglas. Responde solo con el JSON, nada más."},
                {"type": "image_url", "image_url": {"url": f"data:image/png;base64,{b64}"}},
            ]},
        ],
    }
    ultimo_error = None
    for clave in CLAVES_NVIDIA:
        try:
            respuesta = requests.post(NVIDIA_NIM_URL, json=payload, headers={"Authorization": f"Bearer {clave}"}, timeout=30)
            if respuesta.status_code in (401, 403, 429, 503):
                ultimo_error = f"HTTP {respuesta.status_code}"
                continue
            respuesta.raise_for_status()
            texto = respuesta.json()["choices"][0]["message"]["content"]
            if not texto:
                ultimo_error = "sin contenido"
                continue
            datos = _extraer_json(texto)
            campos = _mapear_campos(datos)
            return campos, texto
        except Exception as error:
            ultimo_error = error
            continue
    raise ErrorOCR(f"NVIDIA NIM Vision falló (todas las keys del pool): {ultimo_error}")


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


def consultar_vision_qwen(imagen_bytes):
    """Qwen3-VL-30B-A3B servido local en LAN (vLLM, compatible OpenAI,
    sin autenticación) — motor principal desde el 03/09. Timeout más largo
    que los motores en la nube (90s en vez de 60s): es hardware propio
    compartido, así que se le da más margen antes de pasar a DeepSeek."""
    b64 = base64.b64encode(imagen_bytes).decode()
    payload = {
        "model": MODELO_VISION_QWEN,
        "temperature": 0.0,
        "max_tokens": 4096,
        "response_format": {"type": "json_object"},
        "messages": [
            {"role": "system", "content": PROMPT_SISTEMA},
            {"role": "user", "content": [
                {"type": "text", "text": "Extrae la sede y la lista completa de artículos pedidos de esta imagen siguiendo exactamente las reglas. Responde solo con el JSON, nada más."},
                {"type": "image_url", "image_url": {"url": f"data:image/png;base64,{b64}"}},
            ]},
        ],
    }
    try:
        respuesta = requests.post(QWEN_VL_URL, json=payload, timeout=90)
        respuesta.raise_for_status()
        cuerpo = respuesta.json()
        mensaje = cuerpo["choices"][0]["message"]
        # Bug real encontrado en vivo (08/09): al relanzar el VL MoE en la
        # GDX, esta instancia de vLLM empezó a mandar la respuesta entera
        # dentro de "reasoning" (content queda null) — antes no hacía esto.
        texto = mensaje.get("content") or mensaje.get("reasoning") or ""
    except Exception as error:
        raise ErrorOCR(f"Qwen-VL local falló: {error}") from error

    if not texto:
        raise ErrorOCR("Qwen-VL local no devolvió contenido")

    datos = _extraer_json(texto)
    campos = _mapear_campos(datos)
    return campos, texto


def consultar_vision(imagen_bytes):
    """Qwen-VL local primero (sin costo por token ni rate-limit de nube);
    si el servidor local no responde, se reintenta con NVIDIA NIM
    Vision (11/09, reemplaza a DeepSeek — sin saldo en la cuenta),
    después Gemini, y si ese también falla, con OpenRouter. `motor` en
    el resultado indica cuál de los cuatro respondió."""
    try:
        campos, texto = consultar_vision_qwen(imagen_bytes)
        return campos, texto, "Qwen-VL (local)"
    except ErrorOCR as error_qwen:
        try:
            campos, texto = consultar_vision_nvidia_nim(imagen_bytes)
            return campos, texto, "NVIDIA NIM (respaldo)"
        except ErrorOCR as error_nim:
            try:
                campos, texto = consultar_vision_google(imagen_bytes)
                return campos, texto, "Gemini (respaldo)"
            except ErrorOCR as error_gemini:
                try:
                    campos, texto = consultar_vision_openrouter(imagen_bytes)
                    return campos, texto, "OpenRouter (respaldo)"
                except ErrorOCR as error_openrouter:
                    raise ErrorOCR(f"{error_qwen} — {error_nim} — {error_gemini} — {error_openrouter}") from error_openrouter


# ---------------------------------------------------------------------------
# Notas de voz: transcripción (torchaudio + Whisper) + extracción por texto
# ---------------------------------------------------------------------------

# 11/09, a pedido explícito del usuario: de un audio SOLO interesan
# medicamentos/componentes activos/marcas con su cantidad — nada de sede
# ni pediátrico (una nota de voz casi nunca lo aclara, y no vale la pena
# forzarlo). Si el audio divaga o es muy difícil de entender, el modelo
# tiene que avisarlo ("ambiguo": true) en vez de inventar un pedido —
# eso dispara la tarjeta de "proceda con pedido manual" en la interfaz.
PROMPT_SISTEMA_AUDIO = """Actúa como un sistema de extracción de pedidos a partir de la TRANSCRIPCIÓN de una nota de voz de un cliente/sucursal de una empresa farmacéutica venezolana (Crist Medicals, C.A.).

Tu ÚNICA tarea es identificar los MEDICAMENTOS que se piden — por su nombre comercial (marca) o su principio activo/componente, tal como se digan — junto con la CANTIDAD pedida de cada uno. No extraigas nada más: no inventes sede, no inventes presentación ni miligramaje si no se menciona, no incluyas comentarios del cliente que no sean parte del pedido en sí (saludos, quejas, indicaciones de entrega, etc.).

⚠️ CRITERIO DE AMBIGÜEDAD (más importante que extraer algo a toda costa): la transcripción de una nota de voz puede venir incompleta, con ruido, con partes inconexas, o simplemente el cliente puede divagar sin decir un pedido claro. Si NO podés armar con confianza razonable una lista de artículos y cantidades — porque el audio es muy difícil de entender, se corta, o no queda claro qué se está pidiendo — NO inventes ni fuerces una lista. En ese caso, marca "ambiguo": true y dejá "items" vacío. Es preferible avisar que el audio es ambiguo (para que un humano lo procese a mano) a inventar un pedido equivocado.

REGLAS:
- "articulo": el nombre del medicamento/marca/componente tal como se mencionó, con cualquier detalle que sí se haya dicho (miligramaje, presentación) — sin inventar lo que no se dijo.
- "cantidad": la cantidad pedida de ESE artículo puntual, tal cual se entendió ("10 cajas", "5", "media caja"). Si no se dijo una cantidad clara para ese artículo, usa null — nunca "1" de relleno.
- No repitas el mismo artículo dos veces si el cliente lo repitió sin agregar más pedido (ej. confirmando lo que ya dijo).

FORMATO DE SALIDA:
Devuelve ÚNICAMENTE un JSON válido, sin texto antes ni después, sin bloques de código markdown y sin explicaciones:
{"ambiguo": false, "items": [{"articulo": "...", "cantidad": "..."}]}
Si el audio es demasiado ambiguo para armar un pedido con confianza:
{"ambiguo": true, "items": []}"""


def _mapear_campos_audio(datos):
    """Mismo criterio tolerante a nombres de clave que _mapear_campos,
    adaptado a la forma más chica del audio (sin sede, sin pediátrico ni
    código de barras — nada de eso sale de una nota de voz). Reusa
    _mapear_item: como "pediatrico"/"codigo_barra" no vienen en el JSON
    de audio, _mapear_item ya los deja en False/None por su cuenta, así
    que el resto del sistema (tabla de ítems, lupa, "Montar pedido") no
    necesita saber que este pedido vino de un audio."""
    normalizado = {_normalizar_clave(clave): valor for clave, valor in datos.items()}
    ambiguo = _a_booleano(normalizado.get(_normalizar_clave("ambiguo")))
    items_crudos = normalizado.get(_normalizar_clave("items"))
    items = []
    if isinstance(items_crudos, list):
        for item_crudo in items_crudos:
            item = _mapear_item(item_crudo)
            if item is not None:
                items.append(item)
    return {"sede": None, "items": items, "ambiguo": ambiguo or not items}


def consultar_texto_qwen(texto_transcrito):
    payload = {
        "model": MODELO_VISION_QWEN,
        "temperature": 0.0,
        "max_tokens": 2048,
        "response_format": {"type": "json_object"},
        "messages": [
            {"role": "system", "content": PROMPT_SISTEMA_AUDIO},
            {"role": "user", "content": f'Transcripción de la nota de voz: "{texto_transcrito}"\n\nExtraé la lista de medicamentos y cantidades siguiendo exactamente las reglas. Responde solo con el JSON, nada más.'},
        ],
    }
    try:
        respuesta = requests.post(QWEN_VL_URL, json=payload, timeout=60)
        respuesta.raise_for_status()
        cuerpo = respuesta.json()
        mensaje = cuerpo["choices"][0]["message"]
        texto = mensaje.get("content") or mensaje.get("reasoning") or ""
    except Exception as error:
        raise ErrorOCR(f"Qwen (texto) local falló: {error}") from error
    if not texto:
        raise ErrorOCR("Qwen (texto) local no devolvió contenido")
    return _mapear_campos_audio(_extraer_json(texto)), texto


def consultar_texto_deepseek(texto_transcrito):
    if not clave_deepseek:
        raise ErrorOCR("No hay DEEPSEEK_API_KEY configurada")
    payload = {
        "model": MODELO_VISION_DEEPSEEK,
        "temperature": 0.0,
        "max_tokens": 4096,
        "response_format": {"type": "json_object"},
        "messages": [
            {"role": "system", "content": PROMPT_SISTEMA_AUDIO},
            {"role": "user", "content": f'Transcripción de la nota de voz: "{texto_transcrito}"\n\nExtraé la lista de medicamentos y cantidades siguiendo exactamente las reglas. Responde solo con el JSON, nada más.'},
        ],
    }
    try:
        respuesta = requests.post(DEEPSEEK_URL, json=payload, headers={"Authorization": f"Bearer {clave_deepseek}"}, timeout=60)
        respuesta.raise_for_status()
        texto = respuesta.json()["choices"][0]["message"]["content"]
    except Exception as error:
        raise ErrorOCR(f"DeepSeek (texto) falló: {error}") from error
    if not texto:
        raise ErrorOCR("DeepSeek (texto) no devolvió contenido")
    return _mapear_campos_audio(_extraer_json(texto)), texto


def consultar_texto_google(texto_transcrito):
    url = f"{GOOGLE_BASE_URL}/{MODELO_VISION}:generateContent?key={clave_google}"
    payload = {
        "systemInstruction": {"parts": [{"text": PROMPT_SISTEMA_AUDIO}]},
        "contents": [{"parts": [{"text": f'Transcripción de la nota de voz: "{texto_transcrito}"\n\nExtraé la lista de medicamentos y cantidades siguiendo exactamente las reglas. Responde solo con el JSON, nada más.'}]}],
        "generationConfig": {"temperature": 0.0, "maxOutputTokens": 2048, "responseMimeType": "application/json"},
    }
    try:
        respuesta = requests.post(url, json=payload, timeout=60)
        respuesta.raise_for_status()
        texto = respuesta.json()["candidates"][0]["content"]["parts"][0]["text"]
    except Exception as error:
        raise ErrorOCR(f"Gemini (texto) falló: {error}") from error
    return _mapear_campos_audio(_extraer_json(texto)), texto


def consultar_texto(texto_transcrito):
    """Misma cadena de respaldo que consultar_vision, pero para texto
    (sin imagen — la transcripción del audio ya es texto plano).

    11/09: DeepSeek se sacó de acá también (misma cuenta sin saldo que
    en la cascada de visión, "402 Insufficient Balance") — queda
    Qwen (local) -> Gemini directo, sin escalón intermedio roto."""
    try:
        campos, texto = consultar_texto_qwen(texto_transcrito)
        return campos, texto, "Qwen (texto, local)"
    except ErrorOCR as error_qwen:
        try:
            campos, texto = consultar_texto_google(texto_transcrito)
            return campos, texto, "Gemini (texto, respaldo)"
        except ErrorOCR as error_google:
            raise ErrorOCR(f"{error_qwen} — {error_google}") from error_google


def procesar_audio(audio_bytes, nombre_archivo, numero_pagina):
    """Punto de entrada para una nota de voz: transcribe (torchaudio +
    Whisper) y después extrae medicamentos/cantidades por texto — misma
    forma de resultado que procesar_imagen (campos.items compatible con
    la MISMA tabla/lupa/"Montar pedido" de siempre), más "ambiguo": true
    cuando el audio no da para armar un pedido con confianza (la
    interfaz muestra ahí la tarjeta de "proceda con pedido manual" en
    vez de la tabla de ítems)."""
    from bin.transcribir_audio import transcribir_audio

    resultado_pagina = {"pagina": numero_pagina, "archivo": nombre_archivo, "ok": False}
    resultado_transcripcion = transcribir_audio(audio_bytes, nombre_archivo)
    if not resultado_transcripcion.get("ok"):
        resultado_pagina["error"] = resultado_transcripcion.get("error")
        return resultado_pagina

    texto_transcrito = resultado_transcripcion["texto"]
    try:
        campos, texto_crudo, motor = consultar_texto(texto_transcrito)
    except ErrorOCR as error:
        resultado_pagina["error"] = str(error)
        resultado_pagina["texto_ocr"] = texto_transcrito
        return resultado_pagina

    resultado_pagina.update({
        "ok": True, "campos": campos, "ambiguo": campos.get("ambiguo", False),
        "texto_ocr": f"Transcripción: {texto_transcrito}\n\n{texto_crudo}", "motor": motor,
    })
    return resultado_pagina


def procesar_imagen(imagen_bytes, nombre_archivo, numero_pagina):
    resultado_pagina = {
        "pagina": numero_pagina, "archivo": nombre_archivo, "ok": False,
        "legibilidad": evaluar_legibilidad(imagen_bytes),
    }
    try:
        campos, texto_crudo, motor = consultar_vision(imagen_bytes)
    except ErrorOCR as error:
        resultado_pagina["error"] = str(error)
        return resultado_pagina

    resultado_pagina.update({"ok": True, "campos": campos, "texto_ocr": texto_crudo, "motor": motor})
    return resultado_pagina


def procesar_documento(contenido, nombre_archivo, extension):
    """Punto de entrada de alto nivel: recibe los bytes crudos de un
    archivo (PDF, imagen o audio) y su extensión, y devuelve la lista de
    resultados (uno por página, o uno solo si es audio). No sabe nada de
    HTTP ni de Flask."""
    if extension not in EXTENSIONES_PERMITIDAS:
        raise ValueError(f"Extensión no soportada: .{extension}")

    if extension in EXTENSIONES_AUDIO:
        return [procesar_audio(contenido, nombre_archivo, 1)]

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
