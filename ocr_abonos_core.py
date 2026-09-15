"""
ocr_abonos_core.py — Inteligencia de extracción OCR de "Recibos de Cobro"
(abonos/pagos de clientes contra facturas), sin ningún framework web
encima.

Mismo patrón que ocr_retenciones_core.py (Gemini + respaldo OpenRouter,
mismo manejo de PDF/imagen), pero para un documento distinto: la libreta
de recibos correlativos donde la empresa registra cuánto abonó un cliente
y contra qué factura(s), no el comprobante de retención de IVA.

Campos que extrae, por cada página/imagen:
  - numero_recibo       N° de Recibo de Cobro (correlativo impreso)
  - fecha                Fecha de emisión (DD/MM/YYYY)
  - cliente_nombre        Quién paga ("Hemos recibido de")
  - cliente_codigo_o_rif  El RIF o código de cliente anotado en el recibo
                          (en la práctica suele ser el código PROFIT del
                          cliente, no un RIF fiscal real — se extrae tal
                          cual está escrito, sin intentar validarlo)
  - direccion_fiscal
  - monto                 El valor del recuadro "MONTO" (arriba a la derecha)
  - cantidad_texto        Lo escrito a mano en "LA CANTIDAD DE" (puede
                          mezclar monedas, ej. "44$ y 800COP")
  - concepto              Texto completo del campo "CONCEPTO"
  - facturas_referenciadas Números de factura/nota mencionados en el
                          concepto, separados por coma (ej. "00129005,
                          00379732")
  - forma_pago            Qué casillas están marcadas en "FORMA DE PAGO"
                          (COP / DOLARES / BS), separadas por coma

Motor principal: DeepSeek (deepseek-v4-flash-vision-exp) — modelo de
visión experimental de DeepSeek, comprime cada imagen a ~800x800px/384
tokens antes de leerla, así que en letra manuscrita chica puede rendir
peor que Gemini. Si falla, respaldo automático a Gemini
(gemini-3.6-flash), y si ese también falla, a OpenRouter
(dots-studio/dots-3-note-preview, gratis). El campo "motor" del resultado
indica cuál de los tres respondió.
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
# Configuración (comparte credenciales con ocr_retenciones_core.py)
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
# que ya usa ARA_PROYECT para esto mismo (ara_vision.py).
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
EXTENSIONES_PERMITIDAS = EXTENSIONES_IMAGEN | {"pdf"}


# ---------------------------------------------------------------------------
# PDF -> imágenes (una por página)
# ---------------------------------------------------------------------------

def pdf_a_imagenes(contenido_pdf, dpi=300):
    """Convierte cada página de un PDF a una imagen PNG (bytes).

    dpi=300 (09/09, subido desde 200, ver ocr_retenciones_core.py): más
    resolución al renderizar le da al modelo más detalle para leer
    números chicos."""
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


PROMPT_SISTEMA = """Actúa como un sistema avanzado de OCR contable especializado en documentos manuscritos venezolanos.
Tu tarea es analizar la imagen de un "RECIBO DE COBRO" (una libreta de recibos correlativos, escrita a mano, donde la empresa que la emite registra un abono/pago que le hizo un cliente) y extraer estrictamente los siguientes 10 campos, más 2 puntajes de confianza sobre tu propia extracción.

CONTEXTO DEL DOCUMENTO: es un recibo con un número correlativo impreso ("Recibo de Cobro N°") y una fecha de emisión (DIA/MES/AÑO). El campo "HEMOS RECIBIDO DE" identifica al CLIENTE que está pagando/abonando (no a la empresa que emite el recibo). El campo "CONCEPTO" suele explicar contra qué factura(s) o nota(s) se aplica el abono, y a veces menciona descuentos o porcentajes — de ahí hay que extraer los números de factura mencionados.

PLANTILLA DE REFERENCIA — así se ve un "RECIBO DE COBRO" genuino de esta empresa (Crist Medicals, C.A.), de arriba hacia abajo:
- Logo circular con el nombre "Crist Medicals C.A." y el lema "Comprometidos Con Tu Salud", con el RIF de la empresa (J-41223670-9) y sus datos de contacto en la esquina superior derecha (estos datos de encabezado pueden variar levemente entre distintas libretas impresas — no es señal de documento falso).
- Título "RECIBO DE COBRO N°" con un número correlativo impreso (a veces en tinta roja).
- Un recuadro "FECHA DE EMISION" con tres casillas DIA / MES / AÑO, y junto a él un recuadro "MONTO".
- Fila "HEMOS RECIBIDO DE:" (nombre del cliente) con un campo "RIF:" a la derecha.
- Fila "DIRECCION FISCAL:" con un campo "TELF.:" a la derecha.
- Bloque "LA CANTIDAD DE:" (varias líneas, para escribir el monto en letras).
- Bloque "CONCEPTO:" (varias líneas).
- Fila "FORMA DE PAGO:" con tres casillas/óvalos: "COP", "DOLARES", "BS".
- Campo "OBSERVACIONES:".
- Pie con la nota sobre comprobantes de retención, y dos recuadros de firma: "RECIBI CONFORME POR CRIST MEDICALS CA" y "FIRMA Y SELLO DEL CLIENTE".
Si la imagen no tiene esta estructura reconocible (le faltan bloques característicos como el título "RECIBO DE COBRO", el campo "HEMOS RECIBIDO DE", o las casillas de "FORMA DE PAGO"), es probable que NO sea este tipo de documento — refleja eso en "confianza_formato".

REGLAS DE EXTRACCIÓN:
1. "numero_recibo": el número impreso junto a "RECIBO DE COBRO N°" (arriba a la izquierda).
2. "fecha": el campo "FECHA DE EMISIÓN" (casillas DIA/MES/AÑO). Devuelve en formato DD/MM/YYYY.
3. "cliente_nombre": el nombre/razón social escrito en "HEMOS RECIBIDO DE" — es el CLIENTE que paga, no la empresa que emite el recibo.
4. "cliente_codigo_o_rif": lo que esté escrito en el campo "RIF" del recibo (puede ser un RIF fiscal real tipo J-XXXXXXXX-X, o un código de cliente corto tipo "FAR01547" — transcribe tal cual está escrito, sin intentar corregirlo ni validar el formato).
5. "direccion_fiscal": el campo "DIRECCION FISCAL".
6. "monto": el valor dentro del recuadro "MONTO" (arriba a la derecha, junto a la fecha).
7. "cantidad_texto": lo escrito a mano en el campo "LA CANTIDAD DE" — puede incluir varias monedas en el mismo texto (ej. "44$ y 800COP"), transcribe tal cual.
8. "concepto": el texto completo del campo "CONCEPTO", tal cual está escrito (varias líneas si hace falta, uniéndolas con un espacio).
9. "facturas_referenciadas": SOLO los números de factura o nota que aparezcan mencionados dentro del concepto (suelen venir después de "#" o "Factura" o "Nota"), separados por coma. Si no hay ninguno, usa null.
10. "forma_pago": de las casillas "COP", "DOLARES", "BS" al pie del recibo, cuáles tienen una marca/check — devuelve los nombres marcados separados por coma (ej. "COP, DOLARES"). Si ninguna está marcada, usa null.

PUNTAJES DE CONFIANZA (números enteros de 0 a 100, no strings):
11. "confianza_formato": qué tan seguro estás de que esta imagen ES un "RECIBO DE COBRO" genuino con la estructura descrita arriba. 100 = todos los bloques característicos están presentes y reconocibles con claridad. Baja este número si falta algún bloque clave (el título, "HEMOS RECIBIDO DE", "CONCEPTO", o "FORMA DE PAGO"), si la imagen está cortada/incompleta, borrosa al punto de no distinguir la estructura, o de plano no parece este tipo de documento.
12. "confianza_datos": qué tan seguro estás de haber leído CORRECTAMENTE los datos manuscritos que extrajiste (nombre, montos, concepto, etc.). 100 = letra clara, sin ambigüedad en ningún campo. Baja este número si la letra es difícil de leer, hay tachones/manchas sobre el texto, o tuviste que adivinar el valor de algún campo en vez de leerlo con seguridad.

FORMATO DE SALIDA:
Devuelve ÚNICAMENTE un objeto JSON válido con la siguiente estructura, sin texto adicional, sin bloques de código markdown (```json) y sin explicaciones:
{
  "numero_recibo": "valor",
  "fecha": "DD/MM/YYYY",
  "cliente_nombre": "valor",
  "cliente_codigo_o_rif": "valor",
  "direccion_fiscal": "valor",
  "monto": "valor",
  "cantidad_texto": "valor",
  "concepto": "valor",
  "facturas_referenciadas": "valor",
  "forma_pago": "valor",
  "confianza_formato": 0,
  "confianza_datos": 0
}
Si un dato es completamente ilegible o no existe en el recibo, usa el valor null (los dos puntajes de confianza SIEMPRE deben ser un número, nunca null). La letra es manuscrita — haz tu mejor esfuerzo de lectura antes de rendirte a null."""

CAMPOS_DATOS = (
    "numero_recibo", "fecha", "cliente_nombre", "cliente_codigo_o_rif", "direccion_fiscal",
    "monto", "cantidad_texto", "concepto", "facturas_referenciadas", "forma_pago",
)
CAMPOS_CONFIANZA = ("confianza_formato", "confianza_datos")
CAMPOS_ESPERADOS = CAMPOS_DATOS + CAMPOS_CONFIANZA

# Ponderación del índice de confianza combinado: 70% qué tan reconocible
# es la ESTRUCTURA del documento (evita que un papel cualquiera pase como
# válido), 30% qué tan legible fue lo manuscrito. Igual o por encima de
# este umbral, un recibo se considera validado automáticamente.
PESO_CONFIANZA_FORMATO = 0.7
PESO_CONFIANZA_DATOS = 0.3
UMBRAL_VALIDADO = 70


def _a_entero_0_100(valor):
    """Los puntajes de confianza deben ser 0-100 — tolera que el modelo
    los mande como string ("85"), con símbolo de porcentaje ("85%"), o
    fuera de rango, y nunca deja tumbar el resto de la extracción por
    esto: ante cualquier duda, cae a 0 (peor caso, fuerza revisión
    manual en vez de auto-validar algo que no se pudo confirmar)."""
    if valor is None:
        return 0
    try:
        numero = float(str(valor).strip().rstrip("%"))
    except ValueError:
        return 0
    return max(0, min(100, round(numero)))


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
    comparar "cliente_nombre", "clienteNombre", "cliente nombre" como si
    fueran la misma clave."""
    return re.sub(r"[^a-z0-9]", "", clave.lower())


def _mapear_campos(datos):
    """Arma el diccionario de campos esperados tolerando que el modelo no
    respete el nombre exacto en snake_case (el respaldo de OpenRouter en
    particular tiende a variar mayúsculas/espacios)."""
    normalizado = {_normalizar_clave(clave): valor for clave, valor in datos.items()}
    return {campo: normalizado.get(_normalizar_clave(campo)) for campo in CAMPOS_ESPERADOS}


def consultar_vision_google(imagen_bytes):
    b64 = base64.b64encode(imagen_bytes).decode()
    url = f"{GOOGLE_BASE_URL}/{MODELO_VISION}:generateContent?key={clave_google}"
    payload = {
        "systemInstruction": {"parts": [{"text": PROMPT_SISTEMA}]},
        "contents": [{
            "parts": [
                {"text": "Extrae los 10 campos y los 2 puntajes de confianza de esta imagen siguiendo exactamente las reglas. Responde solo con el JSON, nada más."},
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
                {"type": "text", "text": (
                    "Extrae los 10 campos y los 2 puntajes de confianza de esta imagen siguiendo exactamente las "
                    "reglas. Responde solo con el JSON, nada más. Usa exactamente estos nombres de clave, sin "
                    "traducirlos ni cambiarles el formato: numero_recibo, fecha, cliente_nombre, "
                    "cliente_codigo_o_rif, direccion_fiscal, monto, cantidad_texto, concepto, "
                    "facturas_referenciadas, forma_pago, confianza_formato, confianza_datos."
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
        "max_tokens": 2048,
        "messages": [
            {"role": "system", "content": PROMPT_SISTEMA},
            {"role": "user", "content": [
                {"type": "text", "text": "Extrae los 10 campos y los 2 puntajes de confianza de esta imagen siguiendo exactamente las reglas. Responde solo con el JSON, nada más."},
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
                {"type": "text", "text": "Extrae los 10 campos y los 2 puntajes de confianza de esta imagen siguiendo exactamente las reglas. Responde solo con el JSON, nada más."},
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
        "max_tokens": 8192,
        "response_format": {"type": "json_object"},
        "messages": [
            {"role": "system", "content": PROMPT_SISTEMA},
            {"role": "user", "content": [
                {"type": "text", "text": "Extrae los 10 campos y los 2 puntajes de confianza de esta imagen siguiendo exactamente las reglas. Responde solo con el JSON, nada más."},
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


def calcular_confianza(campos):
    """Separa los 2 puntajes de confianza que vienen mezclados con los
    campos de datos (el modelo los devuelve todos juntos en el mismo
    JSON), y calcula el índice combinado — 70% qué tan reconocible es la
    estructura del documento, 30% qué tan legible fue lo manuscrito.
    Devuelve (campos_de_datos_solamente, info_confianza)."""
    campos_datos = {campo: campos.get(campo) for campo in CAMPOS_DATOS}
    formato = _a_entero_0_100(campos.get("confianza_formato"))
    datos = _a_entero_0_100(campos.get("confianza_datos"))
    indice = round(formato * PESO_CONFIANZA_FORMATO + datos * PESO_CONFIANZA_DATOS)
    confianza = {
        "formato": formato,
        "datos": datos,
        "indice": indice,
        "validado": indice >= UMBRAL_VALIDADO,
    }
    return campos_datos, confianza


def procesar_imagen(imagen_bytes, nombre_archivo, numero_pagina):
    # Puntaje de legibilidad (09/09, sin IA — ver bin/evaluar_legibilidad.py):
    # se calcula SIEMPRE, incluso si la extracción falla, para poder avisar
    # cuándo conviene revisar el documento original contra lo que salió.
    resultado_pagina = {
        "pagina": numero_pagina, "archivo": nombre_archivo, "ok": False,
        "legibilidad": evaluar_legibilidad(imagen_bytes),
    }
    try:
        campos, texto_crudo, motor = consultar_vision(imagen_bytes)
    except ErrorOCR as error:
        resultado_pagina["error"] = str(error)
        return resultado_pagina

    campos_datos, confianza = calcular_confianza(campos)
    resultado_pagina.update({
        "ok": True, "campos": campos_datos, "confianza": confianza,
        "texto_ocr": texto_crudo, "motor": motor,
    })
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
