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

Motor principal: DeepSeek (deepseek-v4-flash-vision-exp) — modelo de
visión experimental de DeepSeek, comprime cada imagen a ~800x800px/384
tokens antes de leerla, pero para un solo número de factura (dato chico,
casi siempre el más prominente del encabezado) ese recorte no debería
pesar tanto como en documentos con letra chica en varios campos. Si
falla, respaldo automático a Gemini (gemini-3.6-flash), y si ese también
falla, a OpenRouter (dots-studio/dots-3-note-preview, gratis). El campo
"motor" del resultado indica cuál de los tres respondió.
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


PROMPT_SISTEMA = """Actúa como un sistema de OCR ultra enfocado. Tu único trabajo es extraer datos de una factura de compra de productos psicotrópicos/farmacéuticos.

## 1. NÚMERO DE FACTURA/DOCUMENTO (campo: numero_factura)
Buscá el número de FACTURA o DOCUMENTO.

⚠️  ADVERTENCIA CRÍTICA ⚠️
Estos SÍ son válidos (buscá cualquiera):
- "N° Factura" / "Factura N°" / "Nro. Factura" / "Numero de Factura"
- "N° Documento" / "Documento N°" / "Nro. Documento" / "Numero de Documento" / "DOCUMENTO:"

Estos NO son válidos (ignorarlos):
- "N° Control" / "Control N°" / "Numero de Control" → NO
- "N° Comprobante" / "Numero de Comprobante" → NO
- "Forma Libre" / "N° Forma Libre" → NO

⚠️ ERROR REAL MÁS FRECUENTE (encontrado en vivo con facturas de Megalabs y
similares): en la esquina superior derecha suele imprimirse primero, en
letra GRANDE y a veces en rojo, algo como "FORMA LIBRE — N° DE CONTROL:
00-293932" — ESE NO ES el número que buscás, aunque sea el más grande y
llamativo de la página. El número correcto está un poco más abajo, casi
siempre en un recuadro más chico junto a la fecha, etiquetado
"DOCUMENTO:" (ej. "DOCUMENTO: 404834"). Regla práctica: si ves DOS
números distintos en el encabezado, uno junto a "N° DE CONTROL" (grande,
arriba del todo) y otro junto a "DOCUMENTO" o "FACTURA" (más abajo, cerca
de la fecha), el que corresponde a "numero_factura" es el de
"DOCUMENTO"/"FACTURA", NUNCA el de "N° DE CONTROL" — sin importar cuál se
vea más grande o más prominente.

Si el documento tiene bloques separados ("Agente de Retención" vs "Proveedor"), el número está en el bloque del PROVEEDOR.

DÓNDE BUSCAR: arriba del documento, junto al encabezado del proveedor.

REGLAS para el número:
1. Devuelve solo los DÍGITOS, sin ceros a la izquierda, sin guiones, puntos ni espacios. Ej: "002540944" → 2540944, "254.094" → 254094.
2. Si hay prefijo alfanumérico (ej. "A-00174857"), devolvé solo la parte numérica.
3. Si no encontrás un número claro de factura o documento, usá null.

## 2. NÚMEROS DE LOTE (campo: lotes)
Buscá TODOS los números de lote que aparezcan en la factura. Los lotes suelen estar en la tabla de detalle (ítems/artículos) en columnas llamadas "N° Lote", "Lote", "Nro. Lote", "Lote N°", "No. Lote", "Lote Fab.", "Lote de Fabricación", "N° de Lote", o simplemente "LOTE" (columna angosta, entre "DESCRIPCION" y "VCTO."/"CANT.").

⚠️ NO TE SALTEES ESTA TABLA: la columna "LOTE" suele traer valores CORTOS
(3-4 dígitos, ej. "0031", "0012", "0052") que a simple vista pueden
parecer poco importantes al lado de columnas más anchas como
"DESCRIPCION" o "PRECIO" — igual hay que leerlos, UNO POR CADA FILA de la
tabla, aunque la hoja tenga sellos, marcas de agua o logos de fondo que
tapen un poco el número. Si la tabla tiene 5 artículos, tiene que haber
hasta 5 lotes en la lista (uno por fila que sí traiga lote visible) — no
te conformes con devolver la lista vacía sin haber revisado cada fila.

REGLAS para los lotes:
1. Cada lote es un código alfanumérico (ej. "A2231", "78B456", "LOT-001", "12345-AB", "0031").
2. Devolvé SOLO el valor del lote, sin la etiqueta "Lote" o "N°".
3. Buscá en TODA la tabla de detalle — cada artículo puede tener su propio lote.
4. Si de verdad no hay ninguna columna de lote en el documento (revisaste cada fila y ninguna trae lote), devolvé un arreglo vacío [].

## FORMATO DE SALIDA
Devuelve ÚNICAMENTE un objeto JSON válido, sin texto adicional, sin bloques de código markdown, sin explicaciones:
{
  "numero_factura": 174857,
  "lotes": ["A2231", "B4456", "78C901"]
}
Si no hay número de factura: {"numero_factura": null, "lotes": []}
Si hay factura pero sin lotes: {"numero_factura": 174857, "lotes": []}
"""

CAMPOS_ESPERADOS = ("numero_factura", "lotes")


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
    ("174857", "N° 174857", "174.857") o con basura alrededor - se queda
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
    resultado = {"numero_factura": _a_numero_factura(crudo)}

    # Lotes: el modelo devuelve un array de strings
    lotes_crudos = normalizado.get(_normalizar_clave("lotes"), [])
    if isinstance(lotes_crudos, list):
        lotes = []
        for l in lotes_crudos:
            lote_str = str(l).strip().strip('"').strip("'")
            # Limpiar etiquetas comunes que el modelo a veces incluye
            for prefijo in ("lote", "n", "no", "n°", "nº", "lot"):
                if lote_str.upper().startswith(prefijo.upper()):
                    lote_str = lote_str[len(prefijo):].strip().strip("- ").strip()
            if lote_str:
                lotes.append(lote_str)
        resultado["lotes"] = lotes
    else:
        resultado["lotes"] = []

    return resultado


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
        "max_tokens": 2048,  # subido de 512: ver comentario en consultar_vision_deepseek
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
                {"type": "text", "text": "Extrae el número de factura de esta imagen siguiendo exactamente las reglas. Responde solo con el JSON, nada más."},
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
        # ~600 tokens de razonamiento incluso con una imagen en blanco) — con 512 el
        # modelo gastaba TODO el tope pensando y nunca llegaba a escribir el JSON:
        # consumía tokens de entrada/salida igual, pero "content" quedaba vacío.
        "response_format": {"type": "json_object"},
        "messages": [
            {"role": "system", "content": PROMPT_SISTEMA},
            {"role": "user", "content": [
                {"type": "text", "text": "Extrae el número de factura de esta imagen siguiendo exactamente las reglas. Responde solo con el JSON, nada más."},
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
        "max_tokens": 2048,
        "response_format": {"type": "json_object"},
        "messages": [
            {"role": "system", "content": PROMPT_SISTEMA},
            {"role": "user", "content": [
                {"type": "text", "text": "Extrae el número de factura de esta imagen siguiendo exactamente las reglas. Responde solo con el JSON, nada más."},
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
