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
  - nombre_proveedor   (14/09: Nombre o Razón Social del bloque
                        "Proveedor"/"Sujeto Retenido" — nuestra propia
                        empresa, normalmente CRIST MEDICALS, C.A. Se usa
                        para verificar que el comprobante identifica bien
                        a la empresa, ver verificar_retencion() más abajo)
  - rif_proveedor      (RIF de ese mismo bloque Proveedor)
  - periodo_fiscal     (caja superior del comprobante, junto a "Fecha de
                        Emisión", ej. "MARZO/2024")
  - fecha_emision_rec  (14/09: "Fecha de Emisión" de la MISMA caja que
                        Periodo Fiscal — la fecha en que el Agente de
                        Retención emitió este comprobante. Campo nuevo,
                        agregado aparte de "fecha" para no redefinir ese
                        campo ya consumido por api_retenciones_ocr.py)
  - detalle            LISTA con una fila por cada factura de la tabla
                        "Compras Internas e Importaciones" del comprobante
                        (antes solo se leía el resumen de arriba y se
                        perdía este desglose — reportado en vivo, 31/08).
                        Cada fila trae: numero_factura, numero_control,
                        numero_nota_debito, numero_nota_credito,
                        tipo_trans, total_compras_con_iva,
                        total_compras_sin_iva, base_imponible,
                        porcentaje_alicuota, impuesto_iva, iva_retenido —
                        tal cual esa fila de la tabla.

Verificación de retención (14/09, a pedido explícito del usuario, según el
Art. 16 de la Providencia SENIAT sobre agentes de retención de IVA): ver
verificar_retencion() más abajo — confirma que el comprobante identifica
bien a nuestra empresa como "Proveedor" (nombre + RIF) y que trae un
Periodo Fiscal legible. servidor_retenciones.py usa esto para NO guardar
en el histórico automático un comprobante que no pasa la verificación,
dejándolo para revisión manual contra el documento original.

Motor principal: DeepSeek (deepseek-v4-flash-vision-exp) — modelo de
visión experimental de DeepSeek, lanzado el 21/08/2026. OJO: comprime cada
imagen a ~800x800px / 384 tokens antes de leerla, así que en comprobantes
con letra chica puede rendir peor que Gemini. Si falla o el resultado no
sirve, respaldo automático a Gemini (gemini-3.6-flash), y si ese también
falla, a OpenRouter (dots-studio/dots-3-note-preview, gratis). El campo
"motor" en cada resultado indica cuál de los tres respondió.
"""

import base64
import io
import json
import os
import re
import time

import pymupdf
import requests
from dotenv import load_dotenv
from PIL import Image, ImageEnhance, ImageOps

from bin.evaluar_legibilidad import evaluar_legibilidad

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

# Motor principal (09/09, tercera vuelta): nvidia/Qwen2.5-VL-7B-Instruct-NVFP4
# (cuantificado, servido vía vLLM en el mismo puerto 8001 — reemplaza al
# Qwen3-VL-30B-A3B que había antes ahí) — sin costo por token ni
# rate-limit de nube, por eso va de primero. Si el servidor local no
# responde (apagado, red caída), cae a DeepSeek -> Gemini -> OpenRouter
# igual que antes.
#
# REVERTIDO 09/09: se probó un puente GOT-OCR (extracción de texto, 8002)
# + Qwen texto (8000) como motor local alternativo mientras este puerto
# 8001 estaba caído — a pedido explícito del usuario, ya que este modelo
# nuevo "es mucho más prometedor", se sacó ese puente por completo (ni de
# primero ni de respaldo) y se volvió al esquema simple de siempre.
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

    dpi=300 (09/09, subido desde 200): en una prueba real con el motor
    nuevo (Qwen2.5-VL-7B), los errores de lectura salieron justo en la
    página que venía de un PDF de peor calidad — más resolución al
    renderizar le da al modelo más detalle para leer números chicos."""
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


# DPI usado cuando el usuario marca el checkbox "Alta resolución" en la
# pantalla (09/09) — para PDFs escaneados con letra chica/borrosa. Más
# alto que el DPI normal (300), pero no se fue a 600: imágenes más
# grandes tardan más en subir y en que el modelo las procese, con
# ganancia marginal más allá de este punto para el tamaño típico de estos
# comprobantes.
DPI_ALTA_RESOLUCION = 400


def mejorar_calidad_imagen(imagen_bytes):
    """Realce de contraste y nitidez para documentos borrosos o
    deteriorados (09/09, checkbox "Alta resolución" en la pantalla) — PIL
    puro, ya es dependencia del proyecto (no se agregó numpy/opencv):

      - autocontraste: estira el histograma de la imagen para que el
        negro salga más negro y el blanco más blanco — un documento
        fotocopiado muchas veces suele salir "lavado", con poco contraste
        real entre la tinta y el papel.
      - nitidez: realza bordes para que los dígitos no se vean
        "esponjosos"/corridos.

    Si algo falla acá (imagen corrupta, formato raro), se devuelve la
    imagen ORIGINAL sin tocar — nunca bloquea el escaneo por esto."""
    try:
        imagen = Image.open(io.BytesIO(imagen_bytes))
        imagen_rgb = imagen.convert("RGB")
        imagen_rgb = ImageOps.autocontrast(imagen_rgb, cutoff=1)
        imagen_rgb = ImageEnhance.Contrast(imagen_rgb).enhance(1.15)
        imagen_rgb = ImageEnhance.Sharpness(imagen_rgb).enhance(1.8)
        buffer = io.BytesIO()
        imagen_rgb.save(buffer, format="PNG")
        return buffer.getvalue()
    except Exception:
        return imagen_bytes


# ---------------------------------------------------------------------------
# Extracción con el modelo de visión
# ---------------------------------------------------------------------------

class ErrorOCR(Exception):
    pass


PROMPT_SISTEMA = """Actúa como un sistema avanzado de OCR y auditoría contable especializado en comprobantes de retención de IVA venezolanos (formulados por el SENIAT).
Tu tarea es analizar la imagen de un comprobante y extraer estrictamente 12 campos.

⚠️ ADVERTENCIA CRÍTICA — ESTE DOCUMENTO TIENE DOS BLOQUES DE DATOS DIFERENTES Y CONFUSIÓN ENTRE ELLOS ES EL ERROR MÁS FRECUENTE:

BLOQUE A — "IDENTIFICACIÓN AGENTE DE RETENCIÓN" (arriba del documento): es la farmacia/empresa que EMITE el comprobante y retiene el IVA. ESTE es TU "cliente". Extrae sus datos para los campos "cliente" y "rif_cliente".

BLOQUE B — "IDENTIFICACIÓN DEL PROVEEDOR" o "SUJETO RETENIDO" (debajo, cerca de la mitad): es tu propia empresa (normalmente CRIST MEDICALS, C.A.). NUNCA uses sus datos para "cliente" ni "rif_cliente".

REGLAS DE EXTRACCIÓN (lee CADA regla completa, no te saltes la parte de "NO uses"):
1. "fecha": Busca "Fecha de emisión". Formato DD/MM/YYYY.

2. "nro_comprobante": Busca el campo "Nro de Comprobante" que aparece en el encabezado del documento (generalmente arriba, a la izquierda o centro). Es un código de ~14 dígitos que suele empezar con el año (ej. 20260800001289). NO confundas con el "Nro de Factura" (que va en la tabla) ni con el "Nro de Control" de una factura. El "Nro de Comprobante" es ÚNICO por documento, no hay varios.

3. "cliente": Busca SOLO en el bloque "IDENTIFICACIÓN AGENTE DE RETENCIÓN" → campo "Nombre o Razón Social". NO uses el nombre del bloque "Proveedor" ni "Sujeto Retenido".

4. "nro_factura": En la tabla "Compras Internas e Importaciones", extrae TODOS los valores de la columna "Nro de Factura", separados por coma en orden.

5. "rif_cliente": Busca SOLO en el bloque "IDENTIFICACIÓN AGENTE DE RETENCIÓN" → campo RIF (formato J-XXXXXXXX-X o V-XXXXXXXX-X). Este RIF pertenece al MISMO emisor que pusiste en "cliente". ⛔ PROHIBIDO usar el RIF del bloque "Proveedor" o "Sujeto Retenido" — ese es el RIF de tu propia empresa, NO es el dato correcto. Si el campo "cliente" es "FARMACIA X", el "rif_cliente" debe ser el RIF de "FARMACIA X", NO el RIF de "CRIST MEDICALS".

6. "monto_retenido": Busca el total de "Impuesto Retenido I.V.A." del comprobante (suma total, no una fila individual).

FORMATO DE MONTOS (venezolano): PUNTO = miles, COMA = decimales (ej. 11.035,00). Transcribe EXACTAMENTE como aparece. NO conviertas.

7. "nombre_proveedor": Busca SOLO en el bloque "IDENTIFICACIÓN DEL PROVEEDOR"/"SUJETO RETENIDO" (BLOQUE B) → campo "Nombre o Razón Social". Es nuestra propia empresa (normalmente CRIST MEDICALS, C.A.) — transcribe EXACTAMENTE lo que veas impreso, no asumas ni completes de memoria.

8. "rif_proveedor": RIF del MISMO bloque Proveedor/Sujeto Retenido (formato J-XXXXXXXX-X). ⛔ NUNCA el RIF del bloque "Agente de Retención" — ese va en "rif_cliente".

9. "periodo_fiscal": Busca el campo "Periodo Fiscal" en la caja superior del comprobante, normalmente junto a "Fecha de Emisión" (ej. "MARZO/2024" o "03/2024").

10. "fecha_emision_rec": Busca "Fecha de Emisión" en la MISMA caja que "Periodo Fiscal" (arriba del comprobante) — es la fecha en que el Agente de Retención emitió este comprobante. Formato DD/MM/YYYY. Es un campo distinto de "fecha" (que sigue existiendo tal cual, no lo reemplaces).

11. "detalle": Tabla "Compras Internas e Importaciones" completa. UNA fila JSON por fila real de la tabla. Claves EXACTAS (no las inventes):
   - "numero_factura", "numero_control", "numero_nota_debito", "numero_nota_credito"
   - "tipo_trans" (ej. "01")
   - "total_compras_con_iva": columna "Total Servicios y Compras Incluyendo el IVA" — la PRIMERA columna de montos de la tabla, la más ancha/completa. Léela con cuidado dígito por dígito, TODO lo demás de esta fila se calcula a partir de este número.
   - "total_compras_sin_iva": columna "Compras sin Derecho a Crédito IVA" (a veces impresa "CSDAC.IVA") — la SEGUNDA columna de montos, inmediatamente después de la anterior. ⚠️ ESTA ES LA COLUMNA QUE MÁS SE CONFUNDE: si la factura NO tiene ningún monto exento (factura 100% gravada), esta columna aparece VACÍA o en "0,00" en el documento — en ese caso usa null o "0,00", NUNCA copies ahí el valor de "total_compras_con_iva" ni inventes un número parecido. Solo transcribe un valor acá si ves un monto realmente impreso en ESA columna específica, distinto al de la columna anterior.
   - "base_imponible", "porcentaje_alicuota" (ej. "16.00"), "impuesto_iva", "iva_retenido": transcribí lo que veas impreso en esas columnas igual, pero no te preocupes si no cuadran con una fórmula — un programa aparte ya recalcula estos cuatro a partir de "total_compras_con_iva" y "total_compras_sin_iva", así que un error acá no rompe el resultado final. Lo que SÍ tiene que ser exacto son esas dos primeras columnas de montos.

12. "legibilidad_ia": tu PROPIA confianza (número de 0 a 100, sin el símbolo %) en que leíste bien los números de ESTA página. No es una nota de qué tan bonito se ve el documento — es qué tan seguro estás de tu propia lectura de los montos y dígitos, considerando nitidez, resolución, manchas, letra chica borrosa o tablas mal recortadas. 100 = no dudaste en ningún número. 50 o menos = tuviste que adivinar o interpretar varios dígitos poco claros. Sé honesto — si tuviste dudas leyendo algún monto, que el número lo refleje, no pongas 100 por defecto.

ESTRUCTURA: las 12 claves (fecha, nro_comprobante, cliente, nro_factura, rif_cliente, monto_retenido, nombre_proveedor, rif_proveedor, periodo_fiscal, fecha_emision_rec, detalle, legibilidad_ia) van SIEMPRE en un ÚNICO objeto JSON. Nunca omitas una clave (usa null o []).

FORMATO DE SALIDA (ÚNICAMENTE JSON, sin ```json ni explicaciones):
{
  "fecha": "DD/MM/YYYY",
  "nro_comprobante": "14 dígitos del comprobante",
  "cliente": "nombre del Agente de Retención (emisor)",
  "nro_factura": "lista separada por comas",
  "rif_cliente": "RIF del Agente de Retención (mismo que cliente, NO del proveedor)",
  "monto_retenido": "total retenido",
  "nombre_proveedor": "nombre del bloque Proveedor/Sujeto Retenido (normalmente CRIST MEDICALS, C.A.)",
  "rif_proveedor": "RIF del bloque Proveedor/Sujeto Retenido",
  "periodo_fiscal": "ej. MARZO/2024",
  "fecha_emision_rec": "DD/MM/YYYY, de la caja junto a Periodo Fiscal",
  "detalle": [ { las 12 claves de la tabla } ],
  "legibilidad_ia": "0 a 100, tu confianza en tu propia lectura"
}"""

CAMPOS_ESPERADOS = (
    "fecha", "nro_comprobante", "cliente", "nro_factura", "rif_cliente", "monto_retenido",
    "nombre_proveedor", "rif_proveedor", "periodo_fiscal", "fecha_emision_rec",
    "legibilidad_ia",
)

CAMPOS_DETALLE = (
    "numero_factura", "numero_control", "numero_nota_debito", "numero_nota_credito",
    "tipo_trans", "total_compras_con_iva", "total_compras_sin_iva",
    "base_imponible", "porcentaje_alicuota", "impuesto_iva", "iva_retenido",
)


# Tasas fijas de estos comprobantes — CORREGIDO 08/09 (segunda vuelta): la
# primera versión de este fix usaba la "porcentaje_alicuota" que el modelo
# lee de la tabla en vez de asumirla, pensando que sería más precisa que un
# valor fijo. Reportado en vivo con capturas reales: el modelo confunde esa
# columna y lee "75" (el % de RETENCIÓN) en vez de "16" (el % de IVA) — con
# eso, la fórmula terminaba aplicando 75%×75% en vez de 16%×75%, dando
# montos varias veces más grandes que los reales. El usuario confirmó que
# la alícuota de IVA de estos comprobantes es SIEMPRE 16%, sin excepción —
# se vuelve a un valor fijo (como hacía originalmente central_telefonica_ia)
# y se ignora por completo lo que el modelo diga en "porcentaje_alicuota"
# para el cálculo Y para lo que se muestra (se fuerza a "16,00" en vez de
# repetir un valor que puede venir mal leído).
TASA_IVA = 0.16
PORCENTAJE_RETENCION_IVA = 0.75


def _a_numero(valor):
    """Convierte a float un monto que puede venir como número o como texto
    con separador de miles/decimales estilo venezolano ("1.234,56") o
    estilo inglés ("1234.56"). Devuelve None si no se puede interpretar —
    nunca lanza, para que el recálculo simplemente se salte esa fila."""
    if valor is None:
        return None
    if isinstance(valor, (int, float)):
        return float(valor)
    texto = re.sub(r"[^0-9.,]", "", str(valor))
    if not texto:
        return None
    if "," in texto and "." in texto:
        if texto.rfind(",") > texto.rfind("."):
            texto = texto.replace(".", "").replace(",", ".")
        else:
            texto = texto.replace(",", "")
    elif "," in texto:
        texto = texto.replace(",", ".")
    try:
        return float(texto)
    except ValueError:
        return None


def _formatear_monto_ve(numero):
    """Arma el string con separadores venezolanos (punto miles, coma
    decimal) — mismo formato en que el modelo transcribe los demás montos,
    para no mezclar formatos en la misma columna."""
    return f"{numero:,.2f}".replace(",", "_").replace(".", ",").replace("_", ".")


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
    fila por fila.

    SOLO EXTRACCIÓN (09/09, a pedido explícito del usuario): esta función
    ya NO calcula la retención — deja base_imponible/porcentaje_alicuota/
    impuesto_iva/iva_retenido/monto_retenido tal cual los transcribió el
    motor de visión/OCR, aunque vengan mal. El cálculo real (fórmula de
    oro) es un paso APARTE que dispara el usuario con el botón "Calcular
    retención" en la pantalla, ver calcular_retencion_documento() más
    abajo — antes esto se hacía automático acá mismo, pero eso mezclaba
    "leer el documento" con "hacer la cuenta" en un solo paso que el
    usuario no podía revisar/corregir en el medio."""
    normalizado = {_normalizar_clave(clave): valor for clave, valor in datos.items()}
    campos = {campo: normalizado.get(_normalizar_clave(campo)) for campo in CAMPOS_ESPERADOS}
    campos["detalle"] = _mapear_detalle(normalizado.get(_normalizar_clave("detalle")))
    return campos


def _mapear_detalle(filas_crudas):
    """Tolera lo mismo que _mapear_campos pero fila por fila — si el
    modelo no devolvió "detalle" o no es una lista, [] (nunca se inventa
    una fila que no vino). Deja los 11 campos TAL CUAL los leyó el motor
    de extracción, sin calcular nada (ver nota en _mapear_campos)."""
    if not isinstance(filas_crudas, list):
        return []
    detalle = []
    for fila in filas_crudas:
        if not isinstance(fila, dict):
            continue
        fila_normalizada = {_normalizar_clave(clave): valor for clave, valor in fila.items()}
        item = {campo: fila_normalizada.get(_normalizar_clave(campo)) for campo in CAMPOS_DETALLE}
        detalle.append(item)
    return detalle


def calcular_fila_retencion(item, porcentaje_retencion=PORCENTAJE_RETENCION_IVA):
    """FÓRMULA DE ORO (08/09, provista por el usuario con ejemplos reales
    verificados a mano; ver también "LEEME RETENCION.txt" en la raíz del
    proyecto, mismos dos ejemplos): la Base Imponible nunca se lee directo
    de la columna del documento (esa es justo la que el modelo confunde
    con la columna vecina) — se DERIVA de "Total Servicios y Compras con
    IVA" y "Compras sin Derecho a Crédito IVA" (CSDAC.IVA), dos columnas
    que el modelo lee mucho mejor:

        Base Imponible = (Total con IVA − CSDAC.IVA) / 1.16
        Impuesto IVA   = Base Imponible × 0.16
        IVA Retenido   = Impuesto IVA × % de retención del cliente

    Caso "factura 100% gravada" (sin CSDAC.IVA, factura sin exento): con
    CSDAC.IVA en 0 la misma fórmula ya da el resultado correcto sin
    necesidad de una rama aparte. Verificado con los dos ejemplos reales
    del usuario: 8.309,11 con_iva sin csdac -> 859,56 retenido; 50.841,90
    con_iva - 39.563,55 csdac -> 1.166,73 retenido. Los dos coinciden
    exacto.

    09/09: separado de _mapear_detalle para que el cálculo sea un paso
    APARTE (botón "Calcular retención" en la pantalla) en vez de correr
    automático al extraer — así el usuario puede revisar/corregir los
    montos leídos antes de que se calcule la retención sobre ellos."""
    item = dict(item)
    con_iva = _a_numero(item.get("total_compras_con_iva"))
    sin_iva = _a_numero(item.get("total_compras_sin_iva"))

    # Bug de duplicación conocido: si "Con IVA" y "Sin IVA"/CSDAC salieron
    # EXACTAMENTE iguales, el modelo copió la misma columna en las dos —
    # nunca pueden ser iguales de verdad. Se trata como si CSDAC.IVA no
    # hubiera venido (0), que es la rama "factura 100% gravada" — más
    # seguro que restar un CSDAC que sabemos que está mal, lo que daría
    # Base Imponible = 0.
    csdac_iva = 0.0 if (con_iva is not None and sin_iva is not None and con_iva == sin_iva) else (sin_iva or 0.0)

    base_imponible = (con_iva - csdac_iva) / (1 + TASA_IVA) if con_iva is not None else None
    # Respaldo final: si no se pudo leer "Total Compras Con IVA" en
    # absoluto (ni siquiera ese, el más fácil de leer), no queda otra que
    # confiar en lo que el motor transcribió directo en "base_imponible"
    # — mejor eso que dejar la fila entera en null.
    if base_imponible is None:
        base_imponible = _a_numero(item.get("base_imponible"))

    impuesto_iva = base_imponible * TASA_IVA if base_imponible is not None else None

    if impuesto_iva is not None:
        item["base_imponible"] = _formatear_monto_ve(base_imponible)
        item["porcentaje_alicuota"] = "16,00"
        item["impuesto_iva"] = _formatear_monto_ve(impuesto_iva)
        item["iva_retenido"] = _formatear_monto_ve(impuesto_iva * porcentaje_retencion)

    return item


def calcular_retencion_documento(campos):
    """Punto de entrada del botón "Calcular retención" (09/09): recibe los
    campos YA EXTRAÍDOS de un comprobante (por GOT-OCR+puente, Qwen-VL, o
    lo que el usuario haya corregido a mano en la pantalla) y les aplica
    la fórmula de oro fila por fila — no vuelve a leer ninguna imagen,
    solo hace la cuenta.

    % de retención SIEMPRE 75% (09/09, revertido a pedido explícito del
    usuario): hubo una versión que buscaba el RIF del cliente en Profit
    (tabla clientes.contribu_e) para aplicar 100% a los "contribuyentes
    especiales" — el usuario indicó que esa regla no tiene prioridad acá,
    así que se sacó por completo. Siempre PORCENTAJE_RETENCION_IVA."""
    campos = dict(campos)
    porcentaje_retencion = PORCENTAJE_RETENCION_IVA

    detalle_crudo = campos.get("detalle") or []
    detalle_calculado = [calcular_fila_retencion(fila, porcentaje_retencion) for fila in detalle_crudo]
    campos["detalle"] = detalle_calculado

    # "monto_retenido" (el total del comprobante) recalculado como la suma
    # de "iva_retenido" de cada fila de detalle, ya recalculadas por
    # fórmula arriba — más confiable que la lectura directa del total
    # (número chico, el que más alucina el modelo). Si ninguna fila tiene
    # datos suficientes, se deja el valor que ya traía (nunca se reemplaza
    # por algo peor).
    montos_detalle = [_a_numero(fila.get("iva_retenido")) for fila in detalle_calculado]
    montos_validos = [monto for monto in montos_detalle if monto is not None]
    if montos_validos:
        campos["monto_retenido"] = _formatear_monto_ve(sum(montos_validos))
    return campos


# Datos fiscales propios (14/09, a pedido explícito del usuario, con el
# Art. 16 de la Providencia SENIAT en mano — ver PDF de referencia que
# mandó): en TODO comprobante que llega a este sistema, "nosotros" somos
# siempre el "Proveedor"/"Sujeto Retenido" (Punto 5 de la providencia),
# nunca el Agente de Retención — si el bloque de Proveedor del documento
# no dice esto, es señal de que se subió el comprobante equivocado, o que
# el modelo confundió los dos bloques (ver advertencia en PROMPT_SISTEMA).
RIF_PROVEEDOR_PROPIO = "J412236709"
NOMBRE_PROVEEDOR_PROPIO = "CRIST MEDICALS"


def _normalizar_rif(valor):
    """"J-41223670-9", "j412236709" y "J 412236709" deben compararse
    iguales — se descarta todo lo que no sea letra/número antes de
    comparar contra RIF_PROVEEDOR_PROPIO."""
    return re.sub(r"[^A-Z0-9]", "", str(valor or "").upper())


def _extraer_anio_mes_fecha(fecha_texto):
    """(año, mes) de un texto de fecha "DD/MM/YYYY" (tolera "-" o "." como
    separador) — None si no se pudo interpretar con confianza (mes fuera
    de 1-12 incluido, que suele delatar que se leyó DD/MM al revés)."""
    coincidencia = re.search(r"(\d{1,2})[/\-.](\d{1,2})[/\-.](\d{4})", str(fecha_texto or ""))
    if not coincidencia:
        return None
    mes, anio = int(coincidencia.group(2)), int(coincidencia.group(3))
    if not (1 <= mes <= 12):
        return None
    return (anio, mes)


def _extraer_anio_mes_comprobante(nro_comprobante):
    """(año, mes) de los primeros 6 dígitos del Nº de Comprobante (formato
    AAAAMMSSSSSSSS, Punto 1 del Art. 16 de la providencia SENIAT: los
    primeros 4 dígitos son el año y los 2 siguientes el mes de emisión) —
    None si no hay al menos 6 dígitos o el mes no es válido."""
    digitos = re.sub(r"\D", "", str(nro_comprobante or ""))
    if len(digitos) < 6:
        return None
    anio, mes = int(digitos[:4]), int(digitos[4:6])
    if not (1 <= mes <= 12):
        return None
    return (anio, mes)


# Nº de Factura (14/09, a pedido explícito del usuario): en estos
# comprobantes viene con el formato típico de numeración fiscal
# venezolana "00-XXXXXX" — dos ceros de prefijo (punto de emisión) más un
# secuencial de 6 dígitos, 8 dígitos en total. Se compara sobre el texto
# sin separadores (guiones/espacios) para tolerar que el modelo lo haya
# transcrito con o sin el guión.
PATRON_NUMERO_FACTURA_VALIDO = re.compile(r"^00\d{6}$")


def _formato_valido_numero_factura(numero_factura):
    digitos = re.sub(r"\D", "", str(numero_factura or ""))
    return bool(PATRON_NUMERO_FACTURA_VALIDO.match(digitos))


def verificar_retencion(campos):
    """Verificación de cumplimiento (14/09, a pedido explícito del
    usuario): un comprobante válido debe identificar correctamente a
    Cristmedicals como "Proveedor" (nombre + RIF, Punto 5 de la
    providencia SENIAT), traer un Periodo Fiscal legible, un Nº de
    Comprobante de 14 dígitos (formato AAAAMMSSSSSSSS, Punto 1) cuyo
    año/mes coincida con su Fecha de Emisión — si no coinciden, es señal
    de que el comprobante está anulado o es inválido (caso real
    reportado 14/09: Nº de Comprobante 20260500001274 (período 2026-05)
    contra Fecha 14/06/2025 (2025-06); la retención efectivamente estaba
    anulada) — y cada Nº de Factura de la tabla de detalle con el formato
    "00-XXXXXX" (dos ceros + 6 dígitos).

    Si algo de esto no cuadra, es señal de que el documento está mal
    leído o es, directamente, un comprobante inválido —
    servidor_retenciones.py usa este resultado para NO dar por bueno el
    registro (no lo guarda en el histórico automático) sin que alguien lo
    revise a mano contra el documento original.

    Devuelve {"ok": True, "errores": []} o {"ok": False, "errores": [...]}
    (una entrada en texto plano por cada chequeo que falló) — nunca lanza
    excepción."""
    errores = []

    nombre_proveedor = (campos.get("nombre_proveedor") or "").strip()
    if not nombre_proveedor:
        errores.append("No se pudo leer el nombre/razón social del Proveedor (Punto 5) en el comprobante.")
    elif NOMBRE_PROVEEDOR_PROPIO not in nombre_proveedor.upper():
        errores.append(
            f"El Proveedor del comprobante es \"{nombre_proveedor}\", no \"{NOMBRE_PROVEEDOR_PROPIO}\" — "
            "puede ser el comprobante de otra empresa, o el modelo confundió el bloque de Proveedor con el de Agente de Retención."
        )

    rif_proveedor = (campos.get("rif_proveedor") or "").strip()
    if not rif_proveedor:
        errores.append("No se pudo leer el RIF del Proveedor (Punto 5) en el comprobante.")
    elif _normalizar_rif(rif_proveedor) != RIF_PROVEEDOR_PROPIO:
        errores.append(
            f"El RIF del Proveedor leído es \"{rif_proveedor}\" y no coincide con el nuestro (J-{RIF_PROVEEDOR_PROPIO[1:-1]}-{RIF_PROVEEDOR_PROPIO[-1]})."
        )

    periodo_fiscal = (campos.get("periodo_fiscal") or "").strip()
    if not periodo_fiscal:
        errores.append("No se pudo leer el Periodo Fiscal del comprobante.")

    nro_comprobante = (campos.get("nro_comprobante") or "").strip()
    digitos_comprobante = re.sub(r"\D", "", nro_comprobante)
    if nro_comprobante and len(digitos_comprobante) != 14:
        errores.append(
            f"El Nº de Comprobante (\"{nro_comprobante}\") tiene {len(digitos_comprobante)} dígito(s), "
            "debería tener 14 (formato AAAAMMSSSSSSSS, Punto 1 de la providencia SENIAT)."
        )

    fecha_referencia = (campos.get("fecha_emision_rec") or campos.get("fecha") or "").strip()
    anio_mes_comprobante = _extraer_anio_mes_comprobante(nro_comprobante)
    anio_mes_fecha = _extraer_anio_mes_fecha(fecha_referencia)
    if anio_mes_comprobante and anio_mes_fecha and anio_mes_comprobante != anio_mes_fecha:
        errores.append(
            f"El Nº de Comprobante ({nro_comprobante}) indica período {anio_mes_comprobante[0]}-{anio_mes_comprobante[1]:02d}, "
            f"pero la Fecha de Emisión ({fecha_referencia}) es de {anio_mes_fecha[0]}-{anio_mes_fecha[1]:02d} — "
            "no coinciden, el comprobante puede estar anulado o ser inválido."
        )

    for indice, fila in enumerate(campos.get("detalle") or [], start=1):
        numero_factura_fila = (fila.get("numero_factura") or "").strip() if isinstance(fila, dict) else ""
        if numero_factura_fila and not _formato_valido_numero_factura(numero_factura_fila):
            digitos_factura = re.sub(r"\D", "", numero_factura_fila)
            errores.append(
                f"El Nº de Factura \"{numero_factura_fila}\" (fila {indice} de la tabla) no tiene el formato esperado "
                f"00-XXXXXX (dos ceros + 6 dígitos) — tiene {len(digitos_factura)} dígito(s)."
            )

    return {"ok": not errores, "errores": errores}


# Esquema forzado para Gemini — DeepSeek rechaza esto ("This response_format
# type is unavailable now", confirmado en vivo el 31/08: solo acepta
# response_format json_object, que garantiza JSON válido pero NO garantiza
# qué claves trae — un campo puede faltar sin que eso sea un error de
# formato). Gemini sí soporta responseSchema con "required": fuerza a que
# las 8 claves de arriba y las 12 de cada fila de "detalle" estén SIEMPRE
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
                {"text": "Extrae todos los campos requeridos de esta imagen (incluida la tabla de detalle completa) siguiendo exactamente las reglas. Responde solo con el JSON, nada más."},
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

    # Bug real encontrado en vivo (09/09): esto quedaba FUERA del try/except
    # de arriba — si Gemini respondía 200 pero con un JSON mal formado
    # (truncado, etc.), el JSONDecodeError crudo se escapaba sin volverse
    # ErrorOCR, y como consultar_vision() solo atrapa ErrorOCR en la cadena
    # de respaldo, esto tumbaba TODA la petición en vez de pasar a
    # OpenRouter — un escaneo completo se caía con "network error" en el
    # navegador aunque OpenRouter sí hubiera podido responder bien.
    try:
        datos = _extraer_json(texto)
        campos = _mapear_campos(datos)
    except Exception as error:
        raise ErrorOCR(f"Gemini devolvió un JSON inválido: {error}") from error
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
                {"type": "text", "text": "Extrae todos los campos requeridos de esta imagen (incluida la tabla de detalle completa) siguiendo exactamente las reglas. Responde solo con el JSON, nada más. Usa exactamente estos nombres de clave, sin traducirlos ni cambiarles el formato: fecha, nro_comprobante, cliente, nro_factura, rif_cliente, monto_retenido, nombre_proveedor, rif_proveedor, periodo_fiscal, fecha_emision_rec, detalle (cada fila con numero_factura, numero_control, numero_nota_debito, numero_nota_credito, tipo_trans, total_compras_con_iva, total_compras_sin_iva, base_imponible, porcentaje_alicuota, impuesto_iva, iva_retenido), legibilidad_ia."},
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

    try:
        datos = _extraer_json(texto)
        campos = _mapear_campos(datos)
    except Exception as error:
        raise ErrorOCR(f"OpenRouter devolvió un JSON inválido: {error}") from error
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
                {"type": "text", "text": "Extrae todos los campos requeridos de esta imagen (incluida la tabla de detalle completa) siguiendo exactamente las reglas. Responde solo con el JSON, nada más."},
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
        # ~600 tokens de razonamiento incluso con una imagen en blanco) — con 2048/4096
        # el modelo podía gastar todo el tope pensando y nunca llegar a escribir el
        # JSON: consumía tokens de entrada/salida igual, pero "content" quedaba vacío.
        # Encima ahora la respuesta trae la tabla de detalle completa, no solo el resumen.
        "response_format": {"type": "json_object"},
        "messages": [
            {"role": "system", "content": PROMPT_SISTEMA},
            {"role": "user", "content": [
                {"type": "text", "text": "Extrae todos los campos requeridos de esta imagen (incluida la tabla de detalle completa) siguiendo exactamente las reglas. Responde solo con el JSON, nada más."},
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

    try:
        datos = _extraer_json(texto)
        campos = _mapear_campos(datos)
    except Exception as error:
        raise ErrorOCR(f"DeepSeek devolvió un JSON inválido: {error}") from error
    return campos, texto


def consultar_vision_qwen(imagen_bytes):
    """Qwen3-VL-30B-A3B servido local en LAN (vLLM, compatible OpenAI,
    sin autenticación) — motor principal desde el 03/09. Timeout más largo
    que los motores en la nube (90s en vez de 60s): es hardware propio
    compartido, sin el mismo SLA que un proveedor cloud, así que se le da
    más margen antes de darlo por caído y pasar a DeepSeek."""
    b64 = base64.b64encode(imagen_bytes).decode()
    payload = {
        "model": MODELO_VISION_QWEN,
        "temperature": 0.0,
        "max_tokens": 8192,
        "response_format": {"type": "json_object"},
        "messages": [
            {"role": "system", "content": PROMPT_SISTEMA},
            {"role": "user", "content": [
                {"type": "text", "text": "Extrae todos los campos requeridos de esta imagen (incluida la tabla de detalle completa) siguiendo exactamente las reglas. Responde solo con el JSON, nada más."},
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
        # Mismo criterio que ya usamos para el reasoning_content de DeepSeek:
        # si content viene vacío, se usa reasoning como respaldo antes de
        # darlo por fallado.
        texto = mensaje.get("content") or mensaje.get("reasoning") or ""
    except Exception as error:
        raise ErrorOCR(f"Qwen-VL local falló: {error}") from error

    if not texto:
        raise ErrorOCR("Qwen-VL local no devolvió contenido")

    try:
        datos = _extraer_json(texto)
        campos = _mapear_campos(datos)
    except Exception as error:
        raise ErrorOCR(f"Qwen-VL local devolvió un JSON inválido: {error}") from error
    return campos, texto


def consultar_vision(imagen_bytes):
    """Qwen-VL local primero (sin costo por token ni rate-limit de nube);
    si el servidor local no responde, se reintenta con NVIDIA NIM
    Vision (11/09, reemplaza a DeepSeek — sin saldo en la cuenta),
    después Gemini, y si ese también falla, con OpenRouter. `motor` en
    el resultado indica cuál de los cuatro respondió.

    IMPORTANTE: ningún motor de este archivo calcula la retención — todos
    devuelven los campos tal cual los leyeron (ver _mapear_campos). El
    cálculo (fórmula de oro) es un paso aparte, ver
    calcular_retencion_documento(), disparado por el botón "Calcular
    retención" en la pantalla, no por esta cascada de extracción."""
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
                    raise ErrorOCR(
                        f"{error_qwen} — {error_nim} — {error_gemini} — {error_openrouter}"
                    ) from error_openrouter


def procesar_imagen(imagen_bytes, nombre_archivo, numero_pagina):
    # Puntaje de legibilidad (09/09, sin IA — ver bin/evaluar_legibilidad.py):
    # se calcula SIEMPRE, incluso si la extracción falla, para que la
    # pantalla pueda avisar cuándo conviene revisar el documento original
    # contra lo que salió — nunca bloquea ni cambia el resultado del OCR.
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
