"""
servidor_ocr_ps.py — OCR-PS: escáner de número de factura para la base de
datos de control de Psicotrópicos.

El gerente de compras sube las facturas de productos psicotrópicos y el
único dato que hace falta cruzar contra esa base de datos es el número de
factura de cada una. Este módulo existe para eso y nada más: subir, leer,
copiar la lista de números.

La inteligencia de extracción (Gemini + respaldo OpenRouter) vive en
ocr_ps_core.py, compartida con la API pública (api_retenciones_ocr.py).
Este archivo solo se encarga de la interfaz web propia: la página HTML,
exportar a Excel al vuelo, y el histórico acumulado en
/data/ocr_ps_historico.xlsx — mismo patrón que servidor_retenciones.py y
servidor_abonos.py.

Ejecutar:
    python servidor_ocr_ps.py
"""

import json
import mimetypes
import os
import re
import threading
import time
import uuid
from datetime import datetime

import jwt
import pymupdf
import requests
from dotenv import load_dotenv
from flask import Flask, Response, jsonify, redirect, render_template, request, send_file, session
from openpyxl import Workbook, load_workbook
from openpyxl.utils import get_column_letter
from io import BytesIO

import bd_ocr
import ocr_pedidos_core as ocr_pedidos
import ocr_ps_core as ocr
from bin.consultar_cliente_profit import buscar_cliente
from bin.consultar_nota_entrega_ara import obtener_notas_de_cotizacion
from bin.escribir_comentario_nota_ara import escribir_comentario_nota
from bin.consultar_factura_pdf_ara import obtener_factura_pdf_de_nota
from bin.detectar_psicotropicos import revisar_items_psicotropicos
from bin.leer_renglones_nota_ara import leer_renglones_nota
from bin.generar_proforma_nota import generar_pdf_proforma_nota
from bin.consultar_precio_profit import obtener_precio, obtener_descuento_cliente
from bin.consultar_producto_ara import buscar_producto
from bin.servir_imagen_producto import buscar_imagen_producto
from bin.consultar_proveedor_profit import buscar_proveedor

# ---------------------------------------------------------------------------
# Configuración
# ---------------------------------------------------------------------------

# override=True: sin esto, load_dotenv() NUNCA pisa una variable que ya
# exista en el entorno del proceso — y ERP_SSO_SECRET real (bug encontrado
# en vivo 07/09) también quedó guardada como variable de USUARIO de Windows
# en esta máquina (la de ARA_PROYECT, distinta a la nuestra) desde que se
# leyó para diagnosticar el SSO. Sin override, el server seguía verificando
# contra el secreto de ARA sin importar lo que dijera este .env.
load_dotenv(override=True)

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 30 * 1024 * 1024  # 30 MB
# Firma la cookie de sesión (login SSO, ver más abajo) — sin esto Flask no
# puede mantener session[...] entre requests. Fija por env var para que la
# sesión sobreviva a un reinicio del proceso; si no está configurada, se usa
# una clave aleatoria por arranque (las sesiones activas se cierran solas en
# cada reinicio, pero el server sigue funcionando).
app.config["SECRET_KEY"] = os.environ.get("FLASK_SECRET_KEY") or os.urandom(32)
# Esta app se embebe como iframe cross-origin dentro del CRM (ver
# agregar_cors más abajo) — sin esto, la cookie de sesión se pone bien en
# el login pero el navegador la descarta en los fetch() posteriores desde
# adentro del iframe (SameSite=Lax, el default de Flask, bloquea cookies
# en contextos "cross-site" como un iframe de otro dominio). Resultado
# real reportado (15/09, migración a la .23): "ya inicié sesión y no lo
# toma" — usuario_activo() siempre None después del primer request.
# SameSite=None EXIGE Secure=True (si no, el navegador rechaza la cookie
# directamente) — el sitio ya se sirve por HTTPS (ocr.cristmedicals.com),
# así que no hay downside.
app.config["SESSION_COOKIE_SAMESITE"] = "None"
app.config["SESSION_COOKIE_SECURE"] = True


@app.after_request
def agregar_cors(respuesta):
    """Permite que el CRM u otros orígenes accedan a los endpoints de este
    servidor. Sin esto, los iframes anidados (CRM → OCR-PS → PDF) fallan
    con 404/CORS en otras máquinas."""
    origen = request.headers.get("Origin", "")
    if origen:
        respuesta.headers["Access-Control-Allow-Origin"] = origen
    respuesta.headers["Access-Control-Allow-Methods"] = "GET, POST, OPTIONS"
    respuesta.headers["Access-Control-Allow-Headers"] = "Content-Type"
    respuesta.headers["Access-Control-Allow-Credentials"] = "true"
    return respuesta

# ---------------------------------------------------------------------------
# SSO del ERP CristMedicals (07/09) — mismo mecanismo que ARA_PROYECT
# (ara/ARA_Brain/sso_erp_routes.py): el ERP firma un JWT HS256 (60s de vida,
# iss="cristmedicals-erp") con los datos del empleado, y redirige acá a
# /auth/sso?token=<JWT>. Se verifica y se guarda el usuario activo en la
# sesión de Flask (cookie firmada del lado del server) — a diferencia de
# ARA (SPA con localStorage), acá no hace falta el paso extra de código de
# canje: esto es una app clásica server-rendered, la sesión ya vive segura
# del lado del servidor.
ERP_SSO_SECRET = os.environ.get("ERP_SSO_SECRET", "").strip()
ERP_SSO_ISSUER = "cristmedicals-erp"

DIRECTORIO_BASE = os.path.dirname(os.path.abspath(__file__))
RUTA_HISTORICO = os.path.join(DIRECTORIO_BASE, "data", "ocr_ps_historico.xlsx")
RUTA_DOCUMENTOS = os.path.join(DIRECTORIO_BASE, "data", "documentos_ps")
ENCABEZADOS_HISTORICO = [
    "Fecha de registro", "Motor", "Numero de Factura",
    "Codigo de Proveedor", "Nombre de Proveedor", "RIF de Proveedor",
    "Ref. Documento", "Lotes",
]
_lock_historico = threading.Lock()

# Pedidos OCR — mismo puerto que OCR-PS (ver servidor_ocr_ps.py arriba),
# módulo aparte para leer notas de entrega manuscritas y pasarlas a una
# tabla limpia (Sede/Articulo/Cantidad/Pediatrico). Una fila del histórico
# = un ARTÍCULO, no una página (una nota trae varios artículos).
RUTA_HISTORICO_PEDIDOS = os.path.join(DIRECTORIO_BASE, "data", "pedidos_ocr_historico.xlsx")
ENCABEZADOS_HISTORICO_PEDIDOS = [
    "Fecha de registro", "Motor", "Pagina", "Archivo", "Sede", "Articulo", "Cantidad", "Pediatrico",
    "Codigo",  # agregado 03/09 (búsqueda contra el maestro de ARA_PROYECT) — va al final para
    # no correr las columnas de un histórico que ya tenía filas reales (ver migración más abajo)
]
_lock_historico_pedidos = threading.Lock()


def _abrir_o_crear_historico():
    if os.path.exists(RUTA_HISTORICO):
        libro = load_workbook(RUTA_HISTORICO)
        hoja = libro.active
        assert hoja is not None
        return libro, hoja

    os.makedirs(os.path.dirname(RUTA_HISTORICO), exist_ok=True)
    libro = Workbook()
    hoja = libro.active
    assert hoja is not None
    hoja.title = "Histórico OCR-PS"
    hoja.append(ENCABEZADOS_HISTORICO)
    for indice, encabezado in enumerate(ENCABEZADOS_HISTORICO, start=1):
        hoja.column_dimensions[get_column_letter(indice)].width = max(16, len(encabezado) + 2)
    hoja.freeze_panes = "A2"
    return libro, hoja


EXTENSIONES_IMAGEN_A_CONVERTIR = {"png", "jpg", "jpeg", "webp", "bmp"}


def _convertir_imagen_a_pdf(contenido, extension):
    """Convierte una imagen (PNG/JPG/WEBP/BMP) a PDF de una página usando
    pymupdf (ya es dependencia del proyecto, no hizo falta instalar nada
    nuevo). Devuelve los bytes del PDF, o el contenido ORIGINAL sin tocar
    si la conversión falla (nunca bloquea el guardado por esto)."""
    try:
        documento = pymupdf.open(stream=contenido, filetype=extension)
        try:
            return documento.convert_to_pdf()
        finally:
            documento.close()
    except Exception as error:
        print(f"[documentos ocr-ps] No se pudo convertir la imagen a PDF: {error}")
        return contenido


def guardar_documento_original(contenido, nombre_archivo):
    """Guarda en disco el documento TAL CUAL se subió (no las páginas ya
    convertidas a imagen para el OCR) — es la evidencia del documento
    original. Un mismo archivo subido genera un solo nombre guardado acá,
    compartido por todas las páginas/filas que salgan de él, para que
    'Cargar al servidor' después lo suba una sola vez a la BD (ver
    bd_ocr.obtener_o_guardar_documento_ps). Devuelve el nombre guardado
    (para anotarlo en el histórico) o None si no se pudo escribir.

    09/09 (a pedido explícito del usuario): si lo que se subió es una
    imagen (PNG/JPG/WEBP/BMP), se convierte a PDF ANTES de guardar — de
    este servidor solo deben salir PDFs, nunca fotos sueltas, para que
    documentos_ps en la BD quede uniforme."""
    extension = nombre_archivo.rsplit(".", 1)[-1].lower() if "." in nombre_archivo else "bin"
    if extension in EXTENSIONES_IMAGEN_A_CONVERTIR:
        contenido = _convertir_imagen_a_pdf(contenido, extension)
        extension = "pdf"
    try:
        os.makedirs(RUTA_DOCUMENTOS, exist_ok=True)
        nombre_guardado = f"{datetime.now().strftime('%Y%m%d_%H%M%S')}_{uuid.uuid4().hex[:8]}.{extension}"
        with open(os.path.join(RUTA_DOCUMENTOS, nombre_guardado), "wb") as f:
            f.write(contenido)
        return nombre_guardado
    except OSError as error:
        print(f"[documentos ocr-ps] No se pudo guardar el archivo original: {error}")
        return None


def guardar_en_historico_global(motor, campos, ref_documento=None, lotes=None):
    """Agrega una fila al Excel histórico en /data. Protegido con un lock
    para que dos páginas procesándose casi al mismo tiempo no se pisen al
    abrir/guardar el archivo. Si falla, no se interrumpe el escaneo — el
    dato sigue disponible igual en la tabla de la página.
    Devuelve el número de fila donde quedó (para poder corregirla después
    desde la página, sin tocar el resto del histórico), o None si falló."""
    with _lock_historico:
        try:
            libro, hoja = _abrir_o_crear_historico()
            hoja.append([
                datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                motor,
                campos.get("numero_factura"),
                campos.get("codigo_proveedor"),
                campos.get("nombre_proveedor"),
                campos.get("rif_proveedor"),
                ref_documento,
                "; ".join(lotes) if lotes else None,
            ])
            numero_fila = hoja.max_row
            libro.save(RUTA_HISTORICO)
            return numero_fila
        except Exception as error:
            print(f"[historico ocr-ps] No se pudo guardar la fila en {RUTA_HISTORICO}: {error}")
            return None
    # Ojo: acá NO se guarda en la base de datos MySQL — eso solo pasa
    # cuando alguien revisa el Excel y pulsa "Cargar al servidor" (ver
    # /api/historico/cargar_al_servidor más abajo). El Excel es el
    # borrador; el servidor solo recibe lo ya revisado.


def procesar_imagen_y_guardar(imagen_bytes, nombre_archivo, numero_pagina, ref_documento=None):
    resultado = ocr.procesar_imagen(imagen_bytes, nombre_archivo, numero_pagina)
    if resultado.get("ok"):
        lotes = resultado.get("campos", {}).get("lotes", [])
        numero_fila = guardar_en_historico_global(resultado["motor"], resultado["campos"], ref_documento, lotes=lotes)
        resultado["fila_historico"] = numero_fila
    return resultado


# ---------------------------------------------------------------------------
# Pedidos OCR — mismo lock/patrón de histórico que arriba, pero una fila
# del histórico es un ARTÍCULO (una nota trae varios), no una página.
# ---------------------------------------------------------------------------

def _abrir_o_crear_historico_pedidos():
    if os.path.exists(RUTA_HISTORICO_PEDIDOS):
        libro = load_workbook(RUTA_HISTORICO_PEDIDOS)
        hoja = libro.active
        assert hoja is not None
        # Migración: la columna "Codigo" se agregó el 03/09 — un histórico
        # creado antes de eso todavía tiene el encabezado viejo (8
        # columnas). Se agrega el título que falta en la 9na para que las
        # filas nuevas (que sí mandan el código) no queden sin encabezado.
        if hoja.cell(row=1, column=len(ENCABEZADOS_HISTORICO_PEDIDOS)).value != "Codigo":
            hoja.cell(row=1, column=len(ENCABEZADOS_HISTORICO_PEDIDOS), value="Codigo")
            libro.save(RUTA_HISTORICO_PEDIDOS)
        return libro, hoja

    os.makedirs(os.path.dirname(RUTA_HISTORICO_PEDIDOS), exist_ok=True)
    libro = Workbook()
    hoja = libro.active
    assert hoja is not None
    hoja.title = "Histórico Pedidos OCR"
    hoja.append(ENCABEZADOS_HISTORICO_PEDIDOS)
    for indice, encabezado in enumerate(ENCABEZADOS_HISTORICO_PEDIDOS, start=1):
        hoja.column_dimensions[get_column_letter(indice)].width = max(16, len(encabezado) + 2)
    hoja.freeze_panes = "A2"
    return libro, hoja


def guardar_en_historico_pedidos(motor, pagina, archivo, sede, items):
    """Agrega UNA fila por artículo al Excel histórico de pedidos — si la
    nota trae 5 artículos, se agregan 5 filas, todas con la misma
    sede/página/archivo. Si `items` viene vacío no agrega nada (no tiene
    sentido una fila de "pedido sin artículos"). Nunca interrumpe el
    escaneo si falla: el dato sigue disponible igual en la pantalla."""
    if not items:
        return
    with _lock_historico_pedidos:
        try:
            libro, hoja = _abrir_o_crear_historico_pedidos()
            fecha = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            for item in items:
                hoja.append([
                    fecha, motor, pagina, archivo, sede,
                    item.get("articulo"), item.get("cantidad"),
                    "Si" if item.get("pediatrico") else "No",
                    item.get("codigo"),
                ])
            libro.save(RUTA_HISTORICO_PEDIDOS)
        except Exception as error:
            print(f"[historico pedidos-ocr] No se pudo guardar en {RUTA_HISTORICO_PEDIDOS}: {error}")


def procesar_imagen_pedido_y_guardar(imagen_bytes, nombre_archivo, numero_pagina):
    """Ya NO guarda automático en el Excel histórico al escanear (pedido
    03/09) — Pedidos OCR guarda en bd_ocr.pedidos_ocr, y solo cuando el
    usuario aprieta "Subir" en /api/pedidos/subir, después de revisar y
    corregir la tarjeta. Guardar acá, apenas sale del OCR, dejaba el
    histórico con el texto crudo aunque el usuario lo corrigiera después
    en pantalla — ver guardar_en_historico_pedidos (ya sin uso, se deja
    documentada por si hace falta volver al Excel)."""
    return ocr_pedidos.procesar_imagen(imagen_bytes, nombre_archivo, numero_pagina)


def procesar_audio_pedido_y_guardar(audio_bytes, nombre_archivo, numero_pagina):
    """Mismo criterio que procesar_imagen_pedido_y_guardar de arriba,
    para notas de voz (11/09, a pedido explícito del usuario)."""
    return ocr_pedidos.procesar_audio(audio_bytes, nombre_archivo, numero_pagina)


# ---------------------------------------------------------------------------
# Rutas
# ---------------------------------------------------------------------------

def _destino_seguro(valor, por_defecto="/pedidos"):
    """Solo se acepta una ruta RELATIVA interna ("/", "/pedidos", "/menu")
    como destino post-login — nunca una URL completa ni algo que empiece
    con "//" (evita que un parámetro "next" armado a mano mande a un
    dominio externo, "open redirect"). Si no es válido, cae al valor de
    siempre (Pedidos OCR)."""
    if valor and valor.startswith("/") and not valor.startswith("//"):
        return valor
    return por_defecto


def _reenviar_token_a_sso():
    """Red de seguridad (mismo criterio que ARA_PROYECT): si el ERP/Auth
    Central manda el JWT a la raíz (o a /menu, /pedidos) en vez de a
    /auth/sso, lo reenviamos a la ruta correcta en vez de ignorar el token y
    mostrar la página sin sesión.

    CORREGIDO 10/09 (bug real reportado desde el CRM): antes esto no
    avisaba desde QUÉ página venía el token, así que sso_callback() no
    tenía forma de saber a dónde volver y mandaba siempre a "/pedidos" —
    entrar desde el CRM al link de OCR-PS terminaba abriendo Pedidos OCR.
    Ahora se manda la página de origen como "next" para volver ahí."""
    token = request.args.get("token", "")
    if token:
        return redirect(f"/auth/sso?token={token}&next={request.path}")
    return None


@app.route("/")
def index():
    reenvio = _reenviar_token_a_sso()
    if reenvio:
        return reenvio
    return render_template("ocr_ps.html")


@app.route("/menu")
def menu():
    reenvio = _reenviar_token_a_sso()
    if reenvio:
        return reenvio
    return render_template("menu.html")


# ---------------------------------------------------------------------------
# Login SSO contra el ERP CristMedicals — ver ERP_SSO_SECRET arriba.
# ---------------------------------------------------------------------------

def _verificar_jwt_erp(token: str) -> dict:
    """Decodifica y verifica el JWT del ERP (firma/exp/iss). Lanza
    jwt.InvalidTokenError (incluye ExpiredSignatureError) si no es válido,
    o RuntimeError si ERP_SSO_SECRET no está configurada — mismo criterio
    que sso_erp_routes.py en ARA_PROYECT."""
    if not ERP_SSO_SECRET:
        raise RuntimeError("ERP_SSO_SECRET no configurada")
    return jwt.decode(token, key=ERP_SSO_SECRET, algorithms=["HS256"], issuer=ERP_SSO_ISSUER)


def usuario_activo() -> dict | None:
    """Devuelve {employee_id, nombre, email} si hay una sesión SSO válida en
    esta request, o None. Nunca lanza — un endpoint que necesite bloquear
    sin sesión decide él mismo qué responder."""
    return session.get("usuario_sso")


# Aislamiento por usuario (10/09 "en espera", 11/09 extendido al
# historial de cotizaciones: "no debo saber las cotizaciones de otro
# usuario, excepto la mía" — a pedido explícito del usuario): cada
# empleado solo puede ver SUS PROPIAS cotizaciones/pedidos — excepto un
# único usuario "global" (configurable acá, sin hardcodear el nombre en
# la lógica) que puede ver las de todos. Coincidencia flexible (sin
# mayúsculas/acentos) porque el nombre que manda el ERP no siempre viene
# escrito exactamente igual (visto en vivo: "QUIÑONEZ" una vez,
# "QUINONEZ" otra).
USUARIO_GLOBAL_ESPERA = os.environ.get("USUARIO_GLOBAL_ESPERA", "JONAIBER").strip()


def _normalizar_nombre(texto: str) -> str:
    texto = (texto or "").strip().upper()
    reemplazos = {"Á": "A", "É": "E", "Í": "I", "Ó": "O", "Ú": "U", "Ñ": "N"}
    for origen, destino in reemplazos.items():
        texto = texto.replace(origen, destino)
    return texto


def _es_usuario_global(usuario: dict | None) -> bool:
    """True para el usuario con acceso a TODO (ver USUARIO_GLOBAL_ESPERA
    arriba) — usado tanto para "en espera" como para el historial de
    cotizaciones."""
    if not usuario:
        return False
    return _normalizar_nombre(USUARIO_GLOBAL_ESPERA) in _normalizar_nombre(usuario.get("nombre") or "")


@app.route("/auth/sso", methods=["GET"])
def sso_callback():
    """Callback real documentado por el ERP: redirige acá con ?token=<JWT>
    recién firmado (60s de vida). Guarda el usuario en la sesión de Flask
    (cookie firmada del lado del servidor) y manda de vuelta a la página
    de origen (`next`, ver _reenviar_token_a_sso — "/pedidos" si no vino
    ninguna, mismo comportamiento de siempre)."""
    destino = _destino_seguro(request.args.get("next", ""))
    token = request.args.get("token", "")
    if not token:
        return redirect(f"{destino}?sso_error=sin_token")

    try:
        payload = _verificar_jwt_erp(token)
    except RuntimeError:
        print("[SSO] ERP_SSO_SECRET no configurada — no se puede verificar el token del ERP.")
        return redirect(f"{destino}?sso_error=sso_no_configurado")
    except jwt.ExpiredSignatureError:
        return redirect(f"{destino}?sso_error=expirado")
    except jwt.InvalidTokenError as error:
        print(f"[SSO] Token inválido: {error}")
        return redirect(f"{destino}?sso_error=invalido")

    employee_id = str(payload.get("employee_id") or "").strip()
    if not employee_id:
        return redirect(f"{destino}?sso_error=sin_employee_id")

    session["usuario_sso"] = {
        "employee_id": employee_id,
        "nombre": str(payload.get("full_name") or employee_id),
        "email": str(payload.get("email") or ""),
    }
    session.permanent = True
    return redirect(destino)


# Bugs reales ya documentados por ARA_PROYECT (mismo ERP, mismo Auth
# Central): a veces el callback llega a /sso (sin el prefijo /auth/) en vez
# de a la ruta documentada — alias en vez de depender de que corrijan la
# configuración del lado del ERP.
@app.route("/sso", methods=["GET"])
def sso_alias_sin_auth():
    token = request.args.get("token", "")
    next_param = request.args.get("next", "")
    if token:
        destino = f"/auth/sso?token={token}"
        if next_param:
            destino += f"&next={next_param}"
        return redirect(destino)
    return redirect("/pedidos?sso_error=sin_token")


# Análogo al alias /ara-inteligente/auth/sso de ARA: Auth Central a veces
# antepone la ruta de la app ("/pedidos") al callback documentado.
@app.route("/pedidos/auth/sso", methods=["GET"])
def sso_alias_pedidos():
    token = request.args.get("token", "")
    next_param = request.args.get("next", "")
    if token:
        destino = f"/auth/sso?token={token}"
        if next_param:
            destino += f"&next={next_param}"
        return redirect(destino)
    return redirect("/pedidos?sso_error=sin_token")


@app.route("/api/sesion", methods=["GET"])
def api_sesion():
    """Para que el frontend sepa si hay un usuario SSO activo y muestre su
    nombre — sin esto, cualquier endpoint que necesite el usuario tendría
    que adivinarlo del lado del cliente."""
    usuario = usuario_activo()
    if not usuario:
        return jsonify({"logueado": False})
    return jsonify({"logueado": True, "nombre": usuario["nombre"], "employee_id": usuario["employee_id"]})


@app.route("/logout", methods=["GET", "POST"])
def logout():
    session.pop("usuario_sso", None)
    return redirect("/pedidos")


@app.route("/api/estado-ocr", methods=["GET"])
def estado_ocr():
    # 15/09: "respaldo" decía "DeepSeek" siempre, hardcodeado — DeepSeek se
    # sacó de la cascada real hace rato (sin saldo en la cuenta,
    # reemplazado por NVIDIA NIM) pero acá nadie había actualizado este
    # texto. consultar_vision() real es Qwen-VL -> NVIDIA NIM -> Gemini ->
    # OpenRouter, ver ocr_ps_core.consultar_vision.
    cadena_respaldo = []
    if ocr.CLAVES_NVIDIA:
        cadena_respaldo.append(f"{ocr.MODELO_VISION_NVIDIA} (NVIDIA NIM)")
    if ocr.clave_google:
        cadena_respaldo.append(f"{ocr.MODELO_VISION} (Gemini)")
    if ocr.clave_openrouter:
        cadena_respaldo.append(f"{ocr.MODELO_VISION_OPENROUTER} (OpenRouter)")
    return jsonify({
        "conectado": True,
        "detalle": {
            "modelo_vision": ocr.MODELO_VISION_QWEN,
            "proveedor": "Qwen-VL (local)",
            "respaldo": " → ".join(cadena_respaldo),
        },
    })


@app.route("/api/escanear", methods=["POST"])
def escanear():
    archivo = request.files.get("archivo")
    if archivo is None or not archivo.filename:
        return jsonify({"error": "Sube un archivo (campo 'archivo')"}), 400

    extension = archivo.filename.rsplit(".", 1)[-1].lower()
    if extension not in ocr.EXTENSIONES_PERMITIDAS:
        return jsonify({"error": f"Extensión no soportada: .{extension}"}), 400

    contenido = archivo.read()
    nombre_archivo = archivo.filename
    ref_documento = guardar_documento_original(contenido, nombre_archivo)

    if extension == "pdf":
        try:
            imagenes = ocr.pdf_a_imagenes(contenido)
        except Exception as error:
            return jsonify({"error": f"No se pudo leer el PDF: {error}"}), 400
        if not imagenes:
            return jsonify({"error": "El PDF no tiene páginas"}), 400
    else:
        imagenes = [contenido]

    def generar():
        try:
            total = len(imagenes)
            yield json.dumps({"evento": "inicio", "total_paginas": total}) + "\n"
            for indice, imagen_bytes in enumerate(imagenes, start=1):
                if indice > 1:
                    time.sleep(4)
                nombre = f"{nombre_archivo}-pagina{indice}.png" if extension == "pdf" else nombre_archivo
                resultado = procesar_imagen_y_guardar(imagen_bytes, nombre, indice, ref_documento)
                yield json.dumps({"evento": "pagina", "resultado": resultado}) + "\n"
            yield json.dumps({"evento": "fin"}) + "\n"
        except (BrokenPipeError, ConnectionAbortedError, ConnectionResetError, GeneratorExit):
            return

    return Response(generar(), mimetype="application/x-ndjson")


@app.route("/api/exportar_excel", methods=["POST"])
def exportar_excel():
    payload = request.get_json(silent=True) or {}
    resultados = payload.get("resultados", [])

    libro = Workbook()
    hoja = libro.active
    assert hoja is not None
    hoja.title = "OCR-PS"

    encabezados = ["Página", "Motor", "Numero de Factura", "Codigo de Proveedor", "Nombre de Proveedor", "RIF de Proveedor"]
    hoja.append(encabezados)

    for fila in resultados:
        campos = fila.get("campos", {})
        hoja.append([
            fila.get("pagina"),
            fila.get("motor"),
            campos.get("numero_factura"),
            campos.get("codigo_proveedor"),
            campos.get("nombre_proveedor"),
            campos.get("rif_proveedor"),
        ])

    for indice, encabezado in enumerate(encabezados, start=1):
        hoja.column_dimensions[get_column_letter(indice)].width = max(16, len(encabezado) + 2)

    buffer = BytesIO()
    libro.save(buffer)
    buffer.seek(0)

    return send_file(
        buffer,
        as_attachment=True,
        download_name="ocr_ps.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


@app.route("/api/proveedores/buscar", methods=["GET"])
def buscar_proveedores():
    """Busca proveedores reales en PROFIT (tabla prov, solo lectura) por
    nombre, código o RIF — el mismo texto se prueba contra los tres, así
    la barra de búsqueda de la tarjeta no necesita saber cuál es cuál.
    Esto es lo que impide que se guarde en la BD un proveedor mal
    tipeado: solo se puede elegir uno de los que devuelve PROFIT."""
    texto = (request.args.get("q") or "").strip()
    if not texto:
        return jsonify({"ok": True, "total": 0, "proveedores": []})
    resultado = buscar_proveedor(nombre=texto, codigo=texto, rif=texto, limite=15)
    if not resultado.get("ok"):
        return jsonify(resultado), 502
    return jsonify(resultado)


@app.route("/api/historico/actualizar_fila", methods=["POST"])
def actualizar_fila_historico():
    """Corrige, en el Excel histórico, una fila que ya se había guardado
    — se usa cuando alguien corrige el número en la tarjeta o elige un
    proveedor después de que el escaneo ya guardó la fila. Sin esto, la
    corrección solo viviría en la pantalla y se perdería apenas se
    recargue o se exporte desde el histórico en vez de la pantalla."""
    payload = request.get_json(silent=True) or {}
    fila = payload.get("fila")
    if not isinstance(fila, int) or fila < 2:
        return jsonify({"error": "Falta 'fila' (número de fila válido del histórico)"}), 400
    if not os.path.exists(RUTA_HISTORICO):
        return jsonify({"error": "Todavía no hay histórico"}), 404

    columnas_editables = {
        "numero_factura": "Numero de Factura",
        "codigo_proveedor": "Codigo de Proveedor",
        "nombre_proveedor": "Nombre de Proveedor",
        "rif_proveedor": "RIF de Proveedor",
        "lotes": "Lotes",
    }
    with _lock_historico:
        try:
            libro = load_workbook(RUTA_HISTORICO)
            hoja = libro.active
            if fila > hoja.max_row:
                return jsonify({"error": "Esa fila ya no existe en el histórico"}), 404
            for campo, encabezado in columnas_editables.items():
                if campo in payload:
                    indice_columna = ENCABEZADOS_HISTORICO.index(encabezado) + 1
                    hoja.cell(row=fila, column=indice_columna).value = payload[campo]
            libro.save(RUTA_HISTORICO)
        except Exception as error:
            return jsonify({"error": f"No se pudo actualizar la fila: {error}"}), 500
    return jsonify({"ok": True})


@app.route("/cotejo")
def cotejo():
    return render_template("cotejo_lotes.html")


@app.route("/api/facturas/buscar", methods=["GET"])
def buscar_facturas():
    """Busca facturas YA CARGADAS al servidor (MySQL) por número — el
    primer paso antes de registrarle un lote a una. Como el mismo número
    de factura puede repetirse entre proveedores distintos, la respuesta
    trae el proveedor de cada una para que la persona elija la correcta."""
    numero_factura = request.args.get("numero_factura", "").strip()
    if not numero_factura or not numero_factura.isdigit():
        return jsonify({"ok": False, "error": "Falta 'numero_factura' (numérico)"}), 400
    facturas = bd_ocr.buscar_facturas_ps(numero_factura=int(numero_factura))
    return jsonify({"ok": True, "facturas": facturas})


@app.route("/api/facturas/<int:numero_factura>/<cod_prov>/documento", methods=["GET"])
def descargar_documento_factura(numero_factura, cod_prov):
    """Recibe numero_factura + código de proveedor por la URL y devuelve
    (descarga) el PDF real vinculado a esa factura.

    CORREGIDO 10/09 (a pedido explícito del usuario): el número de
    factura SOLO no identifica el documento con certeza — distintos
    proveedores pueden usar el mismo número de control/factura, así que
    antes esto podía devolver el PDF de un proveedor equivocado si dos
    facturas coincidían en número. Ahora exige también el código de
    proveedor exacto (mismo criterio ya aplicado en otros módulos: nunca
    confiar en un número corto solo para identificar un documento real).
    El PDF ya queda guardado y vinculado (ocr_ps.documento_id ->
    documentos_ps) automáticamente cuando la factura se sube con "Cargar
    al servidor" — este endpoint solo lo recupera. 404 si no hay ningún
    documento vinculado a esa combinación exacta de número+proveedor."""
    documento = bd_ocr.obtener_documento_ps_por_factura(numero_factura, cod_prov)
    if documento is None:
        return jsonify({
            "ok": False,
            "error": f"No hay ningún documento vinculado a la factura {numero_factura} del proveedor '{cod_prov}'",
        }), 404
    return send_file(
        BytesIO(documento["contenido"]),
        as_attachment=True,
        download_name=documento["nombre_archivo"] or f"factura_{numero_factura}_{cod_prov}.pdf",
        mimetype="application/pdf",
    )


@app.route("/api/lotes/de_factura/<int:factura_id>", methods=["GET"])
def lotes_de_factura(factura_id):
    return jsonify({"ok": True, "lotes": bd_ocr.listar_lotes_de_factura(factura_id)})


@app.route("/api/lotes/registrar", methods=["POST"])
def registrar_lote():
    """La persona asignada anota a mano el lote de un artículo de una
    factura ya cargada. numero_factura se guarda junto con el lote (no
    solo factura_id) para que cotejar_lote pueda buscar directo por la
    clave sin tener que hacer join primero."""
    payload = request.get_json(silent=True) or {}
    factura_id = payload.get("factura_id")
    numero_factura = payload.get("numero_factura")
    lote = (payload.get("lote") or "").strip()
    if not factura_id or not numero_factura or not lote:
        return jsonify({"ok": False, "error": "Faltan 'factura_id', 'numero_factura' o 'lote'"}), 400
    id_nuevo = bd_ocr.guardar_lote(factura_id, numero_factura, lote)
    return jsonify({"ok": True, "id": id_nuevo, "clave": f"{numero_factura}-{lote}"})


@app.route("/api/lotes/cotejar", methods=["GET"])
def cotejar_lote():
    """El cotejamiento: ¿esta combinación factura+lote ya está registrada
    en nuestra base? Si sí, devuelve el contexto completo (proveedor,
    fechas); si no, coincidencias vacío."""
    numero_factura = request.args.get("numero_factura", "").strip()
    lote = request.args.get("lote", "").strip()
    if not numero_factura or not numero_factura.isdigit() or not lote:
        return jsonify({"ok": False, "error": "Faltan 'numero_factura' (numérico) y 'lote'"}), 400
    coincidencias = bd_ocr.cotejar_lote(int(numero_factura), lote)
    return jsonify({"ok": True, "encontrado": len(coincidencias) > 0, "coincidencias": coincidencias})


@app.route("/api/lotes/guardar_con_factura", methods=["POST"])
def guardar_lotes_con_factura():
    """Guarda lotes asociados a una factura por su número. Busca la factura
    más reciente en ocr_ps por numero_factura y crea los registros en
    lotes_ps. Si la factura no existe, se ignora esa entrada (ya no está
    cargada). Devuelve cuántos lotes se registraron por factura."""
    try:
        payload = request.get_json(silent=True) or {}
        facturas_con_lotes = payload.get("facturas_con_lotes", [])
        if not facturas_con_lotes:
            return jsonify({"ok": False, "error": "Faltan 'facturas_con_lotes'"}), 400

        total_registrados = 0
        for entrada in facturas_con_lotes:
            numero_factura = entrada.get("numero_factura")
            lotes = entrada.get("lotes", [])
            if not numero_factura or not lotes:
                continue
            try:
                numero_factura = int(numero_factura)
            except (ValueError, TypeError):
                continue
            conexion = bd_ocr.obtener_conexion()
            try:
                with conexion.cursor() as cursor:
                    cursor.execute(
                        "SELECT id FROM ocr_ps WHERE numero_factura = %s ORDER BY id DESC LIMIT 1",
                        (numero_factura,),
                    )
                    factura = cursor.fetchone()
                if not factura:
                    continue
                for lote in lotes:
                    lote_str = str(lote).strip()
                    if lote_str:
                        try:
                            bd_ocr.guardar_lote(factura["id"], numero_factura, lote_str)
                            total_registrados += 1
                        except Exception as lote_error:
                            print(f"[lotes] Error guardando lote '{lote_str}': {lote_error}")
            finally:
                conexion.close()
        return jsonify({"ok": True, "registrados": total_registrados})
    except Exception as error:
        print(f"[lotes] Error en guardar_con_factura: {error}")
        return jsonify({"ok": False, "error": str(error)}), 500


@app.route("/api/historico/estado", methods=["GET"])
def estado_historico():
    if not os.path.exists(RUTA_HISTORICO):
        return jsonify({"existe": False, "filas": 0, "filas_cargadas": 0, "filas_pendientes": 0})
    with _lock_historico:
        libro = load_workbook(RUTA_HISTORICO, read_only=True)
        hoja = libro.active
        assert hoja is not None
        filas = max(0, hoja.max_row - 1)
        libro.close()
    cargadas = min(bd_ocr.obtener_progreso("ocr_ps"), filas)
    return jsonify({"existe": True, "filas": filas, "filas_cargadas": cargadas, "filas_pendientes": max(0, filas - cargadas)})


@app.route("/api/historico/cargar_al_servidor", methods=["POST"])
def cargar_al_servidor():
    """Lee el Excel histórico completo y sube a la base de datos MySQL
    solo los números de factura que todavía no se habían cargado."""
    if not os.path.exists(RUTA_HISTORICO):
        return jsonify({"ok": True, "filas_nuevas": 0, "filas_totales": 0})

    with _lock_historico:
        libro = load_workbook(RUTA_HISTORICO, read_only=True)
        hoja = libro.active
        assert hoja is not None
        filas = list(hoja.iter_rows(min_row=2, values_only=True))
        libro.close()

    ya_cargadas = bd_ocr.obtener_progreso("ocr_ps")
    filas_nuevas = filas[ya_cargadas:]

    cargadas_ahora = 0
    for fila in filas_nuevas:
        # Columnas: Fecha de registro, Motor, Numero de Factura,
        # Codigo de Proveedor, Nombre de Proveedor, RIF de Proveedor, Ref. Documento, Lotes
        fila_lista = list(fila)
        _fecha_registro, motor, numero_factura, codigo_proveedor, nombre_proveedor, rif_proveedor, ref_documento = fila_lista[:7]
        lotes_celda = fila_lista[7] if len(fila_lista) > 7 else None
        try:
            contenido_pdf = None
            if ref_documento:
                ruta_pdf = os.path.join(RUTA_DOCUMENTOS, ref_documento)
                if os.path.exists(ruta_pdf):
                    with open(ruta_pdf, "rb") as f:
                        contenido_pdf = f.read()
            documento_id = bd_ocr.obtener_o_guardar_documento_ps(ref_documento, contenido_pdf) if ref_documento else None
            campos = {
                "numero_factura": numero_factura, "codigo_proveedor": codigo_proveedor,
                "nombre_proveedor": nombre_proveedor, "rif_proveedor": rif_proveedor,
            }
            factura_id = bd_ocr.guardar_factura_ps(motor, campos, documento_id=documento_id)
            if lotes_celda and numero_factura:
                for lote in str(lotes_celda).split(";"):
                    lote = lote.strip()
                    if lote:
                        bd_ocr.guardar_lote(factura_id, numero_factura, lote)
            cargadas_ahora += 1
        except Exception as error:
            print(f"[bd_ocr] No se pudo cargar una fila del histórico: {error}")
            break

    bd_ocr.actualizar_progreso("ocr_ps", ya_cargadas + cargadas_ahora)
    return jsonify({"ok": True, "filas_nuevas": cargadas_ahora, "filas_totales": len(filas)})


@app.route("/api/historico/descargar", methods=["GET"])
def descargar_historico():
    if not os.path.exists(RUTA_HISTORICO):
        return jsonify({"error": "Todavía no hay extracciones guardadas en el histórico"}), 404
    with _lock_historico:
        return send_file(
            RUTA_HISTORICO,
            as_attachment=True,
            download_name="ocr_ps_historico.xlsx",
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )


@app.route("/api/historico/reiniciar", methods=["POST"])
def reiniciar_historico():
    """Vacía el progreso acumulado del histórico. No lo borra a lo loco:
    renombra el archivo actual con fecha/hora como respaldo, y la próxima
    extracción exitosa arranca un histórico nuevo desde cero."""
    if not os.path.exists(RUTA_HISTORICO):
        return jsonify({"ok": True, "filas_eliminadas": 0})

    with _lock_historico:
        filas_eliminadas = 0
        try:
            libro_actual = load_workbook(RUTA_HISTORICO, read_only=True)
            hoja_actual = libro_actual.active
            filas_eliminadas = max(0, hoja_actual.max_row - 1)
            libro_actual.close()
        except Exception:
            pass

        marca = datetime.now().strftime("%Y%m%d_%H%M%S")
        ruta_respaldo = os.path.join(DIRECTORIO_BASE, "data", f"ocr_ps_historico_respaldo_{marca}.xlsx")
        try:
            os.replace(RUTA_HISTORICO, ruta_respaldo)
        except OSError as error:
            return jsonify({"error": f"No se pudo reiniciar el histórico: {error}"}), 500

    return jsonify({"ok": True, "filas_eliminadas": filas_eliminadas, "respaldo": os.path.basename(ruta_respaldo)})


# ---------------------------------------------------------------------------
# Pedidos OCR — notas de entrega/pedido manuscritas, mismo puerto que
# OCR-PS. La inteligencia vive en ocr_pedidos_core.py; acá solo la
# interfaz web propia (página, streaming del escaneo, exportar a Excel,
# histórico acumulado en /data/pedidos_ocr_historico.xlsx).
# ---------------------------------------------------------------------------

@app.route("/pedidos")
def pedidos_index():
    reenvio = _reenviar_token_a_sso()
    if reenvio:
        return reenvio
    return render_template("ocr_pedidos.html")


@app.route("/api/pedidos/estado-ocr", methods=["GET"])
def pedidos_estado_ocr():
    # 15/09: mismo fix que /api/estado-ocr — "respaldo" decía "DeepSeek"
    # hardcodeado aunque la cascada real ya no lo usa (reemplazado por
    # NVIDIA NIM).
    cadena_respaldo = []
    if ocr_pedidos.CLAVES_NVIDIA:
        cadena_respaldo.append(f"{ocr_pedidos.MODELO_VISION_NVIDIA} (NVIDIA NIM)")
    if ocr_pedidos.clave_google:
        cadena_respaldo.append(f"{ocr_pedidos.MODELO_VISION} (Gemini)")
    if ocr_pedidos.clave_openrouter:
        cadena_respaldo.append(f"{ocr_pedidos.MODELO_VISION_OPENROUTER} (OpenRouter)")
    return jsonify({
        "conectado": True,
        "detalle": {
            "modelo_vision": ocr_pedidos.MODELO_VISION_QWEN,
            "proveedor": "Qwen-VL (local)",
            "respaldo": " → ".join(cadena_respaldo),
        },
    })


@app.route("/api/pedidos/escanear", methods=["POST"])
def pedidos_escanear():
    archivo = request.files.get("archivo")
    if archivo is None or not archivo.filename:
        return jsonify({"error": "Sube un archivo (campo 'archivo')"}), 400

    extension = archivo.filename.rsplit(".", 1)[-1].lower()
    if extension not in ocr_pedidos.EXTENSIONES_PERMITIDAS:
        return jsonify({"error": f"Extensión no soportada: .{extension}"}), 400

    contenido = archivo.read()
    nombre_archivo = archivo.filename

    # Guardar documento original en documentos_ps (para visualizador lateral)
    ref_guardado = guardar_documento_original(contenido, nombre_archivo)

    es_audio = extension in ocr_pedidos.EXTENSIONES_AUDIO
    if extension == "pdf":
        try:
            imagenes = ocr_pedidos.pdf_a_imagenes(contenido)
        except Exception as error:
            return jsonify({"error": f"No se pudo leer el PDF: {error}"}), 400
        if not imagenes:
            return jsonify({"error": "El PDF no tiene páginas"}), 400
    else:
        # Una nota de voz no se parte en "páginas" — es un único archivo,
        # se procesa entero (11/09, a pedido explícito del usuario).
        imagenes = [contenido]

    def generar():
        try:
            total = len(imagenes)
            yield json.dumps({"evento": "inicio", "total_paginas": total, "ref_documento": ref_guardado}) + "\n"
            for indice, imagen_bytes in enumerate(imagenes, start=1):
                if indice > 1:
                    time.sleep(4)
                nombre = f"{nombre_archivo}-pagina{indice}.png" if extension == "pdf" else nombre_archivo
                resultado = (
                    procesar_audio_pedido_y_guardar(imagen_bytes, nombre, indice) if es_audio
                    else procesar_imagen_pedido_y_guardar(imagen_bytes, nombre, indice)
                )
                yield json.dumps({"evento": "pagina", "resultado": resultado}) + "\n"
            yield json.dumps({"evento": "fin"}) + "\n"
        except (BrokenPipeError, ConnectionAbortedError, ConnectionResetError, GeneratorExit):
            return

    return Response(generar(), mimetype="application/x-ndjson")


@app.route("/api/pedidos/exportar_excel", methods=["POST"])
def pedidos_exportar_excel():
    """Arma la tabla limpia final: UNA FILA POR ARTÍCULO (no por página),
    con la sede repetida en cada fila de su nota — es el pedido a la
    factura, "tabla limpia en excel" que pidió el usuario. "Codigo" viene
    vacío salvo que el usuario haya usado la lupa para confirmarlo contra
    el maestro de ARA_PROYECT (agregado 03/09)."""
    payload = request.get_json(silent=True) or {}
    resultados = payload.get("resultados", [])

    libro = Workbook()
    hoja = libro.active
    assert hoja is not None
    hoja.title = "Pedidos OCR"

    encabezados = ["Sede", "Articulo", "Codigo", "Cantidad", "Pediatrico"]
    hoja.append(encabezados)

    for fila in resultados:
        campos = fila.get("campos") or {}
        sede = campos.get("sede")
        for item in campos.get("items") or []:
            hoja.append([
                sede,
                item.get("articulo"),
                item.get("codigo"),
                item.get("cantidad"),
                "Si" if item.get("pediatrico") else "No",
            ])

    anchos = [14, 46, 14, 14, 12]
    for indice, ancho in enumerate(anchos, start=1):
        hoja.column_dimensions[get_column_letter(indice)].width = ancho

    buffer = BytesIO()
    libro.save(buffer)
    buffer.seek(0)

    return send_file(
        buffer,
        as_attachment=True,
        download_name="pedidos_ocr.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


@app.route("/api/pedidos/productos/buscar", methods=["GET"])
def pedidos_buscar_producto():
    """Busca artículos reales en el maestro de stock de ARA_PROYECT (solo
    lectura) — la lupa de cada renglón de artículo en Pedidos OCR. El
    texto es obligatorio (nunca trae "todo el maestro" sin querer).

    09/09 (a pedido explícito del usuario, movido desde una columna en la
    tabla): cada resultado se enriquece con su precio real para mostrarlo
    junto al stock en la lista de la lupa, ANTES de confirmar el artículo
    — así se ve de una vez sin tener que elegirlo primero. `cliente` es
    opcional: sin cliente elegido todavía para este pedido, sale el
    precio de lista (PRECIO 1) sin descuento, informativo nomás — el
    precio real que se manda al subir el pedido siempre se vuelve a
    resolver server-side con el cliente definitivo (ver /api/pedidos/subir).

    14/09 (a pedido explícito del usuario — "fusión"): la BÚSQUEDA
    siempre se hace contra el maestro de ARA_PROYECT
    (bin/consultar_producto_ara.py) — se probó buscar directo contra el
    catálogo de Cristmedicals (bin/consultar_articulos_cliente.py) y su
    relevancia de texto es mala (no encuentra nada o trae resultados muy
    ambiguos, esa API no está pensada para buscar por texto, solo lista
    el catálogo completo del cliente). El precio de cada resultado ya
    encontrado por ARA sale de bin/consultar_precio_profit.obtener_precio,
    que arma las 3 capas reales (tipo_cli + 10%, categoría de la página de
    Cristmedicals, desc_glob) — ver el docstring de ese módulo."""
    texto = (request.args.get("q") or "").strip()
    if not texto:
        return jsonify({"ok": True, "total": 0, "productos": []})
    # "original": el renglón completo que extrajo el OCR (item.articulo) —
    # para calcular el % de coincidencia contra ESE texto, no contra `q`
    # (que puede ser una sola palabra editada a mano en el buscador).
    texto_original = (request.args.get("original") or "").strip()
    cliente = (request.args.get("cliente") or "").strip()

    resultado = buscar_producto(texto=texto, limite=15, texto_original=texto_original)
    if not resultado.get("ok"):
        return jsonify(resultado), 502

    for producto in resultado.get("productos", []):
        precio_resuelto = obtener_precio(co_art=producto.get("codigo") or "", co_cli=cliente)
        producto["precio"] = precio_resuelto.get("precio_con_descuento") if precio_resuelto.get("ok") else None

    return jsonify(resultado)


@app.route("/api/pedidos/productos/imagen/<codigo>", methods=["GET"])
def pedidos_imagen_producto(codigo):
    """Foto del producto para el preview al pasar el cursor en el
    buscador (11/09, a pedido explícito del usuario). Local, no depende
    de ningún CDN externo: bin/servir_imagen_producto.py busca el
    archivo por código en la carpeta de 10.490 fotos ya existente."""
    ruta = buscar_imagen_producto(codigo)
    if not ruta:
        return jsonify({"error": "Sin foto para este código"}), 404
    return send_file(ruta)


@app.route("/api/pedidos/clientes/buscar", methods=["GET"])
def pedidos_buscar_cliente():
    """Busca clientes reales en PROFIT por razón social O por código
    exacto — para elegir a quién se le monta el pedido antes de subirlo.
    El texto es obligatorio."""
    texto = (request.args.get("q") or "").strip()
    if not texto:
        return jsonify({"ok": True, "total": 0, "clientes": []})
    resultado = buscar_cliente(texto=texto, limite=15)
    if not resultado.get("ok"):
        return jsonify(resultado), 502
    return jsonify(resultado)


@app.route("/api/pedidos/clientes/<co_cli>/descuento", methods=["GET"])
def pedidos_descuento_cliente(co_cli):
    """Descuento lineal ("descuento global") del cliente — solo para
    MOSTRARLO apenas se elige el cliente (11/09, a pedido explícito del
    usuario). Nunca modifica nada — es solo referencia general del
    cliente (clientes.desc_glob); el precio real de cada artículo, con
    las 3 capas de descuento completas, sale de
    bin/consultar_precio_profit.obtener_precio (ver /api/pedidos/productos/buscar).

    14/09 (a pedido explícito del usuario: "no tomemos esa api de
    búsqueda de cliente... volvamos a la nuestra de ARA/Profit"): vuelve
    a resolverse directo contra Profit, sin pasar por la API de clientes
    de Cristmedicals."""
    resultado = obtener_descuento_cliente(co_cli)
    return jsonify(resultado), (200 if resultado.get("ok") else 502)


# Endpoint externo real de creación de pedidos — confirmado 04/09 leyendo
# CartService.php completo (el archivo real de la app de cotización): es el
# mismo que llama processCart() de forma asíncrona después de guardar el
# pedido en su propia tabla local. El de 192.168.4.136:8001 que se había
# usado el 03/09 era información de otra fuente, nunca confirmada contra
# código real — se reemplaza acá.
PEDIDOS_ENDPOINT_URL = os.environ.get("PEDIDOS_ENDPOINT_URL", "https://apiweb.cristmedicals.com/api/pedidos/pedido-profit")


def _cantidad_a_entero(texto):
    """La cantidad viene tal cual la transcribió el OCR ("2 cajas", "x3",
    "1 caja y media") — se toma el primer número entero que aparezca. Si no
    hay ninguno, 0: no se inventa una cantidad, el usuario tiene que
    corregirla a mano en la tarjeta antes de subir."""
    if texto is None:
        return 0
    coincidencia = re.search(r"\d+", str(texto))
    return int(coincidencia.group()) if coincidencia else 0


@app.route("/api/pedidos/espera/guardar", methods=["POST"])
def pedidos_espera_guardar():
    """"Dejar en espera" (10/09): guarda el estado completo de una
    tarjeta (sede+items+cliente elegido) para retomarla después, sin
    perder lo hecho — caso típico: el cliente dice que va a agregar algo
    más pero que esperen. Exige sesión SSO — sin saber quién es el
    empleado no hay forma de aplicar el aislamiento por usuario."""
    usuario = usuario_activo()
    if not usuario:
        return jsonify({"ok": False, "error": "No hay usuario activo — entrá desde el sidebar del ERP antes de dejar algo en espera"}), 401

    payload = request.get_json(silent=True) or {}
    campos = payload.get("campos")
    if not isinstance(campos, dict):
        return jsonify({"ok": False, "error": "Falta 'campos'"}), 400

    try:
        id_nuevo = bd_ocr.guardar_pedido_en_espera(
            empleado_id=usuario["employee_id"],
            empleado_nombre=usuario["nombre"],
            campos=campos,
            cliente=payload.get("cliente"),
            etiqueta=(payload.get("etiqueta") or "").strip() or None,
            archivo_origen=(payload.get("archivo_origen") or "").strip() or None,
            motor=payload.get("motor"),
        )
    except Exception as error:
        return jsonify({"ok": False, "error": f"No se pudo dejar en espera: {error}"}), 500
    return jsonify({"ok": True, "id": id_nuevo})


@app.route("/api/pedidos/espera/listar", methods=["GET"])
def pedidos_espera_listar():
    """Lista SOLO las cotizaciones en espera del usuario activo — salvo
    que sea el usuario global (USUARIO_GLOBAL_ESPERA), que ve las de
    todos. Un usuario normal nunca recibe filas ajenas acá."""
    usuario = usuario_activo()
    if not usuario:
        return jsonify({"ok": False, "error": "No hay usuario activo"}), 401
    if _es_usuario_global(usuario):
        pedidos = bd_ocr.listar_pedidos_en_espera()
    else:
        pedidos = bd_ocr.listar_pedidos_en_espera(empleado_id=usuario["employee_id"])
    return jsonify({"ok": True, "pedidos": pedidos, "es_global": _es_usuario_global(usuario)})


@app.route("/api/pedidos/espera/<int:espera_id>", methods=["GET"])
def pedidos_espera_obtener(espera_id):
    """Trae UNA cotización en espera completa para retomarla en pantalla.
    403 si no es del usuario activo y tampoco es el usuario global —
    nunca se filtra el contenido de la cotización de otra persona."""
    usuario = usuario_activo()
    if not usuario:
        return jsonify({"ok": False, "error": "No hay usuario activo"}), 401
    fila = bd_ocr.obtener_pedido_en_espera(espera_id)
    if fila is None:
        return jsonify({"ok": False, "error": "No existe esa cotización en espera"}), 404
    if fila["empleado_id"] != usuario["employee_id"] and not _es_usuario_global(usuario):
        return jsonify({"ok": False, "error": "No tenés acceso a una cotización en espera de otro usuario"}), 403
    return jsonify({"ok": True, "pedido": fila})


@app.route("/api/pedidos/espera/<int:espera_id>", methods=["DELETE"])
def pedidos_espera_eliminar(espera_id):
    """Se llama al retomar una cotización en espera (sale de la lista) o
    al descartarla a mano. Mismo chequeo de dueño/usuario global que el
    GET de arriba."""
    usuario = usuario_activo()
    if not usuario:
        return jsonify({"ok": False, "error": "No hay usuario activo"}), 401
    fila = bd_ocr.obtener_pedido_en_espera(espera_id)
    if fila is None:
        return jsonify({"ok": True})
    if fila["empleado_id"] != usuario["employee_id"] and not _es_usuario_global(usuario):
        return jsonify({"ok": False, "error": "No tenés acceso a una cotización en espera de otro usuario"}), 403
    bd_ocr.eliminar_pedido_en_espera(espera_id)
    return jsonify({"ok": True})


@app.route("/api/pedidos/alertas_stock", methods=["POST"])
def pedidos_guardar_alerta_stock():
    """Botón "Generar alerta" en la lupa de artículos (10/09, a pedido
    explícito del usuario) — antes era decorativo, ahora guarda en la BD
    con el código de cliente de la página (por eso ese botón de elegir
    cliente ahora va primero, antes de "Montar pedido": sin cliente
    elegido no hay a quién asociarle la alerta)."""
    payload = request.get_json(silent=True) or {}
    cod_cliente = (payload.get("cod_cliente") or "").strip()
    codigo_articulo = (payload.get("codigo_articulo") or "").strip()
    if not cod_cliente:
        return jsonify({"ok": False, "error": "Falta el código de cliente — elegí el cliente de la página antes de generar la alerta"}), 400
    if not codigo_articulo:
        return jsonify({"ok": False, "error": "Falta el código del artículo"}), 400

    usuario = usuario_activo()
    try:
        alerta_id = bd_ocr.guardar_alerta_stock(
            cod_cliente=cod_cliente,
            codigo_articulo=codigo_articulo,
            cliente_nombre=(payload.get("cliente_nombre") or "").strip() or None,
            descripcion_articulo=(payload.get("descripcion_articulo") or "").strip() or None,
            solicitado_por=usuario["nombre"] if usuario else None,
        )
    except Exception as error:
        return jsonify({"ok": False, "error": f"No se pudo guardar la alerta: {error}"}), 500
    return jsonify({"ok": True, "id": alerta_id})


def _resolver_notas(numero_cotizacion, pedido_id):
    """Consulta la(s) nota(s) de entrega de cada número de cotización
    (puede haber más de uno si el pedido se dividió) y junta todas las
    que aparezcan. Lista vacía si Profit todavía no generó ninguna."""
    notas_encontradas = []
    for numero in numero_cotizacion.split(","):
        resultado_nota = obtener_notas_de_cotizacion(numero.strip())
        if resultado_nota.get("ok"):
            notas_encontradas.extend(resultado_nota["notas_entrega"])
        else:
            print(f"[pedidos-ocr] Sin nota de entrega para cotización {numero.strip()}: {resultado_nota.get('error')}")
    return notas_encontradas


def _resolver_facturas(notas_encontradas, pedido_id):
    """Resuelve el N° de FACTURA real para cada nota (11/09, cambio de
    plan del usuario: ahora se muestra y se puede descargar por número
    de factura, no de nota — la nota queda como texto informativo nomás
    hasta que haya una API aparte para su propio PDF/proforma). Reusa el
    mismo endpoint que ya baja el PDF (api_publico_routes.py expone el
    número real en el header Content-Disposition de esa respuesta, no
    hace falta una llamada aparte) — se descarta el PDF acá, solo
    interesa el número para guardarlo en pedidos_ocr.numero_factura."""
    facturas = []
    for numero_nota in notas_encontradas:
        resultado = obtener_factura_pdf_de_nota(numero_nota)
        if resultado.get("ok") and resultado.get("numero_factura"):
            facturas.append(resultado["numero_factura"])
        else:
            print(f"[pedidos-ocr] Sin factura real para la nota {numero_nota} del pedido {pedido_id}: {resultado.get('error')}")
    return facturas


def _escribir_comentario_en_notas(notas_encontradas, comentario, pedido_id):
    """Comentario del modal "Montar pedido" -> reng_nde real (11/09, API
    de ARA_PROYECT confirmada: POST /api/publico/nota-entrega/comentario).
    Se agrega al renglón 1 de CADA nota encontrada (si el pedido se
    dividió en varios almacenes, todas quedan con el mismo comentario).
    Nunca bloquea nada más arriba: la cotización y la nota ya existen en
    Profit antes de llegar acá."""
    if not comentario:
        return
    for numero_nota in notas_encontradas:
        resultado_comentario = escribir_comentario_nota(numero_nota, comentario)
        if not resultado_comentario.get("ok"):
            print(f"[pedidos-ocr] No se pudo escribir el comentario en la nota {numero_nota} del pedido {pedido_id}: {resultado_comentario.get('error')}")


def _resolver_facturas_en_background(pedido_id, notas_encontradas):
    """La factura real (nota -> reng_fac -> factura) suele tardar en
    existir en Profit incluso más que la nota misma — por eso esto SIEMPRE
    corre en background, nunca en la respuesta de "Subir pedido" (cada
    intento baja el PDF completo, ~300-500KB, para leer el número del
    header). 11/09, cambio de plan del usuario."""
    esperas_s = (60, 180, 300)  # 1 min, 3 min, 5 min
    for espera in esperas_s:
        time.sleep(espera)
        try:
            facturas = _resolver_facturas(notas_encontradas, pedido_id)
        except Exception as error:
            print(f"[pedidos-ocr] Resolución de factura real del pedido {pedido_id} falló: {error}")
            continue
        if facturas:
            bd_ocr.actualizar_numero_factura(pedido_id, ",".join(facturas))
            print(f"[pedidos-ocr] Factura real encontrada para el pedido {pedido_id}: {','.join(facturas)}")
            return
    print(f"[pedidos-ocr] Pedido {pedido_id}: sin factura real después de todos los reintentos")


def _reintentar_resolver_nota_en_background(pedido_id, numero_cotizacion, comentario):
    """Corre en un hilo aparte, minutos después de haber respondido
    "Subir pedido" (11/09, a pedido explícito del usuario, tras la
    carrera real observada con la cotización 248561): reintenta unas
    pocas veces, espaciado, por si Profit todavía no había generado la
    nota en el momento de la subida. Si a la tercera tampoco aparece, se
    deja como está — bd_ocr.pedidos_ocr.numero_entrega se queda en null
    igual que si esto no existiera, sin romper nada."""
    esperas_s = (180, 300, 600)  # 3 min, 5 min, 10 min
    for espera in esperas_s:
        time.sleep(espera)
        try:
            notas_encontradas = _resolver_notas(numero_cotizacion, pedido_id)
        except Exception as error:
            print(f"[pedidos-ocr] Reintento de fondo del pedido {pedido_id} falló: {error}")
            continue
        if notas_encontradas:
            bd_ocr.actualizar_numero_entrega(pedido_id, ",".join(notas_encontradas))
            _escribir_comentario_en_notas(notas_encontradas, comentario, pedido_id)
            print(f"[pedidos-ocr] Reintento de fondo del pedido {pedido_id}: nota encontrada ({','.join(notas_encontradas)})")
            _resolver_facturas_en_background(pedido_id, notas_encontradas)
            return
    print(f"[pedidos-ocr] Reintento de fondo del pedido {pedido_id}: sin nota de entrega después de todos los reintentos")


@app.route("/api/pedidos/subir", methods=["POST"])
def pedidos_subir():
    """Arma y envía UN pedido real al endpoint externo de creación de
    pedidos. Se llama recién cuando el usuario ya revisó/corrigió la
    tarjeta y eligió el cliente — nunca automático apenas escanea (por
    eso el guardado en pedidos_ocr, más abajo, también queda para acá:
    es el histórico real de Pedidos OCR, ya no un Excel).

    El precio de cada ítem se resuelve en vivo vía
    bin/consultar_precio_profit.obtener_precio (tipo_cli + su 10% cuando
    corresponde, descuento de categoría de la página de Cristmedicals, y
    desc_glob del cliente — ver el docstring de ese módulo) — nunca un
    valor fijo ni 0. Un ítem cuyo precio no se pudo resolver queda afuera
    del pedido, no se manda con precio inventado.

    La respuesta trae "duplicado": true (14/09, caso real reportado:
    cotización #134263) cuando el endpoint externo de Cristmedicals
    reportó este pedido como duplicado de uno viejo y devolvió SU número
    de cotización en vez de crear uno nuevo — el frontend debe mostrar
    una advertencia explícita en ese caso, nunca tratarlo como un éxito
    silencioso idéntico a una cotización recién creada."""
    payload = request.get_json(silent=True) or {}
    cod_cliente = (payload.get("cod_cliente") or "").strip()
    cliente_nombre = (payload.get("cliente_nombre") or "").strip() or None
    items_crudos = payload.get("items") or []
    motor = payload.get("motor")
    pagina_origen = payload.get("pagina_origen")  # para buscar el documento original
    comentario = (payload.get("comentario") or "").strip()

    if not cod_cliente:
        return jsonify({"error": "Hace falta el código del cliente"}), 400

    items_validos = [
        item for item in items_crudos
        if (item.get("articulo") or "").strip() and (item.get("codigo") or "").strip()
    ]
    if not items_validos:
        return jsonify({
            "error": "No hay artículos con código confirmado contra el maestro — buscalos con la lupa antes de subir",
        }), 400

    # Aviso de artículos repetidos (14/09, a pedido explícito del usuario
    # — antes bloqueaba, ahora NO: "que si la deje subir... pero que
    # genere la alerta primero... y guardamos esa alerta para tener
    # respaldo"). Caso real: dos filas del mismo pedido con el MISMO
    # código confirmado — el OCR duplicó un renglón sin que nadie lo
    # notara. Se compara por "codigo" (el confirmado contra el maestro,
    # nunca el texto libre de "articulo") — nunca se auto-combinan las
    # cantidades acá: sumarlas solo sería inventar un número.
    # `items_validos` ya refleja lo que hay en pantalla AL MOMENTO de
    # subir (incluida cualquier cantidad que el usuario haya corregido a
    # mano), nunca la extracción original del OCR. El frontend ya le dio
    # al usuario la opción de montarlo así o de ir a corregir/eliminar la
    # fila ANTES de llegar acá — este chequeo server-side es la fuente de
    # verdad que queda grabada, no depende de que el aviso del navegador
    # se haya visto o no.
    cantidades_por_codigo = {}
    articulo_por_codigo = {}
    for item in items_validos:
        codigo_norm = item["codigo"].strip().upper()
        cantidades_por_codigo.setdefault(codigo_norm, []).append(item.get("cantidad"))
        articulo_por_codigo.setdefault(codigo_norm, item.get("articulo") or codigo_norm)
    repetidos = {codigo: cants for codigo, cants in cantidades_por_codigo.items() if len(cants) > 1}
    aviso_articulos_repetidos = [
        f"{articulo_por_codigo[codigo]} (aparece {len(cants)} veces, cantidades: {', '.join(str(c) for c in cants)})"
        for codigo, cants in repetidos.items()
    ]

    # Bloqueo de psicotrópicos (11/09, a pedido explícito del usuario):
    # Cristmedicals no puede montar estos pedidos por esta vía, ni
    # armados por OCR ni editados/agregados a mano — este chequeo corre
    # ACÁ, sobre los items ya confirmados contra el maestro, así que
    # cubre los dos casos por igual (todo item que llega hasta acá pasó
    # por la lupa, ver elegirProducto en ocr_pedidos.html). Nunca se
    # llega a mandar nada a Profit si hay una coincidencia.
    encontrados_psicotropicos = revisar_items_psicotropicos(items_validos)
    if encontrados_psicotropicos:
        articulos = ", ".join(e["articulo"] for e in encontrados_psicotropicos)
        return jsonify({
            "error": f"Este pedido contiene psicotrópicos ({articulos}). Por favor procese con el departamento correspondiente — no se puede montar por acá.",
        }), 400

    items_payload = []
    items_sin_precio = []
    for item in items_validos:
        co_art = item["codigo"].strip()
        precio_resuelto = obtener_precio(co_art=co_art, co_cli=cod_cliente)
        if not precio_resuelto.get("ok"):
            items_sin_precio.append(f"{item.get('articulo') or co_art} ({precio_resuelto.get('error', 'sin precio')})")
            continue
        cantidad = _cantidad_a_entero(item.get("cantidad"))
        items_payload.append({
            "co_art": co_art,
            "precio": precio_resuelto["precio_con_descuento"],
            # Sin selección de sede (quitada 04/09) — toda la cantidad va a
            # quantitySC, ninguna a quantityBQ. Si más adelante hace falta
            # repartir entre los dos almacenes, hay que reintroducir algún
            # criterio acá.
            "quantitySC": cantidad,
            "quantityBQ": 0,
        })

    if not items_payload:
        return jsonify({
            "error": "No se pudo resolver el precio en Profit de ningún artículo: " + "; ".join(items_sin_precio),
        }), 502

    usuario = usuario_activo()
    nombre_usuario = usuario["nombre"] if usuario else None

    # Buscar documento original por nombre guardado en documentos_ps
    documento_id = None
    archivo_origen = (payload.get("archivo_origen") or "").strip()
    if archivo_origen:
        try:
            conexion = bd_ocr.obtener_conexion()
            try:
                with conexion.cursor() as cursor:
                    # Buscar por nombre_archivo exacto (es el nombre guardado por guardar_documento_original)
                    cursor.execute(
                        "SELECT id FROM documentos_ps WHERE nombre_archivo = %s ORDER BY id DESC LIMIT 1",
                        (archivo_origen,),
                    )
                    doc = cursor.fetchone()
                    if doc:
                        documento_id = doc["id"] if isinstance(doc, dict) else doc[0]
                    else:
                        # Intentar match por sufijo (si el nombre tiene timestamp diferente)
                        sufijo = archivo_origen.split("_")[-1] if "_" in archivo_origen else archivo_origen
                        if sufijo:
                            cursor.execute(
                                "SELECT id FROM documentos_ps WHERE nombre_archivo LIKE %s ORDER BY id DESC LIMIT 1",
                                (f"%{sufijo}",),
                            )
                            doc = cursor.fetchone()
                            if doc:
                                documento_id = doc["id"] if isinstance(doc, dict) else doc[0]
            finally:
                conexion.close()
        except:
            pass

    try:
        pedido_id = bd_ocr.guardar_pedido_ocr(
            cod_cliente=cod_cliente, sede=None, motor=motor, items=items_validos, estado="pendiente",
            cliente_nombre=cliente_nombre, subido_por=nombre_usuario, documento_id=documento_id,
        )
    except Exception as error:
        return jsonify({"error": f"No se pudo iniciar el registro del pedido: {error}"}), 500

    payload_envio = {"cod_pedido": pedido_id, "cod_cliente": cod_cliente, "items": items_payload}
    # NOTA (07/09): "subido por <usuario>" todavía NO se manda dentro de
    # payload_envio — el contrato real de apiweb.cristmedicals.com/api/pedidos/
    # pedido-profit (confirmado leyendo CartService.php completo) solo acepta
    # cod_pedido/cod_cliente/items, sin campo de comentario/observación. Se
    # guarda igual en bd_ocr.pedidos_ocr.subido_por (arriba) para que quede en
    # el historial de ESTE sistema. El comentario del modal (`comentario`,
    # arriba) NO se manda acá tampoco — se inyecta más abajo, después de
    # resolver la(s) nota(s) de entrega reales, directo en
    # reng_nde.comentario (11/09, ver bin/escribir_comentario_nota_ara.py).

    numero_cotizacion = None
    duplicado = False
    try:
        respuesta = requests.post(PEDIDOS_ENDPOINT_URL, json=payload_envio, timeout=20)
        exito = respuesta.ok
        cuerpo_respuesta = respuesta.text[:2000]
        cuerpo_json = respuesta.json() if respuesta.content else {}
        resultado_cuerpo = cuerpo_json.get("resultado") or {}
        # El endpoint devuelve el número de cotización real en
        # resultado.fact_nums — es una LISTA, no un solo número: cuando el
        # pedido supera el máximo de artículos por cotización que maneja
        # Profit (confirmado en vivo, 10/09: un pedido de 21+ artículos ya
        # vuelve partido en dos), esa lista trae 2 o más números. Antes
        # acá se tomaba solo el primero y el resto se perdía sin avisar —
        # ahora se guardan todos, separados por coma (mismo criterio que
        # "nro_factura" en retenciones para varios números en un campo).
        cotizaciones = resultado_cuerpo.get("fact_nums") or []
        if cotizaciones:
            numero_cotizacion = ",".join(str(c) for c in cotizaciones)
        # Mismo criterio que CartService::_sendOrderToExternalEndpoint: un
        # 409 con resultado.duplicado + resultado.pedido_montado en true
        # significa que el pedido YA se procesó del otro lado — es éxito,
        # no falla, aunque el status HTTP diga lo contrario.
        #
        # 14/09 (caso real reportado por el usuario, cotización #134263):
        # esta detección de "duplicado" la hace el endpoint EXTERNO de
        # Cristmedicals (apiweb.cristmedicals.com/api/pedidos/pedido-profit,
        # fuera de este repo) — parece comparar solo cod_cliente + artículos
        # + cantidades, SIN ventana de tiempo, así que un pedido genuinamente
        # NUEVO que por coincidencia repite exactamente el mismo cliente,
        # artículos y cantidades que uno viejo ya facturado puede volver
        # marcado como "duplicado" y devolver el número de cotización VIEJO
        # en vez de crear uno nuevo — antes esto se aceptaba como éxito
        # silencioso, indistinguible de una cotización recién creada, y una
        # ejecutiva terminó confirmándole al cliente el número equivocado.
        # No se puede arreglar la detección en sí (es lógica del lado de
        # Cristmedicals) — lo que sí se puede es dejar de ocultarlo: se
        # guarda la bandera acá y se manda al frontend para que muestre una
        # advertencia explícita en vez de un éxito idéntico al normal.
        if resultado_cuerpo.get("duplicado") and resultado_cuerpo.get("pedido_montado"):
            duplicado = True
            if not exito:
                exito = True
    except requests.RequestException as error:
        exito = False
        cuerpo_respuesta = str(error)
    except ValueError:
        pass  # cuerpo sin JSON válido: se queda como estaba (numero_cotizacion=None), sin forzar nada

    try:
        # Marcas al frente de la respuesta guardada (14/09) — para poder
        # encontrar estos casos después con una simple búsqueda en
        # pedidos_ocr.respuesta_endpoint, sin agregar columna. Esto es el
        # "respaldo" del aviso de artículos repetidos que pidió el
        # usuario: aunque no bloquea la subida, queda grabado quién y qué
        # pedido tuvo un artículo repetido, para auditar después.
        marcas = []
        if duplicado:
            marcas.append("[DUPLICADO DETECTADO POR CRISTMEDICALS — cotización reutilizada, no creada]")
        if aviso_articulos_repetidos:
            marcas.append(f"[ARTÍCULO(S) REPETIDO(S), SUBIDO IGUAL: {'; '.join(aviso_articulos_repetidos)}]")
        respuesta_para_guardar = (f"{' '.join(marcas)} {cuerpo_respuesta}") if marcas else cuerpo_respuesta
        bd_ocr.actualizar_estado_pedido_ocr(
            pedido_id, "enviado" if exito else "error", respuesta_para_guardar, numero_cotizacion,
        )
    except Exception as error:
        print(f"[pedidos-ocr] No se pudo actualizar el estado del pedido {pedido_id}: {error}")

    # N° de Nota de Entrega real (11/09, API de ARA_PROYECT confirmada:
    # ara/ARA_Brain/api_publico_routes.py, /api/publico/cotizaciones/nota).
    # Una cotización puede dar VARIAS notas (una por almacén) — se
    # consulta cada número de cotización (puede haber más de uno si el
    # pedido se dividió) y se juntan todas las notas encontradas. Nunca
    # bloquea la respuesta de "Subir pedido": si esta consulta falla, el
    # pedido ya quedó enviado igual, el N° Entrega se queda en null.
    if exito and numero_cotizacion:
        try:
            notas_encontradas = _resolver_notas(numero_cotizacion, pedido_id)
            if notas_encontradas:
                bd_ocr.actualizar_numero_entrega(pedido_id, ",".join(notas_encontradas))
                _escribir_comentario_en_notas(notas_encontradas, comentario, pedido_id)
                threading.Thread(
                    target=_resolver_facturas_en_background,
                    args=(pedido_id, notas_encontradas),
                    daemon=True,
                ).start()
            else:
                # Carrera real observada en vivo (11/09, cotización 248561):
                # Profit no siempre generó la nota todavía en el instante en
                # que se crea la cotización — acá no hay nada más para
                # intentar sin bloquear la respuesta, así que se reintenta
                # solo, en background, varios minutos después.
                threading.Thread(
                    target=_reintentar_resolver_nota_en_background,
                    args=(pedido_id, numero_cotizacion, comentario),
                    daemon=True,
                ).start()
        except Exception as error:
            print(f"[pedidos-ocr] No se pudo resolver la nota de entrega del pedido {pedido_id}: {error}")

    if not exito:
        return jsonify({"ok": False, "cod_pedido": pedido_id, "error": cuerpo_respuesta}), 502
    return jsonify({
        "ok": True, "cod_pedido": pedido_id, "numero_cotizacion": numero_cotizacion,
        "respuesta": cuerpo_respuesta, "omitidos": items_sin_precio, "duplicado": duplicado,
        "articulos_repetidos": aviso_articulos_repetidos,
    })


@app.route("/api/pedidos/historico/estado", methods=["GET"])
def pedidos_estado_historico():
    """Ya no cuenta filas del Excel (ver procesar_imagen_pedido_y_guardar)
    — cuenta pedidos subidos con éxito en bd_ocr.pedidos_ocr."""
    total = bd_ocr.contar_pedidos_ocr()
    return jsonify({"existe": total > 0, "filas": total})


@app.route("/api/pedidos/historico/lista", methods=["GET"])
def pedidos_lista_historico():
    """Historial de pedidos subidos (éxito o error) — para el "mini
    dashboard" de historial de cotizaciones en la interfaz. Incluye
    numero_entrega y tiene_documento.

    Aislado por usuario (11/09, a pedido explícito del usuario: "no debo
    saber las cotizaciones de otro usuario, excepto la mía") — mismo
    criterio que "en espera": el usuario global ve todo, cualquier otro
    ve solo lo que subió él mismo. A diferencia de pedidos_en_espera,
    pedidos_ocr no guarda un ID estable de quién subió cada fila (solo
    el nombre, en subido_por) — así que se compara por nombre
    normalizado (sin mayúsculas/acentos, igual que _es_usuario_global).

    14/09 (a pedido explícito del usuario: "ver el historial completo...
    y filtrarlas por día y ejecutivo activo de ese día"): `desde`/`hasta`
    (query params, "YYYY-MM-DD") filtran directo en SQL (ver
    bd_ocr.listar_pedidos_ocr) en vez de traer un bloque fijo y filtrar
    en JavaScript — así un día viejo no queda invisible solo por estar
    fuera del bloque de siempre. `ejecutivo` (query param, opcional) deja
    elegir a un usuario puntual DENTRO de lo que ya puede ver — el
    usuario global puede pedir cualquier nombre, uno normal solo puede
    "elegir" el suyo propio (si pide otro, igual se lo pisa por su
    propio nombre — nunca se expone la cotización de otro)."""
    usuario = usuario_activo()
    if not usuario:
        return jsonify({"ok": False, "error": "No hay usuario activo"}), 401
    es_global = _es_usuario_global(usuario)
    desde = (request.args.get("desde") or "").strip() or None
    hasta = (request.args.get("hasta") or "").strip() or None
    ejecutivo_pedido = (request.args.get("ejecutivo") or "").strip()
    try:
        # limite alto siempre (no solo con desde/hasta puestos): un usuario
        # no-global se filtra por nombre DESPUÉS de traer las filas (más
        # abajo) — con un límite chico, alguien con pocas cotizaciones
        # propias mezcladas entre muchas de otros ejecutivos se quedaría
        # sin ver su historial viejo aunque exista en la base.
        filas = bd_ocr.listar_pedidos_ocr(limite=2000, desde=desde, hasta=hasta)
        if not es_global:
            nombre_normalizado = _normalizar_nombre(usuario.get("nombre") or "")
            filas = [f for f in filas if _normalizar_nombre(f.get("subido_por") or "") == nombre_normalizado]
        elif ejecutivo_pedido:
            ejecutivo_normalizado = _normalizar_nombre(ejecutivo_pedido)
            filas = [f for f in filas if _normalizar_nombre(f.get("subido_por") or "") == ejecutivo_normalizado]
    except Exception as error:
        return jsonify({"ok": False, "error": str(error)}), 502
    return jsonify({"ok": True, "pedidos": filas, "es_global": es_global})


@app.route("/api/pedidos/historico/ejecutivos", methods=["GET"])
def pedidos_historico_ejecutivos():
    """Nombres de ejecutivos con al menos un pedido subido en el rango de
    fechas dado (14/09, a pedido explícito del usuario) — para llenar el
    selector del "mini dashboard" con solo quienes estuvieron activos
    ESE día, no una lista fija de todos los usuarios que existieron. Un
    usuario no-global recibe solo su propio nombre (si tiene algo en el
    rango) — nunca la lista completa, mismo criterio de aislamiento que
    /api/pedidos/historico/lista."""
    usuario = usuario_activo()
    if not usuario:
        return jsonify({"ok": False, "error": "No hay usuario activo"}), 401
    desde = (request.args.get("desde") or "").strip() or None
    hasta = (request.args.get("hasta") or "").strip() or None
    try:
        if _es_usuario_global(usuario):
            ejecutivos = bd_ocr.listar_ejecutivos_pedidos_ocr(desde=desde, hasta=hasta)
        else:
            propios = bd_ocr.listar_ejecutivos_pedidos_ocr(desde=desde, hasta=hasta)
            nombre_normalizado = _normalizar_nombre(usuario.get("nombre") or "")
            ejecutivos = [e for e in propios if _normalizar_nombre(e) == nombre_normalizado]
    except Exception as error:
        return jsonify({"ok": False, "error": str(error)}), 502
    return jsonify({"ok": True, "ejecutivos": ejecutivos})


@app.route("/api/pedidos/<int:pedido_id>/documento", methods=["GET"])
def pedidos_descargar_documento(pedido_id):
    """Sirve el PDF/imagen original de un pedido dado por su ID.

    CORREGIDO 10/09: mimetype ya no viene hardcodeado a "application/pdf"
    — una nota de Pedidos OCR casi siempre es una FOTO (JPEG/PNG), no un
    PDF, y forzar el mimetype equivocado rompe el preview inline (el
    navegador espera un PDF y recibe bytes de imagen). Se detecta según
    la extensión real del archivo guardado."""
    try:
        conexion = bd_ocr.obtener_conexion()
        try:
            with conexion.cursor() as cursor:
                cursor.execute(
                    """
                    SELECT d.nombre_archivo, d.contenido
                    FROM pedidos_ocr p
                    JOIN documentos_ps d ON d.id = p.documento_id
                    WHERE p.id = %s AND p.documento_id IS NOT NULL
                    """,
                    (pedido_id,),
                )
                fila = cursor.fetchone()
        finally:
            conexion.close()

        if not fila:
            return jsonify({"error": "No se encontró el documento de este pedido"}), 404
        nombre_archivo = fila["nombre_archivo"] or "pedido.pdf"
        tipo_detectado, _ = mimetypes.guess_type(nombre_archivo)
        return send_file(
            BytesIO(fila["contenido"]),
            mimetype=tipo_detectado or "application/pdf",
            as_attachment=False,
            download_name=nombre_archivo,
        )
    except Exception as error:
        return jsonify({"error": str(error)}), 500


@app.route("/api/pedidos/factura-real/<numero_nota>", methods=["GET"])
def pedidos_factura_pdf_real(numero_nota):
    """PDF real de la FACTURA de Profit asociada a una nota de entrega
    (11/09, a pedido explícito del usuario) — distinto del documento
    original que subió el usuario (ver pedidos_descargar_documento):
    ese es la nota/lista escaneada, este es la factura real que generó
    Profit. La API de ARA_PROYECT resuelve toda la cadena
    nota -> reng_fac -> factura -> servidor de PDFs; acá solo se
    proxea el resultado (la clave de esa API nunca debe llegar al
    navegador)."""
    resultado = obtener_factura_pdf_de_nota(numero_nota)
    if not resultado.get("ok"):
        return jsonify({"error": resultado.get("error")}), 502
    return Response(
        resultado["pdf_bytes"],
        mimetype="application/pdf",
        headers={"Content-Disposition": f'inline; filename="{resultado["nombre_archivo"]}"'},
    )


@app.route("/api/pedidos/<int:pedido_id>/proforma-nota/<numero_nota>", methods=["GET"])
def pedidos_proforma_nota(pedido_id, numero_nota):
    """Proforma PROPIA de la nota de entrega (11/09, a pedido explícito
    del usuario) — Profit sí genera un PDF de esto, pero no sabemos
    cómo lo arma ni dónde lo guarda (no hay servidor de PDFs para notas
    como el que sí existe para facturas). Se reconstruye acá: los
    renglones REALES (qué código quedó, si se anuló, el comentario)
    salen de reng_nde; la cantidad y descripción completa de cada
    artículo se completan cruzando contra lo que este mismo sistema ya
    guardó al montar el pedido (pedidos_ocr.items_json), porque reng_nde
    no expone esos dos campos."""
    pedido = bd_ocr.obtener_pedido_ocr(pedido_id)
    if not pedido:
        return jsonify({"error": "No existe ese pedido"}), 404

    resultado_reng = leer_renglones_nota(numero_nota)
    if not resultado_reng.get("ok"):
        return jsonify({"error": resultado_reng.get("error")}), 502

    items_por_codigo = {
        (item.get("codigo") or "").strip().upper(): item
        for item in pedido.get("items") or []
        if item.get("codigo")
    }
    renglones = []
    for reng in resultado_reng["renglones"]:
        co_art = (reng.get("co_art") or "").strip().upper()
        item_original = items_por_codigo.get(co_art)
        renglones.append({
            "codigo": reng.get("co_art"),
            "descripcion": item_original["articulo"] if item_original else reng.get("co_art"),
            "cantidad": item_original.get("cantidad") if item_original else None,
            "comentario": reng.get("comentario"),
            "anulado": reng.get("anulado"),
        })

    pdf_bytes = generar_pdf_proforma_nota({
        "numero_nota": numero_nota,
        "fecha": pedido.get("fecha_registro"),
        "cod_cliente": pedido.get("cod_cliente"),
        "cliente_nombre": pedido.get("cliente_nombre"),
        "numero_cotizacion": pedido.get("numero_cotizacion"),
        "renglones": renglones,
    })
    return Response(
        pdf_bytes,
        mimetype="application/pdf",
        headers={"Content-Disposition": f'inline; filename="proforma_nota_{numero_nota}.pdf"'},
    )


@app.route("/api/pedidos/historico/descargar", methods=["GET"])
def pedidos_descargar_historico():
    if not os.path.exists(RUTA_HISTORICO_PEDIDOS):
        return jsonify({"error": "Todavía no hay extracciones guardadas en el histórico"}), 404
    with _lock_historico_pedidos:
        return send_file(
            RUTA_HISTORICO_PEDIDOS,
            as_attachment=True,
            download_name="pedidos_ocr_historico.xlsx",
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )


@app.route("/api/pedidos/historico/reiniciar", methods=["POST"])
def pedidos_reiniciar_historico():
    """Mismo criterio que /api/historico/reiniciar de OCR-PS: no borra a
    lo loco, renombra el archivo actual con fecha/hora como respaldo."""
    if not os.path.exists(RUTA_HISTORICO_PEDIDOS):
        return jsonify({"ok": True, "filas_eliminadas": 0})

    with _lock_historico_pedidos:
        filas_eliminadas = 0
        try:
            libro_actual = load_workbook(RUTA_HISTORICO_PEDIDOS, read_only=True)
            hoja_actual = libro_actual.active
            filas_eliminadas = max(0, hoja_actual.max_row - 1)
            libro_actual.close()
        except Exception:
            pass

        marca = datetime.now().strftime("%Y%m%d_%H%M%S")
        ruta_respaldo = os.path.join(DIRECTORIO_BASE, "data", f"pedidos_ocr_historico_respaldo_{marca}.xlsx")
        try:
            os.replace(RUTA_HISTORICO_PEDIDOS, ruta_respaldo)
        except OSError as error:
            return jsonify({"error": f"No se pudo reiniciar el histórico: {error}"}), 500

    return jsonify({"ok": True, "filas_eliminadas": filas_eliminadas, "respaldo": os.path.basename(ruta_respaldo)})


# ---------------------------------------------------------------------------
# Visualizador PDF — buscar facturas, listar recientes, descargar PDF
# ---------------------------------------------------------------------------

@app.route("/api/facturas/visualizador", methods=["GET"])
def facturas_visualizador():
    """Busca facturas por query (nombre/código proveedor) o lista recientes."""
    q = request.args.get("q", "").strip()
    if q:
        filas = bd_ocr.buscar_facturas_por_proveedor(q)
    else:
        filas = bd_ocr.listar_facturas_recientes()

    return jsonify({"ok": True, "facturas": [
        {
            "id": f["id"],
            "fecha": f["fecha_registro"].strftime("%Y-%m-%d %H:%M") if f["fecha_registro"] else "—",
            "numero_factura": f["numero_factura"],
            "codigo_proveedor": f["codigo_proveedor"] or "—",
            "nombre_proveedor": f["nombre_proveedor"] or "—",
            "rif_proveedor": f["rif_proveedor"] or "—",
            "tiene_pdf": bool(f.get("nombre_archivo")),
            # 11/09: el visualizador asumía que todo documento guardado era
            # un PDF y siempre usaba el <iframe> — varias facturas se
            # guardan como foto (.jpeg) y el navegador no puede renderizar
            # eso dentro de un visor de PDF (queda "No se pudo cargar el
            # documento PDF"). El frontend usa esto para elegir <img> en
            # vez de <iframe>.
            "es_imagen": (f.get("nombre_archivo") or "").lower().rsplit(".", 1)[-1] in {"png", "jpg", "jpeg", "webp", "bmp"},
        }
        for f in filas
    ]})


@app.route("/api/facturas/<int:factura_id>/pdf", methods=["GET"])
def descargar_pdf_factura(factura_id):
    """Sirve el documento (PDF o imagen) de una factura dada.

    CORREGIDO 11/09: mimetype ya no viene hardcodeado a "application/pdf"
    — varias facturas se guardan como foto (.jpeg), y forzar ese mimetype
    hacía que el navegador intentara abrir bytes de imagen como PDF y
    fallara con "No se pudo cargar el documento PDF" (mismo bug ya
    corregido en Pedidos OCR, ver pedidos_descargar_documento)."""
    nombre, contenido = bd_ocr.obtener_pdf_factura(factura_id)
    if not contenido:
        return jsonify({"error": "No se encontró el PDF de esta factura"}), 404
    tipo_detectado, _ = mimetypes.guess_type(nombre or "")
    return send_file(BytesIO(contenido), mimetype=tipo_detectado or "application/pdf", as_attachment=False, download_name=nombre or "factura.pdf")


if __name__ == "__main__":
    from waitress import serve

    puerto = int(os.environ.get("PUERTO_OCR_PS", 5042))
    print(f"OCR-PS (números de factura para psicotrópicos) + Pedidos OCR — puerto {puerto}, modelo principal: {ocr.MODELO_VISION_QWEN} (Qwen-VL local), respaldo: {ocr.MODELO_VISION_DEEPSEEK} (DeepSeek) -> {ocr.MODELO_VISION} (Gemini)")
    # Waitress (servidor WSGI de producción) en vez del server de desarrollo
    # de Flask — este es el más usado por clientes reales de todos (04/09).
    print(f"Sirviendo con Waitress en el puerto {puerto}…")
    serve(app, host="0.0.0.0", port=puerto, threads=14)
