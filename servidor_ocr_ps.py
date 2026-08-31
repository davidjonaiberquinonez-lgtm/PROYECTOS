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
import os
import threading
import time
import uuid
from datetime import datetime

from flask import Flask, Response, jsonify, render_template, request, send_file
from openpyxl import Workbook, load_workbook
from openpyxl.utils import get_column_letter
from io import BytesIO

import bd_ocr
import ocr_pedidos_core as ocr_pedidos
import ocr_ps_core as ocr
from bin.consultar_proveedor_profit import buscar_proveedor

# ---------------------------------------------------------------------------
# Configuración
# ---------------------------------------------------------------------------

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 30 * 1024 * 1024  # 30 MB

DIRECTORIO_BASE = os.path.dirname(os.path.abspath(__file__))
RUTA_HISTORICO = os.path.join(DIRECTORIO_BASE, "data", "ocr_ps_historico.xlsx")
RUTA_DOCUMENTOS = os.path.join(DIRECTORIO_BASE, "data", "documentos_ps")
ENCABEZADOS_HISTORICO = [
    "Fecha de registro", "Motor", "Numero de Factura",
    "Codigo de Proveedor", "Nombre de Proveedor", "RIF de Proveedor",
    "Ref. Documento",
]
_lock_historico = threading.Lock()

# Pedidos OCR — mismo puerto que OCR-PS (ver servidor_ocr_ps.py arriba),
# módulo aparte para leer notas de entrega manuscritas y pasarlas a una
# tabla limpia (Sede/Articulo/Cantidad/Pediatrico). Una fila del histórico
# = un ARTÍCULO, no una página (una nota trae varios artículos).
RUTA_HISTORICO_PEDIDOS = os.path.join(DIRECTORIO_BASE, "data", "pedidos_ocr_historico.xlsx")
ENCABEZADOS_HISTORICO_PEDIDOS = [
    "Fecha de registro", "Motor", "Pagina", "Archivo", "Sede", "Articulo", "Cantidad", "Pediatrico",
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


def guardar_documento_original(contenido, nombre_archivo):
    """Guarda en disco el PDF/foto TAL CUAL se subió (no las páginas ya
    convertidas a imagen para el OCR) — es la evidencia del documento
    original. Un mismo archivo subido genera un solo nombre guardado acá,
    compartido por todas las páginas/filas que salgan de él, para que
    'Cargar al servidor' después lo suba una sola vez a la BD (ver
    bd_ocr.obtener_o_guardar_documento_ps). Devuelve el nombre guardado
    (para anotarlo en el histórico) o None si no se pudo escribir."""
    try:
        os.makedirs(RUTA_DOCUMENTOS, exist_ok=True)
        extension = nombre_archivo.rsplit(".", 1)[-1].lower() if "." in nombre_archivo else "bin"
        nombre_guardado = f"{datetime.now().strftime('%Y%m%d_%H%M%S')}_{uuid.uuid4().hex[:8]}.{extension}"
        with open(os.path.join(RUTA_DOCUMENTOS, nombre_guardado), "wb") as f:
            f.write(contenido)
        return nombre_guardado
    except OSError as error:
        print(f"[documentos ocr-ps] No se pudo guardar el archivo original: {error}")
        return None


def guardar_en_historico_global(motor, campos, ref_documento=None):
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
        numero_fila = guardar_en_historico_global(resultado["motor"], resultado["campos"], ref_documento)
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
                ])
            libro.save(RUTA_HISTORICO_PEDIDOS)
        except Exception as error:
            print(f"[historico pedidos-ocr] No se pudo guardar en {RUTA_HISTORICO_PEDIDOS}: {error}")


def procesar_imagen_pedido_y_guardar(imagen_bytes, nombre_archivo, numero_pagina):
    resultado = ocr_pedidos.procesar_imagen(imagen_bytes, nombre_archivo, numero_pagina)
    if resultado.get("ok"):
        campos = resultado["campos"]
        guardar_en_historico_pedidos(resultado["motor"], numero_pagina, nombre_archivo, campos.get("sede"), campos.get("items"))
    return resultado


# ---------------------------------------------------------------------------
# Rutas
# ---------------------------------------------------------------------------

@app.route("/")
def index():
    return render_template("ocr_ps.html")


@app.route("/menu")
def menu():
    return render_template("menu.html")


@app.route("/api/estado-ocr", methods=["GET"])
def estado_ocr():
    cadena_respaldo = []
    if ocr.clave_google:
        cadena_respaldo.append(f"{ocr.MODELO_VISION} (Gemini)")
    if ocr.clave_openrouter:
        cadena_respaldo.append(f"{ocr.MODELO_VISION_OPENROUTER} (OpenRouter)")
    return jsonify({
        "conectado": True,
        "detalle": {
            "modelo_vision": ocr.MODELO_VISION_DEEPSEEK,
            "proveedor": "DeepSeek",
            "respaldo": " → ".join(cadena_respaldo) if cadena_respaldo else None,
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
        # Codigo de Proveedor, Nombre de Proveedor, RIF de Proveedor, Ref. Documento
        _fecha_registro, motor, numero_factura, codigo_proveedor, nombre_proveedor, rif_proveedor, ref_documento = fila
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
            bd_ocr.guardar_factura_ps(motor, campos, documento_id=documento_id)
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
    return render_template("ocr_pedidos.html")


@app.route("/api/pedidos/estado-ocr", methods=["GET"])
def pedidos_estado_ocr():
    cadena_respaldo = []
    if ocr_pedidos.clave_google:
        cadena_respaldo.append(f"{ocr_pedidos.MODELO_VISION} (Gemini)")
    if ocr_pedidos.clave_openrouter:
        cadena_respaldo.append(f"{ocr_pedidos.MODELO_VISION_OPENROUTER} (OpenRouter)")
    return jsonify({
        "conectado": True,
        "detalle": {
            "modelo_vision": ocr_pedidos.MODELO_VISION_DEEPSEEK,
            "proveedor": "DeepSeek",
            "respaldo": " → ".join(cadena_respaldo) if cadena_respaldo else None,
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

    if extension == "pdf":
        try:
            imagenes = ocr_pedidos.pdf_a_imagenes(contenido)
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
                resultado = procesar_imagen_pedido_y_guardar(imagen_bytes, nombre, indice)
                yield json.dumps({"evento": "pagina", "resultado": resultado}) + "\n"
            yield json.dumps({"evento": "fin"}) + "\n"
        except (BrokenPipeError, ConnectionAbortedError, ConnectionResetError, GeneratorExit):
            return

    return Response(generar(), mimetype="application/x-ndjson")


@app.route("/api/pedidos/exportar_excel", methods=["POST"])
def pedidos_exportar_excel():
    """Arma la tabla limpia final: UNA FILA POR ARTÍCULO (no por página),
    con la sede repetida en cada fila de su nota — es el pedido a la
    factura, "tabla limpia en excel" que pidió el usuario, con exactamente
    los 4 campos: Sede, Articulo, Cantidad, Pediatrico."""
    payload = request.get_json(silent=True) or {}
    resultados = payload.get("resultados", [])

    libro = Workbook()
    hoja = libro.active
    assert hoja is not None
    hoja.title = "Pedidos OCR"

    encabezados = ["Sede", "Articulo", "Cantidad", "Pediatrico"]
    hoja.append(encabezados)

    for fila in resultados:
        campos = fila.get("campos") or {}
        sede = campos.get("sede")
        for item in campos.get("items") or []:
            hoja.append([
                sede,
                item.get("articulo"),
                item.get("cantidad"),
                "Si" if item.get("pediatrico") else "No",
            ])

    anchos = [14, 46, 14, 12]
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


@app.route("/api/pedidos/historico/estado", methods=["GET"])
def pedidos_estado_historico():
    if not os.path.exists(RUTA_HISTORICO_PEDIDOS):
        return jsonify({"existe": False, "filas": 0})
    with _lock_historico_pedidos:
        libro = load_workbook(RUTA_HISTORICO_PEDIDOS, read_only=True)
        hoja = libro.active
        assert hoja is not None
        filas = max(0, hoja.max_row - 1)
        libro.close()
    return jsonify({"existe": True, "filas": filas})


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


if __name__ == "__main__":
    puerto = int(os.environ.get("PUERTO_OCR_PS", 5042))
    print(f"OCR-PS (números de factura para psicotrópicos) — puerto {puerto}, modelo principal: {ocr.MODELO_VISION_DEEPSEEK} (DeepSeek), respaldo: {ocr.MODELO_VISION} (Gemini)")
    app.run(host="0.0.0.0", port=puerto, debug=True, threaded=True, use_reloader=False)
