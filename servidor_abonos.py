"""
servidor_abonos.py — Escáner OCR de Recibos de Cobro (abonos de clientes)

Sube una foto o PDF de un recibo de cobro (libreta correlativa manuscrita)
y extrae, por cada página/imagen: número de recibo, fecha, cliente,
código/RIF, dirección fiscal, monto, cantidad en texto, concepto, las
facturas referenciadas en el concepto, y la forma de pago marcada.

La inteligencia de extracción (Gemini + respaldo OpenRouter) vive en
ocr_abonos_core.py. Este archivo solo se encarga de la interfaz web
propia: la página HTML, exportar a Excel al vuelo, y el histórico
acumulado en /data/abonos_historico.xlsx — mismo patrón que
servidor_retenciones.py, para el otro tipo de documento.

Ejecutar:
    python servidor_abonos.py
"""

import json
import os
import threading
import time
from datetime import datetime

from flask import Flask, Response, jsonify, render_template, request, send_file
from openpyxl import Workbook, load_workbook
from openpyxl.utils import get_column_letter
from io import BytesIO

import bd_ocr
import ocr_abonos_core as ocr

# ---------------------------------------------------------------------------
# Configuración
# ---------------------------------------------------------------------------

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 30 * 1024 * 1024  # 30 MB

# ---------------------------------------------------------------------------
# Histórico global: cada extracción exitosa queda guardada acá, no solo en
# la tabla de la página actual — así no se pierde nada aunque nadie pulse
# "Exportar a Excel" en el momento.
# ---------------------------------------------------------------------------

DIRECTORIO_BASE = os.path.dirname(os.path.abspath(__file__))
RUTA_HISTORICO = os.path.join(DIRECTORIO_BASE, "data", "abonos_historico.xlsx")
ENCABEZADOS_HISTORICO = [
    "Fecha de registro", "Archivo", "Motor",
    "N Recibo", "Fecha", "Cliente", "Codigo/RIF", "Direccion Fiscal",
    "Monto", "Cantidad (texto)", "Concepto", "Facturas Referenciadas", "Forma de Pago",
    "Indice de confianza", "Validado automaticamente",
]
_lock_historico = threading.Lock()


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
    hoja.title = "Histórico Abonos"
    hoja.append(ENCABEZADOS_HISTORICO)
    for indice, encabezado in enumerate(ENCABEZADOS_HISTORICO, start=1):
        hoja.column_dimensions[get_column_letter(indice)].width = max(14, len(encabezado) + 2)
    hoja.freeze_panes = "A2"
    return libro, hoja


def guardar_en_historico_global(nombre_archivo, motor, campos, confianza=None, validado_manual=False):
    """Agrega una fila al Excel histórico en /data. Protegido con un lock
    para que dos páginas procesándose casi al mismo tiempo no se pisen al
    abrir/guardar el archivo. Si falla (ej. el archivo está abierto en
    Excel en ese momento), no se interrumpe el escaneo — solo se registra
    en consola, el dato sigue disponible igual en la tabla de la página."""
    with _lock_historico:
        try:
            libro, hoja = _abrir_o_crear_historico()
            hoja.append([
                datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                nombre_archivo,
                motor,
                campos.get("numero_recibo"),
                campos.get("fecha"),
                campos.get("cliente_nombre"),
                campos.get("cliente_codigo_o_rif"),
                campos.get("direccion_fiscal"),
                campos.get("monto"),
                campos.get("cantidad_texto"),
                campos.get("concepto"),
                campos.get("facturas_referenciadas"),
                campos.get("forma_pago"),
                (confianza or {}).get("indice"),
                "No — guardado a mano" if validado_manual else "Sí",
            ])
            libro.save(RUTA_HISTORICO)
        except Exception as error:
            print(f"[historico abonos] No se pudo guardar la fila en {RUTA_HISTORICO}: {error}")
    # Ojo: acá NO se guarda en la base de datos MySQL — eso solo pasa
    # cuando alguien revisa el Excel y pulsa "Cargar al servidor" (ver
    # /api/historico/cargar_al_servidor más abajo). El Excel es el
    # borrador; el servidor solo recibe lo ya revisado.


def procesar_imagen_y_guardar(imagen_bytes, nombre_archivo, numero_pagina):
    """Envoltorio sobre ocr.procesar_imagen que además registra el
    resultado en el histórico local — SOLO si el índice de confianza
    combinado (70% qué tan reconocible fue el formato del recibo, 30% qué
    tan legible salió lo manuscrito) llegó al umbral de auto-validación.
    Si no, el resultado queda disponible en la tabla de la página para
    que un supervisor lo revise/corrija y lo agregue a mano — ver
    /api/historico/guardar_manual — nunca se pierde, solo no se auto-
    confía en él."""
    resultado = ocr.procesar_imagen(imagen_bytes, nombre_archivo, numero_pagina)
    if resultado.get("ok") and resultado.get("confianza", {}).get("validado"):
        guardar_en_historico_global(nombre_archivo, resultado["motor"], resultado["campos"], resultado["confianza"])
        resultado["guardado_en_historico"] = True
    else:
        resultado["guardado_en_historico"] = False
    return resultado


# ---------------------------------------------------------------------------
# Rutas
# ---------------------------------------------------------------------------

@app.route("/")
def index():
    return render_template("abonos.html")


@app.route("/menu")
def menu():
    return render_template("menu.html")


@app.route("/api/estado-ocr", methods=["GET"])
def estado_ocr():
    return jsonify({
        "conectado": True,
        "detalle": {
            "modelo_vision": ocr.MODELO_VISION,
            "proveedor": "Google Gemini",
            "respaldo": ocr.MODELO_VISION_OPENROUTER if ocr.clave_openrouter else None,
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
                    # Freno preventivo: la capa gratuita de Gemini limita a
                    # ~15 peticiones/minuto.
                    time.sleep(4)
                nombre = f"{nombre_archivo}-pagina{indice}.png" if extension == "pdf" else nombre_archivo
                resultado = procesar_imagen_y_guardar(imagen_bytes, nombre, indice)
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
    hoja.title = "Abonos"

    encabezados = [
        "Página", "Archivo", "Motor", "N Recibo", "Fecha", "Cliente", "Codigo/RIF",
        "Direccion Fiscal", "Monto", "Cantidad (texto)", "Concepto", "Facturas Referenciadas", "Forma de Pago",
        "Indice de confianza", "Validado automaticamente",
    ]
    hoja.append(encabezados)

    for fila in resultados:
        campos = fila.get("campos", {})
        confianza = fila.get("confianza") or {}
        hoja.append([
            fila.get("pagina"),
            fila.get("archivo"),
            fila.get("motor"),
            campos.get("numero_recibo"),
            campos.get("fecha"),
            campos.get("cliente_nombre"),
            campos.get("cliente_codigo_o_rif"),
            campos.get("direccion_fiscal"),
            campos.get("monto"),
            campos.get("cantidad_texto"),
            campos.get("concepto"),
            campos.get("facturas_referenciadas"),
            campos.get("forma_pago"),
            confianza.get("indice"),
            "Sí" if confianza.get("validado") else "No",
        ])

    for indice, encabezado in enumerate(encabezados, start=1):
        hoja.column_dimensions[get_column_letter(indice)].width = max(14, len(encabezado) + 2)

    buffer = BytesIO()
    libro.save(buffer)
    buffer.seek(0)

    return send_file(
        buffer,
        as_attachment=True,
        download_name="abonos.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


@app.route("/api/historico/guardar_manual", methods=["POST"])
def guardar_manual():
    """Un recibo que quedó por debajo del umbral de auto-validación (o que
    un supervisor prefirió revisar/corregir antes de confiar en él) — el
    supervisor lo agrega al histórico a mano desde la tabla de la página,
    con los campos ya corregidos si hizo falta. Queda marcado en el Excel
    como "guardado a mano", no como validación automática."""
    payload = request.get_json(silent=True) or {}
    campos = payload.get("campos") or {}
    if not any((campos.get(c) for c in ocr.CAMPOS_DATOS)):
        return jsonify({"error": "No hay campos con datos para guardar"}), 400

    guardar_en_historico_global(
        payload.get("archivo") or "(sin nombre)",
        payload.get("motor") or "manual",
        campos,
        confianza=payload.get("confianza"),
        validado_manual=True,
    )
    return jsonify({"ok": True})


@app.route("/api/historico/estado", methods=["GET"])
def estado_historico():
    if not os.path.exists(RUTA_HISTORICO):
        return jsonify({"existe": False, "filas": 0, "filas_cargadas": 0, "filas_pendientes": 0})
    with _lock_historico:
        libro = load_workbook(RUTA_HISTORICO, read_only=True)
        hoja = libro.active
        assert hoja is not None
        filas = max(0, hoja.max_row - 1)  # menos el encabezado
        libro.close()
    cargadas = min(bd_ocr.obtener_progreso("abonos"), filas)
    return jsonify({"existe": True, "filas": filas, "filas_cargadas": cargadas, "filas_pendientes": max(0, filas - cargadas)})


@app.route("/api/historico/cargar_al_servidor", methods=["POST"])
def cargar_al_servidor():
    """Lee el Excel histórico completo y sube a la base de datos MySQL
    solo las filas que todavía no se habían cargado. Igual que en
    retenciones: el Excel es el borrador que se revisa a mano, este botón
    es el que confirma y manda lo revisado al servidor."""
    if not os.path.exists(RUTA_HISTORICO):
        return jsonify({"ok": True, "filas_nuevas": 0, "filas_totales": 0})

    with _lock_historico:
        libro = load_workbook(RUTA_HISTORICO, read_only=True)
        hoja = libro.active
        assert hoja is not None
        filas = list(hoja.iter_rows(min_row=2, values_only=True))
        libro.close()

    ya_cargadas = bd_ocr.obtener_progreso("abonos")
    filas_nuevas = filas[ya_cargadas:]

    cargadas_ahora = 0
    for fila in filas_nuevas:
        # Columnas: Fecha de registro, Archivo, Motor, N Recibo, Fecha, Cliente, Codigo/RIF,
        # Direccion Fiscal, Monto, Cantidad (texto), Concepto, Facturas Referenciadas,
        # Forma de Pago, Indice de confianza, Validado automaticamente
        (
            _fecha_registro, _archivo, motor, numero_recibo, fecha, cliente_nombre, cliente_codigo_o_rif,
            direccion_fiscal, monto, cantidad_texto, concepto, facturas_referenciadas, forma_pago,
            indice_confianza, validado_texto,
        ) = fila
        campos = {
            "numero_recibo": numero_recibo, "fecha": fecha, "cliente_nombre": cliente_nombre,
            "cliente_codigo_o_rif": cliente_codigo_o_rif, "direccion_fiscal": direccion_fiscal,
            "monto": monto, "cantidad_texto": cantidad_texto, "concepto": concepto,
            "facturas_referenciadas": facturas_referenciadas, "forma_pago": forma_pago,
        }
        try:
            bd_ocr.guardar_abono(
                motor, campos,
                indice_confianza=indice_confianza,
                validado=(validado_texto == "Sí"),
            )
            cargadas_ahora += 1
        except Exception as error:
            print(f"[bd_ocr] No se pudo cargar una fila del histórico: {error}")
            break

    bd_ocr.actualizar_progreso("abonos", ya_cargadas + cargadas_ahora)
    return jsonify({"ok": True, "filas_nuevas": cargadas_ahora, "filas_totales": len(filas)})


@app.route("/api/historico/descargar", methods=["GET"])
def descargar_historico():
    if not os.path.exists(RUTA_HISTORICO):
        return jsonify({"error": "Todavía no hay extracciones guardadas en el histórico"}), 404
    with _lock_historico:
        return send_file(
            RUTA_HISTORICO,
            as_attachment=True,
            download_name="abonos_historico.xlsx",
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )


@app.route("/api/historico/reiniciar", methods=["POST"])
def reiniciar_historico():
    """Vacía el progreso acumulado del histórico. No lo borra a lo loco:
    renombra el archivo actual con fecha/hora como respaldo (queda en
    /data por si hace falta recuperar algo), y la próxima extracción
    exitosa arranca un histórico nuevo desde cero."""
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
        ruta_respaldo = os.path.join(DIRECTORIO_BASE, "data", f"abonos_historico_respaldo_{marca}.xlsx")
        try:
            os.replace(RUTA_HISTORICO, ruta_respaldo)
        except OSError as error:
            return jsonify({"error": f"No se pudo reiniciar el histórico: {error}"}), 500

    return jsonify({"ok": True, "filas_eliminadas": filas_eliminadas, "respaldo": os.path.basename(ruta_respaldo)})


if __name__ == "__main__":
    puerto = int(os.environ.get("PUERTO_ABONOS", 5040))
    print(f"Escáner de recibos de cobro (abonos) — usando modelo de visión Google: {ocr.MODELO_VISION}")
    app.run(host="0.0.0.0", port=puerto, debug=True, threaded=True, use_reloader=False)
