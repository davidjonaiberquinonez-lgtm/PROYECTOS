"""
servidor_retenciones.py — Escáner OCR de Comprobantes de Retención de IVA

Sube un PDF (una o varias páginas) o una imagen de un comprobante de
retención de IVA venezolano y extrae, por cada página/imagen:

  - fecha
  - nro_comprobante
  - cliente (el Agente de Retención que emite el comprobante — no somos
    nosotros, nosotros aparecemos como "Proveedor"/"Sujeto Retenido")
  - nro_factura
  - rif_cliente (RIF de ese mismo Agente de Retención)
  - monto_retenido

La inteligencia de extracción (Gemini + respaldo OpenRouter, prompt,
manejo de PDF) vive en ocr_retenciones_core.py, compartida con
api_retenciones_ocr.py (la versión pública sin interfaz que usa otro
equipo). Este archivo solo se encarga de la interfaz web propia: la
página HTML, exportar a Excel al vuelo, y el histórico acumulado en
/data/retenciones_historico.xlsx.

Ejecutar:
    python servidor_retenciones.py
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

import sys

sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), "bin"))

import bd_ocr
import ocr_retenciones_core as ocr
from consultar_proveedor_profit import buscar_proveedor

# ---------------------------------------------------------------------------
# Configuración
# ---------------------------------------------------------------------------

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 30 * 1024 * 1024  # 30 MB

# ---------------------------------------------------------------------------
# Histórico global: cada extracción exitosa (de cualquier sesión/usuario)
# queda guardada acá, no solo en la tabla de la página actual — así no se
# pierde nada aunque nadie pulse "Exportar a Excel" en el momento.
# ---------------------------------------------------------------------------

DIRECTORIO_BASE = os.path.dirname(os.path.abspath(__file__))
RUTA_HISTORICO = os.path.join(DIRECTORIO_BASE, "data", "retenciones_historico.xlsx")
ENCABEZADOS_HISTORICO = [
    "Fecha de registro", "Archivo", "Motor",
    "Fecha", "Nro Comprobante", "Cliente", "Nro Factura", "RIF Cliente", "Monto Retenido",
]
_lock_historico = threading.Lock()

# Histórico del DETALLE (una fila por factura de la tabla "Compras Internas
# e Importaciones" de cada comprobante) — archivo aparte, no columnas
# nuevas en ENCABEZADOS_HISTORICO de arriba: ese esquema tiene posiciones
# fijas que /api/historico/cargar_al_servidor desempaqueta por índice
# (ver más abajo) y bd_ocr.guardar_retencion espera esos 6 campos exactos
# — agregar columnas ahí correría todo y rompería esa carga a MySQL.
RUTA_HISTORICO_DETALLE = os.path.join(DIRECTORIO_BASE, "data", "retenciones_detalle_historico.xlsx")
ENCABEZADOS_HISTORICO_DETALLE = [
    "Fecha de registro", "Archivo", "Motor", "Nro Comprobante",
    "Numero Factura", "Numero Control", "Numero Nota Debito", "Numero Nota Credito",
    "Tipo Trans", "Documento Afectado", "Total Compras Con Iva", "Total Compras Sin Iva",
    "Base Imponible", "% Alicuota", "Impuesto IVA", "IVA Retenido",
]
_lock_historico_detalle = threading.Lock()


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
    hoja.title = "Histórico OCR"
    hoja.append(ENCABEZADOS_HISTORICO)
    for indice, encabezado in enumerate(ENCABEZADOS_HISTORICO, start=1):
        hoja.column_dimensions[get_column_letter(indice)].width = max(14, len(encabezado) + 2)
    hoja.freeze_panes = "A2"
    return libro, hoja


def guardar_en_historico_global(nombre_archivo, motor, campos):
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
                campos.get("fecha"),
                campos.get("nro_comprobante"),
                campos.get("cliente"),
                campos.get("nro_factura"),
                campos.get("rif_cliente"),
                campos.get("monto_retenido"),
            ])
            libro.save(RUTA_HISTORICO)
        except Exception as error:
            print(f"[historico OCR] No se pudo guardar la fila en {RUTA_HISTORICO}: {error}")
    # Ojo: acá NO se guarda en la base de datos MySQL — eso solo pasa
    # cuando alguien revisa el Excel y pulsa "Cargar al servidor" (ver
    # /api/historico/cargar_al_servidor más abajo). El Excel es el
    # borrador; el servidor solo recibe lo ya revisado.


def _abrir_o_crear_historico_detalle():
    if os.path.exists(RUTA_HISTORICO_DETALLE):
        libro = load_workbook(RUTA_HISTORICO_DETALLE)
        hoja = libro.active
        assert hoja is not None
        return libro, hoja

    os.makedirs(os.path.dirname(RUTA_HISTORICO_DETALLE), exist_ok=True)
    libro = Workbook()
    hoja = libro.active
    assert hoja is not None
    hoja.title = "Detalle facturas"
    hoja.append(ENCABEZADOS_HISTORICO_DETALLE)
    for indice, encabezado in enumerate(ENCABEZADOS_HISTORICO_DETALLE, start=1):
        hoja.column_dimensions[get_column_letter(indice)].width = max(14, len(encabezado) + 2)
    hoja.freeze_panes = "A2"
    return libro, hoja


def guardar_detalle_en_historico(nombre_archivo, motor, nro_comprobante, filas_detalle):
    """Una fila por factura de "Compras Internas e Importaciones" — si el
    comprobante trae 3 facturas, se agregan 3 filas acá (todas con el
    mismo Nro Comprobante). No interrumpe el escaneo si falla."""
    if not filas_detalle:
        return
    with _lock_historico_detalle:
        try:
            libro, hoja = _abrir_o_crear_historico_detalle()
            fecha_registro = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            for fila in filas_detalle:
                hoja.append([
                    fecha_registro, nombre_archivo, motor, nro_comprobante,
                    fila.get("numero_factura"), fila.get("numero_control"),
                    fila.get("numero_nota_debito"), fila.get("numero_nota_credito"),
                    fila.get("tipo_trans"), fila.get("documento_afectado"),
                    fila.get("total_compras_con_iva"), fila.get("total_compras_sin_iva"),
                    fila.get("base_imponible"), fila.get("porcentaje_alicuota"),
                    fila.get("impuesto_iva"), fila.get("iva_retenido"),
                ])
            libro.save(RUTA_HISTORICO_DETALLE)
        except Exception as error:
            print(f"[historico detalle OCR] No se pudo guardar en {RUTA_HISTORICO_DETALLE}: {error}")


def procesar_imagen_y_guardar(imagen_bytes, nombre_archivo, numero_pagina):
    """Envoltorio sobre ocr.procesar_imagen que además registra el
    resultado exitoso en el histórico local — eso es específico de esta
    interfaz, no de la inteligencia en sí (la API pública no lo hace)."""
    resultado = ocr.procesar_imagen(imagen_bytes, nombre_archivo, numero_pagina)
    if resultado.get("ok"):
        campos = resultado["campos"]
        guardar_en_historico_global(nombre_archivo, resultado["motor"], campos)
        guardar_detalle_en_historico(nombre_archivo, resultado["motor"], campos.get("nro_comprobante"), campos.get("detalle"))
    return resultado


# ---------------------------------------------------------------------------
# Rutas
# ---------------------------------------------------------------------------

@app.route("/")
def index():
    return render_template("retenciones.html")


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
                    # ~15 peticiones/minuto. Sin esta pausa, un PDF de
                    # varias páginas dispara todas casi de corrido y choca
                    # con ese límite (429) aunque cada página individual
                    # responda rápido.
                    time.sleep(4)
                nombre = f"{nombre_archivo}-pagina{indice}.png" if extension == "pdf" else nombre_archivo
                resultado = procesar_imagen_y_guardar(imagen_bytes, nombre, indice)
                yield json.dumps({"evento": "pagina", "resultado": resultado}) + "\n"
            yield json.dumps({"evento": "fin"}) + "\n"
        except (BrokenPipeError, ConnectionAbortedError, ConnectionResetError, GeneratorExit):
            # El cliente cerró la conexión (recargó la página, cerró la
            # pestaña) a mitad del stream — no es un error del servidor,
            # simplemente se deja de generar el resto de las páginas.
            return

    return Response(generar(), mimetype="application/x-ndjson")


@app.route("/api/proveedor/buscar", methods=["GET"])
def proveedor_buscar():
    """Busca proveedores reales en Profit (tabla `prov`, CRISTM25) por
    nombre, código o RIF — usado para confirmar/completar el proveedor de
    una factura escaneada antes de guardarla, en vez de dejar el nombre
    crudo que sacó el OCR (que puede venir con variaciones de escritura).
    Al menos un filtro (nombre/codigo/rif) es obligatorio."""
    resultado = buscar_proveedor(
        nombre=request.args.get("nombre", ""),
        codigo=request.args.get("codigo", ""),
        rif=request.args.get("rif", ""),
        incluir_inactivos=request.args.get("incluir_inactivos", "").lower() in ("1", "true", "si"),
        limite=int(request.args.get("limite", 20) or 20),
    )
    return jsonify(resultado), (200 if resultado.get("ok") else 400)


@app.route("/api/exportar_excel", methods=["POST"])
def exportar_excel():
    payload = request.get_json(silent=True) or {}
    resultados = payload.get("resultados", [])

    libro = Workbook()
    hoja = libro.active
    assert hoja is not None
    hoja.title = "Retenciones IVA"

    encabezados = ["Página", "Archivo", "Motor", "Fecha", "Nro Comprobante", "Cliente", "Nro Factura", "RIF Cliente", "Monto Retenido"]
    hoja.append(encabezados)

    for fila in resultados:
        campos = fila.get("campos", {})
        hoja.append([
            fila.get("pagina"),
            fila.get("archivo"),
            fila.get("motor"),
            campos.get("fecha"),
            campos.get("nro_comprobante"),
            campos.get("cliente"),
            campos.get("nro_factura"),
            campos.get("rif_cliente"),
            campos.get("monto_retenido"),
        ])

    for indice, encabezado in enumerate(encabezados, start=1):
        hoja.column_dimensions[get_column_letter(indice)].width = max(14, len(encabezado) + 2)

    # Segunda hoja: el detalle fila-por-factura de "Compras Internas e
    # Importaciones" de cada comprobante — antes se perdía, solo quedaba
    # el resumen de arriba (reportado en vivo, 31/08).
    hoja_detalle = libro.create_sheet("Detalle facturas")
    encabezados_detalle = [
        "Página", "Nro Comprobante", "Numero Factura", "Numero Control",
        "Numero Nota Debito", "Numero Nota Credito", "Tipo Trans", "Documento Afectado",
        "Total Compras Con Iva", "Total Compras Sin Iva", "Base Imponible",
        "% Alicuota", "Impuesto IVA", "IVA Retenido",
    ]
    hoja_detalle.append(encabezados_detalle)
    for fila in resultados:
        campos = fila.get("campos", {})
        for detalle in campos.get("detalle") or []:
            hoja_detalle.append([
                fila.get("pagina"), campos.get("nro_comprobante"),
                detalle.get("numero_factura"), detalle.get("numero_control"),
                detalle.get("numero_nota_debito"), detalle.get("numero_nota_credito"),
                detalle.get("tipo_trans"), detalle.get("documento_afectado"),
                detalle.get("total_compras_con_iva"), detalle.get("total_compras_sin_iva"),
                detalle.get("base_imponible"), detalle.get("porcentaje_alicuota"),
                detalle.get("impuesto_iva"), detalle.get("iva_retenido"),
            ])
    for indice, encabezado in enumerate(encabezados_detalle, start=1):
        hoja_detalle.column_dimensions[get_column_letter(indice)].width = max(14, len(encabezado) + 2)

    buffer = BytesIO()
    libro.save(buffer)
    buffer.seek(0)

    return send_file(
        buffer,
        as_attachment=True,
        download_name="retenciones_iva.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


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
    cargadas = min(bd_ocr.obtener_progreso("retenciones"), filas)
    return jsonify({"existe": True, "filas": filas, "filas_cargadas": cargadas, "filas_pendientes": max(0, filas - cargadas)})


@app.route("/api/historico/cargar_al_servidor", methods=["POST"])
def cargar_al_servidor():
    """Lee el Excel histórico completo y sube a la base de datos MySQL
    solo las filas que todavía no se habían cargado (usa el progreso
    guardado en bd_ocr.progreso_carga) — así revisar/corregir el Excel a
    mano antes de pulsar este botón no se pierde, y pulsarlo de nuevo
    nunca duplica lo que ya se cargó."""
    if not os.path.exists(RUTA_HISTORICO):
        return jsonify({"ok": True, "filas_nuevas": 0, "filas_totales": 0})

    with _lock_historico:
        libro = load_workbook(RUTA_HISTORICO, read_only=True)
        hoja = libro.active
        assert hoja is not None
        filas = list(hoja.iter_rows(min_row=2, values_only=True))
        libro.close()

    ya_cargadas = bd_ocr.obtener_progreso("retenciones")
    filas_nuevas = filas[ya_cargadas:]

    cargadas_ahora = 0
    for fila in filas_nuevas:
        # Columnas: Fecha de registro, Archivo, Motor, Fecha, Nro Comprobante, Cliente, Nro Factura, RIF Cliente, Monto Retenido
        _fecha_registro, _archivo, motor, fecha, nro_comprobante, cliente, nro_factura, rif_cliente, monto_retenido = fila
        campos = {
            "fecha": fecha, "nro_comprobante": nro_comprobante, "cliente": cliente,
            "nro_factura": nro_factura, "rif_cliente": rif_cliente, "monto_retenido": monto_retenido,
        }
        try:
            bd_ocr.guardar_retencion(motor, campos)
            cargadas_ahora += 1
        except Exception as error:
            print(f"[bd_ocr] No se pudo cargar una fila del histórico: {error}")
            break  # se corta acá — el progreso queda hasta la última fila que sí se cargó

    bd_ocr.actualizar_progreso("retenciones", ya_cargadas + cargadas_ahora)
    return jsonify({"ok": True, "filas_nuevas": cargadas_ahora, "filas_totales": len(filas)})


@app.route("/api/historico/descargar", methods=["GET"])
def descargar_historico():
    if not os.path.exists(RUTA_HISTORICO):
        return jsonify({"error": "Todavía no hay extracciones guardadas en el histórico"}), 404
    with _lock_historico:
        return send_file(
            RUTA_HISTORICO,
            as_attachment=True,
            download_name="retenciones_historico.xlsx",
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
            filas_eliminadas = max(0, hoja_actual.max_row - 1)  # menos el encabezado
            libro_actual.close()
        except Exception:
            pass  # si no se pudo leer el conteo, igual se reinicia

        marca = datetime.now().strftime("%Y%m%d_%H%M%S")
        ruta_respaldo = os.path.join(DIRECTORIO_BASE, "data", f"retenciones_historico_respaldo_{marca}.xlsx")
        try:
            os.replace(RUTA_HISTORICO, ruta_respaldo)
        except OSError as error:
            return jsonify({"error": f"No se pudo reiniciar el histórico: {error}"}), 500

    return jsonify({"ok": True, "filas_eliminadas": filas_eliminadas, "respaldo": os.path.basename(ruta_respaldo)})


@app.route("/api/historico_detalle/estado", methods=["GET"])
def estado_historico_detalle():
    if not os.path.exists(RUTA_HISTORICO_DETALLE):
        return jsonify({"existe": False, "filas": 0})
    with _lock_historico_detalle:
        libro = load_workbook(RUTA_HISTORICO_DETALLE, read_only=True)
        hoja = libro.active
        assert hoja is not None
        filas = max(0, hoja.max_row - 1)
        libro.close()
    return jsonify({"existe": True, "filas": filas})


@app.route("/api/historico_detalle/descargar", methods=["GET"])
def descargar_historico_detalle():
    if not os.path.exists(RUTA_HISTORICO_DETALLE):
        return jsonify({"error": "Todavía no hay detalle de facturas guardado en el histórico"}), 404
    with _lock_historico_detalle:
        return send_file(
            RUTA_HISTORICO_DETALLE,
            as_attachment=True,
            download_name="retenciones_detalle_historico.xlsx",
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )


@app.route("/api/historico_detalle/reiniciar", methods=["POST"])
def reiniciar_historico_detalle():
    if not os.path.exists(RUTA_HISTORICO_DETALLE):
        return jsonify({"ok": True, "filas_eliminadas": 0})

    with _lock_historico_detalle:
        filas_eliminadas = 0
        try:
            libro_actual = load_workbook(RUTA_HISTORICO_DETALLE, read_only=True)
            hoja_actual = libro_actual.active
            filas_eliminadas = max(0, hoja_actual.max_row - 1)
            libro_actual.close()
        except Exception:
            pass

        marca = datetime.now().strftime("%Y%m%d_%H%M%S")
        ruta_respaldo = os.path.join(DIRECTORIO_BASE, "data", f"retenciones_detalle_historico_respaldo_{marca}.xlsx")
        try:
            os.replace(RUTA_HISTORICO_DETALLE, ruta_respaldo)
        except OSError as error:
            return jsonify({"error": f"No se pudo reiniciar el histórico de detalle: {error}"}), 500

    return jsonify({"ok": True, "filas_eliminadas": filas_eliminadas, "respaldo": os.path.basename(ruta_respaldo)})


if __name__ == "__main__":
    puerto = int(os.environ.get("PUERTO_RETENCIONES", 5030))
    print(f"Escáner de retenciones IVA — modelo principal: {ocr.MODELO_VISION_DEEPSEEK} (DeepSeek), respaldo: {ocr.MODELO_VISION} (Gemini)")
    app.run(host="0.0.0.0", port=puerto, debug=True, threaded=True, use_reloader=False)
