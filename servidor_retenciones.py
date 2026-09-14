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
import uuid
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

# Fecha de modificación de ESTE archivo en disco, mostrada chiquita en el
# header de la página (14/09) — para poder confirmar de un vistazo, sin
# entrar por SSH/terminal, que el servidor corriendo cargó el código
# actual y no quedó un proceso viejo pegado sirviendo una versión
# anterior (pasó dos veces antes en este proyecto: /pedidos en Pedidos
# OCR y esta misma página de Retenciones, ambas por procesos que llevaban
# días corriendo desde antes del último cambio).
VERSION_SERVIDOR = datetime.fromtimestamp(os.path.getmtime(__file__)).strftime("%d/%m/%Y %H:%M")

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
    "Ref Documento",  # agregado 09/09 (endpoint de descarga del PDF por numero_factura,
    # ver bd_ocr.obtener_documento_retencion_por_factura) — va al final para no correr
    # las columnas de un histórico que ya tenía filas reales (mismo criterio que
    # ENCABEZADOS_HISTORICO_PEDIDOS en servidor_ocr_ps.py).
]
_lock_historico = threading.Lock()

# PDF/foto TAL CUAL se subió (no las páginas ya convertidas a imagen para
# el OCR) — es la evidencia del documento original, y lo que se termina
# guardando como BLOB en MySQL (documentos_retenciones) cuando se pulsa
# "Cargar al servidor". Antes del 09/09 esto no se guardaba en ningún
# lado: una vez procesado el documento, el PDF original se perdía — no
# había forma de recuperarlo después por numero_factura.
RUTA_DOCUMENTOS = os.path.join(DIRECTORIO_BASE, "data", "documentos_retenciones")

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
        # Migración: "Ref Documento" se agregó el 09/09 — un histórico
        # creado antes de eso todavía tiene el encabezado viejo (9
        # columnas). Se agrega el título que falta en la 10ma para que las
        # filas nuevas (que sí mandan la referencia) no queden sin encabezado.
        if hoja.cell(row=1, column=len(ENCABEZADOS_HISTORICO)).value != "Ref Documento":
            hoja.cell(row=1, column=len(ENCABEZADOS_HISTORICO), value="Ref Documento")
            libro.save(RUTA_HISTORICO)
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


def guardar_documento_original(contenido, nombre_archivo):
    """Guarda en disco el PDF/foto TAL CUAL se subió (mismo patrón que
    servidor_ocr_ps.py) — un mismo archivo subido genera un solo nombre
    guardado acá, compartido por todas las páginas que salgan de él, para
    que 'Cargar al servidor' después lo suba una sola vez a la BD (ver
    bd_ocr.obtener_o_guardar_documento_retencion). Devuelve el nombre
    guardado (para anotarlo en el histórico) o None si no se pudo escribir."""
    try:
        os.makedirs(RUTA_DOCUMENTOS, exist_ok=True)
        extension = nombre_archivo.rsplit(".", 1)[-1].lower() if "." in nombre_archivo else "bin"
        nombre_guardado = f"{datetime.now().strftime('%Y%m%d_%H%M%S')}_{uuid.uuid4().hex[:8]}.{extension}"
        with open(os.path.join(RUTA_DOCUMENTOS, nombre_guardado), "wb") as f:
            f.write(contenido)
        return nombre_guardado
    except OSError as error:
        print(f"[documentos retenciones] No se pudo guardar el archivo original: {error}")
        return None


def guardar_en_historico_global(nombre_archivo, motor, campos, ref_documento=None):
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
                ref_documento,
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


def procesar_imagen_y_guardar(imagen_bytes, nombre_archivo, numero_pagina, ref_documento=None):
    """Envoltorio sobre ocr.procesar_imagen que además registra el
    resultado exitoso en el histórico local — eso es específico de esta
    interfaz, no de la inteligencia en sí (la API pública no lo hace).
    `ref_documento` (09/09) es el nombre con el que se guardó el PDF/foto
    original en disco (ver guardar_documento_original) — se anota en el
    histórico para poder recuperarlo después desde "Cargar al servidor".

    Verificación de retención (14/09, a pedido explícito del usuario, ver
    ocr.verificar_retencion): si el comprobante no identifica bien a
    Cristmedicals como Proveedor o no trae Periodo Fiscal, NO se guarda en
    ninguno de los dos históricos — queda solo en pantalla, marcado con el
    motivo del fallo, para que alguien lo revise a mano contra el
    documento original antes de procesarlo por otra vía. `resultado["ok"]`
    sigue siendo True (el OCR sí extrajo algo) — lo que cambia es
    `resultado["verificacion"]`, que el frontend usa para bloquear/avisar."""
    resultado = ocr.procesar_imagen(imagen_bytes, nombre_archivo, numero_pagina)
    if resultado.get("ok"):
        campos = resultado["campos"]
        verificacion = ocr.verificar_retencion(campos)
        resultado["verificacion"] = verificacion
        if verificacion["ok"]:
            guardar_en_historico_global(nombre_archivo, resultado["motor"], campos, ref_documento)
            guardar_detalle_en_historico(nombre_archivo, resultado["motor"], campos.get("nro_comprobante"), campos.get("detalle"))
    return resultado


# ---------------------------------------------------------------------------
# Rutas
# ---------------------------------------------------------------------------

@app.route("/")
def index():
    return render_template("retenciones.html", version_servidor=VERSION_SERVIDOR)


@app.route("/menu")
def menu():
    return render_template("menu.html")


@app.route("/api/estado-ocr", methods=["GET"])
def estado_ocr():
    cadena_respaldo = [f"{ocr.MODELO_VISION_DEEPSEEK} (DeepSeek)"]
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

    # Guarda el PDF/foto TAL CUAL se subió (antes de cualquier realce de
    # calidad) — es la evidencia del documento original, la que se termina
    # subiendo a MySQL cuando se pulsa "Cargar al servidor" (09/09, para
    # poder descargarlo después por numero_factura). Un solo archivo
    # guardado por documento subido, compartido por todas sus páginas.
    ref_documento = guardar_documento_original(contenido, nombre_archivo)

    # Checkbox "Alta resolución" (09/09): para facturas borrosas o mal
    # digitalizadas — sube el DPI de renderizado del PDF (400 en vez de
    # 300) y aplica realce de contraste/nitidez (ver
    # ocr.mejorar_calidad_imagen) tanto al PDF renderizado como a una
    # imagen subida directo. No va prendido por defecto porque el realce
    # tarda un poco más y en un documento ya nítido no aporta nada.
    alta_resolucion = request.form.get("alta_resolucion", "").lower() in ("1", "true", "on", "si")

    if extension == "pdf":
        try:
            dpi = ocr.DPI_ALTA_RESOLUCION if alta_resolucion else 300
            imagenes = ocr.pdf_a_imagenes(contenido, dpi=dpi)
        except Exception as error:
            return jsonify({"error": f"No se pudo leer el PDF: {error}"}), 400
        if not imagenes:
            return jsonify({"error": "El PDF no tiene páginas"}), 400
    else:
        imagenes = [contenido]

    if alta_resolucion:
        imagenes = [ocr.mejorar_calidad_imagen(imagen) for imagen in imagenes]

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


@app.route("/api/calcular_retencion", methods=["POST"])
def calcular_retencion():
    """Botón "Calcular retención" (09/09, a pedido explícito del usuario):
    la extracción (GOT-OCR+puente, Qwen-VL, etc.) solo llena los campos
    tal cual los leyó, sin calcular nada — este endpoint recibe esos
    campos (ya editados/revisados a mano en la pantalla si hizo falta) y
    aplica la fórmula de oro aparte, sin volver a tocar la imagen."""
    payload = request.get_json(silent=True) or {}
    campos = payload.get("campos")
    if not isinstance(campos, dict):
        return jsonify({"error": "Falta 'campos' en el cuerpo de la petición"}), 400
    try:
        campos_calculados = ocr.calcular_retencion_documento(campos)
    except Exception as error:
        return jsonify({"error": f"No se pudo calcular la retención: {error}"}), 500
    return jsonify({"ok": True, "campos": campos_calculados})


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

    # Una fila por cada factura de campos.detalle — un comprobante puede
    # traer varias (mismo criterio que la tabla en pantalla, ver
    # filaResultado()/filaDetalle() en templates/retenciones.html). Antes
    # esto escribía una sola fila por PÁGINA usando campos.nro_factura /
    # campos.monto_retenido (esquema viejo, previo al array "detalle") —
    # con un comprobante de varias facturas, esos campos vienen vacíos o
    # solo reflejan la primera, y el resto se perdía en la exportación
    # aunque la pantalla sí las mostraba todas (reportado en vivo, 08/09).
    encabezados = [
        "Página", "Archivo", "Motor", "Fecha", "Nro Comprobante", "Cliente", "RIF Cliente",
        "Nº Factura", "Nº Control", "Nota Débito", "Nota Crédito", "Tipo Trans", "Doc. Afectado",
        "Total C/IVA", "Total S/IVA", "Base Imponible", "% Alícuota", "Impuesto IVA", "IVA Retenido",
    ]
    hoja.append(encabezados)

    for fila in resultados:
        campos = fila.get("campos", {})
        cabecera = [
            fila.get("pagina"), fila.get("archivo"), fila.get("motor"),
            campos.get("fecha"), campos.get("nro_comprobante"), campos.get("cliente"), campos.get("rif_cliente"),
        ]
        detalle = campos.get("detalle") or [{}]  # sin detalle: una fila vacía, igual que filaDetalle(r, null, ...) en pantalla
        for item in detalle:
            hoja.append(cabecera + [
                item.get("numero_factura"), item.get("numero_control"),
                item.get("numero_nota_debito"), item.get("numero_nota_credito"),
                item.get("tipo_trans"), item.get("documento_afectado"),
                item.get("total_compras_con_iva"), item.get("total_compras_sin_iva"),
                item.get("base_imponible"), item.get("porcentaje_alicuota"),
                item.get("impuesto_iva"), item.get("iva_retenido"),
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
        # Columnas: Fecha de registro, Archivo, Motor, Fecha, Nro Comprobante, Cliente, Nro Factura,
        # RIF Cliente, Monto Retenido, Ref Documento (10ma, agregada 09/09 — filas viejas no la
        # traen, de ahí fila_lista[9:10] en vez de desempaquetado fijo, que rompería con 9 o 10 valores).
        fila_lista = list(fila)
        _fecha_registro, _archivo, motor, fecha, nro_comprobante, cliente, nro_factura, rif_cliente, monto_retenido = fila_lista[:9]
        ref_documento = fila_lista[9] if len(fila_lista) > 9 else None
        campos = {
            "fecha": fecha, "nro_comprobante": nro_comprobante, "cliente": cliente,
            "nro_factura": nro_factura, "rif_cliente": rif_cliente, "monto_retenido": monto_retenido,
        }
        try:
            contenido_pdf = None
            if ref_documento:
                ruta_pdf = os.path.join(RUTA_DOCUMENTOS, ref_documento)
                if os.path.exists(ruta_pdf):
                    with open(ruta_pdf, "rb") as f:
                        contenido_pdf = f.read()
            documento_id = bd_ocr.obtener_o_guardar_documento_retencion(ref_documento, contenido_pdf) if ref_documento else None
            bd_ocr.guardar_retencion(motor, campos, documento_id=documento_id)
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
    from waitress import serve

    puerto = int(os.environ.get("PUERTO_RETENCIONES", 5030))
    print(f"Escáner de retenciones IVA — modelo principal: {ocr.MODELO_VISION_QWEN} (Qwen-VL local), respaldo: {ocr.MODELO_VISION_DEEPSEEK} (DeepSeek) -> {ocr.MODELO_VISION} (Gemini)")
    # Waitress (servidor WSGI de producción) en vez del server de desarrollo
    # de Flask — este panel lo usa más de una persona a la vez (04/09).
    print(f"Sirviendo con Waitress en el puerto {puerto}…")
    serve(app, host="0.0.0.0", port=puerto, threads=14)
