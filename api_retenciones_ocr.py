"""
api_retenciones_ocr.py — API pública de OCR contable, sin interfaz gráfica.

Este servidor es solo la "inteligencia": recibe un PDF o imagen y devuelve
los datos extraídos en JSON. No sirve ninguna página HTML — la interfaz
la construye quien consuma esta API (mismo estilo del adaptador de
consulta de clientes de PROFIT/ARA: endpoints simples, sin login, JSON
plano de entrada y salida).

Cubre TRES tipos de documento, cada uno con su propio endpoint:

  - /api/escanear             Comprobantes de Retención de IVA
                               (inteligencia en ocr_retenciones_core.py)
  - /api/escanear_abono       Recibos de Cobro / abonos de clientes
                               (inteligencia en ocr_abonos_core.py)
  - /api/escanear_factura_ps  Número de factura para la base de datos de
                               Psicotrópicos (inteligencia en ocr_ps_core.py)

Los tres motores son independientes (cada uno con su propio prompt y
campos), pero comparten la misma infraestructura de visión (Gemini +
respaldo OpenRouter) y el mismo manejo de PDF/imagen. También los usan,
cada uno por su lado, servidor_retenciones.py, servidor_abonos.py y
servidor_ocr_ps.py — nuestras interfaces internas, que además guardan un
histórico en Excel; esta API pública no lo hace, es sin estado.

Endpoints:

    GET /api/estado
        Health-check. Responde qué modelo está activo, si hay respaldo
        configurado, y el contrato de cada endpoint de escaneo.

        Respuesta:
        {
          "ok": true,
          "modelo_principal": "gemini-3.6-flash",
          "modelo_respaldo": "dots-studio/dots-3-note-preview:free",
          "formatos_aceptados": ["bmp", "jpeg", "jpg", "pdf", "png", "webp"],
          "endpoints": {
            "retenciones": {
              "endpoint": "/api/escanear",
              "metodo": "POST",
              "campo_archivo": "archivo",
              "campos_extraidos": ["fecha", "nro_comprobante", "cliente", "nro_factura", "rif_cliente", "monto_retenido"]
            },
            "abonos": {
              "endpoint": "/api/escanear_abono",
              "metodo": "POST",
              "campo_archivo": "archivo",
              "campos_extraidos": ["numero_recibo", "fecha", "cliente_nombre", "cliente_codigo_o_rif", "direccion_fiscal", "monto", "cantidad_texto", "concepto", "facturas_referenciadas", "forma_pago"]
            },
            "ocr_ps": {
              "endpoint": "/api/escanear_factura_ps",
              "metodo": "POST",
              "campo_archivo": "archivo",
              "campos_extraidos": ["numero_factura"]
            }
          }
        }

    POST /api/escanear
        Comprobantes de Retención de IVA. Sube un archivo
        (multipart/form-data, campo "archivo": un PDF de una o varias
        páginas, o una imagen suelta) y devuelve TODAS las páginas ya
        procesadas en una sola respuesta (no streaming — más simple de
        consumir desde cualquier lenguaje/herramienta).

        Respuesta:
        {
          "total_paginas": 2,
          "resultados": [
            {
              "pagina": 1,
              "archivo": "comprobante.pdf-pagina1.png",
              "ok": true,
              "motor": "Gemini",
              "campos": {
                "fecha": "03/08/2026",
                "nro_comprobante": "20260800000174",
                "cliente": "FARMACIA YORMEDICAL, C.A.",
                "nro_factura": "00398588",
                "rif_cliente": "J-50069072-0",
                "monto_retenido": "880,58"
              }
            },
            {
              "pagina": 2,
              "archivo": "comprobante.pdf-pagina2.png",
              "ok": false,
              "error": "Gemini falló: ... — OpenRouter también falló: ..."
            }
          ]
        }

    POST /api/escanear_abono
        Recibos de Cobro (abonos de clientes contra facturas). Mismo
        contrato de entrada que /api/escanear (multipart/form-data, campo
        "archivo") y misma forma de respuesta — solo cambian los campos
        dentro de "campos":

        Cada resultado trae además un bloque "confianza" con un índice
        0-100 (70% qué tan reconocible fue el FORMATO del recibo, 30% qué
        tan legible salió lo MANUSCRITO) y "validado": true si ese índice
        llegó a 70 — nuestra propia interfaz (servidor_abonos.py) usa
        exactamente ese umbral para decidir si guarda la fila sola en el
        histórico o si espera a que un humano la revise; el consumidor de
        esta API puede aplicar el mismo criterio (o uno propio) con estos
        mismos números.

        Respuesta:
        {
          "total_paginas": 1,
          "resultados": [
            {
              "pagina": 1,
              "archivo": "recibo058230.jpg",
              "ok": true,
              "motor": "Gemini",
              "campos": {
                "numero_recibo": "058230",
                "fecha": "26/06/2026",
                "cliente_nombre": "Farmacia Animas Benditas 1, C.A",
                "cliente_codigo_o_rif": "FAR01547",
                "direccion_fiscal": "El Abejal de Palmira",
                "monto": "44.245",
                "cantidad_texto": "44$ y 800COP",
                "concepto": "Pago de Facturas #00129005, #00379732 se aplica el 7% por pronto pago y el 19% Descuento Especial",
                "facturas_referenciadas": "00129005, 00379732",
                "forma_pago": "COP, DOLARES"
              },
              "confianza": {
                "formato": 95,
                "datos": 80,
                "indice": 90,
                "validado": true
              }
            }
          ]
        }

    POST /api/escanear_factura_ps
        Número de factura para la base de datos de Psicotrópicos. Mismo
        contrato de entrada (campo "archivo"). "campos" trae un único
        valor, ya como entero:

        Respuesta:
        {
          "total_paginas": 1,
          "resultados": [
            {
              "pagina": 1,
              "archivo": "factura_compra.jpg",
              "ok": true,
              "motor": "Gemini",
              "campos": { "numero_factura": 174857 }
            }
          ]
        }

Sin autenticación (igual que el adaptador de PROFIT) — cualquiera con
acceso a la red donde corre este servicio puede usarlo.

Ejecutar:
    python api_retenciones_ocr.py
"""

import os

from flask import Flask, jsonify, request

import ocr_abonos_core as ocr_abonos
import ocr_ps_core as ocr_ps
import ocr_retenciones_core as ocr_retenciones

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 30 * 1024 * 1024  # 30 MB


@app.after_request
def permitir_cualquier_origen(respuesta):
    # Sin restricción de origen — quien consuma esta API le va a poner su
    # propia interfaz encima, probablemente en otro dominio/puerto.
    respuesta.headers["Access-Control-Allow-Origin"] = "*"
    respuesta.headers["Access-Control-Allow-Methods"] = "GET, POST, OPTIONS"
    respuesta.headers["Access-Control-Allow-Headers"] = "Content-Type"
    return respuesta


@app.route("/api/escanear", methods=["OPTIONS"])
@app.route("/api/escanear_abono", methods=["OPTIONS"])
@app.route("/api/escanear_factura_ps", methods=["OPTIONS"])
def escanear_preflight():
    return "", 204


@app.route("/api/estado", methods=["GET"])
def estado():
    return jsonify({
        "ok": True,
        "modelo_principal": ocr_retenciones.MODELO_VISION,
        "modelo_respaldo": ocr_retenciones.MODELO_VISION_OPENROUTER if ocr_retenciones.clave_openrouter else None,
        "formatos_aceptados": sorted(ocr_retenciones.EXTENSIONES_PERMITIDAS),
        "endpoints": {
            "retenciones": {
                "endpoint": "/api/escanear",
                "metodo": "POST",
                "campo_archivo": "archivo",
                "campos_extraidos": list(ocr_retenciones.CAMPOS_ESPERADOS),
            },
            "abonos": {
                "endpoint": "/api/escanear_abono",
                "metodo": "POST",
                "campo_archivo": "archivo",
                "campos_extraidos": list(ocr_abonos.CAMPOS_ESPERADOS),
            },
            "ocr_ps": {
                "endpoint": "/api/escanear_factura_ps",
                "metodo": "POST",
                "campo_archivo": "archivo",
                "campos_extraidos": list(ocr_ps.CAMPOS_ESPERADOS),
            },
        },
    })


def _escanear_con(motor_ocr):
    archivo = request.files.get("archivo")
    if archivo is None or not archivo.filename:
        return jsonify({"error": "Sube un archivo (campo 'archivo')"}), 400

    extension = archivo.filename.rsplit(".", 1)[-1].lower()
    contenido = archivo.read()

    try:
        resultados = motor_ocr.procesar_documento(contenido, archivo.filename, extension)
    except ValueError as error:
        return jsonify({"error": str(error)}), 400
    except Exception as error:
        return jsonify({"error": f"No se pudo procesar el archivo: {error}"}), 500

    return jsonify({"total_paginas": len(resultados), "resultados": resultados})


@app.route("/api/escanear", methods=["POST"])
def escanear_retenciones():
    return _escanear_con(ocr_retenciones)


@app.route("/api/escanear_abono", methods=["POST"])
def escanear_abono():
    return _escanear_con(ocr_abonos)


@app.route("/api/escanear_factura_ps", methods=["POST"])
def escanear_factura_ps():
    return _escanear_con(ocr_ps)


if __name__ == "__main__":
    puerto = int(os.environ.get("PUERTO_API_RETENCIONES", 5031))
    print(f"API pública de OCR (retenciones + abonos + OCR-PS, sin interfaz) — puerto {puerto}, modelo: {ocr_retenciones.MODELO_VISION}")
    app.run(host="0.0.0.0", port=puerto, debug=True, threaded=True, use_reloader=False)
