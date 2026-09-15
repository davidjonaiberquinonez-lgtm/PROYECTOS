# -*- coding: utf-8 -*-
"""
consultar_factura_pdf_ara.py — Trae el PDF real de la FACTURA de Profit
asociada a una Nota de Entrega, contra la API pública de ARA_PROYECT
(ara/ARA_Brain/api_publico_routes.py, endpoint
GET /api/publico/nota-entrega/factura-pdf).

11/09, a pedido explícito del usuario: hasta ahora "Ver historial" solo
mostraba el documento ORIGINAL que subió el usuario (la nota/lista de
pedido escaneada). Esto trae, además, el PDF de la FACTURA de verdad que
Profit generó — resolviendo del lado de ARA_PROYECT toda la cadena
nota (reng_nde) -> reng_fac (num_doc=nota, tipo_doc='E') -> factura real
-> servidor de PDFs (192.168.4.23:3010) -> PDF. Acá solo se pasa el
número de NOTA (el mismo que ya se guarda en pedidos_ocr.numero_entrega,
ver bin/consultar_nota_entrega_ara.py) — nunca hay que resolver la
factura del lado de este sistema.

Uso como módulo:
    from bin.consultar_factura_pdf_ara import obtener_factura_pdf_de_nota
    resultado = obtener_factura_pdf_de_nota("499165")
    if resultado["ok"]:
        resultado["pdf_bytes"]  # bytes crudos del PDF

Uso como CLI (para probarlo suelto, sin levantar ningún servidor):
    python bin/consultar_factura_pdf_ara.py --nota 499165 --salida factura.pdf
"""
import argparse
import os
import re
import time

import requests

_URL_PRODUCTOS = os.environ.get("ARA_API_PUBLICA_URL", "http://192.168.4.217:4050/api/publico/productos/buscar")
_BASE_URL = _URL_PRODUCTOS.split("/api/publico/", 1)[0]
ARA_API_FACTURA_PDF_URL = f"{_BASE_URL}/api/publico/nota-entrega/factura-pdf"
ARA_API_KEY = os.environ.get("ARA_API_PUBLICA_KEY", "")

# Más largo que el resto de los adaptadores (8s): esta llamada encadena
# Profit + el servidor de PDFs (192.168.4.23:3010) del otro lado — la
# API de ARA_PROYECT ya le da 20s a ese último salto, así que acá hay
# que esperar al menos eso.
CONNECT_TIMEOUT_S = 25
MAX_INTENTOS = 2
ESPERA_REINTENTO_S = 1


def obtener_factura_pdf_de_nota(numero_nota) -> dict:
    """Devuelve {"ok": True, "pdf_bytes": b"...", "nombre_archivo": "...",
    "numero_factura": "416857"} o {"ok": False, "error": "..."}. Nunca
    lanza excepción hacia quien la llama.

    `numero_factura` (11/09) sale del header Content-Disposition que ya
    manda la API de ARA_PROYECT (`filename="factura_<num>.pdf"`) — el
    número REAL de factura, resuelto del lado de ellos vía reng_fac. No
    hace falta pedirles un endpoint aparte solo para el número: cada
    llamada acá ya lo trae, aunque se use nomás para mostrarlo (sin bajar
    el PDF de nuevo después, ver bd_ocr.actualizar_numero_factura)."""
    try:
        fact_num = int(str(numero_nota).strip())
    except (TypeError, ValueError):
        return {"ok": False, "error": f"'{numero_nota}' no es un número de nota válido."}
    if not ARA_API_KEY:
        return {"ok": False, "error": "Falta configurar ARA_API_PUBLICA_KEY en el entorno."}

    intento = 0
    while True:
        intento += 1
        try:
            respuesta = requests.get(
                ARA_API_FACTURA_PDF_URL,
                params={"fact_num": fact_num},
                headers={"X-API-Key": ARA_API_KEY},
                timeout=CONNECT_TIMEOUT_S,
            )
            if respuesta.status_code == 401:
                return {"ok": False, "error": "ARA_API_PUBLICA_KEY inválida (401 de la API de ARA_PROYECT)."}
            if "application/pdf" in (respuesta.headers.get("Content-Type") or ""):
                coincidencia = re.search(r'filename="?factura_(\d+)\.pdf"?', respuesta.headers.get("Content-Disposition") or "")
                numero_factura = coincidencia.group(1) if coincidencia else None
                return {
                    "ok": True,
                    "pdf_bytes": respuesta.content,
                    "nombre_archivo": f"factura_{numero_factura or fact_num}.pdf",
                    "numero_factura": numero_factura,
                }
            datos = respuesta.json() if respuesta.content else {}
            return {"ok": False, "error": datos.get("error") or f"La API de ARA_PROYECT devolvió HTTP {respuesta.status_code} sin PDF."}
        except requests.RequestException as exc:
            if intento < MAX_INTENTOS:
                time.sleep(ESPERA_REINTENTO_S)
                continue
            return {"ok": False, "error": f"No se pudo consultar la API de ARA_PROYECT: {exc}"}
        except Exception as exc:
            return {"ok": False, "error": f"No se pudo consultar la API de ARA_PROYECT: {exc}"}


def _main() -> None:
    parser = argparse.ArgumentParser(description="Trae el PDF real de la factura de Profit asociada a una nota de entrega, vía la API pública de ARA_PROYECT.")
    parser.add_argument("--nota", required=True, help="Número de nota de entrega (fact_num de reng_nde).")
    parser.add_argument("--salida", default=None, help="Archivo donde guardar el PDF (si no se pasa, solo informa OK/error).")
    args = parser.parse_args()
    resultado = obtener_factura_pdf_de_nota(args.nota)
    if not resultado.get("ok"):
        print(f"Error: {resultado.get('error')}")
        return
    print(f"OK — {len(resultado['pdf_bytes'])} bytes ({resultado['nombre_archivo']})")
    if args.salida:
        with open(args.salida, "wb") as f:
            f.write(resultado["pdf_bytes"])
        print(f"Guardado en {args.salida}")


if __name__ == "__main__":
    _main()
