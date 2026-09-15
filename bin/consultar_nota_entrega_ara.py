# -*- coding: utf-8 -*-
"""
consultar_nota_entrega_ara.py — Adaptador de solo lectura: resuelve el/los
número(s) real(es) de Nota de Entrega de una cotización de Pedidos OCR,
contra la API pública de ARA_PROYECT (ara/ARA_Brain/api_publico_routes.py,
endpoint /api/publico/cotizaciones/nota).

Pensado para Pedidos OCR (servidor_ocr_ps.py): apenas se crea una
cotización real en Profit (ver pedidos_subir), se consulta acá para saber
con qué número(s) de Nota de Entrega quedó — se guarda en
pedidos_ocr.numero_entrega para mostrarlo en "Ver historial".

OJO (11/09, confirmado por el equipo de ARA_PROYECT con datos reales): una
sola cotización puede terminar en MÁS de una nota de entrega — una por
cada almacén con el que Profit completó el pedido (ej. cotización 248336
dio 2 notas reales, 72177325 y 499165). Por eso esta función siempre
devuelve una LISTA, nunca un solo número — mismo criterio que ya se usa
acá mismo para numero_cotizacion cuando un pedido se divide en varias
cotizaciones.

Reusa la misma base URL y clave que ya se usa para el maestro de
productos (ARA_API_PUBLICA_URL / ARA_API_PUBLICA_KEY) — es la misma API,
solo cambia el path.

Uso como módulo:
    from bin.consultar_nota_entrega_ara import obtener_notas_de_cotizacion
    obtener_notas_de_cotizacion("248336")

Uso como CLI (para probarlo suelto, sin levantar ningún servidor):
    python bin/consultar_nota_entrega_ara.py --cotizacion 248336
"""
import argparse
import json
import os
import time

import requests

_URL_PRODUCTOS = os.environ.get("ARA_API_PUBLICA_URL", "http://192.168.4.217:4050/api/publico/productos/buscar")
# La API pública vive en la misma base que el maestro de productos — se
# deriva de ahí en vez de pedir una env var nueva solo para esto.
_BASE_URL = _URL_PRODUCTOS.split("/api/publico/", 1)[0]
ARA_API_NOTA_COTIZACION_URL = f"{_BASE_URL}/api/publico/cotizaciones/nota"
ARA_API_KEY = os.environ.get("ARA_API_PUBLICA_KEY", "")

CONNECT_TIMEOUT_S = 8
MAX_INTENTOS = 3
ESPERA_REINTENTO_S = 1


def obtener_notas_de_cotizacion(numero_cotizacion) -> dict:
    """Devuelve {"ok": True, "notas_entrega": ["72177325", "499165"]}
    (siempre una lista de strings, aunque sea una sola nota) o
    {"ok": False, "error": "..."} si no se encontró o falló la consulta.
    Nunca lanza excepción hacia quien la llama."""
    try:
        fact_num = int(str(numero_cotizacion).strip())
    except (TypeError, ValueError):
        return {"ok": False, "error": f"'{numero_cotizacion}' no es un número de cotización válido."}
    if not ARA_API_KEY:
        return {"ok": False, "error": "Falta configurar ARA_API_PUBLICA_KEY en el entorno."}

    intento = 0
    while True:
        intento += 1
        try:
            respuesta = requests.get(
                ARA_API_NOTA_COTIZACION_URL,
                params={"fact_num": fact_num},
                headers={"X-API-Key": ARA_API_KEY},
                timeout=CONNECT_TIMEOUT_S,
            )
            if respuesta.status_code == 401:
                return {"ok": False, "error": "ARA_API_PUBLICA_KEY inválida (401 de la API de ARA_PROYECT)."}
            datos = respuesta.json()
            if not datos.get("ok"):
                return {"ok": False, "error": datos.get("error") or "La API de ARA_PROYECT no encontró ninguna nota."}
            crudas = datos.get("notas_entrega") or []
            # La API puede devolver una lista de números simples (endpoint
            # de facturas) o de objetos {"cod_nota": ...} (endpoint de
            # cotizaciones, con estado de preparación/chequeo/embalaje) —
            # se tolera cualquiera de los dos formatos.
            numeros = [
                str(nota.get("cod_nota")) if isinstance(nota, dict) else str(nota)
                for nota in crudas
            ]
            return {"ok": True, "notas_entrega": numeros}
        except requests.RequestException as exc:
            if intento < MAX_INTENTOS:
                time.sleep(ESPERA_REINTENTO_S)
                continue
            return {"ok": False, "error": f"No se pudo consultar la API de ARA_PROYECT: {exc}"}
        except Exception as exc:
            return {"ok": False, "error": f"No se pudo consultar la API de ARA_PROYECT: {exc}"}


def _main() -> None:
    parser = argparse.ArgumentParser(description="Resuelve la(s) Nota(s) de Entrega real(es) de una cotización, vía la API pública de ARA_PROYECT.")
    parser.add_argument("--cotizacion", required=True, help="Número de cotización (fact_num de cotiz_c).")
    args = parser.parse_args()
    print(json.dumps(obtener_notas_de_cotizacion(args.cotizacion), ensure_ascii=False, indent=2))


if __name__ == "__main__":
    _main()
