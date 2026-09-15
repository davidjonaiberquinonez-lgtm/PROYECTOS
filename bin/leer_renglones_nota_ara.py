# -*- coding: utf-8 -*-
"""
leer_renglones_nota_ara.py — Trae los renglones REALES (co_art, comentario,
anulado) de una Nota de Entrega desde Profit, vía la API pública de
ARA_PROYECT (GET /api/publico/nota-entrega/comentario, sin `reng_num` —
trae TODOS los renglones de esa nota).

11/09, a pedido explícito del usuario: para armar la proforma propia de
la nota (ver bin/generar_proforma_nota.py) hace falta saber EXACTAMENTE
qué artículos quedaron en esa nota puntual — importante cuando un pedido
se dividió en varias notas (una por almacén, ver
bin/consultar_nota_entrega_ara.py): cada nota solo trae SUS renglones,
no los del pedido completo.

Uso como módulo:
    from bin.leer_renglones_nota_ara import leer_renglones_nota
    leer_renglones_nota("499165")

Uso como CLI:
    python bin/leer_renglones_nota_ara.py --nota 499165
"""
import argparse
import json
import os
import time

import requests

_URL_PRODUCTOS = os.environ.get("ARA_API_PUBLICA_URL", "http://192.168.4.217:4050/api/publico/productos/buscar")
_BASE_URL = _URL_PRODUCTOS.split("/api/publico/", 1)[0]
ARA_API_COMENTARIO_NOTA_URL = f"{_BASE_URL}/api/publico/nota-entrega/comentario"
ARA_API_KEY = os.environ.get("ARA_API_PUBLICA_KEY", "")

CONNECT_TIMEOUT_S = 8
MAX_INTENTOS = 3
ESPERA_REINTENTO_S = 1


def leer_renglones_nota(numero_nota) -> dict:
    """Devuelve {"ok": True, "renglones": [{"reng_num", "co_art", "comentario", "anulado"}, ...]}
    o {"ok": False, "error": "..."}. Nunca lanza excepción hacia quien la
    llama."""
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
                ARA_API_COMENTARIO_NOTA_URL,
                params={"fact_num": fact_num},
                headers={"X-API-Key": ARA_API_KEY},
                timeout=CONNECT_TIMEOUT_S,
            )
            if respuesta.status_code == 401:
                return {"ok": False, "error": "ARA_API_PUBLICA_KEY inválida (401 de la API de ARA_PROYECT)."}
            datos = respuesta.json()
            if not datos.get("ok"):
                return {"ok": False, "error": datos.get("error") or "La API de ARA_PROYECT no encontró la nota."}
            return {"ok": True, "renglones": datos.get("renglones") or []}
        except requests.RequestException as exc:
            if intento < MAX_INTENTOS:
                time.sleep(ESPERA_REINTENTO_S)
                continue
            return {"ok": False, "error": f"No se pudo consultar la API de ARA_PROYECT: {exc}"}
        except Exception as exc:
            return {"ok": False, "error": f"No se pudo consultar la API de ARA_PROYECT: {exc}"}


def _main() -> None:
    parser = argparse.ArgumentParser(description="Trae los renglones reales de una nota de entrega, vía la API pública de ARA_PROYECT.")
    parser.add_argument("--nota", required=True, help="Número de nota de entrega (fact_num de reng_nde).")
    args = parser.parse_args()
    print(json.dumps(leer_renglones_nota(args.nota), ensure_ascii=False, indent=2))


if __name__ == "__main__":
    _main()
