# -*- coding: utf-8 -*-
"""
escribir_comentario_nota_ara.py — Inyecta el comentario que el usuario
escribe en el modal "Montar pedido" de Pedidos OCR directamente en
Profit real (reng_nde.comentario, tipo_doc='T'), contra la API pública
de ARA_PROYECT (ara/ARA_Brain/api_publico_routes.py, endpoint
POST /api/publico/nota-entrega/comentario).

11/09, a pedido explícito del usuario: hasta ahora el campo "Comentario
(opcional)" del modal se leía en el frontend pero el backend nunca lo
usaba (el endpoint de creación de cotización, apiweb.cristmedicals.com,
no acepta ningún campo de observación) — quedaba guardado solo en este
sistema, nunca llegaba a Profit. Este adaptador cierra ese hueco
escribiéndolo en el renglón de la NOTA DE ENTREGA ya resuelta (ver
bin/consultar_nota_entrega_ara.py), no en la cotización.

OJO (documentado también del lado de ARA_PROYECT): `fact_num` acá es el
número de la NOTA (el mismo que devuelve obtener_notas_de_cotizacion,
guardado en pedidos_ocr.numero_entrega), no el número de cotización.
Modo por defecto "agregar" (concatena, nunca pisa lo que JVT PEDIDOS u
otro proceso ya haya escrito en ese renglón).

Uso como módulo:
    from bin.escribir_comentario_nota_ara import escribir_comentario_nota
    escribir_comentario_nota("499165", "Lo retiran en la droguería")

Uso como CLI (para probarlo suelto, sin levantar ningún servidor):
    python bin/escribir_comentario_nota_ara.py --nota 499165 --comentario "..."
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


def escribir_comentario_nota(numero_nota, comentario: str, reng_num: int = 1, modo: str = "agregar") -> dict:
    """Devuelve {"ok": True, "comentario_anterior": ..., "comentario_actual": ...}
    o {"ok": False, "error": "..."}. Nunca lanza excepción hacia quien la
    llama — un fallo acá nunca debe tumbar la subida del pedido, que ya
    quedó hecha del lado de Profit antes de llegar a esto."""
    try:
        fact_num = int(str(numero_nota).strip())
    except (TypeError, ValueError):
        return {"ok": False, "error": f"'{numero_nota}' no es un número de nota válido."}
    comentario = (comentario or "").strip()
    if not comentario:
        return {"ok": False, "error": "Comentario vacío — nada que escribir."}
    if not ARA_API_KEY:
        return {"ok": False, "error": "Falta configurar ARA_API_PUBLICA_KEY en el entorno."}

    intento = 0
    while True:
        intento += 1
        try:
            respuesta = requests.post(
                ARA_API_COMENTARIO_NOTA_URL,
                json={"fact_num": fact_num, "reng_num": reng_num, "comentario": comentario, "modo": modo},
                headers={"X-API-Key": ARA_API_KEY},
                timeout=CONNECT_TIMEOUT_S,
            )
            if respuesta.status_code == 401:
                return {"ok": False, "error": "ARA_API_PUBLICA_KEY inválida (401 de la API de ARA_PROYECT)."}
            datos = respuesta.json()
            if not datos.get("ok"):
                return {"ok": False, "error": datos.get("error") or "La API de ARA_PROYECT no pudo escribir el comentario."}
            return datos
        except requests.RequestException as exc:
            if intento < MAX_INTENTOS:
                time.sleep(ESPERA_REINTENTO_S)
                continue
            return {"ok": False, "error": f"No se pudo consultar la API de ARA_PROYECT: {exc}"}
        except Exception as exc:
            return {"ok": False, "error": f"No se pudo consultar la API de ARA_PROYECT: {exc}"}


def _main() -> None:
    parser = argparse.ArgumentParser(description="Agrega un comentario al renglón de una Nota de Entrega real, vía la API pública de ARA_PROYECT.")
    parser.add_argument("--nota", required=True, help="Número de nota de entrega (fact_num de reng_nde, tipo_doc='T').")
    parser.add_argument("--comentario", required=True, help="Texto a agregar.")
    parser.add_argument("--reng-num", type=int, default=1, help="Renglón a modificar (default 1).")
    parser.add_argument("--modo", choices=("agregar", "reemplazar"), default="agregar")
    args = parser.parse_args()
    print(json.dumps(escribir_comentario_nota(args.nota, args.comentario, args.reng_num, args.modo), ensure_ascii=False, indent=2))


if __name__ == "__main__":
    _main()
