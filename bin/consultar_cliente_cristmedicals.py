# -*- coding: utf-8 -*-
"""
consultar_cliente_cristmedicals.py — Adaptador de solo lectura: la ficha
real del cliente tal cual la resuelve el propio sistema de Cristmedicals
(apiweb.cristmedicals.com:10443/api/pagina?co_cli=...) — mismo endpoint que
usa CartService::getClienteInfo() (código real de Cristmedicals que mandó
el usuario, 14/09).

Por qué esto en vez de leer `clientes`/`tipo_cli` directo en Profit (ver
bin/consultar_precio_profit.py): esa consulta directa arrastra un problema
real — bin.consultar_precio_profit.obtener_precio() exige poder mapear
`tipo_cli.precio_a` a una de las 5 listas conocidas (PRECIO 1..5); para
clientes cuyo tipo no resuelve ahí (el LEFT JOIN no encuentra fila, o trae
un precio_a que no está contemplado), CADA ítem del pedido fallaba al
resolver precio y el pedido entero quedaba sin poder montarse (14/09,
reportado en vivo: "no dejaba montar pedidos con algunos [clientes] porque
la api anterior pedía campos de cliente estrictos"). Cristmedicals ya
resuelve esto de su lado — este adaptador trae esa misma ficha ya resuelta
(desc_glob, tipo, razón social, RIF, etc.) sin repetir esa lógica frágil acá.

Uso como módulo:
    from bin.consultar_cliente_cristmedicals import obtener_info_cliente, obtener_descuento_y_tipo_cliente
    info = obtener_info_cliente(co_cli="FAR00499")

Uso como CLI (para probarlo suelto, sin levantar ningún servidor):
    python bin/consultar_cliente_cristmedicals.py --co-cli FAR00499
"""
import argparse
import json
import os
import time

import requests

CLIENTE_INFO_URL = os.environ.get(
    "CRISTMEDICALS_CLIENTE_URL", "https://apiweb.cristmedicals.com:10443/api/pagina"
)

CONNECT_TIMEOUT_S = 30  # mismo timeout que usa CartService.php (Http::timeout(30))
MAX_INTENTOS = 2
ESPERA_REINTENTO_S = 1


def obtener_info_cliente(co_cli: str) -> dict:
    """Trae la ficha real del cliente tal cual la ve Cristmedicals: nombre
    (cli_des), tipo, RIF, dirección, teléfono, desc_glob (descuento
    lineal), crédito, zona, etc. Devuelve {"ok": True, "cliente": {...}}
    o {"ok": False, "error": "..."} — nunca lanza excepción."""
    co_cli = (co_cli or "").strip()
    if not co_cli:
        return {"ok": False, "error": "Hace falta co_cli."}

    intento = 0
    while True:
        intento += 1
        try:
            respuesta = requests.get(
                CLIENTE_INFO_URL, params={"co_cli": co_cli}, timeout=CONNECT_TIMEOUT_S
            )
            if respuesta.status_code == 404:
                return {"ok": False, "error": f"Cliente '{co_cli}' no existe en Cristmedicals."}
            respuesta.raise_for_status()
            datos = respuesta.json()
            # El endpoint real devuelve una LISTA de un elemento (confirmado
            # en vivo, 14/09) — mismo criterio que CartService::getClienteInfo
            # ("si la respuesta es un arreglo con un solo elemento...").
            if isinstance(datos, list):
                if not datos:
                    return {"ok": False, "error": f"Cliente '{co_cli}' no existe en Cristmedicals."}
                cliente = datos[0]
            elif isinstance(datos, dict):
                cliente = datos.get("data", datos)
            else:
                return {"ok": False, "error": "Respuesta inesperada de la API de clientes de Cristmedicals."}
            return {"ok": True, "cliente": cliente}
        except requests.RequestException as exc:
            if intento < MAX_INTENTOS:
                time.sleep(ESPERA_REINTENTO_S)
                continue
            return {"ok": False, "error": f"No se pudo consultar la ficha del cliente en Cristmedicals: {exc}"}
        except ValueError as exc:
            return {"ok": False, "error": f"Cristmedicals devolvió una respuesta inválida: {exc}"}


def obtener_descuento_y_tipo_cliente(co_cli: str) -> dict:
    """Recorte de obtener_info_cliente() con lo que hace falta para el
    chip de "descuento lineal" de Pedidos OCR: descuento global y el tipo
    de cliente que usa la API de precios/artículos para resolver su
    lista. Mismo contrato de salida que
    bin.consultar_precio_profit.obtener_descuento_cliente() (mismas
    claves ok/co_cli/descuento_global_pct) — pensado para reemplazarlo
    como fuente principal, ya que ese cálculo directo contra Profit es el
    que fallaba para clientes con un tipo que no resuelve a ninguna de
    las 5 listas conocidas."""
    resultado = obtener_info_cliente(co_cli)
    if not resultado.get("ok"):
        return resultado
    cliente = resultado["cliente"]
    return {
        "ok": True,
        "co_cli": co_cli,
        "razon_social": cliente.get("cli_des"),
        "rif": cliente.get("rif"),
        "tipo": cliente.get("tipo"),
        "descuento_global_pct": float(cliente.get("desc_glob") or 0),
    }


def _main() -> None:
    parser = argparse.ArgumentParser(description="Ficha real del cliente según la API de Cristmedicals.")
    parser.add_argument("--co-cli", required=True, help="Código exacto del cliente (co_cli).")
    args = parser.parse_args()
    print(json.dumps(obtener_info_cliente(co_cli=args.co_cli), ensure_ascii=False, indent=2, default=str))


if __name__ == "__main__":
    _main()
