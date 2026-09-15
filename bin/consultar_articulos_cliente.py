# -*- coding: utf-8 -*-
"""
consultar_articulos_cliente.py — Adaptador de solo lectura: catálogo de
artículos preciado por Cristmedicals para un cliente puntual
(apiweb.cristmedicals.com/api/pagina/articulos?co_cli=...).

ALCANCE REDUCIDO (14/09, a pedido explícito del usuario): esta API se probó
como fuente de BÚSQUEDA de artículos y dio resultados ambiguos/vacíos — no
está pensada para buscar por texto, solo lista el catálogo completo del
cliente. La búsqueda de artículos (relevancia, stock, código, código de
barra) vuelve a ser 100% bin/consultar_producto_ara.py (maestro de
ARA_PROYECT). Lo único que se sigue usando de ESTA API es el "descuento de
categoría" (`porc_descuento`) que trae cada artículo para un cliente —
ver obtener_precio_articulo_cliente() más abajo — que bin/consultar_precio_profit.py
consume como una capa MÁS del cálculo de precio (además del tipo_cli y el
desc_glob, que se resuelven directo contra Profit, no contra esta API).
buscar_articulos_cliente() se deja definida por si hace falta comparar/
depurar algo suelto, pero servidor_ocr_ps.py ya no la llama.

La API no tiene búsqueda de texto ni paginación real: siempre devuelve el
catálogo COMPLETO de ese cliente (miles de artículos, ~2MB, ~2s en frío).
Por eso acá se cachea en memoria por co_cli con un TTL corto — no tiene
sentido bajar de nuevo los mismos ~5000 artículos en cada tecla que el
usuario escribe en la lupa (medido en vivo: reordenar por relevancia esos
~5000 artículos con rapidfuzz toma ~40ms, así que el costo real está en la
descarga, no en el filtrado), pero tampoco hay que dejarlo pegado por horas
si cambia un precio o el stock.

Uso como módulo (para wirearlo a servidor_ocr_ps.py):
    from bin.consultar_articulos_cliente import buscar_articulos_cliente, obtener_precio_articulo_cliente
    resultado = buscar_articulos_cliente(co_cli="FAR00499", texto="ibuprofeno 600mg cj x 10 tab (genven)")
    precio = obtener_precio_articulo_cliente(co_cli="FAR00499", co_art="MD00639")

Uso como CLI (para probarlo suelto, sin levantar ningún servidor):
    python bin/consultar_articulos_cliente.py --co-cli FAR00499 --texto "ibuprofeno 600mg"
"""
import argparse
import json
import os
import time

import requests
from rapidfuzz import fuzz, utils

# Reutiliza los mismos ajustes de relevancia ya probados en vivo para la
# lupa de artículos (medida/dosis, laboratorio entre paréntesis, penalidad
# por combinaciones) — son funciones puras de texto, no dependen de la
# fuente (ARA o Cristmedicals), así que no hace falta duplicarlas acá.
from bin.consultar_producto_ara import (
    _ajuste_por_combinacion,
    _ajuste_por_laboratorio,
    _ajuste_por_medida,
    _extraer_laboratorio,
)

ARTICULOS_CLIENTE_URL = os.environ.get(
    "CRISTMEDICALS_ARTICULOS_URL", "https://apiweb.cristmedicals.com/api/pagina/articulos"
)

CONNECT_TIMEOUT_S = 20  # el catálogo completo pesa ~2MB, más lento que una consulta puntual
MAX_INTENTOS = 2
ESPERA_REINTENTO_S = 1
TTL_CACHE_S = 180  # 3 min: alcanza para todas las búsquedas de un mismo pedido sin quedar pegado a un precio/stock viejo

_cache: dict = {}  # co_cli -> {"expira": epoch, "articulos": [...]}


def _bajar_catalogo(co_cli: str) -> dict:
    """Trae el catálogo COMPLETO (sin filtrar) de un cliente directo de la
    API de Cristmedicals. Nunca lanza excepción hacia quien llama."""
    intento = 0
    while True:
        intento += 1
        try:
            respuesta = requests.get(
                ARTICULOS_CLIENTE_URL, params={"co_cli": co_cli}, timeout=CONNECT_TIMEOUT_S
            )
            respuesta.raise_for_status()
            datos = respuesta.json()
            return {"ok": True, "articulos": datos.get("articulos", [])}
        except requests.RequestException as exc:
            if intento < MAX_INTENTOS:
                time.sleep(ESPERA_REINTENTO_S)
                continue
            return {"ok": False, "error": f"No se pudo consultar el catálogo de precios de Cristmedicals: {exc}"}
        except Exception as exc:
            return {"ok": False, "error": f"No se pudo consultar el catálogo de precios de Cristmedicals: {exc}"}


def _obtener_catalogo(co_cli: str) -> dict:
    ahora = time.time()
    entrada = _cache.get(co_cli)
    if entrada and entrada["expira"] > ahora:
        return {"ok": True, "articulos": entrada["articulos"]}
    resultado = _bajar_catalogo(co_cli)
    if not resultado.get("ok"):
        # Si falla la red pero había un catálogo previo en caché (aunque ya
        # esté vencido), es preferible usar ese antes que dejar al usuario
        # sin poder buscar o subir el pedido por un corte momentáneo.
        if entrada:
            return {"ok": True, "articulos": entrada["articulos"]}
        return resultado
    _cache[co_cli] = {"expira": ahora + TTL_CACHE_S, "articulos": resultado["articulos"]}
    return {"ok": True, "articulos": resultado["articulos"]}


def _normalizar_item(articulo: dict) -> dict:
    stock_total = sum(int(r.get("disponible") or 0) for r in (articulo.get("stock_por_region") or []))
    return {
        "codigo": (articulo.get("co_art") or "").strip(),
        "descripcion": (articulo.get("des_art") or "").strip(),
        "codigo_barra": articulo.get("codigo_barra"),
        "unidad": articulo.get("unidad"),
        "stock_act": stock_total,
        "precio_base": articulo.get("precio_base"),
        "porc_descuento": articulo.get("porc_descuento"),
        "precio": articulo.get("precio_con_descuento"),
    }


def buscar_articulos_cliente(co_cli: str, texto: str = "", limite: int = 15, texto_original: str = "") -> dict:
    """Busca por coincidencia de texto DENTRO del catálogo ya preciado de
    `co_cli` (no en el maestro genérico). El texto es obligatorio. Devuelve
    {"ok": True, "total": N, "productos": [...]} con cada producto trayendo
    ya "precio" (precio_con_descuento real de Cristmedicals) y "coincidencia"
    (0-100), o {"ok": False, "error": "..."} — nunca lanza excepción."""
    co_cli = (co_cli or "").strip()
    texto = (texto or "").strip()
    texto_original = (texto_original or "").strip() or texto
    if not co_cli:
        return {"ok": False, "error": "Hace falta co_cli."}
    if not texto:
        return {"ok": False, "error": "Hace falta el texto a buscar."}

    catalogo = _obtener_catalogo(co_cli)
    if not catalogo.get("ok"):
        return catalogo

    laboratorio = _extraer_laboratorio(texto_original)
    candidatos = []
    for articulo in catalogo["articulos"]:
        descripcion = (articulo.get("des_art") or "").strip()
        codigo_barra = str(articulo.get("codigo_barra") or "").strip()
        # Mismo criterio que consultar_producto_ara.buscar_producto: código
        # de barras exacto es identificación 100% segura, sin pasar por el
        # puntaje de texto.
        if texto.isdigit() and 8 <= len(texto) <= 13 and codigo_barra == texto:
            coincidencia, orden = 100, 999
        else:
            base = fuzz.WRatio(texto_original, descripcion, processor=utils.default_process)
            ajuste = (
                _ajuste_por_medida(texto_original, descripcion)
                + _ajuste_por_combinacion(texto_original, descripcion)
                + _ajuste_por_laboratorio(laboratorio, descripcion)
            )
            puntaje_bruto = round(base) + ajuste
            coincidencia, orden = max(0, min(100, puntaje_bruto)), puntaje_bruto
        if coincidencia < 40:  # descarta ruido total antes de ordenar miles de artículos
            continue
        item = _normalizar_item(articulo)
        item["coincidencia"] = coincidencia
        item["_orden"] = orden
        candidatos.append(item)

    candidatos.sort(key=lambda p: -p["_orden"])
    candidatos = candidatos[: max(1, min(50, int(limite or 15)))]
    for item in candidatos:
        item.pop("_orden", None)
    return {"ok": True, "total": len(candidatos), "productos": candidatos}


def obtener_precio_articulo_cliente(co_cli: str, co_art: str) -> dict:
    """Precio/descuento de UN artículo puntual para `co_cli`, tal cual los
    trae la página de Cristmedicals — SIN el 10% de tipo_cli ni el
    desc_glob (esos se resuelven contra Profit, ver
    bin/consultar_precio_profit.py, que es quien llama a esta función
    para leer solo "porc_descuento", el descuento de categoría de esta
    API). Devuelve {"ok": False, "error": "..."} si el artículo no
    aparece en el catálogo de ese cliente — nunca inventa un precio."""
    co_cli = (co_cli or "").strip()
    co_art = (co_art or "").strip()
    if not co_cli or not co_art:
        return {"ok": False, "error": "Hace falta co_cli y co_art."}

    catalogo = _obtener_catalogo(co_cli)
    if not catalogo.get("ok"):
        return catalogo

    for articulo in catalogo["articulos"]:
        if (articulo.get("co_art") or "").strip().upper() == co_art.upper():
            return {
                "ok": True,
                "co_art": co_art,
                "co_cli": co_cli,
                "precio_base": articulo.get("precio_base"),
                "porc_descuento": articulo.get("porc_descuento"),
                "precio_con_descuento": articulo.get("precio_con_descuento"),
            }
    return {
        "ok": False,
        "error": f"El artículo '{co_art}' no está en el catálogo de precios de Cristmedicals para el cliente '{co_cli}'.",
    }


def _main() -> None:
    parser = argparse.ArgumentParser(description="Busca artículos ya preciados para un cliente puntual (API de Cristmedicals).")
    parser.add_argument("--co-cli", required=True, help="Código exacto del cliente (co_cli).")
    parser.add_argument("--texto", default="", help="Texto a buscar dentro del catálogo del cliente.")
    parser.add_argument("--co-art", default="", help="Si se manda, resuelve precio puntual en vez de buscar por texto.")
    args = parser.parse_args()

    if args.co_art:
        resultado = obtener_precio_articulo_cliente(co_cli=args.co_cli, co_art=args.co_art)
    else:
        resultado = buscar_articulos_cliente(co_cli=args.co_cli, texto=args.texto)
    print(json.dumps(resultado, ensure_ascii=False, indent=2, default=str))


if __name__ == "__main__":
    _main()
