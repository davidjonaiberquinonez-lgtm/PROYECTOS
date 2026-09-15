# -*- coding: utf-8 -*-
"""
consultar_precio_profit.py — Adaptador de solo lectura: resuelve el precio
REAL de un artículo en Profit (tabla `art`, CRISTM25) para un cliente dado.
Es el ÚNICO cálculo de precio del proyecto (14/09, a pedido explícito del
usuario: "no tomemos esa api [Cristmedicals] de búsqueda de cliente ni de
producto, volvamos a la nuestra de ARA" — la búsqueda de artículos es
bin/consultar_producto_ara.py, esto acá es solo el precio). Tres capas,
en orden:

  1. `clientes.tipo` (RTRIM) se cruza contra `tipo_cli.tip_cli` (Profit,
     consulta directa — no la API de Cristmedicals) para saber qué lista
     de precio corresponde (columna `precio_a`, "PRECIO 1".."PRECIO 5" ->
     art.prec_vta1..art.prec_vta5). Verificado en vivo (03/09): cliente
     FAR00017 (tipo '21T') -> tipo_cli.precio_a = 'PRECIO 4' -> prec_vta4.

  2. 10% DE TIPO_CLI (14/09, explicado por el usuario con ejemplo real):
     PRECIO 3 y PRECIO 4 son precios FULL — a esos se les aplica un 10%
     de descuento fijo. PRECIO 1 y PRECIO 2 YA vienen con ese 10%
     incluido en el propio valor de columna (son los "gemelos C/D" de
     tipo_cli — ej. "18TC/D": "...FACTURAS BCV 10% MENOS" — Profit ya
     tiene precargada esa resta en la columna, no hace falta restarla de
     nuevo). Por eso el 10% acá SOLO se aplica si la lista resuelta en el
     paso 1 es PRECIO 3 o PRECIO 4, nunca si ya es PRECIO 1 o PRECIO 2.

  3. Descuento de categoría de la página de Cristmedicals
     (bin/consultar_articulos_cliente.py, campo "porc_descuento" de
     apiweb.cristmedicals.com/api/pagina/articulos) — este SÍ sigue
     viniendo de esa API, es información que no vive en ninguna tabla de
     Profit a la que tengamos acceso. Se aplica DESPUÉS del paso 2, sobre
     el precio ya resuelto ahí. Si esa API falla, no bloquea el precio —
     sigue sin ese último descuento, mejor un precio aproximado que
     ninguno.
  4. `clientes.desc_glob` (descuento lineal del cliente) se aplica al
     final, sobre el resultado de los pasos 1-3.

Verificado en vivo con el ejemplo real que dio el usuario (14/09):
FAR00003 (tipo '18T' -> PRECIO 4 -> $6.61111) -> 10% tipo_cli -> $5.95 ->
5% de categoría de página -> $5.6525 ≈ $5.65, igual a lo que muestra
Profit al facturar.

Sin `co_cli`, no hay forma de saber la lista de precio del cliente — se
devuelve `art.prec_vta1` (precio base/lista 1) como referencia, igual en
"precio" y "precio_con_descuento" (sin descuento aplicado).

Pensado para Pedidos OCR (servidor_ocr_ps.py): antes de este adaptador se
mandaba precio=0 como placeholder explícito.

Patrón anti-zombi / anti-loop (mismo criterio que consultar_proveedor_profit.py):
  - CONEXIÓN CORTA: abre -> consulta -> cierra siempre en el mismo request,
    nunca una conexión pyodbc viva entre llamadas.
  - WITH (NOLOCK) en toda lectura: no bloquea al ERP mientras alguien está
    facturando.
  - Reintentos ACOTADOS (MAX_INTENTOS) solo ante error de red/enlace ODBC
    (clase '08' o timeout 'HYT00') — un error de sintaxis o de credenciales
    no se reintenta porque reintentarlo no lo arregla.
  - Si el cliente o el mapeo de lista de precio no se puede resolver con
    certeza, se devuelve error explícito en vez de adivinar un precio —
    mandar un precio incorrecto a un pedido real es peor que fallar visible.

Uso como módulo (para wirearlo a un endpoint Flask):
    from bin.consultar_precio_profit import obtener_precio
    resultado = obtener_precio(co_art="JBE00888", co_cli="FAR00017")

Uso como CLI (para probarlo suelto, sin levantar ningún servidor — como
módulo, no como script directo, por el import de bin.consultar_articulos_cliente):
    python -m bin.consultar_precio_profit --co-art JBE00888 --co-cli FAR00017
    python -m bin.consultar_precio_profit --co-art JBE00888
"""
import argparse
import json
import os
import time

import pyodbc

from bin.consultar_articulos_cliente import obtener_precio_articulo_cliente

PROFIT_HOST = os.environ.get("PROFIT_SQL_HOST", "192.168.4.20")
PROFIT_PORT = os.environ.get("PROFIT_SQL_PORT", "1433")
PROFIT_DB = os.environ.get("PROFIT_SQL_NAME", "CRISTM25")
PROFIT_USER = os.environ.get("PROFIT_SQL_USER", "profit")
PROFIT_PASS = os.environ.get("PROFIT_SQL_PASS", "profit")
# "SQL Server" es el nombre del driver ODBC integrado de Windows. En Linux
# (contenedores Docker) el driver de Microsoft se registra con otro nombre
# ("ODBC Driver 18 for SQL Server") — configurable acá para no romper el
# despliegue original en Windows, ver PROFIT_ODBC_DRIVER en docker-compose.
PROFIT_ODBC_DRIVER = os.environ.get("PROFIT_ODBC_DRIVER", "SQL Server")

CONNECT_TIMEOUT_S = 6
MAX_INTENTOS = 3
ESPERA_REINTENTO_S = 2

# tipo_cli.precio_a -> columna real de precio en `art`. Whitelist cerrada
# (nunca se arma el nombre de columna a partir de un valor de la base sin
# pasar por acá) — evita inyectar un identificador de columna arbitrario.
MAPA_LISTA_PRECIO = {
    "PRECIO 1": "prec_vta1",
    "PRECIO 2": "prec_vta2",
    "PRECIO 3": "prec_vta3",
    "PRECIO 4": "prec_vta4",
    "PRECIO 5": "prec_vta5",
}

# PRECIO 3 y 4 son listas FULL (sin el 10% de tipo_cli todavía aplicado);
# PRECIO 1 y 2 son las mismas listas pero YA con ese 10% incluido en el
# valor de columna (ver docstring del módulo) — el 10% de abajo solo se
# aplica cuando cae acá.
LISTAS_PRECIO_FULL = {"prec_vta3", "prec_vta4"}
DESCUENTO_TIPO_CLI = 0.10


def _es_error_reintentable(exc: Exception) -> bool:
    args = getattr(exc, "args", None) or []
    codigo = args[0] if args else ""
    if isinstance(codigo, str):
        return codigo.startswith("08") or codigo == "HYT00"
    return False


def _conectar() -> pyodbc.Connection:
    return pyodbc.connect(
        f"DRIVER={{{PROFIT_ODBC_DRIVER}}};SERVER={PROFIT_HOST},{PROFIT_PORT};DATABASE={PROFIT_DB};"
        f"UID={PROFIT_USER};PWD={PROFIT_PASS};Connection Timeout={CONNECT_TIMEOUT_S}",
        timeout=CONNECT_TIMEOUT_S,
    )


def _resolver(co_art: str, co_cli: str) -> dict:
    """Hace las consultas reales dentro de una única conexión corta. Separado
    de obtener_precio() solo para que el reintento de más arriba pueda volver
    a abrir una conexión limpia ante un error de red, en vez de reintentar
    sobre una conexión que puede haber quedado en mal estado."""
    conn = _conectar()
    try:
        cur = conn.cursor()
        cur.execute(
            "SELECT prec_vta1, prec_vta2, prec_vta3, prec_vta4, prec_vta5, anulado "
            "FROM art WITH (NOLOCK) WHERE RTRIM(co_art) = ?",
            co_art,
        )
        fila_art = cur.fetchone()
        if fila_art is None:
            return {"ok": False, "error": f"Artículo '{co_art}' no existe en Profit (tabla art)."}
        precios_lista = {
            "prec_vta1": float(fila_art[0]),
            "prec_vta2": float(fila_art[1]),
            "prec_vta3": float(fila_art[2]),
            "prec_vta4": float(fila_art[3]),
            "prec_vta5": float(fila_art[4]),
        }

        if not co_cli:
            precio_base = precios_lista["prec_vta1"]
            return {
                "ok": True,
                "co_art": co_art,
                "co_cli": None,
                "precio": precio_base,
                "precio_con_descuento": precio_base,
                "nota": "Sin co_cli: se devuelve PRECIO 1 (lista base), sin descuento de cliente.",
            }

        cur.execute(
            "SELECT RTRIM(c.tipo), c.desc_glob, RTRIM(t.precio_a) "
            "FROM clientes c WITH (NOLOCK) "
            "LEFT JOIN tipo_cli t WITH (NOLOCK) ON RTRIM(t.tip_cli) = RTRIM(c.tipo) "
            "WHERE RTRIM(c.co_cli) = ?",
            co_cli,
        )
        fila_cli = cur.fetchone()
        if fila_cli is None:
            return {"ok": False, "error": f"Cliente '{co_cli}' no existe en Profit (tabla clientes)."}

        tipo, desc_glob, precio_a = fila_cli
        columna_precio = MAPA_LISTA_PRECIO.get((precio_a or "").strip().upper())
        if not columna_precio:
            return {
                "ok": False,
                "error": (
                    f"El tipo de cliente '{tipo}' de '{co_cli}' mapea a precio_a='{precio_a}', "
                    "que no es una de las 5 listas conocidas (PRECIO 1..5) — no se puede "
                    "resolver el precio con certeza."
                ),
            }

        precio_lista = precios_lista[columna_precio]

        # 2. 10% de tipo_cli — SOLO si la lista es FULL (PRECIO 3/4); si ya
        # es PRECIO 1/2 el 10% ya viene incluido en la columna (ver
        # docstring del módulo).
        if columna_precio in LISTAS_PRECIO_FULL:
            precio_lista = precio_lista * (1 - DESCUENTO_TIPO_CLI)

        # 3. Descuento de categoría de la página de Cristmedicals — capa
        # aparte, no vive en ninguna tabla de Profit a la que tengamos
        # acceso. Si esta llamada falla, no bloquea el precio: sigue sin
        # ese último descuento en vez de fallar todo el cálculo.
        porc_categoria = 0.0
        info_pagina = obtener_precio_articulo_cliente(co_cli=co_cli, co_art=co_art)
        if info_pagina.get("ok") and info_pagina.get("porc_descuento") is not None:
            porc_categoria = float(info_pagina["porc_descuento"])
        precio_con_categoria = precio_lista * (1 - porc_categoria / 100)

        # 4. desc_glob (descuento lineal del cliente), al final.
        desc_glob = float(desc_glob or 0)
        precio_con_descuento = precio_con_categoria * (1 - desc_glob / 100)
        return {
            "ok": True,
            "co_art": co_art,
            "co_cli": co_cli,
            "lista_precio": columna_precio,
            "descuento_tipo_cli_pct": DESCUENTO_TIPO_CLI * 100 if columna_precio in LISTAS_PRECIO_FULL else 0.0,
            "descuento_categoria_pct": porc_categoria,
            "descuento_global_pct": desc_glob,
            "precio": precios_lista[columna_precio],
            "precio_con_descuento": round(precio_con_descuento, 2),
        }
    finally:
        conn.close()  # SIEMPRE se cierra — anti-zombi, sin excepción


def obtener_precio(co_art: str, co_cli: str = None) -> dict:
    """Resuelve el precio real de `co_art` en Profit para `co_cli` (opcional).

    Devuelve {"ok": True, "co_art":.., "precio": <float>,
    "precio_con_descuento": <float, igual a precio si no hay co_cli>} o
    {"ok": False, "error": "..."} — nunca lanza excepción hacia quien la
    llama.
    """
    co_art = (co_art or "").strip()
    co_cli = (co_cli or "").strip()
    if not co_art:
        return {"ok": False, "error": "Hace falta co_art."}

    intento = 0
    while True:
        intento += 1
        try:
            return _resolver(co_art, co_cli)
        except Exception as exc:
            if _es_error_reintentable(exc) and intento < MAX_INTENTOS:
                time.sleep(ESPERA_REINTENTO_S)
                continue
            # Anti-loop: nunca más de MAX_INTENTOS, sin importar el error —
            # siempre termina y devuelve una respuesta, nunca cuelga.
            return {"ok": False, "error": f"No se pudo consultar el precio en Profit: {exc}"}


def _resolver_descuento_cliente(co_cli: str) -> dict:
    conn = _conectar()
    try:
        cur = conn.cursor()
        cur.execute(
            "SELECT RTRIM(c.tipo), c.desc_glob, RTRIM(t.precio_a) "
            "FROM clientes c WITH (NOLOCK) "
            "LEFT JOIN tipo_cli t WITH (NOLOCK) ON RTRIM(t.tip_cli) = RTRIM(c.tipo) "
            "WHERE RTRIM(c.co_cli) = ?",
            co_cli,
        )
        fila = cur.fetchone()
        if fila is None:
            return {"ok": False, "error": f"Cliente '{co_cli}' no existe en Profit (tabla clientes)."}
        tipo, desc_glob, precio_a = fila
        return {
            "ok": True,
            "co_cli": co_cli,
            "descuento_global_pct": float(desc_glob or 0),
            "lista_precio": (precio_a or "").strip() or None,
        }
    finally:
        conn.close()


def obtener_descuento_cliente(co_cli: str) -> dict:
    """Devuelve SOLO el descuento global ("descuento lineal") de un
    cliente en Profit (clientes.desc_glob) — sin necesitar saber ningún
    artículo primero, a diferencia de obtener_precio(). Pensado para
    mostrarlo apenas se elige el cliente en Pedidos OCR (11/09, a pedido
    explícito del usuario: "que lo muestre, hasta que yo tenga
    autorización" — de momento es solo informativo, no se deja editar
    acá). Mismo patrón anti-zombi que obtener_precio: nunca lanza
    excepción, reintenta solo errores de red/enlace acotado."""
    co_cli = (co_cli or "").strip()
    if not co_cli:
        return {"ok": False, "error": "Hace falta co_cli."}

    intento = 0
    while True:
        intento += 1
        try:
            return _resolver_descuento_cliente(co_cli)
        except Exception as exc:
            if _es_error_reintentable(exc) and intento < MAX_INTENTOS:
                time.sleep(ESPERA_REINTENTO_S)
                continue
            return {"ok": False, "error": f"No se pudo consultar el descuento en Profit: {exc}"}


def _main() -> None:
    parser = argparse.ArgumentParser(description="Resuelve el precio real de un artículo en Profit (tabla art, CRISTM25).")
    parser.add_argument("--co-art", required=True, help="Código exacto del artículo (co_art).")
    parser.add_argument("--co-cli", default="", help="Código exacto del cliente (co_cli) — opcional.")
    args = parser.parse_args()

    resultado = obtener_precio(co_art=args.co_art, co_cli=args.co_cli)
    print(json.dumps(resultado, ensure_ascii=False, indent=2, default=str))


if __name__ == "__main__":
    _main()
