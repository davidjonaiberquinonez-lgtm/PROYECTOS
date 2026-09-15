# -*- coding: utf-8 -*-
"""
consultar_factura_profit.py — Adaptadores de solo lectura contra Profit
para reforzar el cálculo de retención de IVA en OCR de retenciones (ver
ocr_retenciones_core.py):

  1. obtener_porcentaje_retencion(rif_cliente): el % de retención de IVA que
     aplica un agente de retención NO es siempre 75% — depende de si el
     cliente (el agente de retención, identificado por su RIF) es
     "contribuyente especial" (retiene 100%) o contribuyente ordinario
     (retiene 75%), según la tabla `clientes` de Profit, columna
     `contribu_e` (bit). Verificado en vivo (08/09): 732 clientes reales
     marcados contribu_e=1, 3567 marcados 0 — es un campo real y usado, no
     una columna vacía.

  2. obtener_iva_factura(numero_factura): cuando el modelo de visión no
     pudo leer "base_imponible"/"impuesto_iva" de una fila del comprobante
     (letra chica, tabla mal recortada), pero SÍ pudo leer el número de
     factura, esa factura casi siempre ya está cargada en Profit — se usa
     el IVA que Profit YA calculó (columna `factura.iva`) en vez de
     reconstruirlo a mano desde tot_bruto/exento (probado en vivo: la resta
     tot_bruto - exento no siempre reproduce exacto el iva real de Profit,
     por descuentos/redondeos internos — más seguro usar el valor que
     Profit ya tiene, no reinventarlo).

Patrón anti-zombi (mismo criterio que consultar_proveedor_profit.py):
  - CONEXIÓN CORTA: abre -> consulta -> cierra siempre en el mismo request.
  - WITH (NOLOCK): no bloquea al ERP mientras alguien está facturando.
  - Reintentos ACOTADOS ante error de red/enlace ODBC — nunca en loop
    indefinido.
  - Nunca lanza excepción hacia quien llama — siempre {"ok": ...}.

Uso como módulo:
    from bin.consultar_factura_profit import obtener_porcentaje_retencion, obtener_iva_factura
    obtener_porcentaje_retencion("J-29356328-3")
    obtener_iva_factura("80008108")

Uso como CLI:
    python bin/consultar_factura_profit.py --rif J-29356328-3
    python bin/consultar_factura_profit.py --factura 80008108
"""
import argparse
import json
import os
import time

import pyodbc

PROFIT_HOST = os.environ.get("PROFIT_SQL_HOST", "192.168.4.20")
PROFIT_PORT = os.environ.get("PROFIT_SQL_PORT", "1433")
PROFIT_DB = os.environ.get("PROFIT_SQL_NAME", "CRISTM25")
PROFIT_USER = os.environ.get("PROFIT_SQL_USER", "profit")
PROFIT_PASS = os.environ.get("PROFIT_SQL_PASS", "profit")
PROFIT_ODBC_DRIVER = os.environ.get("PROFIT_ODBC_DRIVER", "SQL Server")

CONNECT_TIMEOUT_S = 6
MAX_INTENTOS = 3
ESPERA_REINTENTO_S = 2

PORCENTAJE_RETENCION_ORDINARIO = 0.75
PORCENTAJE_RETENCION_ESPECIAL = 1.00


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


def _limpiar_rif(valor: str) -> str:
    """RIF sin guiones/espacios/puntos (ej. 'J-29356328-3' -> 'J293563283')
    — en Profit el formato guardado no es consistente (visto en vivo con
    guiones, sin guiones, y hasta con un punto final)."""
    return "".join(c for c in (valor or "") if c.isalnum()).upper()


def obtener_porcentaje_retencion(rif_cliente: str) -> dict:
    """Devuelve {"ok": True, "porcentaje": 0.75 o 1.0, "contribuyente_especial": bool,
    "co_cli": "..."} si encuentra al cliente por RIF, o {"ok": False, "error": "..."}
    si no lo encuentra o el RIF viene vacío — nunca lanza excepción."""
    rif_limpio = _limpiar_rif(rif_cliente)
    if not rif_limpio:
        return {"ok": False, "error": "Hace falta el RIF del cliente."}

    intento = 0
    while True:
        intento += 1
        try:
            conn = _conectar()
            try:
                cur = conn.cursor()
                cur.execute(
                    "SELECT TOP 1 RTRIM(co_cli), contribu_e "
                    "FROM clientes WITH (NOLOCK) "
                    "WHERE REPLACE(REPLACE(REPLACE(RTRIM(rif), '-', ''), ' ', ''), '.', '') = ?",
                    rif_limpio,
                )
                fila = cur.fetchone()
            finally:
                conn.close()
            if fila is None:
                return {"ok": False, "error": f"No se encontró ningún cliente en Profit con RIF '{rif_cliente}'."}
            co_cli, contribu_e = fila
            especial = bool(contribu_e)
            return {
                "ok": True,
                "co_cli": co_cli,
                "contribuyente_especial": especial,
                "porcentaje": PORCENTAJE_RETENCION_ESPECIAL if especial else PORCENTAJE_RETENCION_ORDINARIO,
            }
        except Exception as exc:
            if _es_error_reintentable(exc) and intento < MAX_INTENTOS:
                time.sleep(ESPERA_REINTENTO_S)
                continue
            return {"ok": False, "error": f"No se pudo consultar el % de retención en Profit: {exc}"}


def obtener_iva_factura(numero_factura) -> dict:
    """Devuelve {"ok": True, "iva": <float>, "tot_bruto": <float>, "co_cli": "..."}
    si encuentra la factura en Profit (búsqueda por fact_num exacto), o
    {"ok": False, "error": "..."} si no existe o el número no es válido."""
    try:
        fact_num = int(str(numero_factura).strip())
    except (TypeError, ValueError):
        return {"ok": False, "error": f"'{numero_factura}' no es un número de factura válido."}

    intento = 0
    while True:
        intento += 1
        try:
            conn = _conectar()
            try:
                cur = conn.cursor()
                cur.execute(
                    "SELECT iva, tot_bruto, tot_neto, RTRIM(co_cli) FROM factura WITH (NOLOCK) WHERE fact_num = ?",
                    fact_num,
                )
                fila = cur.fetchone()
            finally:
                conn.close()
            if fila is None:
                return {"ok": False, "error": f"La factura {fact_num} no está cargada en Profit."}
            iva, tot_bruto, tot_neto, co_cli = fila
            # tot_bruto = total ANTES de IVA, tot_neto = total final CON IVA
            # (nombres invertidos respecto a la intuición, pero así están en
            # Profit — verificado en vivo: tot_neto = tot_bruto + iva en la
            # mayoría de las facturas reales, con alguna excepción por
            # descuentos/exentos parciales; se usan los valores tal cual los
            # tiene Profit, nunca se recalculan a mano).
            return {
                "ok": True,
                "iva": float(iva or 0),
                "tot_bruto": float(tot_bruto or 0),
                "tot_neto": float(tot_neto or 0),
                "co_cli": co_cli,
            }
        except Exception as exc:
            if _es_error_reintentable(exc) and intento < MAX_INTENTOS:
                time.sleep(ESPERA_REINTENTO_S)
                continue
            return {"ok": False, "error": f"No se pudo consultar la factura en Profit: {exc}"}


def _main() -> None:
    parser = argparse.ArgumentParser(description="Consultas de solo lectura a Profit para el cálculo de retención de IVA.")
    parser.add_argument("--rif", default="", help="RIF del cliente (agente de retención) para saber su % de retención.")
    parser.add_argument("--factura", default="", help="Número de factura para traer el IVA real ya cargado en Profit.")
    args = parser.parse_args()

    if args.rif:
        print(json.dumps(obtener_porcentaje_retencion(args.rif), ensure_ascii=False, indent=2, default=str))
    if args.factura:
        print(json.dumps(obtener_iva_factura(args.factura), ensure_ascii=False, indent=2, default=str))
    if not args.rif and not args.factura:
        parser.print_help()


if __name__ == "__main__":
    _main()
