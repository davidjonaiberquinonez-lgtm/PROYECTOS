# -*- coding: utf-8 -*-
"""
consultar_proveedor_profit.py — Adaptador de solo lectura: busca proveedores
reales en Profit (tabla `prov`) por nombre, código o RIF, con filtros.

Pensado para el flujo de este servidor: el usuario escanea una factura de
compra (OCR), y en vez de tipear a mano el proveedor en la tabla de
resultados, busca acá para confirmar contra el maestro real de Profit antes
de guardarlo en la base de datos de retenciones — evita nombres mal
tipeados o proveedores duplicados por variaciones de escritura.

Patrón anti-zombi / anti-loop (mismo criterio que ConnectionWrapper.php y
sql_kill_switch.php en ARA_PROYECT):
  - CONEXIÓN CORTA: abre → ejecuta → cierra siempre en el mismo request.
    Nunca queda una conexión pyodbc viva entre búsquedas.
  - WITH (NOLOCK) en la lectura: no toma locks compartidos ni bloquea al
    ERP mientras alguien está facturando.
  - Reintentos ACOTADOS (MAX_INTENTOS): ante una caída de red/ODBC
    reintenta unas pocas veces con espera fija, nunca en loop indefinido —
    si se agotan los intentos, devuelve un error claro y corta ahí mismo.
  - TOP <limite> siempre aplicado: nunca puede devolver una tabla completa
    sin querer.

Uso como módulo (para wirearlo a un endpoint Flask):
    from bin.consultar_proveedor_profit import buscar_proveedor
    resultado = buscar_proveedor(nombre="isoger")

Uso como CLI (para probarlo suelto, sin levantar ningún servidor):
    python bin/consultar_proveedor_profit.py --nombre "isoger"
    python bin/consultar_proveedor_profit.py --codigo "01"
    python bin/consultar_proveedor_profit.py --rif "J-41118496-9"
"""
import argparse
import json
import os
import time

import pyodbc

# CRISTM25 (producción) a propósito: este adaptador alimenta datos reales
# que terminan guardados en la base de retenciones, no es una consulta
# exploratoria — mismo criterio ya establecido en ARA_PROYECT (el código de
# negocio usa CRISTM25, PRUEB25 queda solo para herramientas genéricas).
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
LIMITE_MAX = 50


def _es_error_reintentable(exc: Exception) -> bool:
    """Solo errores de red/enlace ODBC (clase '08', o timeout HYT00) se
    reintentan — un error de sintaxis o de credenciales NO, porque
    reintentarlo no lo va a arreglar (evita loops inútiles)."""
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
    """RIF sin guiones/espacios (ej. 'J-41118496-9' -> 'J411184969') — en
    Profit el formato guardado no siempre es consistente, así que la
    comparación se hace por los caracteres alfanuméricos puros."""
    return "".join(c for c in valor if c.isalnum()).upper()


def buscar_proveedor(
    nombre: str = "",
    codigo: str = "",
    rif: str = "",
    incluir_inactivos: bool = False,
    limite: int = 20,
) -> dict:
    """Busca proveedores reales en Profit (tabla `prov`, CRISTM25) por
    nombre (coincidencia parcial), código exacto o RIF (comparado sin
    guiones/espacios). Al menos un filtro es obligatorio — nunca trae "todos
    los proveedores" sin querer. Solo lectura.

    Devuelve {"ok": True, "total": N, "proveedores": [...]} o
    {"ok": False, "error": "..."} — nunca lanza excepción hacia quien la
    llama.
    """
    nombre = (nombre or "").strip()
    codigo = (codigo or "").strip()
    rif = (rif or "").strip()
    if not nombre and not codigo and not rif:
        return {"ok": False, "error": "Hace falta al menos un filtro: nombre, codigo o rif."}

    limite = max(1, min(LIMITE_MAX, int(limite or 20)))

    condiciones = []
    parametros: list = []
    if codigo:
        condiciones.append("RTRIM(co_prov) = ?")
        parametros.append(codigo)
    if nombre:
        condiciones.append("prov_des LIKE ?")
        parametros.append(f"%{nombre}%")
    if rif:
        condiciones.append("REPLACE(REPLACE(RTRIM(rif), '-', ''), ' ', '') LIKE ?")
        parametros.append(f"%{_limpiar_rif(rif)}%")

    filtro_inactivos = "" if incluir_inactivos else "AND inactivo = 0"
    sql = (
        f"SELECT TOP {limite} "
        "RTRIM(co_prov) AS co_prov, RTRIM(prov_des) AS nombre, RTRIM(rif) AS rif, "
        "RTRIM(nit) AS nit, RTRIM(telefonos) AS telefonos, RTRIM(email) AS email, "
        # direc1 es tipo `text` (legado) — RTRIM no acepta ese tipo directo,
        # hace falta castearlo a varchar primero (bug real atrapado en vivo).
        "RTRIM(CAST(direc1 AS VARCHAR(2000))) AS direccion, inactivo "
        "FROM prov WITH (NOLOCK) "
        f"WHERE ({' OR '.join(condiciones)}) {filtro_inactivos} "
        "ORDER BY prov_des"
    )

    intento = 0
    while True:
        intento += 1
        try:
            conn = _conectar()
            try:
                cur = conn.cursor()
                cur.execute(sql, parametros)
                columnas = [c[0] for c in cur.description]
                filas = [dict(zip(columnas, fila)) for fila in cur.fetchall()]
            finally:
                conn.close()  # SIEMPRE se cierra — anti-zombi, sin excepción
            for fila in filas:
                fila["inactivo"] = bool(fila["inactivo"])
            return {"ok": True, "total": len(filas), "proveedores": filas}
        except Exception as exc:
            if _es_error_reintentable(exc) and intento < MAX_INTENTOS:
                time.sleep(ESPERA_REINTENTO_S)
                continue
            # Anti-loop: nunca más de MAX_INTENTOS, sin importar el error —
            # siempre termina y devuelve una respuesta, nunca cuelga.
            return {"ok": False, "error": f"No se pudo consultar Profit: {exc}"}


def _main() -> None:
    parser = argparse.ArgumentParser(description="Busca proveedores reales en Profit (tabla prov, CRISTM25).")
    parser.add_argument("--nombre", default="", help="Coincidencia parcial por nombre/razón social.")
    parser.add_argument("--codigo", default="", help="Código exacto de proveedor (co_prov).")
    parser.add_argument("--rif", default="", help="RIF (con o sin guiones/espacios).")
    parser.add_argument("--incluir-inactivos", action="store_true", help="Incluye proveedores marcados inactivos.")
    parser.add_argument("--limite", type=int, default=20)
    args = parser.parse_args()

    resultado = buscar_proveedor(
        nombre=args.nombre,
        codigo=args.codigo,
        rif=args.rif,
        incluir_inactivos=args.incluir_inactivos,
        limite=args.limite,
    )
    print(json.dumps(resultado, ensure_ascii=False, indent=2, default=str))


if __name__ == "__main__":
    _main()
