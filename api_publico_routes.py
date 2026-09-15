# -*- coding: utf-8 -*-
"""API pública de solo lectura para servicios EXTERNOS a ARA_PROYECT.

Nace 07/09 a pedido del usuario: Servidor de Retenciones (Pedidos OCR,
`C:\\PROYECTOS\\Servidor de Retenciones\\bin\\consultar_producto_ara.py`)
abría el SQLite de ARA_PROYECT (`proyecto_ara.db`) DIRECTO por ruta de
archivo local — funciona en esta máquina, pero se rompe apenas ese servicio
se mueve a un contenedor/otra máquina (el archivo ya no está ahí). Este
endpoint reemplaza esa lectura directa por HTTP, mismo query/misma forma de
respuesta que ya devolvía `buscar_producto()` de ese lado, para que el
cambio del otro lado sea mínimo.

Seguridad: mismo patrón que watchdog_dashboard_server.py — X-API-Key
obligatorio en todo /api/publico/*, sin clave configurada = todo bloqueado
(nunca abierto por omisión). Es un servicio interno de LAN, no expuesto a
internet — token fijo simple, no JWT de sesión (Pedidos OCR no tiene login
de usuario ARA, no tiene sentido pedirle que inicie sesión solo para esto).
"""
import os
import sqlite3
import time

import requests
from flask import Response, abort, jsonify, request

API_KEY = os.environ.get("ARA_API_PUBLICA_KEY", "")

_BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DB_PATH = os.environ.get("ARA_DB_PATH", os.path.join(_BASE_DIR, "data", "proyecto_ara.db"))

CONNECT_TIMEOUT_S = 6
MAX_INTENTOS = 3
ESPERA_REINTENTO_S = 1
LIMITE_MAX = 50

# ── Profit (SQL Server, CRISTM25 producción) — mismo patrón/mismas env vars
# que vigilar_datos.py: conexión CORTA (abre → fetchall → cierra), NUNCA
# persistente, WITH (NOLOCK) en la lectura, pooling desactivado a nivel
# driver. CRISTM25 es intencional (producción), no PRUEB25.
_SQL_DRIVER = os.environ.get("PROFIT_SQL_DRIVER", os.environ.get("PROFIT_DB_DRIVER", "SQL Server"))
_SQL_HOST = os.environ.get("PROFIT_SQL_HOST", os.environ.get("PROFIT_DB_HOST", "192.168.4.20"))
_SQL_PORT = os.environ.get("PROFIT_SQL_PORT", os.environ.get("PROFIT_DB_PORT", "1433"))
_SQL_DB = os.environ.get("PROFIT_SQL_NAME", os.environ.get("PROFIT_DB_NAME", "CRISTM25"))
_SQL_USER = os.environ.get("PROFIT_DB_USER", "profit")
_SQL_PASS = os.environ.get("PROFIT_DB_PASS", "profit")
_SQL_TIMEOUT_S = int(os.environ.get("PROFIT_DB_TIMEOUT_S", "8"))


def _buscar_nota_entrega(fact_num: int, limite: int) -> dict:
    """Notas de entrega (num_doc, tipo_doc='E') asociadas a una FACTURA —
    mismo esquema/misma tabla que ya usa ARA Coder (reng_fac), confirmado
    en vivo (10/09) con la factura 80007967 real.

    OJO (11/09, hallazgo real): `factura.fact_num` y `cotiz_c.fact_num` son
    secuencias INDEPENDIENTES que pueden coincidir en el mismo número sin
    ser el mismo documento — probado en vivo: fact_num=248336 existe como
    factura real de 2025-07-04 Y como cotización real de 2026-09-10, cada
    una con su propio num_doc en reng_fac. Esta función es SOLO para
    facturas reales — para cotizaciones usar `_buscar_nota_de_cotizacion`
    (no comparten mecanismo)."""
    try:
        import pyodbc  # import local, mismo criterio que vigilar_datos.py

        pyodbc.pooling = False
        conn_str = (
            f"DRIVER={{{_SQL_DRIVER}}};SERVER={_SQL_HOST},{_SQL_PORT};DATABASE={_SQL_DB};"
            f"UID={_SQL_USER};PWD={_SQL_PASS}"
        )
        conn = pyodbc.connect(conn_str, timeout=_SQL_TIMEOUT_S)
        try:
            cur = conn.cursor()
            filas = cur.execute(
                """
                SELECT DISTINCT num_doc
                FROM reng_fac WITH (NOLOCK)
                WHERE fact_num = ? AND tipo_doc = 'E'
                ORDER BY num_doc
                """,
                (fact_num,),
            ).fetchall()
        finally:
            conn.close()
        notas = [int(fila[0]) for fila in filas][:limite]
        return {"ok": True, "fact_num": fact_num, "total": len(notas), "notas_entrega": notas}
    except Exception as exc:
        return {"ok": False, "error": f"No se pudo consultar Profit: {exc}"}


# ── MySQL legacy (barquisimeto, 192.168.4.148) — donde vive el concepto
# real de "nota" (gestion.cd_barr / rep_not.cod_nota), separado por completo
# de Profit. Mismas env vars que preparacion/infrastructure/adapters/
# legacy_mysql_sync.py, para no duplicar un segundo esquema de config.
_MYSQL_HOST = os.environ.get("MYSQL_HOST", "192.168.4.148")
_MYSQL_USER = os.environ.get("MYSQL_USER", "root")
_MYSQL_PASSWORD = os.environ.get("MYSQL_PASSWORD", "")
_MYSQL_PORT = int(os.environ.get("MYSQL_PORT", "3306"))
_MYSQL_DB_LEGACY = os.environ.get("MYSQL_DATABASE_LEGACY", "barquisimeto")


def _buscar_nota_de_cotizacion(fact_num: int, limite: int) -> dict:
    """Nota(s) de entrega asociadas a una COTIZACIÓN.

    Mecanismo real (confirmado en vivo 11/09, formula del usuario, no la
    heurística de cliente+fecha que se había usado antes): `reng_nde` es la
    tabla de renglones de nota de entrega — ahí el número de la COTIZACIÓN
    vive en la columna `num_doc`, y el número de la NOTA real es `fact_num`
    de esa misma fila (tipo_doc='T'). Es decir, los roles de fact_num/num_doc
    se INVIERTEN respecto a como se usan en cotiz_c/reng_fac — no es la
    misma convención, probado y confirmado con la cotización 248336 real
    (dio 2 notas reales, 72177325 y 499165, una por cada almacén con el que
    se completó el pedido — exactamente el caso de "se dividió la nota"
    que describió el usuario).

    Se enriquece cada nota con su estado real (preparación/chequeo/embalaje)
    consultando `gestion` en el MySQL legacy por cd_barr = nota — ya no hace
    falta cruzar por cliente+fecha, el número de nota ya viene exacto."""
    import pyodbc

    pyodbc.pooling = False
    conn_str = (
        f"DRIVER={{{_SQL_DRIVER}}};SERVER={_SQL_HOST},{_SQL_PORT};DATABASE={_SQL_DB};"
        f"UID={_SQL_USER};PWD={_SQL_PASS}"
    )
    try:
        conn = pyodbc.connect(conn_str, timeout=_SQL_TIMEOUT_S)
        try:
            cur = conn.cursor()
            filas = cur.execute(
                """
                SELECT DISTINCT fact_num
                FROM reng_nde WITH (NOLOCK)
                WHERE num_doc = ? AND tipo_doc = 'T'
                ORDER BY fact_num
                """,
                (fact_num,),
            ).fetchall()
        finally:
            conn.close()
    except Exception as exc:
        return {"ok": False, "error": f"No se pudo consultar Profit (reng_nde): {exc}"}

    notas_num = [int(f[0]) for f in filas][:limite]
    if not notas_num:
        return {"ok": False, "error": f"No se encontró ninguna nota de entrega para la cotización {fact_num}."}

    notas = [{"cod_nota": n, "estado_preparacion": None, "estado_chequeo": None, "estado_embalaje": None, "cant_items": None} for n in notas_num]
    try:
        import pymysql
        import pymysql.cursors

        conexion = pymysql.connect(
            host=_MYSQL_HOST, user=_MYSQL_USER, password=_MYSQL_PASSWORD, port=_MYSQL_PORT,
            database=_MYSQL_DB_LEGACY, charset="utf8mb4", cursorclass=pymysql.cursors.DictCursor,
            connect_timeout=6,
        )
        try:
            with conexion.cursor() as cur:
                cur.execute(
                    f"""
                    SELECT cd_barr, verifi_pre, verifi_cheq, verifi_emb, cant_items
                    FROM gestion
                    WHERE cd_barr IN ({",".join(["%s"] * len(notas_num))})
                    """,
                    tuple(notas_num),
                )
                estados = {r["cd_barr"]: r for r in cur.fetchall()}
        finally:
            conexion.close()
        for nota in notas:
            e = estados.get(nota["cod_nota"])
            if e:
                nota["estado_preparacion"] = e["verifi_pre"]
                nota["estado_chequeo"] = e["verifi_cheq"]
                nota["estado_embalaje"] = e["verifi_emb"]
                nota["cant_items"] = e["cant_items"]
    except Exception as exc:
        # El estado es un enriquecimiento — si el legacy no responde, se
        # devuelven igual los números de nota reales (lo confirmado en
        # SQL Server), solo sin el detalle de estado.
        for nota in notas:
            nota["error_estado"] = str(exc)

    return {
        "ok": True,
        "fact_num": fact_num,
        "total": len(notas),
        "notas_entrega": notas,
    }


# ── Comentarios de reng_nde (11/09, a pedido del usuario) — primer endpoint
# de esta API que ESCRIBE en Profit real (CRISTM25). A propósito acotado
# SOLO a reng_nde (nunca factura, el usuario lo pidió explícito) — es donde
# JVT PEDIDOS y el resto del sistema ya escriben comentarios reales por
# renglón (confirmado en vivo: "SOLO NOTA DE ENTREGA", "LO RETIRAN EN LA
# DROGUERIA", etc.). No hay una columna dedicada "jvt_id" — es texto libre
# compartido, por eso el modo por defecto es AGREGAR (concatenar), nunca
# reemplazar de un plumazo lo que JVT u otro proceso ya haya escrito.
_LOG_DIR = os.path.join(_BASE_DIR, "..", "..", "logs")
_LOG_COMENTARIOS = os.path.join(_LOG_DIR, "api_publico_comentarios_reng_nde.txt")
COMENTARIO_MAX_LEN = 1000


def _log_escritura_comentario(fact_num: int, reng_num: int, anterior: str, nuevo: str, modo: str) -> None:
    try:
        os.makedirs(_LOG_DIR, exist_ok=True)
        linea = (
            f"{time.strftime('%Y-%m-%d %H:%M:%S')} | fact_num={fact_num} reng_num={reng_num} "
            f"modo={modo} | antes={anterior!r} -> despues={nuevo!r}\n"
        )
        with open(_LOG_COMENTARIOS, "a", encoding="utf-8") as f:
            f.write(linea)
    except Exception:
        pass  # el log nunca debe tumbar la escritura real


def _leer_comentario_nota(fact_num: int, reng_num: int | None) -> dict:
    import pyodbc

    pyodbc.pooling = False
    conn_str = (
        f"DRIVER={{{_SQL_DRIVER}}};SERVER={_SQL_HOST},{_SQL_PORT};DATABASE={_SQL_DB};"
        f"UID={_SQL_USER};PWD={_SQL_PASS}"
    )
    try:
        conn = pyodbc.connect(conn_str, timeout=_SQL_TIMEOUT_S)
        try:
            cur = conn.cursor()
            if reng_num is not None:
                filas = cur.execute(
                    """
                    SELECT reng_num, co_art, CAST(comentario AS VARCHAR(4000)) AS comentario, anulado
                    FROM reng_nde WITH (NOLOCK)
                    WHERE fact_num = ? AND reng_num = ? AND tipo_doc = 'T'
                    """,
                    (fact_num, reng_num),
                ).fetchall()
            else:
                filas = cur.execute(
                    """
                    SELECT reng_num, co_art, CAST(comentario AS VARCHAR(4000)) AS comentario, anulado
                    FROM reng_nde WITH (NOLOCK)
                    WHERE fact_num = ? AND tipo_doc = 'T'
                    ORDER BY reng_num
                    """,
                    (fact_num,),
                ).fetchall()
        finally:
            conn.close()
    except Exception as exc:
        return {"ok": False, "error": f"No se pudo consultar Profit (reng_nde): {exc}"}

    if not filas:
        return {"ok": False, "error": f"No existe la nota {fact_num} (o el renglón {reng_num}) en reng_nde."}

    return {
        "ok": True,
        "fact_num": fact_num,
        "renglones": [
            {
                "reng_num": f.reng_num,
                "co_art": (f.co_art or "").strip(),
                "comentario": f.comentario or "",
                "anulado": bool(f.anulado),
            }
            for f in filas
        ],
    }


def _escribir_comentario_nota(fact_num: int, reng_num: int, comentario_nuevo: str, modo: str) -> dict:
    import pyodbc

    pyodbc.pooling = False
    conn_str = (
        f"DRIVER={{{_SQL_DRIVER}}};SERVER={_SQL_HOST},{_SQL_PORT};DATABASE={_SQL_DB};"
        f"UID={_SQL_USER};PWD={_SQL_PASS}"
    )
    try:
        conn = pyodbc.connect(conn_str, timeout=_SQL_TIMEOUT_S)
        try:
            cur = conn.cursor()
            fila = cur.execute(
                """
                SELECT CAST(comentario AS VARCHAR(4000)) AS comentario, anulado
                FROM reng_nde WITH (NOLOCK)
                WHERE fact_num = ? AND reng_num = ? AND tipo_doc = 'T'
                """,
                (fact_num, reng_num),
            ).fetchone()
            if fila is None:
                return {"ok": False, "error": f"No existe el renglón {reng_num} de la nota {fact_num} en reng_nde."}
            if fila.anulado:
                return {"ok": False, "error": f"El renglón {reng_num} de la nota {fact_num} está ANULADO — no se escribe."}

            anterior = fila.comentario or ""
            if modo == "reemplazar":
                nuevo = comentario_nuevo
            else:
                nuevo = (anterior.rstrip() + "\r\n" + comentario_nuevo) if anterior.strip() else comentario_nuevo

            cur.execute(
                "UPDATE reng_nde SET comentario = ? WHERE fact_num = ? AND reng_num = ? AND tipo_doc = 'T'",
                (nuevo, fact_num, reng_num),
            )
            conn.commit()
        finally:
            conn.close()
    except Exception as exc:
        return {"ok": False, "error": f"No se pudo escribir en Profit (reng_nde): {exc}"}

    _log_escritura_comentario(fact_num, reng_num, anterior, nuevo, modo)
    return {"ok": True, "fact_num": fact_num, "reng_num": reng_num, "comentario_anterior": anterior, "comentario_actual": nuevo}


# ── PDF de factura, a partir de una NOTA (11/09) ────────────────────────────
# Cadena completa confirmada en vivo: nota (reng_nde.fact_num) -> reng_fac
# (num_doc=nota, tipo_doc='E') -> fact_num real de la FACTURA. El servidor
# de PDFs (192.168.4.23:3010) solo entiende número de factura, nunca de
# nota — por eso hace falta resolver la cadena antes de pedir el PDF.
_PDF_SERVER_URL = os.environ.get("FACTURA_PDF_SERVER_URL", "http://192.168.4.23:3010/pdf")
_PDF_TIMEOUT_S = int(os.environ.get("FACTURA_PDF_TIMEOUT_S", "20"))


def _resolver_factura_de_nota(nota_num: int) -> int | None:
    import pyodbc

    pyodbc.pooling = False
    conn_str = (
        f"DRIVER={{{_SQL_DRIVER}}};SERVER={_SQL_HOST},{_SQL_PORT};DATABASE={_SQL_DB};"
        f"UID={_SQL_USER};PWD={_SQL_PASS}"
    )
    conn = pyodbc.connect(conn_str, timeout=_SQL_TIMEOUT_S)
    try:
        cur = conn.cursor()
        fila = cur.execute(
            "SELECT DISTINCT fact_num FROM reng_fac WITH (NOLOCK) WHERE num_doc = ? AND tipo_doc = 'E'",
            (nota_num,),
        ).fetchone()
    finally:
        conn.close()
    return int(fila[0]) if fila else None


def _conectar() -> sqlite3.Connection:
    uri = f"file:{DB_PATH.replace(os.sep, '/')}?mode=ro"
    conexion = sqlite3.connect(uri, uri=True, timeout=CONNECT_TIMEOUT_S)
    conexion.row_factory = sqlite3.Row
    return conexion


def _buscar_producto(texto: str, limite: int) -> dict:
    """Misma query/mismo armado de patrón que consultar_producto_ara.py del
    lado de Servidor de Retenciones — para que migrar ese lado a llamar acá
    en vez de abrir el archivo sea un cambio mínimo (misma forma de salida).

    CORREGIDO 11/09 (a pedido explícito del usuario, error real "que vale
    oro"): `stock_act` es SOLO el inventario de San Cristóbal (Inv/S.C.
    01 en la Consulta de Precios de Profit) — un artículo con stock real
    en Barquisimeto (despacho_bqto/deposito_bqto) daba stock_act=0 y
    Pedidos OCR lo bloqueaba como "sin stock" aunque hubiera unidades de
    sobra (caso real: MD04928, stock_act=0 pero despacho_bqto=111).
    `stock_maestro` (columna, mismo nombre que la tabla) ya es el
    "depósito general" — la suma de TODOS los depósitos (confirmado en
    vivo: despacho+deposito+despacho_bqto+deposito_bqto = stock_maestro
    para MD04928: 0+0+111+0=111) — ahora se devuelve como "stock_act" en
    la respuesta (mismo nombre de campo, para no romper a nadie que ya
    consuma esta API) y se ordena por ese total real. El valor viejo
    (solo San Cristóbal) queda disponible aparte como "stock_sc" por si
    hace falta distinguirlo."""
    palabras = [p for p in texto.split() if p.lower() != "de"]
    if len(palabras) >= 2:
        patron = f"%{palabras[0]}%{palabras[1]}%"
    else:
        patron = f"%{texto}%"

    intento = 0
    while True:
        intento += 1
        try:
            conexion = _conectar()
            try:
                filas = conexion.execute(
                    """
                    SELECT codigo, descripcion, codigo_barra,
                           stock_maestro AS stock_act, stock_act AS stock_sc
                    FROM stock_maestro
                    WHERE codigo IS NOT NULL
                        AND (descripcion LIKE ? COLLATE NOCASE OR codigo LIKE ? COLLATE NOCASE OR codigo_barra LIKE ? COLLATE NOCASE)
                    ORDER BY
                        CASE WHEN codigo_barra = ? THEN 0 ELSE 1 END,
                        CASE WHEN stock_maestro > 0 THEN 0 ELSE 1 END, stock_maestro DESC, descripcion
                    LIMIT ?
                    """,
                    (patron, patron, patron, texto, limite),
                ).fetchall()
            finally:
                conexion.close()
            return {"ok": True, "total": len(filas), "productos": [dict(fila) for fila in filas]}
        except sqlite3.OperationalError as exc:
            if "locked" in str(exc).lower() and intento < MAX_INTENTOS:
                time.sleep(ESPERA_REINTENTO_S)
                continue
            return {"ok": False, "error": f"No se pudo consultar el maestro: {exc}"}
        except Exception as exc:
            return {"ok": False, "error": f"No se pudo consultar el maestro: {exc}"}


def register_api_publico_routes(app):
    @app.before_request
    def _exigir_api_key_publica():
        if not request.path.startswith("/api/publico/"):
            return None
        if not API_KEY:
            abort(503, description="ARA_API_PUBLICA_KEY no está configurada en el servidor.")
        if request.headers.get("X-API-Key") != API_KEY:
            abort(401, description="Falta o es inválido el header X-API-Key.")
        return None

    @app.route("/api/publico/productos/buscar", methods=["GET"])
    def api_publico_buscar_producto():
        texto = (request.args.get("q") or "").strip()
        if not texto:
            return jsonify({"ok": False, "error": "Hace falta el parámetro 'q' (texto a buscar)."}), 400
        if not os.path.exists(DB_PATH):
            return jsonify({"ok": False, "error": f"No se encontró la base en {DB_PATH}"}), 500
        try:
            limite = max(1, min(LIMITE_MAX, int(request.args.get("limite", 15))))
        except (TypeError, ValueError):
            limite = 15
        resultado = _buscar_producto(texto, limite)
        return jsonify(resultado), (200 if resultado.get("ok") else 502)

    @app.route("/api/publico/facturas/nota-entrega", methods=["GET"])
    def api_publico_nota_entrega():
        try:
            fact_num = int(request.args.get("fact_num", ""))
        except (TypeError, ValueError):
            return jsonify({"ok": False, "error": "Hace falta 'fact_num' (entero) — número de FACTURA."}), 400
        try:
            limite = max(1, min(LIMITE_MAX, int(request.args.get("limite", 15))))
        except (TypeError, ValueError):
            limite = 15
        resultado = _buscar_nota_entrega(fact_num, limite)
        return jsonify(resultado), (200 if resultado.get("ok") else 502)

    @app.route("/api/publico/cotizaciones/nota", methods=["GET"])
    def api_publico_nota_de_cotizacion():
        try:
            fact_num = int(request.args.get("fact_num", ""))
        except (TypeError, ValueError):
            return jsonify({"ok": False, "error": "Hace falta 'fact_num' (entero) — número de COTIZACIÓN."}), 400
        try:
            limite = max(1, min(LIMITE_MAX, int(request.args.get("limite", 15))))
        except (TypeError, ValueError):
            limite = 15
        resultado = _buscar_nota_de_cotizacion(fact_num, limite)
        return jsonify(resultado), (200 if resultado.get("ok") else 502)

    @app.route("/api/publico/nota-entrega/comentario", methods=["GET"])
    def api_publico_leer_comentario_nota():
        try:
            fact_num = int(request.args.get("fact_num", ""))
        except (TypeError, ValueError):
            return jsonify({"ok": False, "error": "Hace falta 'fact_num' (entero) — número de NOTA (reng_nde)."}), 400
        reng_num_raw = request.args.get("reng_num")
        reng_num = None
        if reng_num_raw is not None:
            try:
                reng_num = int(reng_num_raw)
            except ValueError:
                return jsonify({"ok": False, "error": "'reng_num' debe ser entero si se pasa."}), 400
        resultado = _leer_comentario_nota(fact_num, reng_num)
        return jsonify(resultado), (200 if resultado.get("ok") else 502)

    @app.route("/api/publico/nota-entrega/comentario", methods=["POST"])
    def api_publico_escribir_comentario_nota():
        datos = request.get_json(silent=True) or {}
        try:
            fact_num = int(datos.get("fact_num"))
        except (TypeError, ValueError):
            return jsonify({"ok": False, "error": "Hace falta 'fact_num' (entero) en el body — número de NOTA (reng_nde)."}), 400
        try:
            reng_num = int(datos.get("reng_num", 1))
        except (TypeError, ValueError):
            return jsonify({"ok": False, "error": "'reng_num' debe ser entero."}), 400
        comentario = (datos.get("comentario") or "").strip()
        if not comentario:
            return jsonify({"ok": False, "error": "Falta 'comentario' (texto a agregar)."}), 400
        if len(comentario) > COMENTARIO_MAX_LEN:
            return jsonify({"ok": False, "error": f"'comentario' supera el máximo de {COMENTARIO_MAX_LEN} caracteres."}), 400
        modo = datos.get("modo", "agregar")
        if modo not in ("agregar", "reemplazar"):
            return jsonify({"ok": False, "error": "'modo' debe ser 'agregar' (default) o 'reemplazar'."}), 400
        resultado = _escribir_comentario_nota(fact_num, reng_num, comentario, modo)
        return jsonify(resultado), (200 if resultado.get("ok") else 502)

    @app.route("/api/publico/nota-entrega/factura-pdf", methods=["GET"])
    def api_publico_factura_pdf():
        try:
            nota_num = int(request.args.get("fact_num", ""))
        except (TypeError, ValueError):
            return jsonify({"ok": False, "error": "Hace falta 'fact_num' (entero) — número de NOTA (reng_nde)."}), 400

        try:
            factura_num = _resolver_factura_de_nota(nota_num)
        except Exception as exc:
            return jsonify({"ok": False, "error": f"No se pudo consultar Profit (reng_fac): {exc}"}), 502
        if factura_num is None:
            return jsonify({"ok": False, "error": f"No se encontró ninguna factura real para la nota {nota_num}."}), 404

        try:
            resp = requests.get(f"{_PDF_SERVER_URL}/{factura_num}", timeout=_PDF_TIMEOUT_S)
        except Exception as exc:
            return jsonify({"ok": False, "error": f"No se pudo contactar el servidor de PDFs ({_PDF_SERVER_URL}): {exc}"}), 502
        if resp.status_code != 200:
            return jsonify({
                "ok": False,
                "error": f"El servidor de PDFs devolvió HTTP {resp.status_code} para la factura {factura_num}.",
                "fact_num_nota": nota_num,
                "fact_num_factura": factura_num,
            }), 502

        return Response(
            resp.content,
            mimetype="application/pdf",
            headers={"Content-Disposition": f'inline; filename="factura_{factura_num}.pdf"'},
        )
