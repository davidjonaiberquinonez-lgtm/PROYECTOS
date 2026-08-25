"""
bd_ocr.py — Base de datos MySQL (XAMPP en 192.168.4.23) compartida por los
tres módulos de OCR de este servidor (retenciones, abonos, OCR-PS).

Ojo: 192.168.4.23 YA tiene una base de datos llamada "retenciones" — es un
sistema viejo y totalmente distinto (usuarios, permisos, 6000+ comprobantes
reales, otras columnas) que NO tiene nada que ver con este proyecto. Por
eso todo esto vive en su propia base separada: "ocr_scanner". Nunca hay
que apuntar esto a esa otra base.

Flujo: cada extracción exitosa se sigue guardando primero en el Excel
histórico de su módulo (para que alguien la revise/corrija a mano). Esta
base de datos NO se llena sola en cada extracción — se llena solo cuando
alguien pulsa "Cargar al servidor" en la interfaz, después de revisar el
Excel. Ese botón lee el histórico completo y solo inserta las filas que
todavía no se habían cargado — la tabla `progreso_carga` es la que
recuerda, por módulo, hasta qué fila ya se cargó, para no duplicar si se
vuelve a pulsar el botón.

Cotejamiento de lotes (solo OCR-PS): cada factura puede tener varios
artículos, cada uno con su propio número de lote. La persona asignada
registra esos lotes a mano contra una factura ya cargada (tabla
`lotes_ps`), y después puede buscar por "número de factura + lote" para
ver si esa combinación ya está registrada en nuestra base — eso es el
"cotejamiento". La clave de búsqueda se guarda armada como
"<numero_factura>-<lote>" para que la búsqueda sea directa.

Mismo patrón de conexión que bin/consultar_proveedor_profit.py: conexión
corta (abre, ejecuta, cierra) en cada función — nunca queda una conexión
viva entre llamadas.

Ejecutar directo para crear/actualizar el esquema (crea la base de datos
"ocr_scanner" si no existe):
    python bd_ocr.py
"""

import os

import pymysql
import pymysql.cursors
from dotenv import load_dotenv

load_dotenv()

MYSQL_HOST = os.environ.get("XAMPP_MYSQL_HOST", "192.168.4.23")
MYSQL_PORT = int(os.environ.get("XAMPP_MYSQL_PORT", "3306"))
MYSQL_USER = os.environ.get("XAMPP_MYSQL_USER", "root")
MYSQL_PASSWORD = os.environ.get("XAMPP_MYSQL_PASSWORD", "")
MYSQL_DB = os.environ.get("XAMPP_MYSQL_DB", "ocr_scanner")
CONNECT_TIMEOUT_S = 6


def obtener_conexion(usar_bd=True):
    return pymysql.connect(
        host=MYSQL_HOST, port=MYSQL_PORT, user=MYSQL_USER, password=MYSQL_PASSWORD,
        database=MYSQL_DB if usar_bd else None,
        charset="utf8mb4", cursorclass=pymysql.cursors.DictCursor,
        connect_timeout=CONNECT_TIMEOUT_S,
    )


TABLAS = [
    """
    CREATE TABLE IF NOT EXISTS retenciones (
        id INT AUTO_INCREMENT PRIMARY KEY,
        fecha_registro DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP,
        motor VARCHAR(50),
        fecha VARCHAR(30),
        nro_comprobante VARCHAR(50),
        cliente VARCHAR(200),
        nro_factura VARCHAR(50),
        rif_cliente VARCHAR(20),
        monto_retenido VARCHAR(30),
        INDEX idx_retenciones_fecha_registro (fecha_registro)
    ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4
    """,
    """
    CREATE TABLE IF NOT EXISTS abonos (
        id INT AUTO_INCREMENT PRIMARY KEY,
        fecha_registro DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP,
        motor VARCHAR(50),
        numero_recibo VARCHAR(50),
        fecha VARCHAR(30),
        cliente_nombre VARCHAR(200),
        cliente_codigo_o_rif VARCHAR(50),
        direccion_fiscal VARCHAR(400),
        monto VARCHAR(30),
        cantidad_texto VARCHAR(200),
        concepto VARCHAR(200),
        facturas_referenciadas VARCHAR(200),
        forma_pago VARCHAR(100),
        indice_confianza INT,
        validado_automaticamente TINYINT(1) NOT NULL DEFAULT 0,
        INDEX idx_abonos_fecha_registro (fecha_registro)
    ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4
    """,
    """
    CREATE TABLE IF NOT EXISTS documentos_ps (
        id INT AUTO_INCREMENT PRIMARY KEY,
        fecha_registro DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP,
        nombre_archivo VARCHAR(255),
        contenido LONGBLOB
    ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4
    """,
    """
    CREATE TABLE IF NOT EXISTS ocr_ps (
        id INT AUTO_INCREMENT PRIMARY KEY,
        fecha_registro DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP,
        motor VARCHAR(50),
        numero_factura BIGINT,
        codigo_proveedor VARCHAR(20),
        nombre_proveedor VARCHAR(200),
        rif_proveedor VARCHAR(20),
        documento_id INT,
        INDEX idx_ocr_ps_fecha_registro (fecha_registro),
        INDEX idx_ocr_ps_numero_factura (numero_factura),
        FOREIGN KEY (documento_id) REFERENCES documentos_ps(id)
    ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4
    """,
    """
    CREATE TABLE IF NOT EXISTS lotes_ps (
        id INT AUTO_INCREMENT PRIMARY KEY,
        fecha_registro DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP,
        factura_id INT NOT NULL,
        numero_factura BIGINT NOT NULL,
        lote VARCHAR(100) NOT NULL,
        clave VARCHAR(180) NOT NULL,
        INDEX idx_lotes_ps_clave (clave),
        INDEX idx_lotes_ps_factura_id (factura_id),
        FOREIGN KEY (factura_id) REFERENCES ocr_ps(id)
    ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4
    """,
    """
    CREATE TABLE IF NOT EXISTS progreso_carga (
        modulo VARCHAR(50) PRIMARY KEY,
        filas_cargadas INT NOT NULL DEFAULT 0,
        actualizado_en DATETIME
    ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4
    """,
]


def crear_tablas():
    conexion = obtener_conexion(usar_bd=False)
    try:
        with conexion.cursor() as cursor:
            cursor.execute(
                f"CREATE DATABASE IF NOT EXISTS {MYSQL_DB} CHARACTER SET utf8mb4 COLLATE utf8mb4_unicode_ci"
            )
        conexion.commit()
    finally:
        conexion.close()

    conexion = obtener_conexion()
    try:
        with conexion.cursor() as cursor:
            for sentencia in TABLAS:
                cursor.execute(sentencia)
        conexion.commit()
    finally:
        conexion.close()


def guardar_retencion(motor, campos):
    conexion = obtener_conexion()
    try:
        with conexion.cursor() as cursor:
            cursor.execute(
                """
                INSERT INTO retenciones (motor, fecha, nro_comprobante, cliente, nro_factura, rif_cliente, monto_retenido)
                VALUES (%s, %s, %s, %s, %s, %s, %s)
                """,
                (
                    motor, campos.get("fecha"), campos.get("nro_comprobante"), campos.get("cliente"),
                    campos.get("nro_factura"), campos.get("rif_cliente"), campos.get("monto_retenido"),
                ),
            )
        conexion.commit()
    finally:
        conexion.close()


def guardar_abono(motor, campos, indice_confianza=None, validado=False):
    conexion = obtener_conexion()
    try:
        with conexion.cursor() as cursor:
            cursor.execute(
                """
                INSERT INTO abonos (
                    motor, numero_recibo, fecha, cliente_nombre, cliente_codigo_o_rif, direccion_fiscal,
                    monto, cantidad_texto, concepto, facturas_referenciadas, forma_pago,
                    indice_confianza, validado_automaticamente
                ) VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
                """,
                (
                    motor, campos.get("numero_recibo"), campos.get("fecha"), campos.get("cliente_nombre"),
                    campos.get("cliente_codigo_o_rif"), campos.get("direccion_fiscal"), campos.get("monto"),
                    campos.get("cantidad_texto"), campos.get("concepto"), campos.get("facturas_referenciadas"),
                    campos.get("forma_pago"), indice_confianza, 1 if validado else 0,
                ),
            )
        conexion.commit()
    finally:
        conexion.close()


def guardar_factura_ps(motor, campos, documento_id=None):
    conexion = obtener_conexion()
    try:
        with conexion.cursor() as cursor:
            cursor.execute(
                """
                INSERT INTO ocr_ps (motor, numero_factura, codigo_proveedor, nombre_proveedor, rif_proveedor, documento_id)
                VALUES (%s, %s, %s, %s, %s, %s)
                """,
                (
                    motor, campos.get("numero_factura"), campos.get("codigo_proveedor"),
                    campos.get("nombre_proveedor"), campos.get("rif_proveedor"), documento_id,
                ),
            )
            id_nuevo = cursor.lastrowid
        conexion.commit()
        return id_nuevo
    finally:
        conexion.close()


def obtener_o_guardar_documento_ps(nombre_archivo, contenido_bytes):
    """Un PDF de varias páginas genera varias filas en ocr_ps que todas
    apuntan al mismo documento — para no guardar el mismo archivo
    duplicado una vez por fila (y para que tampoco se duplique si
    'Cargar al servidor' se pulsa en más de una tanda para el mismo
    documento), primero busca si ya existe un documento con ese nombre
    y reusa su id; solo inserta si es la primera vez que se ve.
    Si contenido_bytes es None (no se encontró el archivo en disco) y
    tampoco existe ya guardado, no guarda nada y devuelve None."""
    if not nombre_archivo:
        return None
    conexion = obtener_conexion()
    try:
        with conexion.cursor() as cursor:
            cursor.execute(
                "SELECT id FROM documentos_ps WHERE nombre_archivo = %s ORDER BY id DESC LIMIT 1",
                (nombre_archivo,),
            )
            fila = cursor.fetchone()
            if fila:
                return fila["id"]
            if not contenido_bytes:
                return None
            cursor.execute(
                "INSERT INTO documentos_ps (nombre_archivo, contenido) VALUES (%s, %s)",
                (nombre_archivo, contenido_bytes),
            )
            id_nuevo = cursor.lastrowid
        conexion.commit()
        return id_nuevo
    finally:
        conexion.close()


def obtener_progreso(modulo):
    """Cuántas filas del histórico de ese módulo ya se cargaron al
    servidor — 0 si nunca se cargó nada."""
    conexion = obtener_conexion()
    try:
        with conexion.cursor() as cursor:
            cursor.execute("SELECT filas_cargadas FROM progreso_carga WHERE modulo = %s", (modulo,))
            fila = cursor.fetchone()
        return fila["filas_cargadas"] if fila else 0
    finally:
        conexion.close()


def actualizar_progreso(modulo, filas_cargadas):
    conexion = obtener_conexion()
    try:
        with conexion.cursor() as cursor:
            cursor.execute(
                """
                INSERT INTO progreso_carga (modulo, filas_cargadas, actualizado_en)
                VALUES (%s, %s, NOW())
                ON DUPLICATE KEY UPDATE
                    filas_cargadas = VALUES(filas_cargadas),
                    actualizado_en = VALUES(actualizado_en)
                """,
                (modulo, filas_cargadas),
            )
        conexion.commit()
    finally:
        conexion.close()


def buscar_facturas_ps(numero_factura=None, codigo_proveedor=None, limite=20):
    """Busca facturas YA CARGADAS al servidor (tabla ocr_ps) — es el primer
    paso para registrar un lote: hay que encontrar primero a qué factura
    (de qué proveedor) pertenece antes de anotarle el lote."""
    condiciones = []
    parametros = []
    if numero_factura:
        condiciones.append("numero_factura = %s")
        parametros.append(numero_factura)
    if codigo_proveedor:
        condiciones.append("codigo_proveedor = %s")
        parametros.append(codigo_proveedor)
    if not condiciones:
        return []

    conexion = obtener_conexion()
    try:
        with conexion.cursor() as cursor:
            cursor.execute(
                f"""
                SELECT id, fecha_registro, numero_factura, codigo_proveedor, nombre_proveedor, rif_proveedor
                FROM ocr_ps WHERE {' AND '.join(condiciones)}
                ORDER BY fecha_registro DESC LIMIT %s
                """,
                (*parametros, max(1, min(50, limite))),
            )
            return cursor.fetchall()
    finally:
        conexion.close()


def guardar_lote(factura_id, numero_factura, lote):
    """Registra a mano un lote de un artículo/ítem de una factura ya
    cargada. La clave (numero_factura + '-' + lote) es lo que después se
    busca en cotejar_lote — se guarda ya armada para que esa búsqueda sea
    directa y no dependa de reconstruir el formato cada vez."""
    clave = f"{numero_factura}-{lote}"
    conexion = obtener_conexion()
    try:
        with conexion.cursor() as cursor:
            cursor.execute(
                "INSERT INTO lotes_ps (factura_id, numero_factura, lote, clave) VALUES (%s, %s, %s, %s)",
                (factura_id, numero_factura, lote, clave),
            )
            id_nuevo = cursor.lastrowid
        conexion.commit()
        return id_nuevo
    finally:
        conexion.close()


def listar_lotes_de_factura(factura_id):
    conexion = obtener_conexion()
    try:
        with conexion.cursor() as cursor:
            cursor.execute(
                "SELECT id, lote, clave, fecha_registro FROM lotes_ps WHERE factura_id = %s ORDER BY fecha_registro",
                (factura_id,),
            )
            return cursor.fetchall()
    finally:
        conexion.close()


def cotejar_lote(numero_factura, lote):
    """El cotejamiento: busca si esa combinación factura+lote ya está
    registrada en nuestra base, y si sí, devuelve el contexto completo
    (proveedor, fecha de la factura, etc.) para mostrarlo como resultado.
    Compara contra lo que YA está cargado — no contra un registro externo."""
    clave = f"{numero_factura}-{lote}"
    conexion = obtener_conexion()
    try:
        with conexion.cursor() as cursor:
            cursor.execute(
                """
                SELECT l.id, l.lote, l.clave, l.fecha_registro AS fecha_registro_lote,
                       f.id AS factura_id, f.numero_factura, f.codigo_proveedor, f.nombre_proveedor,
                       f.rif_proveedor, f.fecha_registro AS fecha_registro_factura
                FROM lotes_ps l
                JOIN ocr_ps f ON f.id = l.factura_id
                WHERE l.clave = %s
                """,
                (clave,),
            )
            return cursor.fetchall()
    finally:
        conexion.close()


if __name__ == "__main__":
    crear_tablas()
    print(f"Base de datos OCR creada/actualizada en MySQL ({MYSQL_HOST}:{MYSQL_PORT}/{MYSQL_DB})")
