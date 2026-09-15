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

import json
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
    CREATE TABLE IF NOT EXISTS documentos_retenciones (
        id INT AUTO_INCREMENT PRIMARY KEY,
        fecha_registro DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP,
        nombre_archivo VARCHAR(255),
        contenido LONGBLOB
    ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4
    """,
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
        documento_id INT,
        INDEX idx_retenciones_fecha_registro (fecha_registro),
        INDEX idx_retenciones_nro_factura (nro_factura),
        FOREIGN KEY (documento_id) REFERENCES documentos_retenciones(id)
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
    """
    CREATE TABLE IF NOT EXISTS pedidos_ocr (
        id INT AUTO_INCREMENT PRIMARY KEY,
        fecha_registro DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP,
        codigo_pedido VARCHAR(50) NOT NULL UNIQUE,
        cod_cliente VARCHAR(20) NOT NULL,
        cliente_nombre VARCHAR(200),
        sede VARCHAR(20),
        motor VARCHAR(50),
        items_json JSON,
        estado VARCHAR(20) NOT NULL DEFAULT 'pendiente',
        numero_cotizacion VARCHAR(100),
        numero_entrega VARCHAR(100),
        numero_factura VARCHAR(100),
        documento_id INT,
        respuesta_endpoint TEXT,
        subido_por VARCHAR(200),
        INDEX idx_pedidos_ocr_fecha_registro (fecha_registro)
    ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4
    """,
    """
    CREATE TABLE IF NOT EXISTS pedidos_en_espera (
        id INT AUTO_INCREMENT PRIMARY KEY,
        fecha_registro DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP,
        empleado_id VARCHAR(50) NOT NULL,
        empleado_nombre VARCHAR(200),
        etiqueta VARCHAR(200),
        archivo_origen VARCHAR(255),
        motor VARCHAR(50),
        campos_json JSON NOT NULL,
        cliente_json JSON,
        INDEX idx_pedidos_en_espera_empleado (empleado_id),
        INDEX idx_pedidos_en_espera_fecha_registro (fecha_registro)
    ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4
    """,
    """
    CREATE TABLE IF NOT EXISTS alertas_stock (
        id INT AUTO_INCREMENT PRIMARY KEY,
        fecha_registro DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP,
        cod_cliente VARCHAR(20) NOT NULL,
        cliente_nombre VARCHAR(200),
        codigo_articulo VARCHAR(20) NOT NULL,
        descripcion_articulo VARCHAR(300),
        solicitado_por VARCHAR(200),
        INDEX idx_alertas_stock_fecha_registro (fecha_registro),
        INDEX idx_alertas_stock_codigo_articulo (codigo_articulo)
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

    _migrar_columnas_pedidos_ocr()
    _migrar_columnas_retenciones()


def _migrar_columnas_retenciones():
    """documento_id se agregó el 09/09 (endpoint de descarga del PDF de una
    retención por numero_factura) — una base creada antes de esa fecha
    tiene la tabla retenciones sin esta columna. Se agrega acá si todavía
    no está, sin tocar filas existentes (quedan NULL en lo viejo — esas
    retenciones ya cargadas nunca tuvieron el PDF guardado, así que no hay
    nada que vincular retroactivamente)."""
    conexion = obtener_conexion()
    try:
        with conexion.cursor() as cursor:
            cursor.execute(
                "SELECT COLUMN_NAME FROM INFORMATION_SCHEMA.COLUMNS "
                "WHERE TABLE_SCHEMA = %s AND TABLE_NAME = 'retenciones'",
                (MYSQL_DB,),
            )
            existentes = {fila["COLUMN_NAME"] for fila in cursor.fetchall()}
            if "documento_id" not in existentes:
                cursor.execute("ALTER TABLE retenciones ADD COLUMN documento_id INT")
        conexion.commit()
    finally:
        conexion.close()


def _migrar_columnas_pedidos_ocr():
    """cliente_nombre y numero_cotizacion se agregaron el 04/09, subido_por
    el 07/09 (usuario activo del SSO del ERP) — una base ya creada antes de
    esas fechas tiene la tabla pedidos_ocr sin esas columnas. Se agregan acá
    si todavía no están, sin tocar filas existentes (quedan NULL en lo viejo).

    10/09: numero_cotizacion pasó de INT a VARCHAR — el endpoint externo
    de creación de pedidos devuelve una LISTA de números (resultado.fact_nums),
    no uno solo: un pedido de más de ~21 artículos vuelve partido en dos
    cotizaciones. Ahora se guardan todos juntos, separados por coma, así
    que la columna necesita poder guardar texto, no solo un entero."""
    columnas_nuevas = {
        "cliente_nombre": "VARCHAR(200)", "numero_cotizacion": "VARCHAR(100)", "subido_por": "VARCHAR(200)",
        "documento_id": "INT", "numero_entrega": "VARCHAR(100)", "numero_factura": "VARCHAR(100)",
    }
    conexion = obtener_conexion()
    try:
        with conexion.cursor() as cursor:
            cursor.execute(
                "SELECT COLUMN_NAME, DATA_TYPE FROM INFORMATION_SCHEMA.COLUMNS "
                "WHERE TABLE_SCHEMA = %s AND TABLE_NAME = 'pedidos_ocr'",
                (MYSQL_DB,),
            )
            columnas_existentes = {fila["COLUMN_NAME"]: fila["DATA_TYPE"] for fila in cursor.fetchall()}
            for columna, tipo in columnas_nuevas.items():
                if columna not in columnas_existentes:
                    cursor.execute(f"ALTER TABLE pedidos_ocr ADD COLUMN {columna} {tipo}")
            if columnas_existentes.get("numero_cotizacion") == "int":
                cursor.execute("ALTER TABLE pedidos_ocr MODIFY COLUMN numero_cotizacion VARCHAR(100)")
        conexion.commit()
    finally:
        conexion.close()


def guardar_retencion(motor, campos, documento_id=None):
    conexion = obtener_conexion()
    try:
        with conexion.cursor() as cursor:
            cursor.execute(
                """
                INSERT INTO retenciones (motor, fecha, nro_comprobante, cliente, nro_factura, rif_cliente, monto_retenido, documento_id)
                VALUES (%s, %s, %s, %s, %s, %s, %s, %s)
                """,
                (
                    motor, campos.get("fecha"), campos.get("nro_comprobante"), campos.get("cliente"),
                    campos.get("nro_factura"), campos.get("rif_cliente"), campos.get("monto_retenido"),
                    documento_id,
                ),
            )
        conexion.commit()
    finally:
        conexion.close()


def obtener_o_guardar_documento_retencion(nombre_archivo, contenido_bytes):
    """Mismo patrón que obtener_o_guardar_documento_ps: un PDF de varias
    páginas genera varias filas en retenciones que apuntan al mismo
    documento — se reusa el id si ese nombre de archivo ya se guardó
    antes, para no duplicarlo ni al recorrer varias páginas ni si "Cargar
    al servidor" se pulsa más de una vez para el mismo documento."""
    if not nombre_archivo:
        return None
    conexion = obtener_conexion()
    try:
        with conexion.cursor() as cursor:
            cursor.execute(
                "SELECT id FROM documentos_retenciones WHERE nombre_archivo = %s ORDER BY id DESC LIMIT 1",
                (nombre_archivo,),
            )
            fila = cursor.fetchone()
            if fila:
                return fila["id"]
            if not contenido_bytes:
                return None
            cursor.execute(
                "INSERT INTO documentos_retenciones (nombre_archivo, contenido) VALUES (%s, %s)",
                (nombre_archivo, contenido_bytes),
            )
            id_nuevo = cursor.lastrowid
        conexion.commit()
        return id_nuevo
    finally:
        conexion.close()


def obtener_documento_retencion_por_factura(numero_factura):
    """Endpoint GET /api/facturas/<numero_factura>/documento de la API de
    retenciones (09/09): el PDF real del comprobante, por número de
    factura. Coincidencia EXACTA (nro_factura = ese número tal cual está
    guardado) — a pedido explícito del usuario, NO busca dentro de una
    lista de varios números separados por coma en el mismo comprobante.
    Devuelve {'nombre_archivo':..., 'contenido':...} o None si no hay
    ningún documento vinculado a ese número exacto."""
    conexion = obtener_conexion()
    try:
        with conexion.cursor() as cursor:
            cursor.execute(
                """
                SELECT d.nombre_archivo, d.contenido
                FROM retenciones r
                JOIN documentos_retenciones d ON d.id = r.documento_id
                WHERE r.nro_factura = %s AND r.documento_id IS NOT NULL
                ORDER BY r.id DESC
                LIMIT 1
                """,
                (str(numero_factura),),
            )
            return cursor.fetchone()
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


def obtener_documento_ps_por_factura(numero_factura, codigo_proveedor):
    """Endpoint GET /api/facturas/<numero_factura>/<cod_prov>/documento
    (10/09, corregido a pedido explícito del usuario): el PDF real de una
    factura de psicotrópicos. numero_factura SOLO no alcanza para
    identificarla — distintos proveedores pueden compartir el mismo
    número de control/factura (reportado en vivo), así que ahora también
    exige el código de proveedor exacto para desambiguar. Si el mismo par
    número+proveedor tiene más de una fila en ocr_ps, se usa la
    vinculación más reciente. Devuelve {'nombre_archivo':..., 'contenido':...}
    (contenido = bytes del PDF) o None si no hay ningún documento
    vinculado a esa combinación exacta."""
    conexion = obtener_conexion()
    try:
        with conexion.cursor() as cursor:
            cursor.execute(
                """
                SELECT d.nombre_archivo, d.contenido
                FROM ocr_ps o
                JOIN documentos_ps d ON d.id = o.documento_id
                WHERE o.numero_factura = %s AND RTRIM(o.codigo_proveedor) = %s AND o.documento_id IS NOT NULL
                ORDER BY o.id DESC
                LIMIT 1
                """,
                (numero_factura, codigo_proveedor),
            )
            return cursor.fetchone()
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


def guardar_pedido_ocr(cod_cliente, sede, motor, items, estado, respuesta_endpoint=None, cliente_nombre=None, subido_por=None, documento_id=None):
    """Pedidos OCR es distinto a los otros tres módulos: NO pasa primero
    por un Excel histórico — se guarda acá directo, y solo en el momento en
    que el usuario aprieta "Subir" (después de revisar/corregir la tarjeta
    y elegir el cliente), nunca apenas termina el escaneo. Así el registro
    que queda siempre refleja lo que el usuario confirmó, no el borrador
    crudo del OCR.

    El "cod_pedido" que exige el endpoint real de creación de pedidos
    (apiweb.cristmedicals.com/api/pedidos/pedido-profit) tiene que ser un
    ENTERO — en la app real es el id autoincremental de su propia tabla
    Pedido, a la que no tenemos acceso. Acá se usa el id autoincremental de
    ESTA fila como equivalente: se inserta primero (para obtenerlo) y
    "codigo_pedido" queda igual a ese id — llamar a pedidos_subir() con el
    id devuelto, y actualizar_estado_pedido_ocr() después de la respuesta
    del endpoint real."""
    conexion = obtener_conexion()
    try:
        with conexion.cursor() as cursor:
            cursor.execute(
                """
                INSERT INTO pedidos_ocr (codigo_pedido, cod_cliente, cliente_nombre, sede, motor, items_json, estado, respuesta_endpoint, subido_por, documento_id)
                VALUES ('', %s, %s, %s, %s, %s, %s, %s, %s, %s)
                """,
                (cod_cliente, cliente_nombre, sede, motor, json.dumps(items, ensure_ascii=False), estado, respuesta_endpoint, subido_por, documento_id),
            )
            id_nuevo = cursor.lastrowid
            cursor.execute("UPDATE pedidos_ocr SET codigo_pedido = %s WHERE id = %s", (str(id_nuevo), id_nuevo))
        conexion.commit()
        return id_nuevo
    finally:
        conexion.close()


def guardar_pedido_en_espera(empleado_id, empleado_nombre, campos, cliente=None, etiqueta=None, archivo_origen=None, motor=None):
    """"Dejar en espera" (10/09, a pedido explícito del usuario): guarda el
    estado COMPLETO de una tarjeta de Pedidos OCR (sede+items en
    `campos`, cliente elegido si ya había uno) para retomarla después sin
    perder lo hecho — caso típico: el cliente dice que quiere agregar algo
    más pero que esperen. `empleado_id` es el dueño (viene de la sesión
    SSO, nunca del cliente) — es la clave que después decide quién puede
    ver esta fila (ver listar_pedidos_en_espera)."""
    conexion = obtener_conexion()
    try:
        with conexion.cursor() as cursor:
            cursor.execute(
                """
                INSERT INTO pedidos_en_espera
                    (empleado_id, empleado_nombre, etiqueta, archivo_origen, motor, campos_json, cliente_json)
                VALUES (%s, %s, %s, %s, %s, %s, %s)
                """,
                (
                    empleado_id, empleado_nombre, etiqueta, archivo_origen, motor,
                    json.dumps(campos, ensure_ascii=False),
                    json.dumps(cliente, ensure_ascii=False) if cliente else None,
                ),
            )
            id_nuevo = cursor.lastrowid
        conexion.commit()
        return id_nuevo
    finally:
        conexion.close()


def listar_pedidos_en_espera(empleado_id=None):
    """Sin `empleado_id`: TODAS las filas (uso exclusivo del usuario
    global, ver USUARIO_GLOBAL_ESPERA en servidor_ocr_ps.py). Con
    `empleado_id`: solo las de ESE usuario — un usuario normal nunca debe
    poder listar las de otro."""
    conexion = obtener_conexion()
    try:
        with conexion.cursor() as cursor:
            if empleado_id:
                cursor.execute(
                    """
                    SELECT id, fecha_registro, empleado_id, empleado_nombre, etiqueta, archivo_origen, motor, campos_json, cliente_json
                    FROM pedidos_en_espera WHERE empleado_id = %s ORDER BY fecha_registro DESC
                    """,
                    (empleado_id,),
                )
            else:
                cursor.execute(
                    """
                    SELECT id, fecha_registro, empleado_id, empleado_nombre, etiqueta, archivo_origen, motor, campos_json, cliente_json
                    FROM pedidos_en_espera ORDER BY fecha_registro DESC
                    """
                )
            filas = cursor.fetchall()
        resultado = []
        for fila in filas:
            campos = json.loads(fila["campos_json"]) if fila["campos_json"] else {"sede": None, "items": []}
            items = campos.get("items") or []
            resultado.append({
                "id": fila["id"],
                "fecha_registro": fila["fecha_registro"].isoformat(sep=" ") if fila["fecha_registro"] else None,
                "empleado_id": fila["empleado_id"],
                "empleado_nombre": fila["empleado_nombre"],
                "etiqueta": fila["etiqueta"],
                "archivo_origen": fila["archivo_origen"],
                "total_items": len(items),
                "cliente": json.loads(fila["cliente_json"]) if fila["cliente_json"] else None,
            })
        return resultado
    finally:
        conexion.close()


def obtener_pedido_en_espera(espera_id):
    """Fila completa (con campos_json/cliente_json ya parseados) para
    retomar la tarjeta, o None si no existe. El chequeo de "es tuya o sos
    el usuario global" se hace en la ruta Flask, acá solo se lee."""
    conexion = obtener_conexion()
    try:
        with conexion.cursor() as cursor:
            cursor.execute(
                """
                SELECT id, fecha_registro, empleado_id, empleado_nombre, etiqueta, archivo_origen, motor, campos_json, cliente_json
                FROM pedidos_en_espera WHERE id = %s
                """,
                (espera_id,),
            )
            fila = cursor.fetchone()
        if fila is None:
            return None
        return {
            "id": fila["id"],
            "fecha_registro": fila["fecha_registro"].isoformat(sep=" ") if fila["fecha_registro"] else None,
            "empleado_id": fila["empleado_id"],
            "empleado_nombre": fila["empleado_nombre"],
            "etiqueta": fila["etiqueta"],
            "archivo_origen": fila["archivo_origen"],
            "motor": fila["motor"],
            "campos": json.loads(fila["campos_json"]) if fila["campos_json"] else {"sede": None, "items": []},
            "cliente": json.loads(fila["cliente_json"]) if fila["cliente_json"] else None,
        }
    finally:
        conexion.close()


def eliminar_pedido_en_espera(espera_id):
    conexion = obtener_conexion()
    try:
        with conexion.cursor() as cursor:
            cursor.execute("DELETE FROM pedidos_en_espera WHERE id = %s", (espera_id,))
        conexion.commit()
    finally:
        conexion.close()


def guardar_alerta_stock(cod_cliente, codigo_articulo, cliente_nombre=None, descripcion_articulo=None, solicitado_por=None):
    """Botón "Generar alerta" en la lupa de Pedidos OCR (10/09, a pedido
    explícito del usuario) — antes era decorativo. Deja registro de qué
    cliente pidió un artículo que en ese momento no tenía stock en el
    maestro de ARA_PROYECT, para poder avisarle a compras. Exige
    cod_cliente y codigo_articulo; el resto es informativo."""
    conexion = obtener_conexion()
    try:
        with conexion.cursor() as cursor:
            cursor.execute(
                """
                INSERT INTO alertas_stock (cod_cliente, cliente_nombre, codigo_articulo, descripcion_articulo, solicitado_por)
                VALUES (%s, %s, %s, %s, %s)
                """,
                (cod_cliente, cliente_nombre, codigo_articulo, descripcion_articulo, solicitado_por),
            )
            id_nuevo = cursor.lastrowid
        conexion.commit()
        return id_nuevo
    finally:
        conexion.close()


def actualizar_estado_pedido_ocr(id_pedido, estado, respuesta_endpoint=None, numero_cotizacion=None):
    """Se llama después de la respuesta del endpoint real de creación de
    pedidos, para dejar en el histórico si terminó enviado o en error, y
    con qué número de cotización quedó asociado (si vino en la respuesta)."""
    conexion = obtener_conexion()
    try:
        with conexion.cursor() as cursor:
            cursor.execute(
                "UPDATE pedidos_ocr SET estado = %s, respuesta_endpoint = %s, numero_cotizacion = %s WHERE id = %s",
                (estado, respuesta_endpoint, numero_cotizacion, id_pedido),
            )
        conexion.commit()
    finally:
        conexion.close()


def actualizar_numero_entrega(id_pedido, numero_entrega):
    """El número de Nota de Entrega real de la cotización (10/09), ver
    bin/consultar_nota_entrega_ara.py."""
    conexion = obtener_conexion()
    try:
        with conexion.cursor() as cursor:
            cursor.execute(
                "UPDATE pedidos_ocr SET numero_entrega = %s WHERE id = %s",
                (numero_entrega, id_pedido),
            )
        conexion.commit()
    finally:
        conexion.close()


def actualizar_numero_factura(id_pedido, numero_factura):
    """El número real de FACTURA de Profit resuelto a partir de la nota
    de entrega (11/09, cambio de plan del usuario: el N° de nota ya no
    es clickeable, ahora lo es el N° de factura — ver
    bin/consultar_factura_pdf_ara.py). Puede traer varios números
    separados por coma, uno por cada nota del pedido."""
    conexion = obtener_conexion()
    try:
        with conexion.cursor() as cursor:
            cursor.execute(
                "UPDATE pedidos_ocr SET numero_factura = %s WHERE id = %s",
                (numero_factura, id_pedido),
            )
        conexion.commit()
    finally:
        conexion.close()


def vincular_documento_pedido(id_pedido, documento_id):
    """Asocia un documento YA guardado (ver bd_ocr.obtener_o_guardar_documento_ps
    — pedidos_ocr.documento_id apunta a esa misma tabla `documentos_ps`,
    compartida con OCR-PS, no una tabla aparte) a esta fila de
    pedidos_ocr. Pendiente de conectar (10/09): todavía no hay ningún
    flujo que llame a esto — falta la API que va a pasar el usuario para
    traer la preforma apenas se crea la cotización."""
    conexion = obtener_conexion()
    try:
        with conexion.cursor() as cursor:
            cursor.execute(
                "UPDATE pedidos_ocr SET documento_id = %s WHERE id = %s",
                (documento_id, id_pedido),
            )
        conexion.commit()
    finally:
        conexion.close()


def listar_pedidos_ocr(limite=200, desde=None, hasta=None):
    """Historial de pedidos subidos (cualquier estado) — para el
    "mini dashboard" de historial en la interfaz de Pedidos OCR (14/09,
    a pedido explícito del usuario: "ver el historial completo... y
    filtrarlas por día"). `desde`/`hasta` (fechas "YYYY-MM-DD") filtran
    directo en SQL — antes esto traía un bloque fijo (máx. 500) y el
    filtro de fecha se aplicaba en JavaScript sobre ESE bloque, así que
    un día viejo fuera del bloque quedaba invisible aunque existiera en
    la base. Con filtro de fecha en SQL, el límite de arriba es solo un
    tope de seguridad (nadie pide "un día" y espera 2000 filas), no un
    techo real del historial. El filtro por usuario (`subido_por`) sigue
    aplicándose en Python, ver servidor_ocr_ps.pedidos_lista_historico —
    ahí ya se normaliza mayúsculas/acentos, cosa que un WHERE exacto en
    SQL no haría con la misma tolerancia."""
    limite = max(1, min(2000, int(limite or 200)))
    condiciones = []
    parametros = []
    if desde:
        condiciones.append("fecha_registro >= %s")
        parametros.append(f"{desde} 00:00:00")
    if hasta:
        condiciones.append("fecha_registro <= %s")
        parametros.append(f"{hasta} 23:59:59")
    where = f"WHERE {' AND '.join(condiciones)}" if condiciones else ""
    conexion = obtener_conexion()
    try:
        with conexion.cursor() as cursor:
            cursor.execute(
                f"""
                SELECT id, fecha_registro, cod_cliente, cliente_nombre, estado, numero_cotizacion,
                       numero_entrega, numero_factura, documento_id, items_json, subido_por
                FROM pedidos_ocr {where} ORDER BY fecha_registro DESC LIMIT %s
                """,
                (*parametros, limite),
            )
            filas = cursor.fetchall()
        for fila in filas:
            try:
                fila["total_articulos"] = len(json.loads(fila.pop("items_json") or "[]"))
            except (TypeError, ValueError):
                fila["total_articulos"] = 0
                fila.pop("items_json", None)
            # "tiene_documento" nomás (nunca el id crudo) — el frontend solo
            # necesita saber si hay un PDF para armar el link de descarga.
            fila["tiene_documento"] = fila.pop("documento_id") is not None
        return filas
    finally:
        conexion.close()


def listar_ejecutivos_pedidos_ocr(desde=None, hasta=None):
    """Nombres distintos de "subido_por" con al menos un pedido en el
    rango de fechas dado (14/09, a pedido explícito del usuario: "filtrar
    por... ejecutivo activo de ese día") — para llenar el selector de
    ejecutivo con solo quienes realmente subieron algo en ese rango, no
    una lista fija de todos los usuarios que existieron alguna vez."""
    condiciones = ["subido_por IS NOT NULL", "subido_por != ''"]
    parametros = []
    if desde:
        condiciones.append("fecha_registro >= %s")
        parametros.append(f"{desde} 00:00:00")
    if hasta:
        condiciones.append("fecha_registro <= %s")
        parametros.append(f"{hasta} 23:59:59")
    where = f"WHERE {' AND '.join(condiciones)}"
    conexion = obtener_conexion()
    try:
        with conexion.cursor() as cursor:
            cursor.execute(
                f"SELECT DISTINCT subido_por FROM pedidos_ocr {where} ORDER BY subido_por",
                tuple(parametros),
            )
            return [fila["subido_por"] for fila in cursor.fetchall()]
    finally:
        conexion.close()


def obtener_pedido_ocr(id_pedido):
    """UN pedido completo por id, con items_json ya parseado en
    "items" (11/09, para armar la proforma de la nota — ver
    bin/generar_proforma_nota.py: ahí hace falta la cantidad/descripción
    que se guardó al montar el pedido, reng_nde no las tiene)."""
    conexion = obtener_conexion()
    try:
        with conexion.cursor() as cursor:
            cursor.execute(
                """
                SELECT id, fecha_registro, cod_cliente, cliente_nombre, estado, numero_cotizacion,
                       numero_entrega, numero_factura, items_json
                FROM pedidos_ocr WHERE id = %s
                """,
                (id_pedido,),
            )
            fila = cursor.fetchone()
        if not fila:
            return None
        try:
            fila["items"] = json.loads(fila.pop("items_json") or "[]")
        except (TypeError, ValueError):
            fila["items"] = []
            fila.pop("items_json", None)
        return fila
    finally:
        conexion.close()


def contar_pedidos_ocr():
    """Cuántos pedidos se subieron con éxito hasta ahora — reemplaza el
    contador de filas del Excel histórico (ya no existe para Pedidos OCR,
    ver servidor_ocr_ps.py)."""
    conexion = obtener_conexion()
    try:
        with conexion.cursor() as cursor:
            cursor.execute("SELECT COUNT(*) AS total FROM pedidos_ocr WHERE estado = 'enviado'")
            fila = cursor.fetchone()
        return fila["total"] if fila else 0
    finally:
        conexion.close()


def consultar_pedidos_ocr(numero_cotizacion=None, cod_cliente=None, limite=20):
    """Consulta de SOLO LECTURA para otros sistemas (ej. el motor/ara_coder
    de ARA_PROYECT): devuelve quién subió cada pedido (usuario del ERP vía
    SSO, en subido_por), cuándo (fecha_registro) y el número de cotización.
    Sirve para responder "quién sacó esta cotización" con datos que solo
    viven acá (el OCR + esta BD). Filtra por número de cotización y/o
    código de cliente; sin filtro devuelve los últimos `limite` pedidos."""
    limite = max(1, min(100, int(limite or 20)))
    condiciones = []
    parametros = []
    if numero_cotizacion is not None:
        # FIND_IN_SET en vez de "=": un pedido de más de ~21 artículos
        # queda con VARIAS cotizaciones guardadas juntas separadas por
        # coma (ej. "1234,1235") — con "=" nunca matchearía buscando por
        # una sola de esas dos (10/09).
        condiciones.append("FIND_IN_SET(%s, numero_cotizacion)")
        parametros.append(str(numero_cotizacion))
    if cod_cliente:
        condiciones.append("cod_cliente = %s")
        parametros.append(cod_cliente)
    where = ("WHERE " + " AND ".join(condiciones)) if condiciones else ""
    conexion = obtener_conexion()
    try:
        with conexion.cursor() as cursor:
            cursor.execute(
                f"""
                SELECT numero_cotizacion, subido_por, fecha_registro, cod_cliente, cliente_nombre, estado
                FROM pedidos_ocr {where} ORDER BY fecha_registro DESC LIMIT %s
                """,
                (*parametros, limite),
            )
            filas = cursor.fetchall()
        resultado = []
        for fila in filas:
            resultado.append({
                "numero_cotizacion": fila["numero_cotizacion"],
                "subido_por": fila["subido_por"],
                "fecha_registro": fila["fecha_registro"].isoformat(sep=" ") if fila["fecha_registro"] else None,
                "cod_cliente": fila["cod_cliente"],
                "cliente_nombre": fila["cliente_nombre"],
                "estado": fila["estado"],
            })
        return resultado
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


def buscar_facturas_por_proveedor(query, limite=30):
    """Busca facturas por nombre o código de proveedor (LIKE).
    Devuelve lista con campos básicos + documento_id para cargar el PDF."""
    query = f"%{query}%"
    conexion = obtener_conexion()
    try:
        with conexion.cursor() as cursor:
            cursor.execute(
                """
                SELECT o.id, o.fecha_registro, o.numero_factura, o.codigo_proveedor,
                       o.nombre_proveedor, o.rif_proveedor, o.motor, d.nombre_archivo
                FROM ocr_ps o
                LEFT JOIN documentos_ps d ON o.documento_id = d.id
                WHERE o.nombre_proveedor LIKE %s OR o.codigo_proveedor LIKE %s
                ORDER BY o.fecha_registro DESC LIMIT %s
                """,
                (query, query, max(1, min(50, limite))),
            )
            return cursor.fetchall()
    finally:
        conexion.close()


def listar_facturas_recientes(limite=20):
    """Lista facturas recientes con info básica del proveedor y documento."""
    conexion = obtener_conexion()
    try:
        with conexion.cursor() as cursor:
            cursor.execute(
                """
                SELECT o.id, o.fecha_registro, o.numero_factura, o.codigo_proveedor,
                       o.nombre_proveedor, o.rif_proveedor, d.nombre_archivo
                FROM ocr_ps o
                LEFT JOIN documentos_ps d ON o.documento_id = d.id
                ORDER BY o.fecha_registro DESC LIMIT %s
                """,
                (max(1, min(50, limite)),),
            )
            return cursor.fetchall()
    finally:
        conexion.close()


def obtener_pdf_factura(factura_id):
    """Devuelve (nombre_archivo, contenido_bytes) del PDF de una factura
    dada. Si no tiene documento asociado, devuelve (None, None)."""
    conexion = obtener_conexion()
    try:
        with conexion.cursor() as cursor:
            cursor.execute(
                """
                SELECT d.nombre_archivo, d.contenido
                FROM ocr_ps o
                JOIN documentos_ps d ON o.documento_id = d.id
                WHERE o.id = %s
                """,
                (factura_id,),
            )
            fila = cursor.fetchone()
            if fila:
                return fila["nombre_archivo"], fila["contenido"]
            return None, None
    finally:
        conexion.close()


if __name__ == "__main__":
    crear_tablas()
    print(f"Base de datos OCR creada/actualizada en MySQL ({MYSQL_HOST}:{MYSQL_PORT}/{MYSQL_DB})")
