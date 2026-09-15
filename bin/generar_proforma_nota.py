# -*- coding: utf-8 -*-
"""
generar_proforma_nota.py — Genera NUESTRO propio PDF de "Nota de Entrega
(proforma)" a partir de los renglones reales de Profit (reng_nde).

11/09, a pedido explícito del usuario: a diferencia de la factura (ver
bin/consultar_factura_pdf_ara.py), Profit SÍ genera un PDF de la nota de
entrega, pero no sabemos cómo lo arma ni dónde lo guarda — no hay ningún
servidor de PDFs para notas como el que sí existe para facturas
(192.168.4.23:3010). Este módulo arma un PROTOTIPO propio, calcado del
estilo visual real de las facturas de Crist Medicals (verde/blanco,
mismo bloque de datos fiscales, QR + código de barras con el número de
documento) — el usuario mandó una factura real como referencia de
estilo (11/09). Se deja EXPLÍCITO en el propio documento que es una
reconstrucción nuestra, no el documento oficial que emite Profit, para
no confundir a nadie que la reciba.

Fuente de los datos:
  - Los renglones REALES de esa nota puntual (qué código quedó, si algún
    renglón fue anulado, el comentario si se le escribió uno) salen de
    Profit vía bin/leer_renglones_nota_ara.py (reng_nde).
  - La cantidad y la descripción completa de cada artículo NO están en
    ese endpoint (reng_nde ahí solo expone co_art/comentario/anulado) —
    se completan cruzando contra lo que este mismo sistema ya guardó al
    montar el pedido (pedidos_ocr.items_json), que es de donde salió esa
    nota en primer lugar. No hay precio ni lote acá (una nota de entrega
    no es un documento fiscal como la factura, y pedidos_ocr tampoco
    guarda lote) — si el negocio necesita esos datos en la proforma más
    adelante, hay que sumarlos a lo que se guarda al montar el pedido.

Uso como módulo:
    from bin.generar_proforma_nota import generar_pdf_proforma_nota
    pdf_bytes = generar_pdf_proforma_nota(datos)

Uso como CLI (para probarlo suelto, con datos de ejemplo):
    python bin/generar_proforma_nota.py --salida proforma_ejemplo.pdf
"""
import argparse
from datetime import datetime
from io import BytesIO
from xml.sax.saxutils import escape as _esc

from reportlab.lib import colors
from reportlab.lib.pagesizes import letter
from reportlab.lib.units import cm
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, HRFlowable
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_RIGHT, TA_CENTER
from reportlab.graphics.shapes import Drawing
from reportlab.graphics.barcode import code128
from reportlab.graphics.barcode.qr import QrCodeWidget
from reportlab.graphics import renderPDF  # noqa: F401 (fuerza el registro de los widgets de barcode/qr)

# Mismo verde/gris que se ve en las facturas reales de Crist Medicals
# (referencia mandada por el usuario, 11/09) — antes esto usaba un
# morado inventado, sin relación con la identidad visual real.
_COLOR_MARCA = colors.HexColor("#1f8a4c")
_COLOR_MARCA_OSCURO = colors.HexColor("#166534")
_COLOR_GRIS = colors.HexColor("#6b7280")
_COLOR_FONDO_ENCABEZADO = colors.HexColor("#eaf7ef")
_COLOR_ANULADO = colors.HexColor("#b91c1c")


def generar_pdf_proforma_nota(datos: dict) -> bytes:
    """`datos` = {
        "numero_nota": "499165",
        "fecha": "2026-09-11 10:30" (o datetime, o None),
        "cod_cliente": "FAR01808",
        "cliente_nombre": "FARMACIA RENNY&ROS, C.A",
        "numero_cotizacion": "248336" (opcional, informativo),
        "renglones": [
            {"codigo": "MD02900", "descripcion": "...", "cantidad": "30",
             "comentario": "...", "anulado": False},
            ...
        ],
    }
    Devuelve los bytes del PDF armado. No consulta nada por su cuenta —
    quien llama ya le pasa los datos resueltos (ver
    servidor_ocr_ps.pedidos_proforma_nota)."""
    buffer = BytesIO()
    documento = SimpleDocTemplate(
        buffer, pagesize=letter,
        topMargin=1.3 * cm, bottomMargin=1.3 * cm, leftMargin=1.3 * cm, rightMargin=1.3 * cm,
    )
    estilos = getSampleStyleSheet()
    estilo_logo = ParagraphStyle("logo", parent=estilos["Normal"], fontSize=9, textColor=colors.white, alignment=TA_CENTER, leading=11)
    estilo_marca = ParagraphStyle("marca", parent=estilos["Heading1"], fontSize=14, textColor=colors.HexColor("#1a1a1a"), spaceAfter=1, alignment=TA_RIGHT)
    estilo_dato_empresa = ParagraphStyle("dato_empresa", parent=estilos["Normal"], fontSize=8, textColor=_COLOR_MARCA_OSCURO, leading=10.5, alignment=TA_RIGHT)
    estilo_etiqueta = ParagraphStyle("etiqueta", parent=estilos["Normal"], fontSize=8.5, textColor=_COLOR_GRIS, leading=13)
    estilo_valor = ParagraphStyle("valor", parent=estilos["Normal"], fontSize=9, textColor=colors.HexColor("#1a1a1a"), leading=13)
    estilo_titulo_doc = ParagraphStyle("titulo_doc", parent=estilos["Heading2"], fontSize=13, alignment=TA_RIGHT, textColor=_COLOR_MARCA_OSCURO, spaceAfter=2)
    estilo_dato_doc = ParagraphStyle("dato_doc", parent=estilos["Normal"], fontSize=9, alignment=TA_RIGHT, leading=13)
    estilo_aviso = ParagraphStyle("aviso", parent=estilos["Normal"], fontSize=7.5, textColor=colors.HexColor("#8a6d00"), alignment=TA_CENTER, leading=10)
    estilo_celda = ParagraphStyle("celda", parent=estilos["Normal"], fontSize=8.5, leading=11)
    estilo_celda_header = ParagraphStyle("celda_header", parent=estilos["Normal"], fontSize=8.5, leading=11, textColor=colors.white, fontName="Helvetica-Bold")
    estilo_celda_anulada = ParagraphStyle("celda_anulada", parent=estilos["Normal"], fontSize=8.5, leading=11, textColor=_COLOR_ANULADO)
    estilo_pie = ParagraphStyle("pie", parent=estilos["Normal"], fontSize=7, textColor=_COLOR_GRIS, alignment=TA_CENTER, leading=9.5)

    elementos = []

    # --- Encabezado: "logo" (izquierda) + datos de la empresa (derecha) ---
    # No hay un archivo de logo real disponible acá — placeholder verde
    # con el nombre, no el isotipo real de Crist Medicals.
    logo_placeholder = Table([[Paragraph("CRIST<br/>MEDICALS", estilo_logo)]], colWidths=[3.2 * cm], rowHeights=[1.6 * cm])
    logo_placeholder.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), _COLOR_MARCA),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("ALIGN", (0, 0), (-1, -1), "CENTER"),
        ("ROUNDEDCORNERS", [8, 8, 8, 8]),
    ]))
    bloque_empresa = [
        Paragraph("Crist Medicals, C.A.", estilo_marca),
        Paragraph("RIF: J-41223670-9", estilo_dato_empresa),
        Paragraph("Calle Pasaje Barcelona, Local Galpón Nro 3", estilo_dato_empresa),
        Paragraph("Zona Industrial De Puente Real, San Cristóbal, Táchira", estilo_dato_empresa),
    ]
    tabla_encabezado = Table([[logo_placeholder, bloque_empresa]], colWidths=[3.5 * cm, 15 * cm])
    tabla_encabezado.setStyle(TableStyle([("VALIGN", (0, 0), (-1, -1), "MIDDLE")]))
    elementos.append(tabla_encabezado)
    elementos.append(Spacer(1, 8))
    elementos.append(HRFlowable(width="100%", thickness=1.4, color=_COLOR_MARCA))
    elementos.append(Spacer(1, 8))

    # --- Cliente (izquierda) + documento + QR/código de barras (derecha) ---
    bloque_cliente = [
        Paragraph("Razón Social:", estilo_etiqueta),
        Paragraph(f"<b>{_esc(_texto(datos.get('cliente_nombre')) or '—')}</b>", estilo_valor),
        Spacer(1, 4),
        Paragraph("Código de Cliente:", estilo_etiqueta),
        Paragraph(_esc(_texto(datos.get("cod_cliente")) or "—"), estilo_valor),
    ]
    if datos.get("numero_cotizacion"):
        bloque_cliente += [
            Spacer(1, 4),
            Paragraph("Cotización de origen:", estilo_etiqueta),
            Paragraph(f"#{_esc(_texto(datos['numero_cotizacion']))}", estilo_valor),
        ]

    numero_nota = _texto(datos.get("numero_nota")) or "—"
    dibujo_codigos = _dibujar_qr_y_barras(numero_nota)

    bloque_documento = [
        Paragraph("PROFORMA — NOTA<br/>DE ENTREGA", estilo_titulo_doc),
        Paragraph(f"N° de Nota: <b>{_esc(numero_nota)}</b>", estilo_dato_doc),
        Paragraph(f"Fecha: {_esc(_formatear_fecha(datos.get('fecha')))}", estilo_dato_doc),
    ]

    tabla_cliente_doc = Table(
        [[bloque_cliente, dibujo_codigos, bloque_documento]],
        colWidths=[6.5 * cm, 6 * cm, 6 * cm],
    )
    tabla_cliente_doc.setStyle(TableStyle([("VALIGN", (0, 0), (-1, -1), "TOP"), ("ALIGN", (1, 0), (1, 0), "CENTER")]))
    elementos.append(tabla_cliente_doc)
    elementos.append(Spacer(1, 10))

    # --- Aviso: esto es un prototipo propio, no el documento oficial ---
    elementos.append(Paragraph(
        "Documento generado por Servidor de Retenciones (Pedidos OCR) a partir de los renglones reales de Profit — "
        "NO es el documento oficial que emite Profit, es una reconstrucción propia para uso interno mientras no exista "
        "una fuente oficial de la nota en PDF.",
        estilo_aviso,
    ))
    elementos.append(Spacer(1, 12))

    # --- Renglones ---
    filas_tabla = [[
        Paragraph("Código", estilo_celda_header),
        Paragraph("Descripción", estilo_celda_header),
        Paragraph("Cantidad", estilo_celda_header),
        Paragraph("Observación", estilo_celda_header),
    ]]
    for renglon in (datos.get("renglones") or []):
        anulado = bool(renglon.get("anulado"))
        estilo_fila = estilo_celda_anulada if anulado else estilo_celda
        descripcion = _esc(_texto(renglon.get("descripcion")) or "—")
        if anulado:
            descripcion = f"<b>[ANULADO]</b> {descripcion}"
        filas_tabla.append([
            Paragraph(_esc(_texto(renglon.get("codigo")) or "—"), estilo_fila),
            Paragraph(descripcion, estilo_fila),
            Paragraph(_esc(_texto(renglon.get("cantidad")) or "—"), estilo_fila),
            Paragraph(_esc(_texto(renglon.get("comentario")) or ""), estilo_fila),
        ])

    if len(filas_tabla) == 1:
        filas_tabla.append([Paragraph("—", estilo_celda), Paragraph("Sin renglones para mostrar", estilo_celda), Paragraph("—", estilo_celda), Paragraph("", estilo_celda)])

    tabla_items = Table(filas_tabla, colWidths=[2.6 * cm, 8.4 * cm, 2.3 * cm, 4.2 * cm], repeatRows=1)
    tabla_items.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), _COLOR_MARCA),
        ("ALIGN", (2, 0), (2, -1), "CENTER"),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#c9ebd6")),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, _COLOR_FONDO_ENCABEZADO]),
        ("TOPPADDING", (0, 0), (-1, -1), 5),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
        ("LEFTPADDING", (0, 0), (-1, -1), 6),
        ("RIGHTPADDING", (0, 0), (-1, -1), 6),
    ]))
    elementos.append(tabla_items)

    elementos.append(Spacer(1, 20))
    elementos.append(HRFlowable(width="100%", thickness=0.5, color=colors.HexColor("#c9ebd6")))
    elementos.append(Spacer(1, 6))
    elementos.append(Paragraph(
        f"Generado por Servidor de Retenciones (Pedidos OCR) el {datetime.now().strftime('%d/%m/%Y %H:%M')}. "
        "Documento interno — no reemplaza ningún comprobante fiscal oficial.",
        estilo_pie,
    ))

    documento.build(elementos)
    return buffer.getvalue()


def _dibujar_qr_y_barras(numero_nota: str):
    """QR + código de barras Code128 con el N° de nota, apilados — mismo
    criterio visual que las facturas reales (referencia mandada por el
    usuario, 11/09): ambos códigos apuntan al mismo número, no a una
    URL, porque todavía no hay una página pública de consulta para
    esto. Devuelve una lista de flowables (no un solo Drawing): el QR
    (widget) necesita su transform en un Drawing propio, escalado a su
    tamaño natural — combinarlo con el código de barras en un único
    Drawing con coordenadas a mano es mucho más frágil."""
    valor = numero_nota or "0"

    qr_widget = QrCodeWidget(valor)
    bounds = qr_widget.getBounds()
    ancho_original = bounds[2] - bounds[0]
    alto_original = bounds[3] - bounds[1]
    lado_qr = 2.6 * cm
    dibujo_qr = Drawing(lado_qr, lado_qr, transform=[lado_qr / ancho_original, 0, 0, lado_qr / alto_original, 0, 0])
    dibujo_qr.add(qr_widget)

    # Code128 ya es un Flowable (no un Shape) — va directo a la lista,
    # sin envolverlo en un Drawing propio.
    codigo_barras = code128.Code128(valor, barHeight=0.9 * cm, barWidth=0.8)

    return [dibujo_qr, Spacer(1, 6), codigo_barras]


def _texto(valor):
    return str(valor).strip() if valor is not None else ""


def _formatear_fecha(valor):
    if not valor:
        return datetime.now().strftime("%d/%m/%Y")
    if isinstance(valor, datetime):
        return valor.strftime("%d/%m/%Y %H:%M")
    return str(valor)


def _main() -> None:
    parser = argparse.ArgumentParser(description="Genera un PDF de prueba de la proforma de nota de entrega con datos de ejemplo.")
    parser.add_argument("--salida", default="proforma_ejemplo.pdf")
    args = parser.parse_args()
    datos_ejemplo = {
        "numero_nota": "499165",
        "numero_cotizacion": "248336",
        "cod_cliente": "FAR01808",
        "cliente_nombre": "FARMACIA RENNY&ROS, C.A",
        "renglones": [
            {"codigo": "MD02900", "descripcion": "LEVOTIROXINA SODICA 50MCG (EUTIROX) CJ X 50 TABL (MERCK)", "cantidad": "1", "comentario": "", "anulado": False},
            {"codigo": "JBE00285", "descripcion": "AMOXICILINA 250MG + ACIDO CLAV. 62.5MG/5ML SUSP X 60ML (ANGELUS)", "cantidad": "6", "comentario": "Retira en la droguería", "anulado": False},
            {"codigo": "MQ00176", "descripcion": "AGUA OXIGENADA 3% 120ML (EL GUARDIAN)", "cantidad": "12", "comentario": "", "anulado": True},
        ],
    }
    pdf_bytes = generar_pdf_proforma_nota(datos_ejemplo)
    with open(args.salida, "wb") as f:
        f.write(pdf_bytes)
    print(f"OK — {len(pdf_bytes)} bytes, guardado en {args.salida}")


if __name__ == "__main__":
    _main()
