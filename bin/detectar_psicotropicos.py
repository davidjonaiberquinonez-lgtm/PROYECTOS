# -*- coding: utf-8 -*-
"""
detectar_psicotropicos.py — Bloqueo de cumplimiento para Pedidos OCR
(11/09, a pedido explícito del usuario): Cristmedicals no puede montar
pedidos de productos psicotrópicos/controlados por esta vía — tienen que
ir por el departamento correspondiente. Ni el pedido armado por OCR ni
uno armado/editado a mano puede pasar esto por alto, así que el chequeo
real que importa es el de servidor_ocr_ps.pedidos_subir (server-side,
imposible de saltear); el de la interfaz (ocr_pedidos.html) es solo para
avisar antes, no reemplaza a este.

No hay ningún campo "psicotrópico"/"controlado" en el maestro de
ARA_PROYECT ni en Profit (solo código, descripción y stock) — así que la
detección es por PALABRA en la descripción del artículo ya confirmado
contra el maestro (nunca el texto crudo del OCR sin confirmar: ese nunca
llega a pedidos_subir sin pasar antes por la lupa, ver elegirProducto en
ocr_pedidos.html).

IMPORTANTE: esta lista es un punto de partida (principios activos +
algunas marcas venezolanas conocidas de benzodiacepinas/hipnóticos/
opioides/estimulantes controlados) armada con conocimiento general de
farmacología — NO es la lista oficial de la Oficina Nacional
Antidrogas/Ley Orgánica de Drogas. El departamento de cumplimiento de
Cristmedicals tiene que revisarla y completarla; agregar una palabra acá
es tan simple como sumarla al set de abajo (sin tildes, en mayúsculas).
"""
import re

PALABRAS_PSICOTROPICOS = {
    # Benzodiacepinas (principio activo)
    "ALPRAZOLAM", "BROMAZEPAM", "CLONAZEPAM", "DIAZEPAM", "LORAZEPAM",
    "MIDAZOLAM", "CLORDIAZEPOXIDO", "FLUNITRAZEPAM", "TRIAZOLAM",
    "CLOBAZAM", "FLURAZEPAM", "NITRAZEPAM", "CLORAZEPATO",
    # Hipnóticos no benzodiacepínicos ("Z-drugs")
    "ZOLPIDEM", "ZOPICLONA", "ESZOPICLONA",
    # Barbitúricos
    "FENOBARBITAL", "PENTOBARBITAL", "TIOPENTAL", "SECOBARBITAL",
    # Opioides
    "TRAMADOL", "MORFINA", "CODEINA", "OXICODONA", "FENTANILO",
    "METADONA", "MEPERIDINA", "PETIDINA", "BUPRENORFINA", "HIDROCODONA",
    "HIDROMORFONA",
    # Estimulantes controlados
    "METILFENIDATO", "ANFETAMINA", "DEXTROANFETAMINA",
    "LISDEXANFETAMINA", "FENTERMINA",
    # Otros
    "KETAMINA", "PREGABALINA",
    # Marcas venezolanas conocidas (el nombre en la factura/pedido casi
    # siempre es la marca, no el principio activo)
    "ANSILAN", "CLONATRIL", "ZOLPIDEX", "RIVOTRIL", "LEXOTANIL",
    "TRANXENE", "XANAX", "VALIUM",
}


def _normalizar(texto: str) -> str:
    texto = (texto or "").strip().upper()
    reemplazos = {"Á": "A", "É": "E", "Í": "I", "Ó": "O", "Ú": "U", "Ñ": "N"}
    for origen, destino in reemplazos.items():
        texto = texto.replace(origen, destino)
    return texto


def detectar_psicotropico(texto: str) -> str | None:
    """Devuelve la palabra de PALABRAS_PSICOTROPICOS que aparece en
    `texto` (como palabra completa, no substring de otra palabra — para
    que "DIAZEPAM" no dispare con "DIAZEPAMOL" si algún día existiera),
    o None si no matchea ninguna."""
    normalizado = _normalizar(texto)
    if not normalizado:
        return None
    for palabra in PALABRAS_PSICOTROPICOS:
        if re.search(rf"\b{re.escape(palabra)}\b", normalizado):
            return palabra
    return None


def revisar_items_psicotropicos(items) -> list[dict]:
    """Revisa una lista de items de pedido (cada uno con 'articulo') y
    devuelve los que matchearon, como [{"articulo": ..., "palabra": ...}, ...].
    Lista vacía si ninguno matchea."""
    encontrados = []
    for item in items or []:
        palabra = detectar_psicotropico(item.get("articulo") if isinstance(item, dict) else None)
        if palabra:
            encontrados.append({"articulo": item.get("articulo"), "palabra": palabra})
    return encontrados
