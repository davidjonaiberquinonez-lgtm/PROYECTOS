# -*- coding: utf-8 -*-
"""
consultar_producto_ara.py — Adaptador de solo lectura: busca artículos
reales en el maestro de stock de ARA_PROYECT (tabla `stock_maestro`).

Hasta el 07/09 esto abría directo el SQLite de ARA_PROYECT
(C:/ARA_PROYECT/ARA/ARA_BRAIN/DATA/proyecto_ara.db) — un archivo LOCAL de
esta máquina, que rompía la lupa de artículos apenas Pedidos OCR se movía a
otra máquina (bloqueante real, encontrado al armar los contenedores Docker).
ara-proyect-3e armó del otro lado una API pública de solo lectura
(192.168.4.217:4050) que devuelve exactamente la misma forma de respuesta —
este archivo ahora le pega a esa API en vez de abrir el archivo, así que
sirve desde cualquier máquina con línea de red hacia 192.168.4.217, sin
depender de un path local.

Pensado para Pedidos OCR (servidor_ocr_ps.py): el OCR lee a mano alzada el
nombre de un artículo pedido, y en vez de dejar esa transcripción cruda tal
cual, el usuario busca acá contra el maestro real para confirmar el código y
la descripción oficial antes de exportar — evita artículos mal tipeados o
sin código de catálogo.

Patrón anti-zombi (mismo criterio que consultar_proveedor_profit.py):
  - Reintentos ACOTADOS ante error de red/timeout — nunca en loop indefinido.
  - LIMIT <limite> siempre aplicado.
  - Nunca lanza excepción hacia quien llama.

Uso como módulo (para wirearlo a un endpoint Flask):
    from bin.consultar_producto_ara import buscar_producto
    resultado = buscar_producto(texto="amoxicilina")

Uso como CLI (para probarlo suelto, sin levantar ningún servidor):
    python bin/consultar_producto_ara.py --texto "amoxicilina"
"""
import argparse
import json
import os
import re
import time

import requests
from rapidfuzz import fuzz, utils

ARA_API_URL = os.environ.get(
    "ARA_API_PUBLICA_URL", "http://192.168.4.217:4050/api/publico/productos/buscar"
)
ARA_API_KEY = os.environ.get("ARA_API_PUBLICA_KEY", "")

CONNECT_TIMEOUT_S = 6
MAX_INTENTOS = 3
ESPERA_REINTENTO_S = 1
LIMITE_MAX = 50

# ---------------------------------------------------------------------------
# Ajuste de coincidencia por dosis/presentación (08/09) — caso real reportado
# por el usuario: "Budeconit Sole x 15ml gotas" contra el maestro daba 86% de
# coincidencia por igual a "BUDESONIDA ... SOL P/NEBULIZAR X 15ML" (correcto)
# Y a "BUDESONIDA ... INHALADOR X 100 DOSIS" (presentación totalmente
# distinta) — WRatio detecta el principio activo/marca (Budesonida/Budecort)
# pero trata los números y unidades como un token más, sin pesarlos. Esto
# extrae pares (número, unidad) de ambos textos con regex y suma/resta puntos
# según coincidan o no — desempata a favor del candidato con la MISMA medida
# exacta, no solo el mismo principio activo.
# ---------------------------------------------------------------------------
PATRON_MEDIDA = re.compile(
    r"(\d+(?:[.,]\d+)?)\s*(mg/ml|mcg|mg|ml|ui|dosis|caps?|tabl?(?:etas?)?|amp(?:ollas?)?|fr(?:ascos?)?|gotas?|und?|unidades?)",
    re.IGNORECASE,
)
_MAPA_UNIDADES = {"tabl": "tab", "tableta": "tab", "ampolla": "amp", "frasco": "fr", "unidad": "und"}


def _normalizar_unidad(unidad: str) -> str:
    unidad = unidad.lower().rstrip("s")
    return _MAPA_UNIDADES.get(unidad, unidad)


def _extraer_medidas(texto: str) -> dict:
    """{'ml': {15.0}, 'mg': {250.0}, ...} — todas las medidas número+unidad
    que aparecen en el texto, agrupadas por unidad normalizada."""
    medidas: dict = {}
    for numero, unidad in PATRON_MEDIDA.findall(texto or ""):
        valor = float(numero.replace(",", "."))
        medidas.setdefault(_normalizar_unidad(unidad), set()).add(valor)
    return medidas


def _ajuste_por_medida(texto_original: str, descripcion: str) -> int:
    """+10 por cada unidad donde el número coincide exacto (misma dosis real);
    -12 si la unidad aparece en los dos pero con número distinto (presentación
    distinta, como 15ML vs 100 DOSIS en el caso real); -6 si el candidato ni
    siquiera menciona esa unidad. Si el texto original no tiene ninguna
    medida reconocible, no ajusta nada (0) — no todo renglón de OCR trae una
    dosis explícita."""
    medidas_original = _extraer_medidas(texto_original)
    if not medidas_original:
        return 0
    medidas_candidato = _extraer_medidas(descripcion)
    ajuste = 0
    for unidad, numeros_originales in medidas_original.items():
        numeros_candidato = medidas_candidato.get(unidad)
        if numeros_candidato is None:
            ajuste -= 6
        elif numeros_originales & numeros_candidato:
            ajuste += 10
        else:
            ajuste -= 12
    return ajuste


PATRON_LABORATORIO = re.compile(r"\(([A-Za-zÁÉÍÓÚÑáéíóúñ0-9 .&'-]{2,40})\)")


def _extraer_laboratorio(texto: str) -> str | None:
    """El laboratorio/marca casi siempre viene entre paréntesis, al final
    (ej. "...600mg cj x 10 tab (GENVEN)") — se toma el ÚLTIMO paréntesis
    del texto (por si hay otro antes, como una presentación) porque el
    laboratorio/marca suele ser el último dato de la línea."""
    coincidencias = PATRON_LABORATORIO.findall(texto or "")
    return coincidencias[-1].strip() if coincidencias else None


def _extraer_farmaco_y_cantidad(texto: str) -> tuple:
    """(fármaco, cantidad) — el principio activo (primera palabra real,
    saltando "de") y la primera medida número+unidad reconocida (ej.
    "600mg"), NO necesariamente la segunda palabra: un renglón real como
    "ibuprofeno tabletas 600mg" tenía la cantidad en la TERCERA palabra,
    no en la segunda, y el patrón viejo ("palabra1%palabra2") armaba
    "ibuprofeno%tabletas" en vez de "ibuprofeno%600mg"."""
    palabras = [p for p in (texto or "").split() if p.lower() != "de"]
    farmaco = palabras[0] if palabras else (texto or "").strip()
    coincidencia_medida = PATRON_MEDIDA.search(texto or "")
    cantidad = f"{coincidencia_medida.group(1)}{coincidencia_medida.group(2)}" if coincidencia_medida else None
    return farmaco, cantidad


def _ajuste_por_laboratorio(laboratorio: str, descripcion: str) -> int:
    """+15 si se detectó un laboratorio/marca entre paréntesis en lo que
    se buscó Y aparece en la descripción del candidato — sin esto, dos
    candidatos con la MISMA coincidencia de texto (100%) quedaban
    empatados aunque uno tuviera justo la marca pedida y el otro no
    (12/09, caso real: "...(genven)" empataba 100% contra productos sin
    ningún parecido a esa marca). No penaliza si no aparece — el
    laboratorio puede venir mal transcripto por el OCR, no hay que
    descartar el producto correcto solo por eso."""
    if not laboratorio:
        return 0
    if laboratorio.upper() in (descripcion or "").upper():
        return 15
    return 0


def _ajuste_por_combinacion(texto_original: str, descripcion: str) -> int:
    """-20 si el candidato es una COMBINACIÓN (trae "+" en la
    descripción, ej. "IBUPROFENO 600MG + TIOCOLCHICOSIDO 4MG") pero lo
    que se buscó no menciona ningún "+" — señal de que se está pidiendo
    el producto de UN solo principio activo y el candidato tiene uno de
    más (11/09, caso real: "ibuprofeno 600mg cj x 10 tab (genven)" daba
    100% de coincidencia por igual contra el Ibuprofeno solo Y contra
    más de diez combinaciones con Tiocolchicosido — un producto
    completamente distinto — enterrando el resultado exacto entre
    combinaciones que nadie pidió). Si el texto buscado SÍ trae "+", no
    se penaliza nada — ahí sí se está pidiendo una combinación."""
    if "+" in descripcion and "+" not in texto_original:
        return -20
    return 0


def buscar_producto(texto: str = "", limite: int = 15, texto_original: str = "") -> dict:
    """Busca en `stock_maestro` por coincidencia parcial en descripción o
    código (case-insensitive). El texto es obligatorio — nunca trae "todo
    el maestro" sin querer. Solo lectura.

    `texto_original` (opcional) es el renglón COMPLETO que extrajo el OCR
    (item.articulo) — se usa solo para calcular el % de coincidencia de cada
    resultado, en vez de `texto` (que puede ser una sola palabra editada a
    mano en el buscador, muy poco discriminante para el puntaje). Si no se
    manda, se usa `texto` para las dos cosas.

    Devuelve {"ok": True, "total": N, "productos": [...]} (cada producto con
    "coincidencia": 0-100) o {"ok": False, "error": "..."} — nunca lanza
    excepción hacia quien la llama.
    """
    texto = (texto or "").strip()
    texto_original = (texto_original or "").strip() or texto
    if not texto:
        return {"ok": False, "error": "Hace falta el texto a buscar."}
    if not ARA_API_KEY:
        return {"ok": False, "error": "Falta configurar ARA_API_PUBLICA_KEY en el entorno."}

    limite = max(1, min(LIMITE_MAX, int(limite or 15)))
    # Se le pide a la API de ARA_PROYECT el MÁXIMO (50), no el `limite`
    # que pidió quien llama (11/09, caso real: buscando "ibuprofeno
    # 600mg cj x 10 tab (genven)" el resultado EXACTO quedaba afuera del
    # todo con limite=15 — esa API ordena por stock, no por relevancia,
    # y hay más de 15 combinaciones con tiocolchicosido (otro producto)
    # con más stock que la marca puntual que se pidió). Acá se reordena
    # por coincidencia sobre el pool completo y recién ahí se recorta al
    # `limite` real — así el mejor resultado nunca queda afuera solo por
    # tener menos stock que candidatos menos relevantes.
    limite_pedido = limite
    limite = LIMITE_MAX
    # Patrón por TRES campos (12/09, a pedido explícito del usuario):
    # fármaco (principio activo) + cantidad (dosis) + laboratorio/marca
    # — en vez de asumir a ciegas "primera y segunda palabra", que fallaba
    # cuando la cantidad no era la 2da palabra (ej. "ibuprofeno tabletas
    # 600mg"). El patrón que va al SQL sigue siendo fármaco+cantidad nomás
    # (laboratorio se usa para desempatar el puntaje más abajo, no en el
    # SQL — filtrar por marca en el WHERE dejaría afuera candidatos donde
    # el OCR leyó mal la marca pero el resto está bien). La API pública de
    # ARA_PROYECT ya envuelve `q` en %...% de su lado, así que acá no hace
    # falta agregar los % de los extremos, solo el del medio.
    farmaco, cantidad = _extraer_farmaco_y_cantidad(texto)
    laboratorio = _extraer_laboratorio(texto_original)
    if cantidad:
        patron = f"{farmaco}%{cantidad}"
    else:
        palabras = [p for p in texto.split() if p.lower() != "de"]
        patron = f"{palabras[0]}%{palabras[1]}" if len(palabras) >= 2 else texto

    intento = 0
    while True:
        intento += 1
        try:
            respuesta = requests.get(
                ARA_API_URL,
                params={"q": patron, "limite": limite},
                headers={"X-API-Key": ARA_API_KEY},
                timeout=CONNECT_TIMEOUT_S,
            )
            if respuesta.status_code == 401:
                return {"ok": False, "error": "ARA_API_PUBLICA_KEY inválida (401 de la API de ARA_PROYECT)."}
            respuesta.raise_for_status()
            datos = respuesta.json()
            if not datos.get("ok"):
                return {"ok": False, "error": datos.get("error") or "La API de ARA_PROYECT devolvió un error."}
            productos = datos.get("productos", [])
            # Puntaje de coincidencia (0-100) del texto ORIGINAL (no el patrón
            # con %) contra la descripción real del maestro — token_set_ratio
            # tolera palabras de más/de menos y orden distinto (ej. "amoxicilina
            # susp" vs "AMOXICILINA 250MG SUSP X 60ML"), a pedido del usuario
            # para pintar la lupa en verde/amarillo según qué tan buena es cada
            # coincidencia.
            for producto in productos:
                # Coincidencia por CÓDIGO DE BARRAS exacto (10/09, a pedido
                # explícito del usuario): si lo que se buscó es un código de
                # barras y el producto devuelto tiene ESE MISMO código de
                # barras, la identificación ya es 100% segura — no tiene
                # sentido bajarle el puntaje por comparar texto contra la
                # descripción que transcribió el OCR, que puede venir con
                # otro orden de palabras/redacción para el mismo producto
                # exacto (reportado en vivo: "NISTATINA SUSP ORAL 100.000UI
                # /5ML X 60ML(H&M)" del OCR vs "NISTATINA 100.000 UI/5ML
                # SUSP X 60ML (H&M)" del maestro — mismo producto, pero el
                # cálculo por texto daba 92% en vez de 100%). Si la
                # descripción está mal, el error está en el NOMBRE
                # transcripto, no en si es el producto correcto.
                codigo_barra_producto = str(producto.get("codigo_barra") or "").strip()
                if texto.isdigit() and 8 <= len(texto) <= 13 and codigo_barra_producto == texto:
                    producto["coincidencia"] = 100
                    producto["_orden"] = 999  # código de barras exacto: siempre primero
                    continue
                # WRatio (no token_set_ratio): probado en vivo con casos
                # reales — token_set_ratio da ~100 para CUALQUIER candidato que
                # contenga las palabras buscadas (por diseño, compara por
                # conjunto de palabras), sin distinguir uno bueno de uno
                # mediocre. WRatio sí separa bien (~86 para una coincidencia
                # real con ruido de OCR, ~57 para un candidato irrelevante).
                # processor=utils.default_process normaliza mayúsculas/espacios
                # (sin esto el puntaje sale cerca de 0 aunque coincida perfecto,
                # bug real encontrado probando esto en vivo).
                descripcion = producto.get("descripcion") or ""
                base = fuzz.WRatio(texto_original, descripcion, processor=utils.default_process)
                ajuste = (
                    _ajuste_por_medida(texto_original, descripcion)
                    + _ajuste_por_combinacion(texto_original, descripcion)
                    + _ajuste_por_laboratorio(laboratorio, descripcion)
                )
                # "coincidencia" (lo que se ve en pantalla) queda topado en
                # 100 — es un porcentaje, no puede pasarse. "_orden" NO se
                # topa: es solo para desempatar el ORDEN de la lista, para
                # que el ajuste por laboratorio no se pierda entre dos
                # candidatos que de todas formas iban a mostrar "100%"
                # (12/09, caso real: "...(genven)" empataba 100% contra
                # candidatos sin ningún parecido a esa marca).
                puntaje_bruto = round(base) + ajuste
                producto["coincidencia"] = max(0, min(100, puntaje_bruto))
                producto["_orden"] = puntaje_bruto
            # Reordenar por coincidencia (11/09, caso real reportado: buscando
            # "ibuprofeno 600mg cj x 10 tab (genven)" el resultado EXACTO
            # (100%, la marca que se pidió) quedaba último en la lista,
            # detrás de combinaciones con tiocolchicosido que ni siquiera
            # tienen esa marca — porque la API ordena por stock, no por
            # relevancia, y acá nunca se reordenaba con el puntaje ya
            # calculado. sort() es estable: entre dos productos con el
            # MISMO puntaje, se mantiene el orden por stock que ya traía.
            productos.sort(key=lambda p: -p["_orden"])
            productos = productos[:limite_pedido]
            for producto in productos:
                producto.pop("_orden", None)
            return {"ok": True, "total": datos.get("total", 0), "productos": productos}
        except requests.RequestException as exc:
            if intento < MAX_INTENTOS:
                time.sleep(ESPERA_REINTENTO_S)
                continue
            return {"ok": False, "error": f"No se pudo consultar el maestro de ARA_PROYECT: {exc}"}
        except Exception as exc:
            return {"ok": False, "error": f"No se pudo consultar el maestro de ARA_PROYECT: {exc}"}


def _main() -> None:
    parser = argparse.ArgumentParser(description="Busca artículos reales en el maestro de stock de ARA_PROYECT.")
    parser.add_argument("--texto", default="", help="Coincidencia parcial por descripción o código.")
    parser.add_argument("--limite", type=int, default=15)
    args = parser.parse_args()

    resultado = buscar_producto(texto=args.texto, limite=args.limite)
    print(json.dumps(resultado, ensure_ascii=False, indent=2, default=str))


if __name__ == "__main__":
    _main()
