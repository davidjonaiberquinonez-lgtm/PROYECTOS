# -*- coding: utf-8 -*-
"""
evaluar_legibilidad.py — Puntaje de legibilidad (0-100) de una imagen de
comprobante/factura, SIN usar ningún modelo de IA — pura heurística de
procesamiento de imagen con PIL (Pillow, ya es dependencia del proyecto vía
pymupdf, no se agregó nada nuevo).

Motivación (09/09): en una prueba real, el mismo motor de visión
(Qwen2.5-VL-7B) sacó 100% de aciertos en una factura nítida y 6 errores de
dígitos + un caracter de más en el RIF en una factura borrosa — el problema
no es el modelo, es la calidad de la imagen de entrada. Este puntaje no
bloquea ni corrige nada — solo le avisa al usuario en pantalla cuál
resultado conviene revisar contra el documento original antes de confiar
en él a ojo cerrado.

Técnica: varianza del filtro Laplaciano sobre la imagen en escala de
grises — el mismo método estándar que usa OpenCV/scikit-image para
detectar desenfoque (a mayor varianza, bordes más marcados = imagen más
nítida), acá implementado con el kernel de convolución de PIL
(ImageFilter.Kernel) para no depender de numpy/opencv. Combinado con una
penalización si la imagen es de resolución muy chica (letra chica se
pierde sin importar qué tan "nítida" salga la varianza).

Uso:
    from evaluar_legibilidad import evaluar_legibilidad
    puntaje = evaluar_legibilidad(imagen_bytes)  # 0.0 a 100.0
"""
import io

from PIL import Image, ImageFilter, ImageStat

# Calibración empírica (documentos escaneados típicos, no un estudio
# formal): por debajo de este valor de varianza, el ojo humano ya nota la
# imagen borrosa; por encima, se ve nítida. Puede necesitar ajuste con más
# casos reales — es un punto de partida, no un número exacto.
UMBRAL_VARIANZA_BORROSA = 40.0
UMBRAL_VARIANZA_NITIDA = 600.0

# Lado mayor (en píxeles) por debajo del cual la resolución en sí ya es un
# problema para leer letra chica, sin importar qué tan nítida esté.
LADO_MINIMO_BUENA_RESOLUCION = 900

# Umbral por debajo del cual se considera "legibilidad baja" y conviene
# avisar en pantalla (ver retenciones.html, icono de advertencia).
UMBRAL_LEGIBILIDAD_BAJA = 50.0


def evaluar_legibilidad(imagen_bytes):
    """Devuelve un puntaje 0.0-100.0. Si la imagen no se puede ni abrir,
    devuelve 100.0 (no se puede evaluar, así que no se bloquea ni se
    marca nada por las dudas — nunca lanza excepción)."""
    try:
        imagen = Image.open(io.BytesIO(imagen_bytes)).convert("L")
    except Exception:
        return 100.0

    ancho, alto = imagen.size
    lado_mayor = max(ancho, alto) or 1

    # Achicar imágenes grandes antes del filtro: el puntaje de nitidez no
    # mejora por analizar más píxeles de los necesarios, y esto lo hace
    # rápido incluso con escaneos de alta resolución.
    if lado_mayor > 1400:
        factor = 1400 / lado_mayor
        imagen = imagen.resize((max(1, int(ancho * factor)), max(1, int(alto * factor))))

    laplaciano = imagen.filter(ImageFilter.Kernel((3, 3), [0, -1, 0, -1, 4, -1, 0, -1, 0], scale=1))
    varianza_nitidez = ImageStat.Stat(laplaciano).var[0]

    rango = UMBRAL_VARIANZA_NITIDA - UMBRAL_VARIANZA_BORROSA
    puntaje_nitidez = (varianza_nitidez - UMBRAL_VARIANZA_BORROSA) / rango * 100
    puntaje_nitidez = max(0.0, min(100.0, puntaje_nitidez))

    penalizacion_resolucion = 1.0 if lado_mayor >= LADO_MINIMO_BUENA_RESOLUCION else max(0.5, lado_mayor / LADO_MINIMO_BUENA_RESOLUCION)

    return round(puntaje_nitidez * penalizacion_resolucion, 1)
