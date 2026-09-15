# -*- coding: utf-8 -*-
"""
servir_imagen_producto.py — Resuelve la ruta local de la foto de un
producto, a partir de su código de artículo, contra la misma carpeta de
10.490 fotos que usa el Visor de Artículos de ARA_PROYECT
(C:\\Users\\Personal\\Pictures\\Nueva carpeta\\productos, un .jpg por
código, ej. ALIM0001.jpg). Reemplaza al CDN externo
(imagenes.cristmedicals.com) como fuente del preview al pasar el cursor
en el buscador de productos de Pedidos OCR (11/09, a pedido explícito
del usuario: "esta mucho mejor").

Uso como módulo:
    from bin.servir_imagen_producto import buscar_imagen_producto
    buscar_imagen_producto("ALIM0001")  # -> Path o None

Uso como CLI (para probarlo suelto, sin levantar ningún servidor):
    python bin/servir_imagen_producto.py --codigo ALIM0001
"""
import argparse
import os
from pathlib import Path
from typing import Optional

CARPETA_IMAGENES = Path(
    os.environ.get(
        "CARPETA_IMAGENES_PRODUCTOS",
        r"C:\Users\Personal\Pictures\Nueva carpeta\productos",
    )
)

# Por si algún día aparecen fotos en otro formato — hoy la carpeta es
# 100% .jpg (10.490 archivos, confirmado), pero probar unas pocas
# extensiones más no cuesta nada y evita tener que tocar esto después.
EXTENSIONES = (".jpg", ".jpeg", ".png", ".webp")


def buscar_imagen_producto(codigo: str) -> Optional[Path]:
    """Devuelve el Path de la foto de `codigo` si existe en la carpeta
    local, o None si no hay ninguna. Nunca lanza excepción."""
    codigo = (codigo or "").strip()
    if not codigo:
        return None
    try:
        for extension in EXTENSIONES:
            ruta = CARPETA_IMAGENES / f"{codigo}{extension}"
            if ruta.is_file():
                return ruta
        return None
    except OSError:
        return None


def _main() -> None:
    parser = argparse.ArgumentParser(description="Busca la foto local de un producto por su código de artículo.")
    parser.add_argument("--codigo", required=True, help="Código de artículo (nombre del archivo, sin extensión).")
    args = parser.parse_args()
    ruta = buscar_imagen_producto(args.codigo)
    print(str(ruta) if ruta else f"Sin foto para '{args.codigo}' en {CARPETA_IMAGENES}")


if __name__ == "__main__":
    _main()
