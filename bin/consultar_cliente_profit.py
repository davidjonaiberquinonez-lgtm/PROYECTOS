# -*- coding: utf-8 -*-
"""
consultar_cliente_profit.py — Adaptador de solo lectura: busca clientes
reales en PROFIT por nombre/razón social O por código exacto, a través del
adaptador REST que ya usa Central Telefónica IA (mismo patrón que
profit_conector.py de ese proyecto — acá se replica en vez de importar
entre carpetas, para que este módulo siga siendo standalone como los demás
en bin/).

Pensado para Pedidos OCR: al montar un pedido desde una nota escaneada hace
falta el código del cliente (cod_cliente) que exige el endpoint de creación
de pedidos — el usuario busca por nombre O por código, sin tener que saber
de antemano cuál de los dos está tipeando.

Uso como módulo:
    from bin.consultar_cliente_profit import buscar_cliente
    resultado = buscar_cliente(texto="farmacia san miguel")   # por nombre
    resultado = buscar_cliente(texto="FAR00877")               # por código
"""
import os

import requests

PROFIT_BASE_URL = os.environ.get("PROFIT_BASE_URL", "http://192.168.4.217:4050")
TIMEOUT_S = 15


def _buscar_por_codigo(texto):
    """Ficha exacta por co_cli — el adaptador devuelve un objeto plano
    (cliente/co_cli/razon_social sueltos), no una lista, y "ok": false con
    un error de SQL si el código no existe (no es una falla real, solo
    "no hay resultado" — se traga acá, nunca sube como excepción)."""
    try:
        respuesta = requests.get(
            f"{PROFIT_BASE_URL}/api/cliente/consulta", params={"co_cli": texto}, timeout=TIMEOUT_S,
        )
        cuerpo = respuesta.json()
    except (requests.RequestException, ValueError):
        return []
    if not cuerpo.get("ok") or not cuerpo.get("co_cli"):
        return []
    return [{"co_cli": cuerpo["co_cli"], "razon_social": cuerpo.get("razon_social") or cuerpo.get("cliente")}]


def _buscar_por_nombre(texto):
    try:
        respuesta = requests.get(
            f"{PROFIT_BASE_URL}/api/cliente/consulta", params={"busqueda": texto}, timeout=TIMEOUT_S,
        )
        cuerpo = respuesta.json()
    except (requests.RequestException, ValueError):
        return []
    return [
        {"co_cli": c.get("co_cli"), "razon_social": c.get("razon_social")}
        for c in (cuerpo.get("clientes") or [])
        if c.get("co_cli")
    ]


def buscar_cliente(texto: str = "", limite: int = 15) -> dict:
    """Busca clientes reales en PROFIT por nombre/razón social (coincidencia
    parcial) O por código exacto — se prueban las dos búsquedas siempre
    (no hay forma confiable de adivinar cuál tipeó el usuario) y se
    combinan los resultados, sin duplicar por co_cli. El texto es
    obligatorio. Solo lectura — nunca lanza excepción hacia quien la llama.

    Devuelve {"ok": True, "total": N, "clientes": [{"co_cli":.., "razon_social":..}, ...]}
    o {"ok": False, "error": "..."}.
    """
    texto = (texto or "").strip()
    if not texto:
        return {"ok": False, "error": "Hace falta el texto a buscar."}

    limite = max(1, min(50, int(limite or 15)))

    try:
        por_codigo = _buscar_por_codigo(texto)
        por_nombre = _buscar_por_nombre(texto)
    except Exception as error:
        return {"ok": False, "error": f"No se pudo consultar PROFIT: {error}"}

    vistos = set()
    clientes = []
    for candidato in por_codigo + por_nombre:
        if candidato["co_cli"] in vistos:
            continue
        vistos.add(candidato["co_cli"])
        clientes.append(candidato)
        if len(clientes) >= limite:
            break

    return {"ok": True, "total": len(clientes), "clientes": clientes}
