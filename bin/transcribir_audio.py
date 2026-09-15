# -*- coding: utf-8 -*-
"""
transcribir_audio.py — Convierte un audio de pedido hablado (nota de voz)
a texto, para que Pedidos OCR pueda extraer artículos/cantidades igual
que de una foto o un PDF (11/09, a pedido explícito del usuario).

Pipeline: torchaudio (carga y decodifica el audio — a pedido explícito
del usuario: "con la librería de torchaudio"; soporta cualquier formato
real que llegue, .opus/.ogg de WhatsApp incluido, vía su backend de
FFmpeg) -> resample a 16kHz mono (lo que espera Whisper) -> Whisper
(modelo de voz-a-texto de OpenAI, forzado a español).

OJO de instalación en Windows: desde torchaudio 2.x, `torchaudio.load`
decodifica a través de TorchCodec, que necesita las DLLs de FFmpeg
"shared" (no el build "full" estático que se usa para todo lo demás en
este proyecto — ese no sirve acá). Se descargó/instaló aparte en
C:\\tools\\ffmpeg-shared (build shared de gyan.dev) — la carpeta bin se
agrega al PATH del proceso más abajo, sin tocar el PATH del sistema.

Este archivo NO decide qué hacer con el texto transcrito (eso es
ocr_pedidos_core.procesar_audio, que le pasa el texto al modelo de
extracción de artículos) — solo entrega la transcripción.

Uso como módulo:
    from bin.transcribir_audio import transcribir_audio
    resultado = transcribir_audio(audio_bytes, "nota_voz.opus")
    if resultado["ok"]:
        resultado["texto"]

Uso como CLI (para probarlo suelto, sin levantar ningún servidor):
    python bin/transcribir_audio.py --archivo nota_voz.opus
"""
import argparse
import os
import tempfile

# Tiene que pasar ANTES de importar torchaudio (busca las DLLs de FFmpeg
# al cargar el backend de TorchCodec). Ver nota de instalación arriba.
_CARPETA_FFMPEG_SHARED = os.environ.get(
    "FFMPEG_SHARED_BIN_DIR",
    r"C:\tools\ffmpeg-shared\ffmpeg-9.0.1-full_build-shared\bin",
)
if os.path.isdir(_CARPETA_FFMPEG_SHARED) and _CARPETA_FFMPEG_SHARED not in os.environ.get("PATH", ""):
    os.environ["PATH"] = _CARPETA_FFMPEG_SHARED + os.pathsep + os.environ.get("PATH", "")

import torch
import torchaudio
import whisper

MODELO_WHISPER = os.environ.get("MODELO_WHISPER", "small")
IDIOMA_AUDIO = os.environ.get("IDIOMA_AUDIO_PEDIDOS", "es")
TASA_MUESTREO_WHISPER = 16000

# Se cargan una sola vez (perezoso, en el primer audio) — cargar el
# modelo de Whisper toma varios segundos, no tiene sentido repetirlo por
# cada nota de voz que llegue mientras el servidor sigue corriendo.
_modelo_whisper = None
_resampler_por_tasa: dict = {}


class ErrorTranscripcion(Exception):
    pass


def _obtener_modelo():
    global _modelo_whisper
    if _modelo_whisper is None:
        _modelo_whisper = whisper.load_model(MODELO_WHISPER)
    return _modelo_whisper


def _resamplear(forma_onda, tasa_origen):
    if tasa_origen == TASA_MUESTREO_WHISPER:
        return forma_onda
    if tasa_origen not in _resampler_por_tasa:
        _resampler_por_tasa[tasa_origen] = torchaudio.transforms.Resample(orig_freq=tasa_origen, new_freq=TASA_MUESTREO_WHISPER)
    return _resampler_por_tasa[tasa_origen](forma_onda)


def transcribir_audio(audio_bytes, nombre_archivo="audio") -> dict:
    """Devuelve {"ok": True, "texto": "..."} o {"ok": False, "error": "..."}.
    Nunca lanza excepción hacia quien la llama."""
    extension = (nombre_archivo.rsplit(".", 1)[-1] if "." in nombre_archivo else "bin").lower()
    with tempfile.TemporaryDirectory() as carpeta_temp:
        ruta_entrada = os.path.join(carpeta_temp, f"entrada.{extension}")
        try:
            with open(ruta_entrada, "wb") as f:
                f.write(audio_bytes)

            # torchaudio decodifica el audio ORIGINAL directo (cualquier
            # formato: .opus/.ogg/.m4a/.mp3/.wav) — no hace falta
            # convertirlo a mano antes.
            forma_onda, tasa_muestreo = torchaudio.load(ruta_entrada)
            if forma_onda.shape[0] > 1:
                forma_onda = forma_onda.mean(dim=0, keepdim=True)
            forma_onda = _resamplear(forma_onda, tasa_muestreo)
            audio_np = forma_onda.squeeze(0).to(torch.float32).numpy()

            modelo = _obtener_modelo()
            resultado = modelo.transcribe(audio_np, language=IDIOMA_AUDIO, fp16=False)
            texto = (resultado.get("text") or "").strip()
        except Exception as error:
            return {"ok": False, "error": f"No se pudo transcribir el audio: {error}"}

    if not texto:
        return {"ok": False, "error": "El audio no tiene voz reconocible (transcripción vacía)."}
    return {"ok": True, "texto": texto}


def _main() -> None:
    parser = argparse.ArgumentParser(description="Transcribe un audio de pedido hablado a texto (torchaudio + Whisper).")
    parser.add_argument("--archivo", required=True, help="Ruta del archivo de audio a transcribir.")
    args = parser.parse_args()
    with open(args.archivo, "rb") as f:
        audio_bytes = f.read()
    resultado = transcribir_audio(audio_bytes, os.path.basename(args.archivo))
    if resultado["ok"]:
        print(resultado["texto"])
    else:
        print(f"Error: {resultado['error']}")


if __name__ == "__main__":
    _main()
