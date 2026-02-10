"""
Módulo de compatibilidad para mantener el antiguo namespace `v1.descargar_archivos`.
Toda la lógica real vive ahora en `core.descargar_archivos`.
"""
from core.descargar_archivos import *  # noqa: F401,F403
