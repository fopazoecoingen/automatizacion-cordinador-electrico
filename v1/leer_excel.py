"""
Módulo de compatibilidad para mantener el antiguo namespace `v1.leer_excel`.
Toda la lógica real vive ahora en `core.leer_excel`.
"""
from core.leer_excel import *  # noqa: F401,F403
