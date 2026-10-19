# -*- coding: utf-8 -*-
"""
Logo ENSA - Carga desde archivo
"""
import base64
from pathlib import Path

LOGO_FILE = "logo_ensa.png"

def logo_base64(filepath: str) -> str:
    """Convierte imagen a base64 para embeber en HTML."""
    path = Path(filepath)
    if not path.exists():
        return ""
    with open(path, "rb") as f:
        return "data:image/png;base64," + base64.b64encode(f.read()).decode()

LOGO_URL = logo_base64(LOGO_FILE)