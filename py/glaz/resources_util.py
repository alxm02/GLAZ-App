"""Pfadauflösung für gebündelte Ressourcen (z. B. die Excel-Vorlage).

Funktioniert sowohl im Entwicklungsbetrieb als auch nach dem Einfrieren mit
PyInstaller (``--onefile``): PyInstaller entpackt sich zur Laufzeit in einen
temporären Ordner und hinterlegt dessen Pfad in ``sys._MEIPASS``.
"""

from __future__ import annotations

import sys
from pathlib import Path


def ressource(name: str) -> Path:
    """Liefert den absoluten Pfad zu einer Datei unter ``glaz/resources/``."""
    basis = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent))
    return basis / "resources" / name
