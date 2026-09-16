"""Umrechnung zwischen Python-Datentypen und Excel-Serialwerten.

Excel speichert Datumsangaben als Tage seit dem 30.12.1899 (Windows-Epoche) und
Uhrzeiten als Tagesbruchteil. Beim Serialisieren nach XML verwendet Excel 17
signifikante Stellen (``%.17g``) -- Pythons ``repr()`` erzeugt teilweise kuerzere
Darstellungen, die zwar denselben Double bezeichnen, aber nicht byte-gleich sind.
"""

from __future__ import annotations

from datetime import date, time

# Excel-Epoche: der beruehmte Off-by-two aus der Lotus-1-2-3-Kompatibilitaet
# (Excel behandelt 1900 faelschlich als Schaltjahr).
EXCEL_EPOCH = date(1899, 12, 30)

SECONDS_PER_DAY = 86400


def date_to_serial(d: date) -> int:
    """Wandelt ein Datum in das Excel-Datumsserial (z. B. 28.08.2026 -> 46262)."""
    if not isinstance(d, date):
        raise TypeError(f"date erwartet, nicht {type(d).__name__}")
    if d < EXCEL_EPOCH:
        raise ValueError(f"Datum {d.isoformat()} liegt vor der Excel-Epoche")
    return (d - EXCEL_EPOCH).days


def serial_to_date(serial: int) -> date:
    """Umkehrung von :func:`date_to_serial`."""
    from datetime import timedelta

    return EXCEL_EPOCH + timedelta(days=int(serial))


def time_to_fraction(t: time) -> float:
    """Wandelt eine Uhrzeit in den Excel-Tagesbruchteil (06:45 -> 0.28125)."""
    if not isinstance(t, time):
        raise TypeError(f"time erwartet, nicht {type(t).__name__}")
    seconds = t.hour * 3600 + t.minute * 60 + t.second
    return seconds / SECONDS_PER_DAY


def fraction_to_time(fraction: float) -> time:
    """Umkehrung von :func:`time_to_fraction` (auf Sekunden gerundet)."""
    if not 0.0 <= fraction < 1.0:
        raise ValueError(f"Tagesbruchteil ausserhalb [0,1): {fraction}")
    total = round(fraction * SECONDS_PER_DAY)
    return time(total // 3600, (total % 3600) // 60, total % 60)


def format_number(value: float | int) -> str:
    """Serialisiert eine Zahl so, wie Excel sie in die XML schreibt.

    Ganzzahlen ohne Nachkommastellen werden als Integer geschrieben, alles
    andere mit ``%.17g``. Das reproduziert z. B. ``0.67708333333333337``
    fuer 16:15 -- exakt die Schreibweise des Referenzdokuments.
    """
    if isinstance(value, int) or (isinstance(value, float) and value.is_integer()):
        return str(int(value))
    return "%.17g" % value
