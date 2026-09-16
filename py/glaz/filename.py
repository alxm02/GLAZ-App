"""Bildet dateisystemsichere Dateinamen fuer die generierte GLAZ-Liste.

Der Dateiname wird aus Profil- und Vorgangsdaten zusammengesetzt und muss auf
allen Zielsystemen (insbesondere Windows-Fileshares) unproblematisch sein --
daher die Transliteration von Sonderzeichen und das Abfangen reservierter
Windows-Geraetenamen in :func:`slugify`.
"""

from __future__ import annotations

import re
from pathlib import Path

from .model import Vorgang

DEFAULT_SCHEMA = "{datum}_GLAZ-Korrektur_{nachname}_{einsatzart}.xlsx"

#: Deutsche Umlaute/Eszett -> ASCII-Transliteration, vor dem eigentlichen Slugify.
_UMLAUT_MAP = {
    "ä": "ae", "ö": "oe", "ü": "ue", "ß": "ss",
    "Ä": "Ae", "Ö": "Oe", "Ü": "Ue",
}

#: Unter Windows reservierte Geraetenamen -- als Dateiname (ohne Endung) unzulaessig.
_RESERVIERTE_NAMEN = {
    "CON", "PRN", "AUX", "NUL",
    *(f"COM{i}" for i in range(1, 10)),
    *(f"LPT{i}" for i in range(1, 10)),
}

_MAX_LAENGE = 60

#: Erlaubte Platzhalter im Dateinamen-Schema. Einzige Quelle -- auch fuer den
#: Text, den die Anwenderin bei einem kaputten Schema zu sehen bekommt.
_PLATZHALTER = ("datum", "nachname", "einsatzart")

#: Unter Windows in Dateinamen verbotene Zeichen sowie Steuerzeichen.
#: ``/`` und ``\`` fehlen bewusst -- die trennen Pfadanteile ab und werden
#: in :func:`_saeubere_dateinamen` frueher behandelt.
_VERBOTENE_ZEICHEN_RE = re.compile(r'[<>:"|?*\x00-\x1f]')


def slugify(text: str) -> str:
    """Wandelt beliebigen Text in ein dateisystemsicheres Fragment.

    Umlaute werden transliteriert, alles ausser ``[A-Za-z0-9-]`` wird zu einem
    Bindestrich, Mehrfach-Bindestriche werden zusammengefasst und das Ergebnis
    wird auf 60 Zeichen gekuerzt. Reservierte Windows-Namen (CON, PRN, ...)
    werden durch Anhaengen eines Bindestrichs entschaerft.
    """
    for umlaut, ersatz in _UMLAUT_MAP.items():
        text = text.replace(umlaut, ersatz)

    text = re.sub(r"[^A-Za-z0-9-]+", "-", text)
    text = re.sub(r"-{2,}", "-", text)
    text = text.strip("-")
    # Nach dem Kuerzen kann erneut ein Bindestrich am Rand entstehen.
    text = text[:_MAX_LAENGE].strip("-")

    if not text:
        text = "unbenannt"

    if text.upper() in _RESERVIERTE_NAMEN:
        text = f"{text}-"

    return text


def _saeubere_dateinamen(name: str) -> str:
    """Macht aus einem Format-Ergebnis einen reinen, sicheren Dateinamen.

    :func:`slugify` laeuft nur ueber die eingesetzten Fragmente -- der
    Schema-Rahmen selbst kommt ungeprueft aus settings.json. Ein Schema mit
    Pfadanteil (``..\\``) legte die Datei sonst ausserhalb des gewaehlten
    Zielordners ab, und :func:`eindeutiger_pfad` pruefte die Kollision im
    falschen Verzeichnis -- eine dortige Datei waere ohne Rueckfrage
    ueberschrieben worden.
    """
    name = re.split(r"[\\/]", name)[-1]
    name = _VERBOTENE_ZEICHEN_RE.sub("-", name)

    if name.lower().endswith(".xlsx"):
        name = name[: -len(".xlsx")]
    # Windows schneidet Punkte und Leerzeichen am Namensende stillschweigend ab.
    name = name.strip().strip(".").strip()
    name = name or "unbenannt"

    # Dieselbe Entschaerfung wie in slugify, aber eine Ebene hoeher: slugify
    # laeuft nur ueber die eingesetzten Fragmente, der Geraetename kann auch
    # im Schema-Rahmen stehen ("NUL.xlsx" aus settings.json). Ein Schreiben
    # nach NUL laeuft scheinbar durch, die Datei existiert danach aber nicht
    # -- und die Mail bekaeme einen Anhang, der ins Leere zeigt.
    if name.upper() in _RESERVIERTE_NAMEN:
        name = f"{name}-"

    return f"{name}.xlsx"


def baue_dateinamen(vorgang: Vorgang, schema: str = DEFAULT_SCHEMA) -> str:
    """Baut den Dateinamen aus dem Schema. ``{datum}`` ist der erste Reisetag."""
    erster_tag, _ = vorgang.zeitraum()
    if erster_tag is None:
        raise ValueError(
            "Vorgang enthaelt keine Zeile mit Datum -- Dateiname kann nicht "
            "gebildet werden."
        )

    erlaubt = ", ".join("{%s}" % p for p in _PLATZHALTER)
    try:
        name = schema.format(
            datum=erster_tag.strftime("%Y-%m-%d"),
            nachname=slugify(vorgang.profil.nachname),
            einsatzart=slugify(vorgang.einsatzart),
        )
    except KeyError as exc:
        # Das Schema steht in settings.json und kann dort haendisch oder von
        # einer anderen App-Version gesetzt worden sein. Roh durchgereicht
        # schluege der Fehler bis in die Qt-Eventschleife durch, denn die
        # Aufrufer fangen nur ValueError.
        raise ValueError(
            f"Das Dateinamen-Schema enthaelt den unbekannten Platzhalter "
            f"'{{{exc.args[0]}}}'. Erlaubt sind: {erlaubt}."
        ) from exc
    except (IndexError, ValueError) as exc:
        raise ValueError(
            f"Das Dateinamen-Schema ist ungueltig formatiert ({exc}). "
            f"Erlaubt sind die Platzhalter: {erlaubt}."
        ) from exc

    return _saeubere_dateinamen(name)


def eindeutiger_pfad(ordner: Path, dateiname: str) -> Path:
    """Liefert einen freien Pfad in ``ordner``, ueberschreibt nie eine bestehende Datei.

    Existiert ``dateiname`` bereits, wird vor der Endung ``_2``, ``_3`` usw.
    angehaengt, bis ein freier Pfad gefunden ist.
    """
    pfad = ordner / dateiname
    if not pfad.exists():
        return pfad

    stem = pfad.stem
    endung = pfad.suffix
    zaehler = 2
    while True:
        kandidat = ordner / f"{stem}_{zaehler}{endung}"
        if not kandidat.exists():
            return kandidat
        zaehler += 1
