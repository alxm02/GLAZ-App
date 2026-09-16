"""Bruecke zwischen JavaScript und dem Fachkern -- die Schicht der Web-App.

Die Progressive Web App laedt CPython als WebAssembly (Pyodide) ins Handy und
ruft von dort dieselben Module auf, die auch die Desktop-App benutzt:
:mod:`glaz.model`, :mod:`glaz.validation`, :mod:`glaz.excel_engine`,
:mod:`glaz.filename`, :mod:`glaz.mailtext`. Damit gilt auf dem Handy exakt
dieselbe Fachlogik wie am Arbeitsplatz -- es gibt keine zweite, in JavaScript
nachgebaute Wahrheit, die auseinanderlaufen koennte.

Zwei Regeln bestimmen den Zuschnitt dieses Moduls:

**Rein und raus gehen nur JSON-taugliche Werte.** Zwischen Python und
JavaScript liegt eine Typgrenze: ``date``, ``time``, ``Enum`` und Dataclasses
kommen auf der anderen Seite nicht als solche an. Alles, was diese Funktionen
annehmen und liefern, besteht deshalb aus ``dict``, ``list``, ``str``,
``int``, ``bool`` und ``None``.

**Nichts darf werfen, nur weil die Eingabe Unfug ist.** Die Daten kommen aus
dem ``localStorage`` eines Handys. Sie koennen von einer aelteren App-Version
stammen, mitten im Schreiben vom Browser abgeschnitten worden sein oder von
Hand verbogen. Eine Ausnahme aus Pyodide hinterlaesst eine weisse Seite ohne
jede Meldung -- fuer die Nutzerin sieht das aus wie ein kaputtes Programm.
Ein unlesbares Datum wird deshalb ``None``, kein Absturz; die Validierung
sagt danach mit einer verstaendlichen Meldung, was fehlt.

Die Importliste ist bewusst kurz und wird von einem Test bewacht: kein
:mod:`glaz.settings` (``%APPDATA%``), kein :mod:`glaz.session`
(Dateisystem), kein :mod:`glaz.mail` (Outlook/COM), kein :mod:`glaz.ui` (Qt).
Nichts davon existiert im Browser.
"""

from __future__ import annotations

from dataclasses import asdict, fields
from datetime import date, time
from pathlib import Path

from .excel_engine import fill_template, haenge_an, lies_belegung
from .filename import baue_dateinamen
from .mailtext import baue_mail_texte
from .model import AusgabeModus, Profile, TravelRow, TravelType, Vorgang
from .validation import Issue, Severity, ist_absendbar, pruefe_zieldatei, validiere
from .vorgaben import (
    DEFAULT_BETREFF_VORLAGE,
    DEFAULT_BODY_VORLAGE,
    DEFAULT_DATEINAMEN_SCHEMA,
    DEFAULT_EMPFAENGER,
)

#: Die Zeitfelder von :class:`~glaz.model.TravelRow` in der Reihenfolge der
#: Dataclass -- dieselbe Liste wie in :mod:`glaz.session`.
_ZEIT_FELDER = (
    "passiv_von", "passiv_bis",
    "aktiv_von", "aktiv_bis",
    "grenz_anreise", "grenz_rueckreise",
    "arbeit_von", "arbeit_bis",
)

#: Uebersetzung der Schweregrade an die Oberflaeche. Deutsch, weil die
#: JavaScript-Seite sie unveraendert in CSS-Klassen und Texte weiterreicht.
_SCHWERE = {Severity.ERROR: "fehler", Severity.WARNING: "warnung"}


# --- Umwandlung einzelner Werte ----------------------------------------------
#
# Dieselbe Konvention wie in glaz.session (ISO-Datum, "HH:MM" mit Sekunden nur,
# wenn welche gesetzt sind) -- bewusst noch einmal ausgeschrieben statt von
# dort importiert: glaz.session haengt ueber glaz.settings am Dateisystem und
# an ``sys.platform`` und laesst sich im Browser gar nicht laden. Wer hier
# etwas am Format aendert, muss es dort mitaendern; Web- und Desktop-App
# tauschen ihre Staende ueber genau dieses Format aus.


def _als_text(wert: object) -> str:
    """Zwingt einen Wert aus JSON/JavaScript auf ``str``.

    Dataclasses pruefen ihre Feldtypen nicht: ein ``null`` oder eine Zahl an
    Stelle eines Strings laedt sonst anstandslos und stuerzt erst spaeter beim
    ersten ``.strip()`` ab -- dann aber ohne erkennbaren Zusammenhang zur
    eigentlichen Ursache.
    """
    if isinstance(wert, str):
        return wert
    if wert is None:
        return ""
    return str(wert)


def _als_bool(wert: object, default: bool = False) -> bool:
    """Nimmt nur echte Booleans an.

    Kein ``bool(wert)``: der String ``"false"`` waere damit ``True`` -- ein
    Fehler, der beim Weg durch JavaScript und JSON leicht entsteht.
    """
    return wert if isinstance(wert, bool) else default


def _datum_zu_text(d: date | None) -> str | None:
    """``date`` -> ``"2026-09-16"``. ``None`` bleibt ``None`` (JSON: ``null``)."""
    if not isinstance(d, date):
        return None
    return d.isoformat()


def _datum_aus_text(wert: object) -> date | None:
    """Umkehrung von :func:`_datum_zu_text`, unlesbare Werte werden ``None``."""
    if not isinstance(wert, str) or not wert.strip():
        return None
    try:
        return date.fromisoformat(wert.strip())
    except ValueError:
        return None


def _zeit_zu_text(t: time | None) -> str | None:
    """``time`` -> ``"07:45"``, mit Sekunden nur, wenn welche gesetzt sind.

    Die kurze Form ist die, die die Nutzerin auch eingetippt hat; Sekunden
    werden trotzdem mitgeschrieben, falls doch welche im Wert stecken -- sonst
    waere die Rundreise ``Vorgang -> dict -> Vorgang`` stillschweigend ungenau.
    """
    if not isinstance(t, time):
        return None
    if t.second or t.microsecond:
        return t.isoformat()
    return t.strftime("%H:%M")


def _zeit_aus_text(wert: object) -> time | None:
    """Umkehrung von :func:`_zeit_zu_text`, unlesbare Werte werden ``None``."""
    if not isinstance(wert, str) or not wert.strip():
        return None
    try:
        return time.fromisoformat(wert.strip())
    except ValueError:
        return None


# --- Vorgang <-> dict ---------------------------------------------------------


def _zeile_aus_dict(d: object) -> TravelRow:
    """Baut eine ``TravelRow`` aus einem Dict -- unlesbare Felder werden leer.

    Es wird bewusst nie ``None`` geliefert: eine einzelne kaputte Zeile darf
    die Zeilennummerierung der uebrigen nicht verschieben, sonst zeigen die
    Zeilennummern in den Pruefmeldungen auf die falsche Zeile im Formular.
    """
    if not isinstance(d, dict):
        return TravelRow()
    zeiten = {name: _zeit_aus_text(d.get(name)) for name in _ZEIT_FELDER}
    return TravelRow(
        datum=_datum_aus_text(d.get("datum")),
        ist_folgezeile=_als_bool(d.get("ist_folgezeile")),
        **zeiten,
    )


def _zeile_zu_dict(zeile: TravelRow) -> dict:
    """Serialisiert eine ``TravelRow`` -- alle Felder, auch die leeren.

    Leere Felder werden ``None`` und nicht weggelassen: die JavaScript-Seite
    baut ihre Eingabefelder anhand dieser Schluessel auf und muesste sonst
    ihrerseits raten, welche es gibt.
    """
    d: dict = {"datum": _datum_zu_text(zeile.datum)}
    for name in _ZEIT_FELDER:
        d[name] = _zeit_zu_text(getattr(zeile, name))
    d["ist_folgezeile"] = bool(zeile.ist_folgezeile)
    return d


def _profil_aus_dict(d: object) -> Profile:
    """Baut ein ``Profile`` robust aus einem Dict -- unbekannte Keys ignorieren.

    Unbekannte Keys sind der Normalfall und kein Fehler: im ``localStorage``
    kann noch der Stand einer aelteren oder neueren App-Version liegen.
    """
    if not isinstance(d, dict):
        return Profile()
    bekannt = {f.name for f in fields(Profile)}
    gefiltert = {k: _als_text(v) for k, v in d.items() if k in bekannt}
    return Profile(**gefiltert)


def vorgang_aus_dict(d: dict) -> Vorgang:
    """Baut einen :class:`~glaz.model.Vorgang` aus der dict-Form der Web-App.

    Nachgiebig auf allen Ebenen: fehlt ``d`` ganz oder ist es gar kein Dict,
    entsteht ein leerer Vorgang; fehlende Schluessel bekommen die Defaults des
    Modells, unlesbare Werte werden leer. Was danach fachlich nicht reicht,
    meldet :func:`pruefe` -- mit einem Satz, den die Nutzerin lesen kann.

    Zu viele Zeilen werden absichtlich NICHT abgeschnitten: dass die Vorlage
    nur 14 Datenzeilen hat, ist eine fachliche Regel, die die Validierung mit
    einer Meldung durchsetzt. Stilles Abschneiden liesse Reisezeilen
    verschwinden, ohne dass es jemand erfaehrt.
    """
    if not isinstance(d, dict):
        d = {}

    profil = _profil_aus_dict(d.get("profil"))

    try:
        reisetyp = TravelType(d.get("reisetyp"))
    except (ValueError, TypeError):
        # TypeError faengt die unhashbaren Werte ab (``{}``, ``[]``): Enum
        # schlaegt den Wert intern in einem Dict nach und wirft dort, noch
        # bevor es zu einem sauberen ValueError kommt.
        # Unbekannter Reisetyp -> Inland. Das ist der engere der beiden Faelle
        # (Grenzuebertrittszeiten sind dort unzulaessig), also der, bei dem
        # ein Irrtum auffaellt statt stillschweigend durchzugehen.
        reisetyp = TravelType.INLAND

    rohe_zeilen = d.get("zeilen")
    zeilen = (
        [_zeile_aus_dict(z) for z in rohe_zeilen]
        if isinstance(rohe_zeilen, list)
        else []
    )

    return Vorgang(
        profil=profil,
        einsatzart=_als_text(d.get("einsatzart")),
        reisetyp=reisetyp,
        zeilen=zeilen,
        # Die dict-Form kennt nur einen Zielordner, und der steht im Profil.
        # Beide Felder tragen deshalb denselben Wert -- ein halb gefuellter
        # Vorgang, bei dem Profil und Vorgang verschiedene Ordner nennen,
        # waere durch nichts zu erklaeren.
        zielordner=profil.zielordner,
        wiederhole_stammdaten=_als_bool(d.get("wiederhole_stammdaten")),
    )


def dict_aus_vorgang(v: Vorgang) -> dict:
    """Umkehrung von :func:`vorgang_aus_dict` -- die Form fuer den localStorage.

    ``dict_aus_vorgang(vorgang_aus_dict(d)) == d`` gilt fuer jedes vollstaendig
    befuellte ``d``: Die Web-App speichert damit genau den Stand wieder ab, den
    sie geladen hat, auch ueber ein App-Update hinweg.
    """
    return {
        "profil": asdict(v.profil),
        "einsatzart": v.einsatzart,
        "reisetyp": TravelType(v.reisetyp).value,
        "wiederhole_stammdaten": bool(v.wiederhole_stammdaten),
        "zeilen": [_zeile_zu_dict(z) for z in v.zeilen],
    }


# --- Pruefung -----------------------------------------------------------------


def _issue_zu_dict(issue: Issue) -> dict:
    """Ein :class:`~glaz.validation.Issue` in der Form, die die Web-App anzeigt.

    ``severity`` heisst hier ``schwere`` und traegt deutsche Werte: die
    JavaScript-Seite reicht sie unveraendert an CSS-Klassen und Vorlesetexte
    weiter, und ein englisches ``warning`` mitten in einer deutschen
    Oberflaeche waere genau die Art Detail, die spaeter niemand mehr
    zurechtruecken will.
    """
    return {
        "feld": issue.feld,
        "meldung": issue.meldung,
        "schwere": _SCHWERE[issue.severity],
        "zeile": issue.zeile_index,
    }


def _sicherer_dateiname(vorgang: Vorgang) -> str:
    """Dateiname des Vorgangs, oder ``""`` solange er sich nicht bilden laesst.

    :func:`~glaz.filename.baue_dateinamen` wirft, solange keine Zeile ein
    Datum traegt -- also waehrend der gesamten Eingabe. Diese Funktion laeuft
    aber bei jedem Tastendruck, um den Namen in der Oberflaeche anzuzeigen.
    Ein leerer Name bedeutet dort schlicht "noch nicht bekannt"; eine Ausnahme
    bedeutete eine weisse Seite.
    """
    try:
        return baue_dateinamen(vorgang, DEFAULT_DATEINAMEN_SCHEMA)
    except ValueError:
        return ""


def pruefe(d: dict) -> dict:
    """Prueft den Vorgang und liefert alles, was die Oberflaeche daraus zeigt.

    Ein einziger Aufruf statt fuenf: Die Web-App ruft ihn bei jeder Eingabe
    auf, und jeder Uebergang zwischen JavaScript und Pyodide kostet Zeit, die
    man beim Tippen merkt.

    :returns: ``{"issues": [...], "absendbar": bool,
        "hat_arbeit_vor_aktivreise": bool, "zeitraum_text": str,
        "dateiname": str}`` -- jedes Issue in der Form
        ``{"feld", "meldung", "schwere", "zeile"}``.
    """
    vorgang = vorgang_aus_dict(d)
    issues = validiere(vorgang)
    return {
        "issues": [_issue_zu_dict(i) for i in issues],
        "absendbar": ist_absendbar(issues),
        "hat_arbeit_vor_aktivreise": vorgang.hat_arbeit_vor_aktivreise(),
        "zeitraum_text": vorgang.zeitraum_text(),
        "dateiname": _sicherer_dateiname(vorgang),
    }


def pruefe_ziel(d: dict, datei_pfad: str, modus: str) -> dict:
    """Prueft die gewaehlte Zieldatei fuer ``"ergaenzen"``/``"nur_versenden"``.

    Die Zieldatei gehoert nicht zum Vorgang (sie beschreibt die Ausgabe, nicht
    die Reise) und wird deshalb getrennt geprueft -- genau wie in der
    Desktop-App. Aus dem Vorgang kommt nur die Zahl der Zeilen, die noch in
    die Datei passen muessen.

    Ein unbekannter Modus wird wie ``"neu"`` behandelt und meldet nichts: dort
    entsteht die Datei erst, es gibt also nichts zu pruefen.

    :returns: ``{"issues": [...], "absendbar": bool}`` -- dieselbe issues-Form
        wie :func:`pruefe`.
    """
    vorgang = vorgang_aus_dict(d)
    try:
        ausgabe_modus = AusgabeModus(modus)
    except ValueError:
        ausgabe_modus = AusgabeModus.NEU

    issues = pruefe_zieldatei(
        ausgabe_modus, _als_text(datei_pfad), len(vorgang.aktive_zeilen())
    )
    return {
        "issues": [_issue_zu_dict(i) for i in issues],
        "absendbar": ist_absendbar(issues),
    }


# --- Excel --------------------------------------------------------------------


def _datei_info(pfad: Path) -> dict:
    """Pfad, Name und Groesse einer geschriebenen Datei.

    Die Groesse ist kein Schmuck: Im Browser liegt die Datei anschliessend im
    virtuellen Dateisystem von Pyodide und muss von dort als Blob zum
    Herunterladen herausgereicht werden -- dafuer will die JavaScript-Seite
    wissen, wie viele Bytes sie erwartet.
    """
    return {
        "pfad": str(pfad),
        "dateiname": pfad.name,
        "bytes": pfad.stat().st_size,
    }


def erzeuge_xlsx(d: dict, vorlage_pfad: str, ziel_pfad: str) -> dict:
    """Erzeugt eine frische Liste aus der Vorlage (Modus ``NEU``).

    Beide Pfade zeigen im Browser in das virtuelle Dateisystem von Pyodide:
    die Vorlage wurde beim Start dorthin entpackt, das Ergebnis holt die
    JavaScript-Seite anschliessend von dort ab.

    :returns: ``{"pfad", "dateiname", "bytes"}``
    :raises ExcelEngineError: wenn Vorlage oder Ziel nicht lesbar/schreibbar
        sind -- anders als bei den Pruefungen ist das hier eine bewusste
        Handlung der Nutzerin, die eine echte Fehlermeldung verdient.
    """
    ziel = fill_template(
        Path(vorlage_pfad), Path(ziel_pfad), vorgang_aus_dict(d)
    )
    return _datei_info(ziel)


def ergaenze_xlsx(
    d: dict,
    datei_pfad: str,
    start_row: int | None = None,
    ueberschreibe_zeilen: int = 0,
) -> dict:
    """Ergaenzt eine vorhandene Liste um die aktiven Zeilen (Modus ``ERGAENZEN``).

    ``start_row`` und ``ueberschreibe_zeilen`` sind dieselben Parameter wie in
    :func:`glaz.excel_engine.haenge_an` und dienen dem *Aktualisieren*: Mit den
    Werten aus dem Rueckgabewert eines vorangegangenen Laufs laesst sich
    derselbe Block noch einmal schreiben, nachdem im Formular etwas geaendert
    wurde -- statt die Zeilen ein zweites Mal anzuhaengen.

    :returns: ``{"pfad", "dateiname", "bytes", "start_row",
        "geschriebene_zeilen"}`` -- die beiden letzten Werte gehoeren in den
        naechsten Aufruf, wenn derselbe Block aktualisiert werden soll.
    """
    ergebnis = haenge_an(
        Path(datei_pfad),
        vorgang_aus_dict(d),
        start_row=start_row,
        ueberschreibe_zeilen=ueberschreibe_zeilen,
    )
    return {
        **_datei_info(ergebnis.pfad),
        "start_row": ergebnis.start_row,
        "geschriebene_zeilen": ergebnis.geschriebene_zeilen,
    }


def belegung(datei_pfad: str) -> dict:
    """Wie voll der Datenbereich einer vorhandenen Datei schon ist.

    Die Oberflaeche zeigt damit an, wie viele Zeilen noch hineinpassen, bevor
    die Nutzerin eine Datei zum Ergaenzen auswaehlt.

    :returns: ``{"belegte_zeilen", "erste_freie_zeile", "freie_zeilen"}`` --
        ``erste_freie_zeile`` ist ``None``, wenn die Datei voll ist.
    """
    return asdict(lies_belegung(Path(datei_pfad)))


# --- Mail ---------------------------------------------------------------------


def mailtexte(
    d: dict,
    empfaenger: str = "",
    betreff_vorlage: str = "",
    body_vorlage: str = "",
) -> dict:
    """Baut Adressen, Betreff und Text der Mail -- ohne Anhang.

    Leere Angaben fallen auf die Vorgaben aus :mod:`glaz.vorgaben` zurueck.
    Die Web-App kann damit beim ersten Start sofort eine vollstaendige Mail
    anbieten, noch bevor in den Einstellungen irgendetwas hinterlegt ist.

    Den Anhang haengt auf dem Handy die Nutzerin selbst an: Ein Browser darf
    einem ``mailto:``-Link keine Datei mitgeben. Die erzeugte .xlsx wird
    deshalb getrennt heruntergeladen.

    :returns: ``{"an", "cc", "betreff", "body"}``
    :raises MailtextFehler: wenn eine selbst gesetzte Vorlage einen unbekannten
        Platzhalter enthaelt -- die Meldung nennt die erlaubten.
    """
    return baue_mail_texte(
        vorgang_aus_dict(d),
        _als_text(empfaenger) or DEFAULT_EMPFAENGER,
        _als_text(betreff_vorlage) or DEFAULT_BETREFF_VORLAGE,
        _als_text(body_vorlage) or DEFAULT_BODY_VORLAGE,
    )
