"""Fachliche Pruefung eines :class:`~glaz.model.Vorgang` vor dem Absenden.

Prueft Profil-Stammdaten, Zeitlogik je Zeile und Zeilenanzahl gegen die
Fachspezifikation der GLAZ-Korrekturbuchungsliste. Liefert eine flache Liste
von :class:`Issue` -- die UI entscheidet selbst, wie sie ERROR/WARNING anzeigt.

Die Zieldatei der Ausgabemodi ``ERGAENZEN``/``NUR_VERSENDEN`` gehoert nicht
zum ``Vorgang`` (sie beschreibt die Ausgabe, nicht die Reise) und wird
deshalb getrennt ueber :func:`pruefe_zieldatei` geprueft.
"""

from __future__ import annotations

from dataclasses import dataclass, replace
from enum import Enum
from pathlib import Path

from .model import (
    MAX_ROWS,
    AusgabeModus,
    TravelRow,
    TravelType,
    Vorgang,
    ist_personalnummer,
)

#: Feldname aller Zieldatei-Meldungen -- die UI haengt daran ihr Fehlerlabel.
FELD_ZIELDATEI = "abschluss.zieldatei"

#: Pflichtfelder im Profil, Attributname -> Anzeigename fuer die Meldung.
_PROFIL_PFLICHTFELDER = {
    "personalnummer": "Personalnummer",
    "vorname": "Vorname",
    "nachname": "Nachname",
    "abteilung": "Abteilung",
    "gruppenleiter_name": "Gruppenleiter (Name)",
    "gruppenleiter_email": "Gruppenleiter (E-Mail)",
}

#: Die vier Zeitpaare einer Zeile: (Attributpraefix, Anzeigename).
_ZEITPAARE = [
    ("passiv", "passive Reisezeit"),
    ("aktiv", "aktive Reisezeit"),
    ("arbeit", "Arbeitszeit"),
]


class Severity(Enum):
    ERROR = "error"      # blockiert den Absende-Button
    WARNING = "warning"  # nur Hinweis, blockiert nicht


@dataclass(frozen=True)
class Issue:
    feld: str        # z. B. "zeile[2].arbeit_bis" oder "profil.gruppenleiter_email"
    meldung: str      # deutscher, handlungsleitender Klartext
    severity: Severity
    zeile_index: int | None = None   # Index in vorgang.zeilen, sonst None


def _ist_zeilenfeld(feld: str) -> bool:
    """True fuer alles, was an den erfassten Reisezeilen haengt.

    Beides zaehlt dazu: die Meldungen einzelner Zeilen (``zeile[2].datum``)
    und die der Zeilenmenge (``vorgang.zeilen`` -- "keine Reisezeile
    erfasst"). Gerade Letztere ist der Regelfall dessen, was
    :func:`entschaerfe_zeilen_fuer_versand` durchlassen soll.
    """
    return feld.startswith("zeile[") or feld == "vorgang.zeilen"


def _pruefe_profil(vorgang: Vorgang) -> list[Issue]:
    issues: list[Issue] = []
    profil = vorgang.profil

    for attribut, anzeigename in _PROFIL_PFLICHTFELDER.items():
        if not getattr(profil, attribut).strip():
            issues.append(
                Issue(
                    feld=f"profil.{attribut}",
                    meldung=f"{anzeigename} fehlt. Bitte im Profil ausfüllen.",
                    severity=Severity.ERROR,
                )
            )

    # Nicht str.isdigit(): das haelt auch Hochzahlen und arabisch-indische
    # Ziffern fuer Ziffern und liesse eine Excel-Datei entstehen, die Excel
    # nicht mehr oeffnen kann.
    if profil.personalnummer.strip() and not ist_personalnummer(profil.personalnummer):
        issues.append(
            Issue(
                feld="profil.personalnummer",
                meldung="Personalnummer darf nur Ziffern enthalten.",
                severity=Severity.ERROR,
            )
        )

    if profil.gruppenleiter_email.strip() and not profil.email_gueltig():
        issues.append(
            Issue(
                feld="profil.gruppenleiter_email",
                meldung="E-Mail-Adresse des Gruppenleiters ist ungültig.",
                severity=Severity.ERROR,
            )
        )

    return issues


def _pruefe_vorgang_kopf(vorgang: Vorgang) -> list[Issue]:
    issues: list[Issue] = []

    if not vorgang.einsatzart.strip():
        issues.append(
            Issue(
                feld="vorgang.einsatzart",
                meldung="Einsatzart fehlt. Bitte auswählen.",
                severity=Severity.ERROR,
            )
        )

    return issues


def _pruefe_zeitpaar(
    zeile: TravelRow, praefix: str, anzeigename: str, zeile_index: int
) -> list[Issue]:
    """Prueft `von`/`bis` eines Zeitpaars: beide gesetzt und von < bis."""
    von = getattr(zeile, f"{praefix}_von")
    bis = getattr(zeile, f"{praefix}_bis")

    if von is None and bis is None:
        return []

    if von is None or bis is None:
        fehlend = f"{praefix}_von" if von is None else f"{praefix}_bis"
        return [
            Issue(
                feld=f"zeile[{zeile_index}].{fehlend}",
                meldung=f"Zeile {zeile_index + 1}: Gegenstück zu {anzeigename} fehlt.",
                severity=Severity.ERROR,
                zeile_index=zeile_index,
            )
        ]

    if von >= bis:
        # Eine Zeile der Vorlage traegt genau ein Datum, Uhrzeiten sind reine
        # Tagesbruchteile -- ein Intervall ueber Mitternacht laesst sich darin
        # nicht abbilden. Statt die Eingabe nur abzulehnen, nennt die Meldung
        # den einzigen Ausweg: aufteilen auf zwei Zeilen.
        if bis < von:
            hinweis = (
                "Zeiten über Mitternacht bitte auf zwei Zeilen aufteilen "
                "(bis 23:59 und ab 00:00 am Folgetag)."
            )
        else:
            hinweis = "Ende muss nach dem Beginn liegen."
        return [
            Issue(
                feld=f"zeile[{zeile_index}].{praefix}_bis",
                meldung=f"Zeile {zeile_index + 1}: {anzeigename} -- {hinweis}",
                severity=Severity.ERROR,
                zeile_index=zeile_index,
            )
        ]

    return []


def _intervall_ueberschneidung(
    a_von, a_bis, b_von, b_bis
) -> bool:
    """True, wenn sich zwei Zeitintervalle echt schneiden (Schnitt > 0).

    Der direkte Vergleich auf ``time`` genuegt, weil ein Intervall ueber
    Mitternacht bereits in :func:`_pruefe_zeitpaar` als ERROR haengen bleibt
    und hier deshalb nie ankommt.

    ``<`` statt ``<=`` ist tragend, nicht nachlaessig: eine nahtlose Abfolge
    (Arbeit bis 12:00, aktive Reise ab 12:00) beruehrt sich nur in einem Punkt
    und ist der Normalfall "erst arbeiten, dann losfahren" -- siehe
    :meth:`~glaz.model.TravelRow.arbeit_vor_aktivreise`. Wer hier auf ``<=``
    verschaerft, warnt bei jedem solchen Tag.
    """
    if None in (a_von, a_bis, b_von, b_bis):
        return False
    spaeter_von = max(a_von, b_von)
    frueher_bis = min(a_bis, b_bis)
    return spaeter_von < frueher_bis


def _pruefe_zeile(
    vorgang: Vorgang, zeile: TravelRow, zeile_index: int
) -> list[Issue]:
    issues: list[Issue] = []

    for praefix, anzeigename in _ZEITPAARE:
        issues.extend(_pruefe_zeitpaar(zeile, praefix, anzeigename, zeile_index))

    # Datum ist Pflicht, sobald Zeiten erfasst sind -- ausser in Folgezeilen,
    # die laut Referenzdokument bewusst kein eigenes Datum tragen.
    if zeile.datum is None and zeile.hat_zeiten() and not zeile.ist_folgezeile:
        issues.append(
            Issue(
                feld=f"zeile[{zeile_index}].datum",
                meldung=f"Zeile {zeile_index + 1}: Datum fehlt.",
                severity=Severity.ERROR,
                zeile_index=zeile_index,
            )
        )

    if vorgang.reisetyp == TravelType.INLAND and (
        zeile.grenz_anreise is not None or zeile.grenz_rueckreise is not None
    ):
        issues.append(
            Issue(
                feld=f"zeile[{zeile_index}].grenz_anreise",
                meldung=(
                    f"Zeile {zeile_index + 1}: Grenzübertrittszeiten sind bei "
                    "Reisetyp Inland nicht zulässig."
                ),
                severity=Severity.ERROR,
                zeile_index=zeile_index,
            )
        )

    # Ueberschneidung von Arbeitszeit mit Reisezeit ist plausibel (z. B. Arbeit
    # waehrend der Zugfahrt), aber pruefenswert -- daher nur WARNING.
    if _intervall_ueberschneidung(
        zeile.arbeit_von, zeile.arbeit_bis, zeile.passiv_von, zeile.passiv_bis
    ):
        issues.append(
            Issue(
                feld=f"zeile[{zeile_index}].arbeit_von",
                meldung=(
                    f"Zeile {zeile_index + 1}: Arbeitszeit überschneidet sich "
                    "mit der passiven Reisezeit."
                ),
                severity=Severity.WARNING,
                zeile_index=zeile_index,
            )
        )

    if _intervall_ueberschneidung(
        zeile.arbeit_von, zeile.arbeit_bis, zeile.aktiv_von, zeile.aktiv_bis
    ):
        issues.append(
            Issue(
                feld=f"zeile[{zeile_index}].arbeit_von",
                meldung=(
                    f"Zeile {zeile_index + 1}: Arbeitszeit überschneidet sich "
                    "mit der aktiven Reisezeit."
                ),
                severity=Severity.WARNING,
                zeile_index=zeile_index,
            )
        )

    # Anders als bei der Arbeitszeit gibt es hier keinen plausiblen Fall:
    # passiv befoerdert werden und gleichzeitig selbst fahren geht nicht.
    # Unbemerkt zaehlt die GLAZ-Auswertung die Stunden doppelt -- daher ERROR.
    if _intervall_ueberschneidung(
        zeile.passiv_von, zeile.passiv_bis, zeile.aktiv_von, zeile.aktiv_bis
    ):
        issues.append(
            Issue(
                feld=f"zeile[{zeile_index}].aktiv_von",
                meldung=(
                    f"Zeile {zeile_index + 1}: passive und aktive Reisezeit "
                    "überschneiden sich."
                ),
                severity=Severity.ERROR,
                zeile_index=zeile_index,
            )
        )

    return issues


def _pruefe_zeilen(vorgang: Vorgang) -> list[Issue]:
    issues: list[Issue] = []
    aktive = vorgang.aktive_zeilen()

    if not aktive:
        issues.append(
            Issue(
                feld="vorgang.zeilen",
                meldung="Es ist keine Reisezeile erfasst.",
                severity=Severity.ERROR,
            )
        )

    if len(aktive) > MAX_ROWS:
        issues.append(
            Issue(
                feld="vorgang.zeilen",
                meldung=(
                    f"Es sind {len(aktive)} Zeilen erfasst, die Vorlage fasst "
                    f"maximal {MAX_ROWS}."
                ),
                severity=Severity.ERROR,
            )
        )

    for index, zeile in enumerate(vorgang.zeilen):
        if zeile.ist_leer():
            continue
        issues.extend(_pruefe_zeile(vorgang, zeile, index))

    # Eine Folgezeile setzt eine vorangehende Hauptzeile voraus -- die erste
    # aktive Zeile als Folgezeile ist daher fachlich unplausibel.
    if aktive and aktive[0].ist_folgezeile:
        erste_index = next(
            i for i, z in enumerate(vorgang.zeilen) if not z.ist_leer()
        )
        issues.append(
            Issue(
                feld=f"zeile[{erste_index}].ist_folgezeile",
                meldung=(
                    f"Zeile {erste_index + 1}: als Folgezeile markiert, aber es "
                    "geht keine Hauptzeile voraus."
                ),
                severity=Severity.WARNING,
                zeile_index=erste_index,
            )
        )

    return issues


def validiere(vorgang: Vorgang) -> list[Issue]:
    """Prueft den gesamten Vorgang und liefert alle gefundenen Probleme."""
    issues: list[Issue] = []
    issues.extend(_pruefe_profil(vorgang))
    issues.extend(_pruefe_vorgang_kopf(vorgang))
    issues.extend(_pruefe_zeilen(vorgang))
    return issues


def _zieldatei_issue(meldung: str) -> list[Issue]:
    """Kurzform fuer die immer gleiche Huelle der Zieldatei-Meldungen."""
    return [Issue(feld=FELD_ZIELDATEI, meldung=meldung, severity=Severity.ERROR)]


def pruefe_zieldatei(
    modus: AusgabeModus, pfad: str, benoetigte_zeilen: int
) -> list[Issue]:
    """Prueft die gewaehlte Zieldatei fuer ``ERGAENZEN``/``NUR_VERSENDEN``.

    Fuer ``NEU`` immer leer: dort entsteht die Datei erst, geprueft wird
    stattdessen der Zielordner.

    Die Funktion haengt bei ``ERGAENZEN`` an der Excel-Engine, laeuft aber im
    GUI-Thread bei jedem Tastendruck. Sie faengt deshalb jeden
    ``ExcelEngineError`` ab und uebersetzt ihn in ein Issue -- eine
    durchschlagende Ausnahme wuerde die Eingabe zum Absturz bringen.
    """
    if modus is AusgabeModus.NEU:
        return []

    pfad = pfad.strip()
    if not pfad:
        return _zieldatei_issue("Bitte eine vorhandene Excel-Datei auswählen.")

    ziel = Path(pfad)
    # Verzeichnis vor Existenz: ein Ordner "existiert" auch, und "Datei nicht
    # gefunden" waere dann eine Meldung, die vor dem sichtbar vorhandenen
    # Eintrag im Dateidialog ratlos macht.
    if ziel.is_dir():
        return _zieldatei_issue(
            f"Das ist ein Ordner, keine Excel-Datei: {pfad}"
        )
    if not ziel.is_file():
        return _zieldatei_issue(f"Datei nicht gefunden: {pfad}")

    # Nur .xlsx: Die Engine patcht das OOXML-Paket der Datei an Ort und Stelle
    # (siehe Modul-Docstring von excel_engine). .xls ist gar kein OOXML,
    # sondern ein altes Binaerformat; .xlsm ist zwar eines, hat aber weder
    # denselben Blattaufbau der Vorlage noch Makros, die ein Patch heil laesst.
    if ziel.suffix.lower() != ".xlsx":
        return _zieldatei_issue(
            "Nur .xlsx-Dateien werden unterstützt. Bitte eine Excel-Arbeitsmappe "
            "auswählen."
        )

    if modus is not AusgabeModus.ERGAENZEN:
        # NUR_VERSENDEN fasst die Datei nicht an -- wie voll sie ist, geht die
        # Pruefung nichts an.
        return []

    # Erst hier importiert: validation wird von Modulen ohne Excel-Bezug
    # geladen, und der Import der Engine kostet mehr als diese Pruefung.
    from .excel_engine import ExcelEngineError, lies_belegung

    try:
        belegung = lies_belegung(ziel)
    except ExcelEngineError as exc:
        return _zieldatei_issue(f"Die Datei lässt sich nicht lesen: {exc}")

    if belegung.freie_zeilen < benoetigte_zeilen:
        return _zieldatei_issue(
            f"In der Datei sind noch {belegung.freie_zeilen} Zeilen frei, "
            f"benötigt werden {benoetigte_zeilen}."
        )

    return []


def entschaerfe_zeilen_fuer_versand(issues: list[Issue]) -> list[Issue]:
    """Stuft die Zeilen-Fehler auf Warnungen herab -- nur fuer ``NUR_VERSENDEN``.

    In diesem Modus wird keine Zeile in eine Datei geschrieben -- eine
    unvollstaendige Zeile kann also nichts kaputt machen, sie beeinflusst nur
    den Zeitraum im Betreff. Wer eine vorab ausgefuellte Datei nachtraeglich
    verschickt, soll dafuer nicht das ganze Formular neu tippen muessen.

    Profil-Issues bleiben Fehler: sie speisen CC und Betreff der Mail, die
    hier ja gerade entsteht. Warnungen bleiben Warnungen -- herabgestuft wird,
    nie herauf.
    """
    return [
        replace(issue, severity=Severity.WARNING)
        if _ist_zeilenfeld(issue.feld) and issue.severity is Severity.ERROR
        else issue
        for issue in issues
    ]


def ist_absendbar(issues: list[Issue]) -> bool:
    """True, wenn keine der Meldungen ein ERROR ist."""
    return not any(issue.severity is Severity.ERROR for issue in issues)
