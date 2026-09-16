"""Datenmodell der GLAZ-Korrekturbuchungsliste.

Die Struktur bildet die Excel-Vorlage ab: ein Kopfbereich (Profil + Vorgang)
und maximal 14 Datenzeilen (Zeilen 7-20 im Blatt ``Tabelle1``).
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field
from datetime import date, time
from enum import Enum

#: Erste Datenzeile im Excel-Blatt.
FIRST_DATA_ROW = 7
#: Letzte Datenzeile im Excel-Blatt.
LAST_DATA_ROW = 20
#: Harte Obergrenze an Datenzeilen, die die Vorlage hergibt.
MAX_ROWS = LAST_DATA_ROW - FIRST_DATA_ROW + 1  # 14

EMAIL_RE = re.compile(r"^[^@\s]+@[^@\s]+\.[A-Za-z]{2,}$")

#: Personalnummern bestehen ausschliesslich aus ASCII-Ziffern.
PERSONALNUMMER_RE = re.compile(r"[0-9]+")


def ist_personalnummer(text: str) -> bool:
    """True, wenn ``text`` nur aus ASCII-Ziffern besteht.

    Gemeinsame Quelle fuer Validierung und Excel-Erzeugung, damit beide nicht
    auseinanderlaufen. ``str.isdigit()`` genuegt hier nicht: es liefert auch
    fuer Hochzahlen und arabisch-indische Ziffern True. Solche Werte landeten
    als Zahlzelle in der XLSX und machten die erzeugte Datei unlesbar.
    """
    return bool(PERSONALNUMMER_RE.fullmatch(text.strip()))


class TravelType(str, Enum):
    """Reisetyp. Steuert, ob Grenzuebertrittszeiten zulaessig sind."""

    INLAND = "inland"
    AUSLAND = "ausland"


class AusgabeModus(str, Enum):
    """Was mit dem Vorgang beim Abschluss geschehen soll.

    Trennt bewusst die *Datei*-Frage von der *Mail*-Frage: Ob eine Mail
    entsteht, entscheidet weiterhin der gedrueckte Knopf. Dieser Modus
    entscheidet nur, woher die Zieldatei kommt.

    ``NEU``
        Eine frische Datei aus der mitgelieferten Vorlage -- das bisherige
        und weiterhin voreingestellte Verhalten.
    ``ERGAENZEN``
        Eine bereits vorhandene Datei wird um die erfassten Reisezeilen
        erweitert. So laesst sich eine Liste ueber mehrere Sitzungen hinweg
        vorab befuellen und erst spaeter abschicken.
    ``NUR_VERSENDEN``
        Eine vorhandene Datei wird unveraendert versendet. Der Gegenpart zu
        ``ERGAENZEN``: erst vorab ausfuellen, dann zu einem spaeteren
        Zeitpunkt verschicken, ohne die Datei noch einmal anzufassen.
    """

    NEU = "neu"
    ERGAENZEN = "ergaenzen"
    NUR_VERSENDEN = "nur_versenden"

    def braucht_zieldatei(self) -> bool:
        """True, wenn der Modus eine vorhandene Datei als Eingabe braucht."""
        return self is not AusgabeModus.NEU

    def schreibt_zeilen(self) -> bool:
        """True, wenn der Modus Reisezeilen in eine Datei schreibt."""
        return self is not AusgabeModus.NUR_VERSENDEN


@dataclass
class Profile:
    """Personenbezogene Stammdaten. Werden lokal gespeichert, nie ins Repo."""

    name: str = ""                 # Anzeigename des Profils (Dropdown)
    personalnummer: str = ""
    vorname: str = ""
    nachname: str = ""
    abteilung: str = ""            # Org-Kuerzel, z. B. "XX YY ZZZ QQQ"
    gruppenleiter_name: str = ""
    gruppenleiter_email: str = ""
    zielordner: str = ""

    @property
    def anzeigename(self) -> str:
        """Zusammengesetzter Name fuer Spalte B (``Vorname Nachname``)."""
        return f"{self.vorname} {self.nachname}".strip()

    def anzeige_label(self) -> str:
        """Beschriftung des Profils in der Auswahlliste.

        Einzige Quelle fuer diesen Text: ohne eigenen Namen faellt die Anzeige
        auf den Personennamen zurueck, damit dasselbe Profil nicht je nach
        Aufrufweg mal "Erika Mustermann" und mal "(ohne Namen)" heisst.
        """
        return self.name.strip() or self.anzeigename or "(ohne Namen)"

    def email_gueltig(self) -> bool:
        return bool(EMAIL_RE.match(self.gruppenleiter_email.strip()))


@dataclass
class Gruppenleiter:
    """Ein Eintrag der gepflegten Gruppenleiter-Stammdatenliste.

    Dient als Komfort-Vorrat: waehlt man einen Eintrag im UI aus, werden
    Name, Abteilung und E-Mail in die Profilfelder uebernommen, statt sie
    bei jedem Vorgang erneut einzutippen. Die eigentliche Quelle fuer die
    Excel-Zellen (``A3``/``D3``) und das Mail-CC bleibt weiterhin
    ``Profile`` -- diese Liste ist nur der Vorrat, aus dem man uebernimmt.
    """

    name: str = ""
    abteilung: str = ""      # Abteilungskuerzel, z. B. "XX YY ZZZ"
    email: str = ""

    def email_gueltig(self) -> bool:
        return bool(EMAIL_RE.match(self.email.strip()))

    def anzeige(self) -> str:
        """Anzeigetext fuer Auswahllisten: ``"Name – Abteilung"``, bzw. nur
        der Name, wenn keine Abteilung hinterlegt ist."""
        if self.abteilung.strip():
            return f"{self.name} – {self.abteilung}"
        return self.name


@dataclass
class TravelRow:
    """Eine Datenzeile der Liste.

    Eine Dienstreise belegt typischerweise zwei Zeilen: Hinfahrt und Rueckfahrt.
    In der Folgezeile bleiben Personalnummer, Name und Datum leer -- das
    entspricht dem Referenzdokument. ``ist_folgezeile`` markiert diesen Fall,
    damit die Validierung dort kein Datum einfordert.
    """

    datum: date | None = None
    passiv_von: time | None = None
    passiv_bis: time | None = None
    aktiv_von: time | None = None
    aktiv_bis: time | None = None
    grenz_anreise: time | None = None
    grenz_rueckreise: time | None = None
    arbeit_von: time | None = None
    arbeit_bis: time | None = None
    ist_folgezeile: bool = False

    def hat_zeiten(self) -> bool:
        """True, wenn irgendeine Uhrzeit gesetzt ist."""
        return any(
            t is not None
            for t in (
                self.passiv_von, self.passiv_bis,
                self.aktiv_von, self.aktiv_bis,
                self.grenz_anreise, self.grenz_rueckreise,
                self.arbeit_von, self.arbeit_bis,
            )
        )

    def arbeit_vor_aktivreise(self) -> bool:
        """True, wenn in dieser Zeile die Arbeitszeit vor der aktiven Reise liegt.

        "Erst arbeiten, dann selbst ans Steuer" ist fuer die GLAZ-Auswertung
        ein eigener Fall: die Arbeitszeit ist regulaer gebucht, die
        anschliessende aktive Reisezeit kommt obendrauf. Ein Tag mit mehreren
        solchen Bloecken (arbeiten, reisen, arbeiten, reisen) laesst sich nicht
        in einer Zeile abbilden -- die Vorlage hat pro Zeile genau EIN
        Arbeitszeit-Paar. Er wird deshalb auf zwei Zeilen verteilt, von denen
        dann jede hier True liefert.

        Nahtlos (``arbeit_bis == aktiv_von``) zaehlt bewusst mit: direkt nach
        Feierabend losfahren ist der Normalfall, nicht die Ausnahme. Eine echte
        Ueberschneidung liefert dagegen False -- die ist nicht "Arbeit vor der
        Reise", sondern "Arbeit waehrend der Reise", und die Validierung meldet
        sie getrennt als Warnung.
        """
        # Gleiche Vorsicht wie in validation._intervall_ueberschneidung: ein
        # halb gefuelltes Paar ist noch kein Muster, sondern eine Eingabe
        # mitten im Tippen. Ein Vergleich gegen None wuerde hier werfen.
        if None in (self.arbeit_von, self.arbeit_bis, self.aktiv_von, self.aktiv_bis):
            return False
        return self.arbeit_bis <= self.aktiv_von

    def ist_leer(self) -> bool:
        """True, wenn die Zeile fuer eine Korrekturbuchung nichts hergibt.

        Allein die Uhrzeiten entscheiden; das Datum bleibt bewusst aussen vor.
        Korrigiert werden Zeiten -- ein Datum ohne jede Uhrzeit korrigiert
        nichts. Praktisch faellt das ins Gewicht, weil das Datumsfeld im
        Formular mit dem heutigen Tag vorbelegt ist: jede noch unbenutzte
        Zeile traegt damit ein Datum. Zaehlte das als Inhalt, wanderte jede
        Leerzeile als uhrzeitlose Datenzeile in die Excel, und die Regel
        "Es ist keine Reisezeile erfasst" schluege nie mehr an.
        """
        return not self.hat_zeiten()


@dataclass
class Vorgang:
    """Ein auszufuellender Vorgang: Profil + Einsatzart + Reisezeilen."""

    profil: Profile = field(default_factory=Profile)
    einsatzart: str = ""                      # -> Zelle J3
    reisetyp: TravelType = TravelType.INLAND
    zeilen: list[TravelRow] = field(default_factory=list)
    zielordner: str = ""

    #: Wiederholung von Pers.-Nr./Name/Datum in Folgezeilen.
    #: Default False -- so wie im Referenzdokument.
    wiederhole_stammdaten: bool = False

    def aktive_zeilen(self) -> list[TravelRow]:
        return [z for z in self.zeilen if not z.ist_leer()]

    def hat_arbeit_vor_aktivreise(self) -> bool:
        """True, wenn mindestens eine erfasste Zeile Arbeit vor aktiver Reise zeigt.

        Fragt nur die aktiven Zeilen: eine leere Zeile kann das Muster gar
        nicht erfuellen, und die Frage gilt dem, was tatsaechlich in die Liste
        geschrieben wird.
        """
        return any(z.arbeit_vor_aktivreise() for z in self.aktive_zeilen())

    def zeitraum(self) -> tuple[date | None, date | None]:
        """Erster und letzter Reisetag ueber alle AKTIVEN Zeilen mit Datum.

        Bewusst nicht ueber ``self.zeilen``: das Datumsfeld ist vorbelegt, eine
        unbenutzte Zeile traegt also das heutige Datum. Ueber alle Zeilen
        gerechnet zoege eine einzige Leerzeile den Zeitraum bis heute auf --
        und dieser Zeitraum steht in Betreff und Text der Mail und im
        Dateinamen.
        """
        daten = sorted(z.datum for z in self.aktive_zeilen() if z.datum is not None)
        if not daten:
            return None, None
        return daten[0], daten[-1]

    def zeitraum_text(self) -> str:
        """Zeitraum fuer Betreff/Body: ``TT.MM.JJJJ`` bzw. ``TT.MM.-TT.MM.JJJJ``."""
        von, bis = self.zeitraum()
        if von is None:
            return ""
        if von == bis:
            return von.strftime("%d.%m.%Y")
        if von.year == bis.year:
            return f"{von.strftime('%d.%m.')}–{bis.strftime('%d.%m.%Y')}"
        return f"{von.strftime('%d.%m.%Y')}–{bis.strftime('%d.%m.%Y')}"
