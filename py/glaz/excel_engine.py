"""Erzeugt die ausgefuellte GLAZ-Liste durch Patchen der Original-Vorlage.

Bewusst KEIN openpyxl: openpyxl verwirft beim Speichern Teile des OOXML-Pakets,
darunter ``docMetadata/LabelInfo.xml`` (Siemens-Sensitivitaetslabel),
``xl/featurePropertyBag/`` und ``xl/printerSettings/printerSettings1.bin``.

Stattdessen wird die .xlsx als ZIP geoeffnet, ausschliesslich
``xl/worksheets/sheet1.xml`` und ``xl/sharedStrings.xml`` veraendert und jeder
andere Part byte-identisch durchkopiert. Das garantiert, dass Styles, Merges,
Spaltenbreiten, Druckeinstellungen und das Sensitivitaetslabel exakt erhalten
bleiben -- die farbige Kennzeichnung der Eingabefelder steckt in den Zellformaten.
"""

from __future__ import annotations

import os
import re
import zipfile
import zlib
from dataclasses import dataclass
from datetime import date, time
from pathlib import Path

from .model import FIRST_DATA_ROW, LAST_DATA_ROW, MAX_ROWS, Vorgang, ist_personalnummer
from .timeconv import date_to_serial, format_number, time_to_fraction

SHEET_PART = "xl/worksheets/sheet1.xml"
SHARED_PART = "xl/sharedStrings.xml"

#: Parts, die dieser Generator veraendern darf. Alles andere wird durchkopiert.
MUTABLE_PARTS = frozenset({SHEET_PART, SHARED_PART})

#: Spaltenzuordnung des Datenbereichs (Zeilen 7-20).
COL_PERSONALNUMMER = "A"
COL_NAME = "B"
COL_DATUM = "C"
COL_PASSIV_VON = "D"
COL_PASSIV_BIS = "E"
COL_AKTIV_VON = "F"
COL_AKTIV_BIS = "G"
COL_GRENZ_ANREISE = "H"
COL_GRENZ_RUECKREISE = "I"
COL_ARBEIT_VON = "J"
COL_ARBEIT_BIS = "K"

#: Alle Spalten, die dieser Generator im Datenbereich befuellt. Nur diese
#: entscheiden darueber, ob eine Zeile als belegt gilt -- die Spalten rechts
#: davon traegt die Vorlage bereits mit Formeln bzw. Rahmen.
DATEN_SPALTEN = (
    COL_PERSONALNUMMER, COL_NAME, COL_DATUM,
    COL_PASSIV_VON, COL_PASSIV_BIS,
    COL_AKTIV_VON, COL_AKTIV_BIS,
    COL_GRENZ_ANREISE, COL_GRENZ_RUECKREISE,
    COL_ARBEIT_VON, COL_ARBEIT_BIS,
)

#: Kopfzellen: Abteilung, Gruppenleiter, Einsatzart.
KOPF_ABTEILUNG = "A3"
KOPF_GRUPPENLEITER = "D3"
KOPF_EINSATZART = "J3"


class ExcelEngineError(RuntimeError):
    """Fehler beim Erzeugen der Ausgabedatei."""


@dataclass(frozen=True)
class Belegung:
    """Belegung des Datenbereichs (Zeilen 7-20) einer vorhandenen Datei."""

    belegte_zeilen: int
    erste_freie_zeile: int | None
    freie_zeilen: int


@dataclass(frozen=True)
class Schreibergebnis:
    """Wohin geschrieben wurde -- Grundlage fuer ein spaeteres Aktualisieren."""

    pfad: Path
    start_row: int
    geschriebene_zeilen: int


@dataclass(frozen=True)
class CellValue:
    """Ein zu schreibender Zellwert. ``is_string`` steuert den sharedStrings-Pool."""

    value: str | float | int
    is_string: bool


#: In XML 1.0 unzulaessige Zeichen. Aus Word, PDF oder Outlook kopierter Text
#: bringt regelmaessig \x0b (in Word der weiche Zeilenumbruch) oder \x0c mit;
#: einmal in die sharedStrings geschrieben ist die .xlsx nicht mehr parsebar.
#: \t, \n und \r sind zulaessig und bleiben deshalb aussen vor.
_UNGUELTIG_RE = re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f\x7f]")


def _xml_escape(text: str) -> str:
    text = _UNGUELTIG_RE.sub("", text)
    return (
        text.replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
    )


class SharedStrings:
    """Verwaltet den sharedStrings-Pool und haengt neue Eintraege hinten an.

    Bestehende Eintraege behalten ihren Index -- damit bleiben alle statischen
    Zellbezuege (Ueberschriften, Fussnoten) unveraendert gueltig.
    """

    def __init__(self, xml: str):
        self._xml = xml
        self._items: list[str] = re.findall(r"<si>(.*?)</si>", xml, re.S)
        self._index: dict[str, int] = {}
        for i, item in enumerate(self._items):
            text = "".join(re.findall(r"<t[^>]*>(.*?)</t>", item, re.S))
            self._index.setdefault(text, i)
        self._appended: list[str] = []

    def intern(self, text: str) -> int:
        """Liefert den Index fuer ``text`` und legt ihn bei Bedarf neu an."""
        escaped = _xml_escape(text)
        if escaped in self._index:
            return self._index[escaped]
        idx = len(self._items) + len(self._appended)
        self._index[escaped] = idx
        # xml:space="preserve" ist noetig, sobald Rand-Leerzeichen vorkommen.
        preserve = ' xml:space="preserve"' if text != text.strip() else ""
        self._appended.append(f"<si><t{preserve}>{escaped}</t></si>")
        return idx

    def to_xml(self, ref_count: int) -> str:
        """Serialisiert den Pool und aktualisiert count/uniqueCount."""
        xml = self._xml
        if self._appended:
            insert_at = xml.rindex("</sst>")
            xml = xml[:insert_at] + "".join(self._appended) + xml[insert_at:]
        unique = len(self._items) + len(self._appended)
        xml = re.sub(
            r'(<sst[^>]*?)\scount="\d+"\s+uniqueCount="\d+"',
            lambda m: f'{m.group(1)} count="{ref_count}" uniqueCount="{unique}"',
            xml,
            count=1,
        )
        return xml


def _set_cell(sheet_xml: str, ref: str, cell: CellValue | None) -> str:
    """Schreibt einen Wert in eine bestehende ``<c>``-Zelle.

    Das ``s``-Attribut (Style) bleibt unangetastet -- daher wird nur der
    Elementkoerper ersetzt, nie die Zelle neu erzeugt. Fehlt die Zelle in der
    Vorlage, ist das ein Strukturfehler und keine stille Ergaenzung wert.
    """
    pattern = re.compile(r'<c r="%s"((?:\s+[a-zA-Z:]+="[^"]*")*)\s*(?:/>|>.*?</c>)' % re.escape(ref), re.S)
    match = pattern.search(sheet_xml)
    if not match:
        raise ExcelEngineError(
            f"Zelle {ref} existiert nicht in der Vorlage - Vorlage veraendert?"
        )

    attrs = match.group(1)
    # Ein evtl. vorhandenes t-Attribut wird neu gesetzt, alles andere bleibt.
    attrs = re.sub(r'\s+t="[^"]*"', "", attrs)

    if cell is None:
        replacement = f'<c r="{ref}"{attrs}/>'
    elif cell.is_string:
        replacement = f'<c r="{ref}"{attrs} t="s"><v>{cell.value}</v></c>'
    else:
        replacement = f'<c r="{ref}"{attrs}><v>{cell.value}</v></c>'

    return sheet_xml[: match.start()] + replacement + sheet_xml[match.end():]


def _hat_wert(sheet_xml: str, ref: str) -> bool:
    """True, wenn die Zelle ``ref`` im Blatt einen ``<v>``-Wert traegt.

    Leere Zellen kommen in freier Wildbahn in beiden Schreibweisen vor --
    ``<c r="A7" s="1"/>`` und ``<c r="A7" s="1"></c>``. Beide zaehlen als leer,
    weil erst der ``<v>``-Inhalt einen Wert ausmacht.
    """
    pattern = re.compile(
        r'<c r="%s"((?:\s+[a-zA-Z:]+="[^"]*")*)\s*(?:/>|>(.*?)</c>)' % re.escape(ref), re.S
    )
    match = pattern.search(sheet_xml)
    if not match or match.group(2) is None:
        return False
    return "<v>" in match.group(2)


def _collect_values(
    vorgang: Vorgang,
    shared: SharedStrings,
    sheet_xml: str,
    *,
    start_row: int = FIRST_DATA_ROW,
    kopfzellen_nur_wenn_leer: bool = False,
) -> dict[str, CellValue]:
    """Baut die Zuordnung Zelle -> Wert fuer einen Vorgang.

    A2 (Bereichskuerzel) bleibt bewusst leer, ebenso P3/Q3 (Monat/Jahr-Dropdowns,
    deren definedNames in der Vorlage auf #REF! zeigen).

    ``start_row`` verschiebt den Datenblock nach unten -- beim Anhaengen an eine
    schon befuellte Datei beginnt er hinter den vorhandenen Zeilen.
    ``kopfzellen_nur_wenn_leer`` schuetzt A3/D3/J3 vor dem Ueberschreiben: in
    einer vorhandenen Datei steht dort bereits, was der Anwender bewusst
    hinterlegt hat, und das aktuelle Profil muss nicht dasselbe sagen.
    """
    values: dict[str, CellValue] = {}
    profil = vorgang.profil

    def put_kopf(ref: str, text: str) -> None:
        if not text:
            return
        if kopfzellen_nur_wenn_leer and _hat_wert(sheet_xml, ref):
            return
        values[ref] = CellValue(shared.intern(text), True)

    if profil.abteilung.strip():
        put_kopf(KOPF_ABTEILUNG, f"Abteilung: {profil.abteilung.strip()}")
    if profil.gruppenleiter_name.strip():
        put_kopf(KOPF_GRUPPENLEITER, f"Gruppenleiter: {profil.gruppenleiter_name.strip()}")
    if vorgang.einsatzart.strip():
        put_kopf(KOPF_EINSATZART, vorgang.einsatzart.strip())

    zeilen = vorgang.aktive_zeilen()
    platz = LAST_DATA_ROW - start_row + 1
    if len(zeilen) > platz:
        raise ExcelEngineError(
            f"{len(zeilen)} Datenzeilen - ab Zeile {start_row} passen maximal "
            f"noch {platz}."
        )

    def put_time(row_no: int, col: str, t: time | None) -> None:
        if t is not None:
            values[f"{col}{row_no}"] = CellValue(format_number(time_to_fraction(t)), False)

    for offset, zeile in enumerate(zeilen):
        row_no = start_row + offset
        # Stammdaten nur in der ersten Zeile einer Reise -- so wie im Original.
        zeige_stammdaten = not zeile.ist_folgezeile or vorgang.wiederhole_stammdaten
        if zeige_stammdaten:
            if profil.personalnummer.strip():
                pnr = profil.personalnummer.strip()
                # Nur reine ASCII-Ziffern ohne fuehrende Null duerfen in eine
                # Zahlzelle: alles andere wuerde entweder kein gueltiger
                # xsd:double sein (Hochzahlen, arabisch-indische Ziffern) oder
                # in Excel die fuehrenden Nullen verlieren.
                if ist_personalnummer(pnr) and not pnr.startswith("0"):
                    values[f"{COL_PERSONALNUMMER}{row_no}"] = CellValue(pnr, False)
                else:
                    values[f"{COL_PERSONALNUMMER}{row_no}"] = CellValue(shared.intern(pnr), True)
            if profil.anzeigename:
                values[f"{COL_NAME}{row_no}"] = CellValue(shared.intern(profil.anzeigename), True)
            if zeile.datum is not None:
                values[f"{COL_DATUM}{row_no}"] = CellValue(date_to_serial(zeile.datum), False)

        put_time(row_no, COL_PASSIV_VON, zeile.passiv_von)
        put_time(row_no, COL_PASSIV_BIS, zeile.passiv_bis)
        put_time(row_no, COL_AKTIV_VON, zeile.aktiv_von)
        put_time(row_no, COL_AKTIV_BIS, zeile.aktiv_bis)
        put_time(row_no, COL_GRENZ_ANREISE, zeile.grenz_anreise)
        put_time(row_no, COL_GRENZ_RUECKREISE, zeile.grenz_rueckreise)
        put_time(row_no, COL_ARBEIT_VON, zeile.arbeit_von)
        put_time(row_no, COL_ARBEIT_BIS, zeile.arbeit_bis)

    return values


def _lies_paket(datei: Path, bezeichnung: str) -> tuple[list[zipfile.ZipInfo], dict[str, bytes]]:
    """Laedt eine .xlsx komplett in den Speicher.

    Vollstaendig einlesen statt streamen, damit die Quelle beim Schreiben des
    Ergebnisses schon wieder geschlossen ist -- ``haenge_an`` ueberschreibt
    genau diese Datei.
    """
    datei = Path(datei)
    if not datei.is_file():
        raise ExcelEngineError(f"{bezeichnung} nicht gefunden: {datei}")

    try:
        with zipfile.ZipFile(datei, "r") as src:
            infos = src.infolist()
            parts = {info.filename: src.read(info.filename) for info in infos}
    except (zipfile.BadZipFile, zlib.error, KeyError, EOFError) as exc:
        # Nicht nur ein fehlendes Zip-Verzeichnis: Auch ein beschaedigter
        # Datenstrom (zlib.error) oder ein abgeschnittenes Paket muessen als
        # ExcelEngineError ankommen. Die Aufrufer -- auch die laufende Pruefung
        # der Web-App -- fangen nur diesen, alles andere schluege roh durch.
        raise ExcelEngineError(
            f"{bezeichnung} ist keine gueltige Excel-Datei: {datei}"
        ) from exc
    except OSError as exc:
        # Keine Leserechte, eine nicht synchronisierte OneDrive-Datei, ein
        # getrenntes Netzlaufwerk.
        raise ExcelEngineError(
            f"{bezeichnung} laesst sich nicht lesen: {datei} ({exc})"
        ) from exc

    for pflicht in (SHEET_PART, SHARED_PART):
        if pflicht not in parts:
            raise ExcelEngineError(f"{pflicht} fehlt - keine gueltige {bezeichnung}: {datei}")

    return infos, parts


def _befuelle_parts(
    parts: dict[str, bytes],
    vorgang: Vorgang,
    *,
    start_row: int,
    kopfzellen_nur_wenn_leer: bool,
) -> None:
    """Patcht sheet1.xml und sharedStrings.xml in ``parts`` mit den Vorgangsdaten."""
    shared = SharedStrings(parts[SHARED_PART].decode("utf-8"))
    sheet_xml = parts[SHEET_PART].decode("utf-8")
    values = _collect_values(
        vorgang,
        shared,
        sheet_xml,
        start_row=start_row,
        kopfzellen_nur_wenn_leer=kopfzellen_nur_wenn_leer,
    )

    for ref, cell in values.items():
        sheet_xml = _set_cell(sheet_xml, ref, cell)

    # count = Anzahl aller String-Zellbezuege im Blatt (Excel toleriert Abweichung,
    # wir halten es dennoch konsistent).
    ref_count = len(re.findall(r'<c [^>]*t="s"', sheet_xml))
    parts[SHEET_PART] = sheet_xml.encode("utf-8")
    parts[SHARED_PART] = shared.to_xml(ref_count).encode("utf-8")


def _schreibe_paket(target: Path, infos: list[zipfile.ZipInfo], parts: dict[str, bytes]) -> Path:
    """Schreibt das Paket ueber eine ``.tmp``-Datei nach ``target``.

    Alle Parts ausser sheet1.xml und sharedStrings.xml gehen byte-identisch
    hinaus (inkl. Kompressionsart und Zeitstempel).
    """
    target.parent.mkdir(parents=True, exist_ok=True)
    tmp = target.with_suffix(target.suffix + ".tmp")
    try:
        with zipfile.ZipFile(tmp, "w") as dst:
            for info in infos:
                new_info = zipfile.ZipInfo(info.filename, date_time=info.date_time)
                new_info.compress_type = info.compress_type
                new_info.external_attr = info.external_attr
                new_info.internal_attr = info.internal_attr
                new_info.create_system = info.create_system
                dst.writestr(new_info, parts[info.filename])
        # os.replace statt shutil.move: Unter Windows scheitert os.rename an
        # einem vorhandenen Ziel, shutil.move faellt dann auf copy2 zurueck --
        # und das kuerzt die Zieldatei zuerst. Bricht die Kopie ab (Netzlaufwerk,
        # voller Datentraeger), bliebe von einer ueber Wochen ergaenzten Liste
        # ein Rumpf. os.replace tauscht in einem Schritt.
        os.replace(tmp, target)
    except PermissionError as exc:
        tmp.unlink(missing_ok=True)
        raise ExcelEngineError(
            f"Datei ist gesperrt oder der Ordner nicht beschreibbar: {target}\n"
            "Ist die Datei noch in Excel geoeffnet?"
        ) from exc
    except OSError as exc:
        tmp.unlink(missing_ok=True)
        raise ExcelEngineError(f"Schreiben fehlgeschlagen: {target} ({exc})") from exc

    return target


def _belegung_aus_sheet(sheet_xml: str) -> Belegung:
    """Wertet den Datenbereich eines schon geladenen Blattes aus."""
    belegte = [
        row
        for row in range(FIRST_DATA_ROW, LAST_DATA_ROW + 1)
        if any(_hat_wert(sheet_xml, f"{col}{row}") for col in DATEN_SPALTEN)
    ]

    # Bewusst hinter der LETZTEN belegten Zeile weiterschreiben, nicht in die
    # erste Luecke: eine Reise belegt Hin- und Rueckfahrt in zwei aufeinander
    # folgenden Zeilen. Wuerde eine Luecke mittendrin gefuellt, risse der neue
    # Eintrag eine bestehende Reise auseinander.
    naechste = (belegte[-1] + 1) if belegte else FIRST_DATA_ROW
    if naechste > LAST_DATA_ROW:
        return Belegung(belegte_zeilen=len(belegte), erste_freie_zeile=None, freie_zeilen=0)
    return Belegung(
        belegte_zeilen=len(belegte),
        erste_freie_zeile=naechste,
        freie_zeilen=LAST_DATA_ROW - naechste + 1,
    )


def fill_template(template: Path, target: Path, vorgang: Vorgang) -> Path:
    """Erzeugt ``target`` aus ``template``, befuellt mit den Daten aus ``vorgang``.

    Alle Parts ausser sheet1.xml und sharedStrings.xml werden byte-identisch
    uebernommen (inkl. Kompressionsart und Zeitstempel).
    """
    infos, parts = _lies_paket(Path(template), "Vorlage")
    _befuelle_parts(parts, vorgang, start_row=FIRST_DATA_ROW, kopfzellen_nur_wenn_leer=False)
    return _schreibe_paket(Path(target), infos, parts)


def lies_belegung(datei: Path) -> Belegung:
    """Ermittelt, wie viele Datenzeilen in ``datei`` schon belegt sind."""
    _, parts = _lies_paket(Path(datei), "Datei")
    return _belegung_aus_sheet(parts[SHEET_PART].decode("utf-8"))


def haenge_an(
    datei: Path,
    vorgang: Vorgang,
    start_row: int | None = None,
    ueberschreibe_zeilen: int = 0,
) -> Schreibergebnis:
    """Ergaenzt ``datei`` in place um die aktiven Zeilen aus ``vorgang``.

    Damit laesst sich eine Liste ueber mehrere Sitzungen hinweg fuellen und
    erst spaeter versenden. Vorhandene Zeilen und Kopfangaben bleiben
    unangetastet.

    ``start_row`` und ``ueberschreibe_zeilen`` dienen dem Aktualisieren: mit dem
    ``Schreibergebnis`` eines vorangegangenen Laufs laesst sich derselbe Block
    noch einmal schreiben, nachdem der Anwender im Formular etwas geaendert hat.
    """
    datei = Path(datei)
    infos, parts = _lies_paket(datei, "Datei")
    sheet_xml = parts[SHEET_PART].decode("utf-8")
    zeilen = vorgang.aktive_zeilen()

    if start_row is None:
        belegung = _belegung_aus_sheet(sheet_xml)
        if belegung.erste_freie_zeile is None:
            raise ExcelEngineError(
                f"{datei.name} ist voll - alle {MAX_ROWS} Datenzeilen sind belegt, "
                f"{len(zeilen)} Zeilen sollen aber noch angehaengt werden."
            )
        start_row = belegung.erste_freie_zeile
    elif not FIRST_DATA_ROW <= start_row <= LAST_DATA_ROW:
        raise ExcelEngineError(
            f"Startzeile {start_row} liegt ausserhalb des Datenbereichs "
            f"{FIRST_DATA_ROW}-{LAST_DATA_ROW}."
        )

    # Beim Aktualisieren steht der zuvor beschriebene Block wieder zur
    # Verfuegung -- massgeblich ist deshalb allein, was ab ``start_row`` passt.
    platz = LAST_DATA_ROW - start_row + 1
    if len(zeilen) > platz:
        raise ExcelEngineError(
            f"{len(zeilen)} Zeilen sollen ab Zeile {start_row} geschrieben werden, "
            f"dort ist aber nur noch Platz fuer {platz}."
        )

    # Schrumpft der Vorgang gegenueber dem letzten Lauf, muessen die
    # ueberzaehligen Altzeilen weg -- sonst blieben Reste einer Reise stehen,
    # die der Anwender gerade geloescht hat.
    for offset in range(min(max(ueberschreibe_zeilen, len(zeilen)), platz)):
        for col in DATEN_SPALTEN:
            sheet_xml = _set_cell(sheet_xml, f"{col}{start_row + offset}", None)
    parts[SHEET_PART] = sheet_xml.encode("utf-8")

    _befuelle_parts(
        parts,
        vorgang,
        start_row=start_row,
        kopfzellen_nur_wenn_leer=True,
    )
    return Schreibergebnis(
        pfad=_schreibe_paket(datei, infos, parts),
        start_row=start_row,
        geschriebene_zeilen=len(zeilen),
    )
