"""Erzeugt Betreff, Text und Adressen der Mail -- ohne Anhang, ohne Versand.

Abgetrennt von :mod:`glaz.mail`, weil dort zwei sehr verschiedene Dinge
zusammenlagen: das Fuellen von Textvorlagen (reine Zeichenkettenarbeit) und
das Ansprechen des lokalen Mailprogramms (``sys.platform``, COM,
``subprocess``, Dateipfade). Nur das Erste laesst sich in Pyodide ausfuehren.

Die Progressive Web App erzeugt denselben Text wie die Desktop-App und
uebergibt ihn der Outlook-App als ``ms-outlook://compose``-Aufruf; den Anhang
haengt dort die Nutzerin selbst an, weil ein Browser das nicht darf. Damit beide Wege
garantiert denselben Wortlaut liefern, gibt es diese Logik genau einmal --
hier. :func:`glaz.mail.baue_mail` ergaenzt lediglich den Anhang.

Deshalb importiert dieses Modul ausser :mod:`glaz.model` nichts: jeder
weitere Import koennte plattformabhaengigen Code nachziehen und die PWA beim
Start zum Absturz bringen.
"""

from __future__ import annotations

from .model import Vorgang


class MailtextFehler(RuntimeError):
    """Eine Vorlage laesst sich nicht fuellen.

    Eigene Klasse statt ``glaz.mail.MailFehler``: Die haengt am Mailversand
    und damit an der Plattform, die es hier gerade nicht geben soll.
    :mod:`glaz.mail` faengt diesen Fehler und reicht ihn als ``MailFehler``
    weiter -- fuer die Desktop-App aendert sich dadurch nichts.

    Die Meldung ist deutsch und handlungsleitend: sie wird der Nutzerin
    direkt angezeigt.
    """


#: Zusatzabsatz, sobald Arbeitszeit vor der aktiven Reisezeit erfasst ist.
#: Er haengt bewusst NICHT in der Body-Vorlage: Die ist frei editierbar, und
#: ein Hinweis, der nur manchmal gilt, liesse sich dort weder sinnvoll
#: formulieren noch bedingt ausblenden. Das Gegenstueck in der Oberflaeche
#: (``ARBEITSZEIT_HINWEIS_UI`` in glaz/ui/main_window.py) kuendigt genau
#: diesen Absatz an.
ARBEITSZEIT_HINWEIS = (
    "Hinweis: An mindestens einem Tag liegt Arbeitszeit vor der aktiven "
    "Reisezeit (erst gearbeitet, dann selbst gefahren). Bitte bei der "
    "Eintragung zusätzlich berücksichtigen — die aktive Reisezeit kommt hier "
    "zur regulär gebuchten Arbeitszeit hinzu. Tage mit mehreren solchen "
    "Abschnitten sind in der Liste auf mehrere Zeilen aufgeteilt, weil die "
    "Vorlage je Zeile nur eine Arbeitszeit vorsieht."
)


def _zusatzhinweise(vorgang: Vorgang) -> list[str]:
    """Absaetze, die je nach erfassten Zeiten an den Body angehaengt werden.

    Als Liste und nicht als einzelner String, damit ein zweiter Sonderfall
    hier dazukommen kann, ohne dass der Aufrufer sich aendert.
    """
    hinweise: list[str] = []
    if vorgang.hat_arbeit_vor_aktivreise():
        hinweise.append(ARBEITSZEIT_HINWEIS)
    return hinweise


def _platzhalter_werte(vorgang: Vorgang) -> dict[str, str]:
    """Werte fuer die Platzhalter in Betreff-/Body-Vorlagen."""
    profil = vorgang.profil
    return {
        "name": profil.anzeigename,
        "zeitraum": vorgang.zeitraum_text(),
        "einsatzart": vorgang.einsatzart,
        "personalnummer": profil.personalnummer,
        "abteilung": profil.abteilung,
    }


def _fuelle_vorlage(vorlage: str, werte: dict[str, str]) -> str:
    """Fuellt Platzhalter wie ``{name}`` in ``vorlage`` mit ``werte``.

    Eine unbekannte oder fehlerhafte Vorlage (z. B. benutzerdefiniert von
    der Anwenderin geaendert) fuehrt zu einem :class:`MailtextFehler` statt zu
    einem rohen ``KeyError``/``ValueError``.
    """
    try:
        return vorlage.format(**werte)
    except KeyError as exc:
        unbekannter_platzhalter = exc.args[0]
        raise MailtextFehler(
            f"Die Vorlage enthaelt den unbekannten Platzhalter "
            f"'{{{unbekannter_platzhalter}}}'. Erlaubt sind: "
            f"{', '.join(f'{{{k}}}' for k in werte)}."
        ) from exc
    except (IndexError, ValueError) as exc:
        raise MailtextFehler(f"Die Vorlage ist ungueltig formatiert: {exc}") from exc


def baue_mail_texte(
    vorgang: Vorgang,
    empfaenger: str,
    betreff_vorlage: str,
    body_vorlage: str,
) -> dict[str, str]:
    """Erzeugt Adressen, Betreff und Text -- alles, was keinen Anhang braucht.

    Fuellt die Platzhalter ``{name}``, ``{zeitraum}``, ``{einsatzart}``,
    ``{personalnummer}`` und ``{abteilung}`` aus ``vorgang`` in
    ``betreff_vorlage`` und ``body_vorlage``. CC wird aus
    ``vorgang.profil.gruppenleiter_email`` gesetzt.

    Haengt ausserdem die Zusatzhinweise an, die sich aus den erfassten Zeiten
    ergeben (siehe :func:`_zusatzhinweise`) -- angehaengt statt in die Vorlage
    eingesetzt, damit sie auch in einer von der Anwenderin umgeschriebenen
    Vorlage erscheinen.

    Der Rueckgabewert ist bewusst ein Dict aus reinen Zeichenketten und kein
    Dataclass-Objekt: Er wird in der PWA unveraendert nach JavaScript
    durchgereicht.

    :returns: ``{"an", "cc", "betreff", "body"}``
    :raises MailtextFehler: wenn eine Vorlage nicht gefuellt werden kann.
    """
    werte = _platzhalter_werte(vorgang)
    betreff = _fuelle_vorlage(betreff_vorlage, werte)
    body = _fuelle_vorlage(body_vorlage, werte)

    hinweise = _zusatzhinweise(vorgang)
    if hinweise:
        # Der Body endet je nach Vorlage mit oder ohne Zeilenumbruch --
        # rstrip zuerst, sonst entstehen hier mal zwei und mal vier Leerzeilen.
        body = body.rstrip() + "\n\n" + "\n\n".join(hinweise) + "\n"

    return {
        "an": empfaenger,
        "cc": vorgang.profil.gruppenleiter_email,
        "betreff": betreff,
        "body": body,
    }
