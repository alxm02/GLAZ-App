"""Fachliche Vorgabewerte: Empfaenger, Mailvorlagen, Dateinamen-Schema.

Diese vier Konstanten sind reine Fachdaten -- sie beantworten die Frage
"womit starten wir, solange niemand etwas anderes eingestellt hat?" und haben
mit *Speichern* nichts zu tun. Bis zur Einfuehrung der Progressive Web App
standen sie trotzdem in :mod:`glaz.settings`, weil das der einzige Ort war,
der sie brauchte.

Die PWA laeuft in Pyodide im Browser: dort gibt es kein ``%APPDATA%``, keine
``settings.json`` und kein ``sys.platform == "win32"``. Ein Import von
:mod:`glaz.settings` allein fuer diese Zeichenketten zoege die gesamte
Dateisystem- und Plattformlogik mit -- die dort nicht laufen kann und auch
nicht laufen soll (die Einstellungen der PWA leben im ``localStorage``).

Deshalb liegen die Werte hier, in einem Modul ohne jeden Import.
:mod:`glaz.settings` holt sie von hier und reicht sie unveraendert weiter,
damit vorhandener Code und vorhandene Tests weiterhin
``from glaz.settings import DEFAULT_EMPFAENGER`` schreiben duerfen.
"""

from __future__ import annotations

#: Fachbereich, der die ausgefuellte Liste entgegennimmt.
DEFAULT_EMPFAENGER = ""

#: Betreffzeile. Platzhalter: ``{name}``, ``{zeitraum}``, ``{einsatzart}``,
#: ``{personalnummer}``, ``{abteilung}`` -- siehe
#: :func:`glaz.mailtext._platzhalter_werte`.
DEFAULT_BETREFF_VORLAGE = "GLAZ-Korrekturbuchungsliste – {name} – {zeitraum}"

#: Mailtext. Dieselben Platzhalter wie im Betreff.
DEFAULT_BODY_VORLAGE = """Hallo zusammen,

anbei die GLAZ-Korrekturbuchungsliste für meine Dienstreise ({einsatzart}) vom {zeitraum}.

Personalnummer: {personalnummer}
Abteilung: {abteilung}

Vielen Dank und viele Grüße
{name}
"""

#: Dateiname der erzeugten Liste. Platzhalter: ``{datum}``, ``{nachname}``,
#: ``{einsatzart}`` -- siehe :func:`glaz.filename.baue_dateinamen`.
DEFAULT_DATEINAMEN_SCHEMA = "{datum}_GLAZ-Korrektur_{nachname}_{einsatzart}.xlsx"
