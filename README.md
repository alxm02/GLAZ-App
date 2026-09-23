# GLAZ-Korrekturbuchungsliste — Web-App

Dies ist **kein Quellrepository**, sondern der fertig gebaute Stand der
Web-App. Er wird von `tools/veroeffentliche_app.py` aus dem privaten
Hauptrepository hierher geschrieben; Änderungen von Hand gehen beim
nächsten Lauf verloren.

## Benutzung

Die App läuft unter **https://alxm02.github.io/GLAZ-App/**

Auf dem iPhone in **Safari** öffnen (nicht Chrome), warten bis sie
bedienbar ist, dann *Teilen → Zum Home-Bildschirm*. Danach arbeitet sie
vollständig offline: Die Prüfung der Eingaben und die Erzeugung der
Excel-Datei laufen als Python (via Pyodide) im Gerät.

## Daten

Es gibt keinen Server. Profile, erfasste Zeiten und Einstellungen liegen
im `localStorage` des jeweiligen Geräts und verlassen es nicht. Die
fertige Datei geht über die Outlook-App (Empfänger und CC vorbelegt)
oder über das Teilen-Menü des Telefons an die Mail-App.
