DOSBox Spiele-Manager - Beschreibung der UI
===========================================

Kurzbeschreibung
----------------

Benötigt das original DosBox : https://www.dosbox.com/ ein tolles stück software unterstützt die Jungs !

Der DOSBox Spiele-Manager ist eine Windows-Oberflaeche (PowerShell + WinForms),
mit der du deine DOS-Spielebibliothek in einem Hauptordner verwaltest.
Die UI hilft dir beim Scannen, Sortieren, Starten und Optimieren deiner Spiele
sowie beim Verwalten von Vorschaubildern und Favoriten.

Was die UI macht
----------------
1. Sie merkt sich deinen zuletzt verwendeten Spieleordner.
2. Sie scannt Unterordner nach Spielen und ermittelt pro Eintrag:
   - Spielname/Ordner
   - moegliche Startdatei (.exe/.com/.bat)
   - Optimierungsstatus (z. B. vorhandene .conf oder DOSBox-Aufruf in .bat)
3. Sie zeigt alle Spiele als Liste oder Kachelansicht an.
4. Sie stellt Detailinformationen und Aktionen fuer das gewaehlte Spiel bereit.
5. Sie kann Daten und Vorschaubilder zwischenspeichern (Cache), damit beim
   naechsten Start nicht alles neu eingelesen werden muss.

UI-Bereiche im Ueberblick
-------------------------
Kopfbereich (oben):
- Ordnerpfad (Textfeld): Spieleordner direkt eingeben
- Ordner waehlen: Ordnerdialog oeffnen
- Neu scannen: Scan manuell starten
- Ansicht: Liste/Kacheln umschalten
- Daten speichern: Eintragsdaten im Cache speichern
- Daten laden: Eintragsdaten aus Cache laden
- Vorschauen erstellen: Vorschaubilder fuer alle geladenen Spiele erstellen/aktualisieren
- Suche + X: Filtertext eingeben und schnell loeschen
- Nur Favoriten / Alle zeigen: Favoritenfilter umschalten
- Sortierung: A-Z, Z-A, Status, Favoriten
- Header ausblenden/einblenden: mehr Platz im Arbeitsbereich

Linke Seite (Bibliothek):
- Spielansicht als Tabelle oder Kachel
- In Kachelansicht werden Vorschaubilder angezeigt, falls vorhanden

Rechte Seite (Details + Aktionen):
- Detailfelder zu Name, Ordner, Startdatei, Optimierung, Erkennung, Config
- Favorit setzen/entfernen
- Spiel starten
- Optimieren (Konfigurationsdatei und Starter anlegen)
- Anpassen (Config) in Notepad
- DOSBox Optionen oeffnen
- Bilder hinzufuegen (lokal)
- Google-Bild (ein Spiel)
- Google-Bilder (ALLE) fuer Batch-Download
- Log-Fenster mit Status- und Fehlermeldungen

Wichtige Funktionen im Detail
-----------------------------
Spiel starten:
- Der Manager startet bevorzugt vorhandene Startskripte direkt.
- Falls noetig wird DOSBox mit passender Konfiguration aufgerufen.

Optimieren:
- Erstellt eine basisnahe DOSBox-Konfiguration im Spielordner.
- Legt zusaetzlich einen passenden Starter an.
- Markiert den Eintrag danach als optimiert.

Vorschaubilder:
- Automatisch aus vorhandenen Bilddateien im Spielordner erstellbar.
- Manuell aus einer lokalen Bilddatei setzbar.
- Online-Suche (z. B. Cover/Screenshot) fuer ein Spiel oder als Batch fuer alle.

Favoriten:
- Spiele koennen als Favorit markiert werden.
- Favoriten koennen gefiltert und bevorzugt sortiert werden.

Cache und Persistenz:
- Einstellungen (zuletzt genutzter Root-Ordner) werden gespeichert.
- Eintraege inklusive Favoriten/Preview-Infos koennen als Cache gesichert werden.
- Beim naechsten Start kann der Cache geladen werden.

Technik: Pfade im Code (2-Spalten-Uebersicht)
---------------------------------------------
Spalte 1: Bereich/Funktion
Spalte 2: Wo im Code und welche Pfadlogik genutzt wird

- Basisordner des Managers
   -> DOSBox-Spiele-Manager.ps1: `$script:BaseDir = Split-Path -Parent $MyInvocation.MyCommand.Path`
   -> Bedeutet: Alle relativen Pfade werden vom Ordner des Scripts aus aufgebaut.

- DOSBox-Hauptprogramm
   -> DOSBox-Spiele-Manager.ps1: `$script:DosBoxExe = Join-Path $script:BaseDir 'DOSBox.exe'`
   -> Pflichtpfad: `DOSBox.exe` muss im gleichen Hauptordner wie der Manager liegen.

- Optionen-Verknuepfung/Starter (technisch BAT)
   -> DOSBox-Spiele-Manager.ps1: `$script:OptionsBat = Join-Path $script:BaseDir 'DOSBox 0.74-3 Options.bat'`
   -> Die UI-Schaltflaeche `DOSBox Optionen` startet genau diese BAT-Datei.

- Gespeicherter letzter Spieleordner
   -> DOSBox-Spiele-Manager.ps1: `$script:SettingsPath = Join-Path $script:BaseDir 'dosbox-game-manager.settings.json'`
   -> Hier wird der zuletzt benutzte Root-Ordner persistent abgelegt.

- Cache-Datei pro Spiele-Root
   -> DOSBox-Spiele-Manager.ps1: `Get-CachePath([string]$rootFolder)`
   -> Ergebnis: `<Spiele-Root>\\dosbox-manager-cache.json`

- Vorschaubild pro Spielordner
   -> DOSBox-Spiele-Manager.ps1: `Get-PreviewCachePath([string]$folderPath)`
   -> Ergebnis: `<Spielordner>\\.dosbox-preview.png`

- Harte Startpruefung fuer DOSBox
   -> DOSBox-Spiele-Manager.ps1: `if (-not (Test-Path $script:DosBoxExe)) { ... exit 1 }`
   -> Wenn `DOSBox.exe` fehlt, startet der Manager nicht weiter.

Wichtig zu Verknuepfungen (.lnk)
--------------------------------
Die UI arbeitet intern mit echten Dateipfaden zu `DOSBox.exe` und `.bat/.conf`.
`.lnk`-Dateien im Paket sind Komfort-Shortcuts fuer Windows, aber nicht die
primaere technische Grundlage fuer den Spielstart in der UI.

Typischer Ablauf
----------------
1. DOSBox-Spiele-Manager.bat starten
2. Spieleordner auswaehlen
3. Neu scannen
4. Optional: Sortieren/Filtern/Favoriten setzen
5. Spiel auswaehlen und starten
6. Optional: Optimieren, Config anpassen, Vorschaubilder pflegen
7. Daten speichern

Hinweise
--------
- Fuer den Betrieb muss DOSBox.exe im Hauptordner vorhanden sein.
- Bei geblockten Skripten startet die mitgelieferte BAT den Manager mit
  ExecutionPolicy Bypass fuer diesen Start.
- Beim Batch-Bilddownload kann der Vorgang je nach Internetverbindung
  laenger dauern.

Wie DOSBox genau aufgesetzt sein muss
-------------------------------------
Minimal (Pflicht):
- `DOSBox.exe` liegt im gleichen Hauptordner wie `DOSBox-Spiele-Manager.ps1`.
- Der Spiele-Root enthaelt pro Spiel einen Unterordner mit mindestens einer
   startbaren Datei (`.exe`, `.com` oder `.bat`).

Empfohlen (fuer volle UI-Funktionen):
- `DOSBox 0.74-3 Options.bat` liegt ebenfalls im Hauptordner
   (fuer die Schaltflaeche `DOSBox Optionen`).
- Schreibrechte im Spiele-Root und in den Spielordnern sind vorhanden,
   damit Cache und Vorschaubilder gespeichert werden koennen.
- Optional pro Spiel eine `.conf` oder ein DOSBox-Aufruf in `.bat`, damit der
   Eintrag als optimiert erkannt wird.

Optional (Komfort):
- `.lnk`-Dateien in `Options/` und `Extras/` koennen genutzt werden, sind aber
   nicht zwingend fuer die interne Funktionsweise der Manager-UI.
