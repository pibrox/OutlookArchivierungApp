# Autostart-Funktion der Outlook Archivierungs-App

## Überblick

Die Autostart-Funktion ermöglicht es der Anwendung, automatisch beim Start von Windows zu starten, ohne dass Administratorrechte erforderlich sind.

## Funktionsweise

Die Implementierung nutzt den Windows-Startordner (`%APPDATA%\Microsoft\Windows\Start Menu\Programs\Startup`), in dem eine Verknüpfung zur Anwendung erstellt wird. Dies ist eine benutzerfreundliche Methode, die keine Registrierungsänderungen oder Administratorrechte erfordert.

## Implementierungsdetails

### Aktivieren des Autostarts

Wenn der Benutzer die Option "Mit Windows starten" aktiviert:

1. Es wird eine Verknüpfung (`.lnk`-Datei) zur Anwendung im Windows-Startordner erstellt
2. Die Einstellung wird in der `settings.json`-Datei gespeichert

### Deaktivieren des Autostarts

Wenn der Benutzer die Option deaktiviert:

1. Die Verknüpfung wird aus dem Startordner entfernt
2. Die Einstellung wird in der `settings.json`-Datei aktualisiert

### Überprüfung des Status

Beim Programmstart wird überprüft, ob die Verknüpfung im Startordner existiert, und die Checkbox entsprechend gesetzt.

## Fehlerbehebung

Falls die Autostart-Funktion nicht wie erwartet funktioniert:

- Prüfen Sie, ob die Verknüpfung im Startordner vorhanden ist
- Stellen Sie sicher, dass der Pfad in der Verknüpfung korrekt ist
- Überprüfen Sie die Berechtigungen für den Startordner

Der Startordner kann über die Windows-Ausführen-Funktion mit `shell:startup` geöffnet werden.
