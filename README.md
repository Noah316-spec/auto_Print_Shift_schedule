# Projektbeschreibung

Dieses Projekt enthält eine Reihe von Funktionen, die verschiedene Aufgaben ausführen, wie z.B. das Drucken von Excel-Dateien und das Durchführen von Abfragen basierend auf dem aktuellen Datum und Wochentag.

## Hauptfunktionen

1. `Form1()`: Diese Funktion initialisiert die Komponenten und startet einen Timer, der alle Sekunden ein Ereignis auslöst.

2. `GetQuarterEnd(DateTime date)`: Diese Funktion berechnet das Ende des aktuellen Quartals basierend auf dem übergebenen Datum.

3. `druck(string pfad)`: Diese Funktion öffnet eine Excel-Datei am angegebenen Pfad, druckt das gesamte Arbeitsbuch aus und schließt dann das Arbeitsbuch und die Anwendung.

4. `abfragefrüMO(string pfad)`: Diese Funktion führt eine Abfrage für die Frühschicht am Montag durch und druckt eine Excel-Datei basierend auf dem aktuellen Wochentag.

5. `btnDruck_Click(object sender, EventArgs e)`: Diese Funktion wird ausgeführt, wenn der Druckknopf geklickt wird. Sie führt verschiedene Aktionen aus, abhängig vom ausgewählten Element in einer ComboBox und dem aktuellen Datum.

6. `timer1_Tick(object sender, EventArgs e)`: Diese Funktion wird bei jedem Tick des Timers aufgerufen und aktualisiert die aktuellen Datumswerte.

Bitte ersetzen Sie `"ihrpfad"` durch den tatsächlichen Pfad zu Ihrer Excel-Datei, bevor Sie diese Funktionen verwenden.
