# QuestionSheet2LMS

**QuestionSheet2LMS** ist ein Werkzeug zur Erstellung importierbarer
LMS-Fragenpools und vollständiger Testobjekte aus tabellarisch erfassten
Prüfungs- und Übungsfragen.

Aufgaben werden in XLSX-, ODS- oder CSV-Dateien erfasst. QuestionSheet2LMS
prüft die Eingaben, bereitet die Fragetypen auf und erzeugt daraus
importierbare ILIAS-Fragenpool-Dateien oder vollständige ILIAS-Testobjekte.

Der Schwerpunkt der aktuellen Alpha-Version liegt auf ILIAS.

## Download für Windows

[Aktuelle Windows-Version herunterladen](https://github.com/oliverriedel/QuestionSheet2LMS/releases)

Die Windows-Version enthält eine direkt ausführbare EXE sowie Beispieldateien, PDF-Anleitungen und Lizenzinformationen. **Eine Python-Installation ist nicht erforderlich.**

ZIP-Datei herunterladen, vollständig entpacken und QuestionSheet2LMS.exe starten.

## 📜 Hintergrund & Entstehung

Der Ausgangspunkt für QuestionSheet2LMS war die langjährige Nutzung und
Anpassung des R-Pakets
[RIlias_public von jpburgard](https://github.com/jpburgard/RIlias_public).

RIlias ermöglicht die programmgesteuerte Erzeugung von ILIAS-Fragen aus R.
Im Laufe der praktischen Nutzung entstand daraus ein eigener Workflow, bei
dem Fragen zunehmend tabellarisch vorbereitet und anschließend automatisiert
in ILIAS-Importdateien überführt wurden.

Dabei zeigte sich insbesondere bei der Zusammenarbeit mit Kolleginnen und
Kollegen, dass eine Tabelle eine sehr niedrige Einstiegshürde für die
Erstellung größerer Fragenbestände bietet.

QuestionSheet2LMS greift diese Erfahrung auf, ist jedoch eine eigenständige
Python-Implementierung. Die tabellarische Eingabe, die Verarbeitung von
XLSX-, ODS- und CSV-Dateien, die Validierung der Eingaben sowie die
grafische Benutzeroberfläche wurden unabhängig neu entwickelt.

Die Entwicklung, das Refactoring und die Umsetzung einzelner Logiken wurden
interaktiv durch generative KI unterstützt. Die erzeugten Importstrukturen
werden anhand nativer Exportdateien der jeweiligen LMS sowie durch praktische
Import- und Funktionstests entwickelt.

## ⚙️ Features

- **Tabellarische Eingabe:** Fragen können in XLSX, ODS oder CSV erfasst werden.
- **Eingabeprüfung:** Typische logische Fehler werden bereits vor dem Export
  erkannt und möglichst mit der betroffenen `Frage_ID` ausgegeben.
- **ILIAS-QTI/QPL-Export:** Erzeugung kompakter ILIAS-Fragenpoolimporte.
- **ILIAS-Testexport:** Erzeugung vollständiger ILIAS-Testobjekte mit den
  enthaltenen Fragen und vordefinierten Prüfungseinstellungen. Prüfungsindividuelle
  Angaben wie Online-Status, Testpasswort und zeitliche Verfügbarkeit werden
  anschließend in ILIAS gesetzt.
- **Mehrere Fragetypen:** SC, MC, KPrim, Freitext, vertikale und horizontale
  Anordnung sowie Lückentextfragen mit Text-, Auswahl- und Zahllücken.
- **Mehrere MC-Bewertungsverfahren:** All-or-Nothing, Nur-Richtige,
  Richtig-minus-Falsch und Bewertung jeder Aussage.
- **Bilder:** Abbildungen können automatisch als ILIAS Media Objects in den
  Fragenpool aufgenommen werden. Die Bilddateien müssen im selben Verzeichnis
  wie die ODS/XLSX/CSV-Datei liegen.
- **Audio-Dateien:** Audio-Dateien können automatisch als ILIAS Media Objects
  in den Fragenpool aufgenommen werden. Die Audio-Dateien müssen im selben 
  Verzeichnis wie die ODS/XLSX/CSV-Datei liegen und im mp3-Format vorliegen.
- **Video-Dateien:** Video-Dateien können automatisch als ILIAS Media Objects
  in den Fragenpool aufgenommen werden. Die Video-Dateien müssen im selben 
  Verzeichnis wie die ODS/XLSX/CSV-Datei liegen und im mp4- oder webm-Format
  vorliegen.
- **Sonderzeichen und Formeln:** Unterstützung von Texten mit Sonderzeichen
  und für ILIAS geeigneten LaTeX-/MathJax-Inhalten.
- **Grafische Benutzeroberfläche:** Nutzung unter Windows ohne installierte
  Python-Umgebung über eine eigenständige Anwendung. Die GUI unterstützt die
  Erstellung eines Fragenpools, eines vollständigen Tests oder beider Exporte
  und zeigt vor dem Export die erkannte Fragenzahl sowie die Gesamtpunktzahl an.

## 🧪 Getestete ILIAS-Versionen

| Fragetyp | ILIAS 5 | ILIAS 9 | ILIAS 11 |
| --- | :---: | :---: | :---: |
| Single Choice | ✅ | ✅ | ✅ |
| Multiple Choice | ✅ | ✅ | ✅ |
| KPrim | ✅ | ✅ | ✅ |
| Freitext | ✅ | ✅ | ✅ |
| Vertikale Anordnung | ✅ | ✅ | ✅ |
| Horizontale Anordnung | ✅ | ✅ | ✅ |
| Lückentext | ✅ | ✅ | ✅ |
| Bild im Fragetext | ✅ | ✅ | ✅ |
| Audio im Fragetext | ✅ | ✅ | ✅ |
| Video im Fragetext | ✅ | ✅ | ✅ |


Die Tests beziehen sich auf die jeweils von QuestionSheet2LMS erzeugten
Importdateien und die im Projekt dokumentierten Testfälle.

Der vollständige Testexport wurde zusätzlich mit ILIAS 5, ILIAS 9
und ILIAS 11 getestet. Unter ILIAS 5 muss nach dem Import lediglich
die Testbeschreibung angepasst werden.

## 🚧 Weitere Zielsysteme

Die interne Datenstruktur ist so angelegt, dass zukünftig weitere
Exportformate ergänzt werden können.

Vorgesehen bzw. untersucht werden insbesondere:

- Moodle XML
- EvaExam CSV

Diese Exportwege sind derzeit noch nicht Bestandteil der Alpha-Version.

## 📖 Dokumentation

Beispieldateien im XLSX- und ODS-Format zeigen die unterstützten Felder,
Fragetypen sowie die Verwendung von Bildern, Audio- und Videodateien.

Zusätzlich steht eine bebilderte Anleitung zur Verfügung. Sie zeigt
für die unterstützten Fragetypen die benötigten Eingabefelder und die
daraus resultierende Darstellung in ILIAS 
(Übersicht_über_die_Aufgabentypen.pdf).

Eine weitere Kurzanleitung beschreibt den Import der mit
QuestionSheet2LMS erzeugten vollständigen Tests und Fragenpools in ILIAS
(Import_von_QuestionsSheet2LMS_in_ILIAS.pdf).

## 🔒 Datenschutz & Sicherheit

QuestionSheet2LMS verarbeitet die ausgewählten XLSX-, ODS- und CSV-Dateien
lokal auf dem Rechner des Anwenders. Der Anwendungscode überträgt dabei
keine Fragen, Tabelleninhalte oder Nutzungsdaten an externe Server und
enthält keine Telemetrie.

Das optionale Feld `Autor` wird in die erzeugten ILIAS-Importdateien
übernommen. Wer dort Klarnamen einträgt, sollte dies insbesondere bei der
Weitergabe erzeugter ZIP-Dateien berücksichtigen.

Verwendete Bilddateien müssen sich im selben Verzeichnis wie die
Eingabedatei befinden. Absolute Pfade, Unterverzeichnisse und Pfadangaben
außerhalb dieses Verzeichnisses werden nicht akzeptiert.

Wie bei anderen lokal verarbeiteten Dokumenten sollten Dateien aus
fremden oder nicht vertrauenswürdigen Quellen vor der Verwendung geprüft
werden.

## ⚖️ Lizenz

Copyright © 2026 Oliver Riedel

QuestionSheet2LMS ist freie Software und wird unter den Bedingungen der
GNU General Public License Version 3 oder jeder späteren Version
veröffentlicht.

Siehe [LICENSE](LICENSE).

Das Programm wird ohne Gewährleistung bereitgestellt.

## Markenhinweis

ILIAS, Moodle und EvaExam sind Marken ihrer jeweiligen Rechteinhaber.

QuestionSheet2LMS ist ein unabhängiges Open-Source-Projekt und steht in
keiner organisatorischen oder wirtschaftlichen Verbindung mit den
Herstellern bzw. Betreibern der genannten Systeme.
