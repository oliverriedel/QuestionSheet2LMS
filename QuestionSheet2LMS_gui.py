# QuestionSheet2LMS
#
# Copyright (C) 2026 Oliver Riedel
#
# This program is free software: you can redistribute it and/or modify
# it under the terms of the GNU General Public License as published by
# the Free Software Foundation, either version 3 of the License, or
# (at your option) any later version.
#
# This program is distributed in the hope that it will be useful,
# but WITHOUT ANY WARRANTY; without even the implied warranty of
# MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.
#
# See the GNU General Public License for more details.

import tkinter as tk
from tkinter import filedialog, messagebox
from pathlib import Path
import os
import subprocess
import sys

import QuestionSheet2LMS

VERSION = "0.3.0-alpha"


class QuestionSheet2LMSGUI:
    def __init__(self, root):
        self.root = root

        self.root.title(
            f"QuestionSheet2LMS {VERSION}"
        )
        self.root.geometry("760x430")
        self.root.resizable(False, False)

        self.eingabedatei = tk.StringVar()
        self.letzte_zip_dateien = []
        self.letzter_zielordner = None

        self.zusammenfassung = tk.StringVar(
            value="Noch keine Datei geprüft."
        )

        self.status = tk.StringVar(
            value="Bitte eine XLSX-, ODS- oder CSV-Datei auswählen."
        )

        self.baue_gui()

    def baue_gui(self):
        titel = tk.Label(
            self.root,
            text="QuestionSheet2LMS – ILIAS-Export",
            font=("Arial", 16, "bold")
        )
        titel.pack(pady=(18, 6))

        untertitel = tk.Label(
            self.root,
            text="Fragenpool oder vollständigen Test aus einer Tabelle erzeugen",
            font=("Arial", 10)
        )
        untertitel.pack(pady=(0, 14))

        datei_frame = tk.Frame(self.root)
        datei_frame.pack(
            padx=20,
            fill="x"
        )

        datei_label = tk.Label(
            datei_frame,
            text="Eingabedatei:"
        )
        datei_label.pack(anchor="w")

        eingabe_frame = tk.Frame(datei_frame)
        eingabe_frame.pack(
            fill="x",
            pady=(5, 10)
        )

        datei_feld = tk.Entry(
            eingabe_frame,
            textvariable=self.eingabedatei,
            width=75,
            state="readonly"
        )
        datei_feld.pack(
            side="left",
            fill="x",
            expand=True
        )

        auswahl_button = tk.Button(
            eingabe_frame,
            text="Auswählen",
            command=self.datei_auswaehlen
        )
        auswahl_button.pack(
            side="left",
            padx=(10, 0)
        )

        zusammenfassung_label = tk.Label(
            self.root,
            textvariable=self.zusammenfassung,
            font=("Arial", 10, "bold"),
            justify="left"
        )
        zusammenfassung_label.pack(
            padx=20,
            anchor="w",
            pady=(4, 12)
        )

        export_label = tk.Label(
            self.root,
            text="Export erstellen:",
            font=("Arial", 10, "bold")
        )
        export_label.pack(
            padx=20,
            anchor="w",
            pady=(0, 5)
        )

        button_frame = tk.Frame(self.root)
        button_frame.pack(
            padx=20,
            fill="x",
            pady=(0, 12)
        )

        self.pool_button = tk.Button(
            button_frame,
            text="Fragenpool erstellen",
            width=20,
            height=2,
            state="disabled",
            command=lambda: self.export_erstellen(
                export_pool=True,
                export_test=False
            )
        )
        self.pool_button.pack(
            side="left",
            padx=(0, 7)
        )

        self.test_button = tk.Button(
            button_frame,
            text="Test erstellen",
            width=20,
            height=2,
            state="disabled",
            command=lambda: self.export_erstellen(
                export_pool=False,
                export_test=True
            )
        )
        self.test_button.pack(
            side="left",
            padx=7
        )

        self.beide_button = tk.Button(
            button_frame,
            text="Beides erstellen",
            width=20,
            height=2,
            state="disabled",
            command=lambda: self.export_erstellen(
                export_pool=True,
                export_test=True
            )
        )
        self.beide_button.pack(
            side="left",
            padx=7
        )

        self.zielordner_button = tk.Button(
            button_frame,
            text="Zielordner öffnen",
            width=18,
            height=2,
            state="disabled",
            command=self.zielordner_oeffnen
        )
        self.zielordner_button.pack(
            side="left",
            padx=(7, 0)
        )

        hinweis = tk.Label(
            self.root,
            text=(
                "Beim Testexport bitte nach dem Import in ILIAS noch "
                "Online-Status, Testpasswort und zeitliche Verfügbarkeit festlegen."
            ),
            justify="left",
            wraplength=715
        )
        hinweis.pack(
            padx=20,
            fill="x",
            pady=(0, 12)
        )

        info_button = tk.Button(
            self.root,
            text="Über QuestionSheet2LMS",
            command=self.zeige_info
        )
        info_button.pack(pady=(0, 10))

        status_label = tk.Label(
            self.root,
            textvariable=self.status,
            anchor="w",
            justify="left",
            wraplength=715
        )
        status_label.pack(
            padx=20,
            fill="x"
        )

    def setze_export_buttons(self, state):
        for button in (
            self.pool_button,
            self.test_button,
            self.beide_button
        ):
            button.config(state=state)

    def datei_auswaehlen(self):
        datei = filedialog.askopenfilename(
            title="Aufgabendatei auswählen",
            filetypes=[
                ("Tabellendateien", "*.xlsx *.ods *.csv"),
                ("Excel-Dateien", "*.xlsx"),
                ("OpenDocument-Tabellen", "*.ods"),
                ("CSV-Dateien", "*.csv"),
                ("Alle Dateien", "*.*")
            ]
        )

        if not datei:
            return

        self.eingabedatei.set(datei)
        self.letzte_zip_dateien = []
        self.letzter_zielordner = None
        self.zielordner_button.config(state="disabled")
        self.setze_export_buttons("disabled")

        try:
            self.status.set(
                "Datei wird geprüft ..."
            )
            self.root.update_idletasks()

            zeilen, fragen, gesamtpunkte = (
                QuestionSheet2LMS.bereite_export_vor(
                    Path(datei),
                    logger=lambda _text="": None
                )
            )

            self.zusammenfassung.set(
                f"{len(zeilen)} Tabellenzeilen eingelesen. "
                f"Daraus wurden {len(fragen)} Fragen erkannt. "
                f"Gesamtpunktzahl: "
                f"{QuestionSheet2LMS.format_punkte(gesamtpunkte)} Punkte."
            )

            self.status.set(
                "Datei erfolgreich geprüft. Export kann erstellt werden."
            )

            self.setze_export_buttons("normal")

        except Exception as e:
            self.zusammenfassung.set(
                "Die Datei konnte nicht erfolgreich geprüft werden."
            )
            self.status.set(
                "Fehler bei der Prüfung der Eingabedatei."
            )
            messagebox.showerror(
                "Fehler bei der Eingabedatei",
                str(e)
            )

    def zeige_info(self):
        
        messagebox.showinfo(
            "Über QuestionSheet2LMS",
            f"QuestionSheet2LMS {VERSION}\n\n"
            "Erstellt aus XLSX-, ODS- und CSV-Dateien "
            "ILIAS-Fragenpool- und Test-Importdateien.\n\n"
            "Unterstützt Bilder sowie lokale Audio- "
            "und Videodateien.\n\n"
            "Copyright (C) 2026 Oliver Riedel\n"
            "Lizenz: GPL-3.0-or-later"
        )        
        
        

    def zielordner_oeffnen(self):
        if self.letzter_zielordner is None:
            return

        zielordner = Path(
            self.letzter_zielordner
        )

        if not zielordner.is_dir():
            messagebox.showerror(
                "Zielordner nicht gefunden",
                "Der Zielordner existiert nicht mehr."
            )
            return

        try:
            if sys.platform == "win32":
                os.startfile(
                    str(zielordner)
                )

            elif sys.platform == "darwin":
                subprocess.Popen(
                    ["open", str(zielordner)]
                )

            else:
                subprocess.Popen(
                    ["xdg-open", str(zielordner)]
                )

        except Exception as e:
            messagebox.showerror(
                "Zielordner konnte nicht geöffnet werden",
                str(e)
            )

    def export_erstellen(
        self,
        export_pool,
        export_test
    ):
        dateiname = self.eingabedatei.get().strip()

        if not dateiname:
            messagebox.showwarning(
                "Keine Datei",
                "Bitte zuerst eine Eingabedatei auswählen."
            )
            return

        datei = Path(dateiname)

        if not datei.exists():
            messagebox.showerror(
                "Datei nicht gefunden",
                f"Die Datei wurde nicht gefunden:\n\n{datei}"
            )
            return

        self.setze_export_buttons("disabled")

        try:
            self.status.set(
                "Export wird erstellt ..."
            )
            self.root.update_idletasks()

            ergebnis = QuestionSheet2LMS.exportiere_datei(
                datei,
                export_pool=export_pool,
                export_test=export_test,
                logger=lambda _text="": None
            )

            self.letzte_zip_dateien = ergebnis["dateien"]
            self.letzter_zielordner = datei.parent

            self.zielordner_button.config(
                state="normal"
            )

            self.zusammenfassung.set(
                f"{ergebnis['tabellenzeilen']} Tabellenzeilen eingelesen. "
                f"Daraus wurden {ergebnis['fragen']} Fragen erkannt. "
                f"Gesamtpunktzahl: "
                f"{QuestionSheet2LMS.format_punkte(ergebnis['gesamtpunkte'])} Punkte."
            )

            dateinamen = "\n".join(
                p.name
                for p in ergebnis["dateien"]
            )

            self.status.set(
                "Export erfolgreich erstellt."
            )

            messagebox.showinfo(
                "Export erfolgreich",
                f"{ergebnis['fragen']} Fragen wurden erfolgreich exportiert.\n"
                f"Gesamtpunktzahl: "
                f"{QuestionSheet2LMS.format_punkte(ergebnis['gesamtpunkte'])} Punkte.\n\n"
                "Ausgabedatei(en):\n"
                f"{dateinamen}"
            )

        except Exception as e:
            self.status.set(
                "Fehler bei der Verarbeitung."
            )

            messagebox.showerror(
                "Fehler bei der Verarbeitung",
                "Die Eingabedatei konnte nicht verarbeitet werden.\n\n"
                f"{e}"
            )

        finally:
            self.setze_export_buttons("normal")


def main():
    root = tk.Tk()

    QuestionSheet2LMSGUI(
        root
    )

    root.mainloop()


if __name__ == "__main__":
    main()
