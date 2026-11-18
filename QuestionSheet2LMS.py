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

import csv
import sys
import warnings
from pathlib import Path
from datetime import datetime
from zipfile import ZipFile, ZIP_DEFLATED
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP
import xml.etree.ElementTree as ET
import shutil
import base64

try:
    from openpyxl import load_workbook
except ImportError:
    load_workbook = None


try:
    from odf.opendocument import load as load_ods
    from odf.table import Table, TableRow
except ImportError:
    load_ods = None
    Table = None
    TableRow = None



ANTWORT_BUCHSTABEN = "abcdefghijklmnopqrstuvwxyz"

# Großzügige Obergrenze für lokale Eingabedateien.
# Verhindert versehentliche bzw. missbräuchlich riesige Dateien.
MAX_EINGABEDATEI_GROESSE = 100 * 1024 * 1024  # 100 MiB


def textwert(wert):
    """
    Wandelt Zellinhalte robust in Text um.
    Leere Zellen werden zu "".
    """
    if wert is None:
        return ""

    return str(wert).strip()


def format_punkte(wert):
    """
    Formatiert Punktwerte ohne binäre Float-Rundungsartefakte.

    Intern werden Punktwerte bevorzugt als Decimal verarbeitet.
    Ganzzahlige Werte werden ohne Nachkommastellen ausgegeben.
    """
    if isinstance(wert, Decimal):
        dezimal = wert
    else:
        dezimal = Decimal(str(wert))

    return format_dezimal(dezimal)

def dezimalwert(wert, feldname):
    """
    Liest eine Dezimalzahl robust ein.

    Akzeptiert sowohl Punkt als auch Komma
    als Dezimaltrennzeichen.
    """


    text = textwert(
        wert
    )

    if not text:

        raise ValueError(
            f"{feldname} darf nicht leer sein."
        )

    #
    # Tausendertrennzeichen werden nicht unterstützt.
    #
    # Beispiele:
    # 1.234,56
    # 1,234.56
    #
    # Solche Schreibweisen enthalten Punkt und Komma
    # gleichzeitig und sind daher nicht eindeutig erkennbar.
    #
    if "." in text and "," in text:

        raise ValueError(
            f"Ungültiger Zahlenwert in '{feldname}': {text}. "
            "Bitte Zahlen ohne Tausendertrennzeichen eingeben, "
            "z. B. 1234,56 oder 1234.56."
        )

    #
    # Deutsches Dezimalkomma für die weitere
    # Verarbeitung in einen Dezimalpunkt umwandeln.
    #
    text = text.replace(
        ",",
        "."
    )


    try:

        return Decimal(
            text
        )

    except InvalidOperation:

        raise ValueError(
            f"Ungültiger Zahlenwert in '{feldname}': {wert}"
        )



def punktewert(wert):
    """
    Liest eine Punktzahl robust ein.

    Akzeptiert Punkt oder Komma als Dezimaltrennzeichen.
    Die Punktzahl muss größer als 0 sein.
    """

    punkte = dezimalwert(
        wert,
        "Punkte"
    )

    if punkte <= 0:

        raise ValueError(
            "Die Punktzahl muss größer als 0 sein."
        )

    # Punktwerte als Decimal beibehalten, damit Teilungen
    # nicht durch binäre Float-Rundungsfehler verfälscht werden.
    return punkte


def format_dezimal(wert):
    """
    Formatiert Decimal-Werte ohne unnötige
    Nachkommastellen.
    """

    text = format(
        wert,
        "f"
    )

    if "." in text:

        text = (
            text
            .rstrip("0")
            .rstrip(".")
        )

    if text == "-0":
        text = "0"

    return text


def verteile_punkte(gesamtpunkte, anzahl, nachkommastellen=3):
    """
    Verteilt eine Gesamtpunktzahl möglichst gleichmäßig
    auf mehrere Teilwerte.

    Die Teilwerte werden auf die gewünschte Zahl von
    Nachkommastellen gerundet. Ein Rundungsrest wird
    dem letzten Teilwert zugeschlagen, sodass die Summe
    exakt der Gesamtpunktzahl entspricht.
    """

    if anzahl < 1:
        raise ValueError(
            "Die Anzahl der Teilwerte muss mindestens 1 sein."
        )

    quantisierung = Decimal(
        "1." + ("0" * nachkommastellen)
    )

    basiswert = (
        gesamtpunkte
        / Decimal(anzahl)
    ).quantize(
        quantisierung,
        rounding=ROUND_HALF_UP
    )

    werte = [
        basiswert
        for _ in range(anzahl)
    ]

    werte[-1] = (
        gesamtpunkte
        - sum(werte[:-1], Decimal("0"))
    )

    return werte



def format_prozentwert(wert):
    """
    Formatiert einen intern als Dezimalzahl gespeicherten
    Prozentwert für die tabellarische Weiterverarbeitung.

    Beispiel:
        0.186 -> "18,6 %"
        0.2   -> "20 %"
    """

    prozentwert = float(wert) * 100

    if prozentwert.is_integer():

        return f"{int(prozentwert)} %"

    return (
        f"{prozentwert:.10f}"
        .rstrip("0")
        .rstrip(".")
        .replace(".", ",")
        + " %"
    )

def roemische_zahl(zahl):
    """
    Wandelt eine positive ganze Zahl
    in eine kleine römische Zahl um.
    """

    if zahl < 1:

        raise ValueError(
            "Römische Nummerierung beginnt bei 1."
        )

    werte = (
        (1000, "m"),
        (900, "cm"),
        (500, "d"),
        (400, "cd"),
        (100, "c"),
        (90, "xc"),
        (50, "l"),
        (40, "xl"),
        (10, "x"),
        (9, "ix"),
        (5, "v"),
        (4, "iv"),
        (1, "i")
    )

    rest = zahl
    ergebnis = ""

    for wert, zeichen in werte:

        while rest >= wert:

            ergebnis += zeichen
            rest -= wert

    return ergebnis


def titel_mit_punkten(zeile):
    """
    Ergänzt bei Punktanzeige = ja die Gesamtpunktzahl
    in der Überschrift.
    """

    titel = textwert(
        zeile.get("Ueberschrift", "")
    )

    punktanzeige = textwert(
        zeile.get("Punktanzeige", "")
    ).lower()

    if punktanzeige != "ja":
        return titel


    punkte = punktewert(
        zeile.get(
            "Punkte",
            ""
        )
    )


    if punkte == 1:
        punktetext = "1 Punkt"

    else:
        punktetext = (
            format_punkte(punkte)
            .replace(".", ",")
            + " Punkte"
        )


    return (
        f"{titel} ({punktetext})"
    )



def lese_csv(datei):
    """
    Liest CSV anhand der Spaltenüberschriften ein.
    Trenner ; , oder Tab wird automatisch erkannt.
    """
    with open(datei, "r", encoding="utf-8-sig", newline="") as f:
        probe = f.read(4096)
        f.seek(0)

        dialect = csv.Sniffer().sniff(
            probe,
            delimiters=";,\t"
        )

        reader = csv.DictReader(f, dialect=dialect)

        daten = []

        for zeile in reader:
            neue_zeile = {}

            for key, value in zeile.items():
                if key is not None:
                    neue_zeile[key.strip()] = textwert(value)

            if any(neue_zeile.values()):
                daten.append(neue_zeile)

        return daten



def lese_xlsx(datei):
    """
    Liest das Tabellenblatt 'Aufgaben' einer XLSX-Datei ein.
    Prozentformatierte Zellen werden als Prozenttext übernommen.
    """
    if load_workbook is None:
        raise RuntimeError(
            "Für XLSX-Dateien wird openpyxl benötigt.\n"
            "Installation: python3 -m pip install openpyxl"
        )


    warnings.filterwarnings(
        "ignore",
        message=(
            "Data Validation extension is not supported "
            "and will be removed"
        ),
        category=UserWarning,
        module="openpyxl"
    )



    wb = load_workbook(
        datei,
        read_only=True,
        data_only=True
    )

    if "Aufgaben" not in wb.sheetnames:

        raise ValueError(
            "Die XLSX-Datei benötigt ein "
            "Tabellenblatt mit dem Namen 'Aufgaben'."
        )

    ws = wb["Aufgaben"]
    
    zeilen = ws.iter_rows()

    try:
        header_zeile = next(zeilen)
    except StopIteration:
        return []

    header = [
        textwert(zelle.value)
        for zelle in header_zeile
    ]

    daten = []

    for zellen in zeilen:
        zeile = {}

        for spalte, zelle in zip(header, zellen):
            if not spalte:
                continue

            wert = zelle.value

            if wert is None:
                zeile[spalte] = ""
                continue

            #
            # Prozentformat erkennen.
            #
            # Excel speichert z. B. 18,6 % intern als 0.186.
            #
            if (
                isinstance(wert, (int, float))
                and "%"
                in str(zelle.number_format)
            ):
            
                zeile[spalte] = format_prozentwert(
                    wert
                )            
            
            else:
                zeile[spalte] = textwert(wert)

        if any(zeile.values()):
            daten.append(zeile)

    return daten



def lese_ods(datei):
    """
    Liest das Tabellenblatt 'Aufgaben' einer ODS-Datei ein.

    Prozentformatierte Zellen werden analog zu XLSX
    als Prozenttext übernommen.
    """

    if load_ods is None:

        raise RuntimeError(
            "Für ODS-Dateien wird odfpy benötigt.\n"
            "Installation: python3 -m pip install odfpy"
        )

    dokument = load_ods(
        str(datei)
    )

    #
    # Tabellenblatt "Aufgaben" suchen.
    #
    aufgaben_tabelle = None

    for tabelle in dokument.spreadsheet.getElementsByType(
        Table
    ):

        if tabelle.getAttribute("name") == "Aufgaben":

            aufgaben_tabelle = tabelle
            break

    if aufgaben_tabelle is None:

        raise ValueError(
            "Die ODS-Datei benötigt ein "
            "Tabellenblatt mit dem Namen 'Aufgaben'."
        )

    #
    # Text aus einer ODS-Zelle auslesen.
    #


    def zelltext(zelle):
        """
        Liest den sichtbaren Zelltext aus.

        ODS-Kommentare/Notizen werden ignoriert,
        damit z. B. kommentierte Spaltenüberschriften
        weiterhin exakt als 'Frage_ID' erkannt werden.
        """

        teile = []

        def sammle_text(knoten):

            #
            # Kommentare in ODS liegen als
            # office:annotation innerhalb der Zelle.
            # Deren Inhalt gehört nicht zum Zellwert.
            #
            if (
                getattr(knoten, "qname", None)
                and knoten.qname[0]
                == "urn:oasis:names:tc:opendocument:xmlns:office:1.0"
                and knoten.qname[1] == "annotation"
            ):

                return

            for kind in knoten.childNodes:

                if getattr(kind, "data", None) is not None:

                    teile.append(
                        kind.data
                    )

                sammle_text(
                    kind
                )

        sammle_text(
            zelle
        )

        return "".join(
            teile
        ).strip()


    #
    # Eine Tabellenzeile auslesen.
    #
    # ODS kann gleiche leere Zellen intern über
    # number-columns-repeated zusammenfassen.
    #
    def lese_ods_zeile(
        tabellenzeile,
        maximale_spalten=None
    ):

        werte = []

        for zelle in tabellenzeile.childNodes:

            if (
                getattr(zelle, "qname", None) is None
                or zelle.qname[1] != "table-cell"
            ):

                continue

            wiederholungen = (
                zelle.getAttribute(
                    "numbercolumnsrepeated"
                )
                or 1
            )

            wiederholungen = int(
                wiederholungen
            )

            werttyp = zelle.getAttribute(
                "valuetype"
            )

            #
            # Zahlen werden aus dem eigentlichen
            # numerischen ODS-Wert gelesen.
            #
            if werttyp in (
                "float",
                "currency"
            ):

                wert = textwert(
                    zelle.getAttribute(
                        "value"
                    )
                )

            #
            # Prozentwerte:
            # ODS speichert z. B. 18,6 % als 0.186.
            #
            elif werttyp == "percentage":

                rohwert = zelle.getAttribute(
                    "value"
                )

                if rohwert is None:

                    wert = zelltext(
                        zelle
                    )

                else:


                    wert = format_prozentwert(
                        rohwert
                    )


            #
            # Sonstige Inhalte als dargestellten
            # Zelltext übernehmen.
            #
            else:

                wert = zelltext(
                    zelle
                )

            #
            # Wiederholte ODS-Zellen auflösen.
            #
            for _ in range(
                wiederholungen
            ):

                werte.append(
                    wert
                )

                if (
                    maximale_spalten is not None
                    and len(werte) >= maximale_spalten
                ):

                    return werte

        return werte

    tabellenzeilen = (
        aufgaben_tabelle.getElementsByType(
            TableRow
        )
    )

    if not tabellenzeilen:

        return []

    #
    # Erste Zeile = Spaltenüberschriften.
    #
    header = [
        textwert(wert)
        for wert in lese_ods_zeile(
            tabellenzeilen[0]
        )
    ]

    #
    # Hintere leere ODS-Spalten entfernen.
    #
    while (
        header
        and not header[-1]
    ):

        header.pop()

    if not header:

        return []

    daten = []

    #
    # Datenzeilen verarbeiten.
    #
    for tabellenzeile in tabellenzeilen[1:]:

        werte = lese_ods_zeile(
            tabellenzeile,
            maximale_spalten=len(header)
        )

        #
        # Fehlende hintere Zellen auffüllen.
        #
        while len(werte) < len(header):

            werte.append(
                ""
            )

        zeile = {}

        for spalte, wert in zip(
            header,
            werte
        ):

            if not spalte:

                continue

            zeile[spalte] = textwert(
                wert
            )

        if any(
            zeile.values()
        ):

            daten.append(
                zeile
            )

    return daten



def lese_datei(datei):
    """
    Entscheidet anhand der Dateiendung,
    ob CSV, ODS oder XLSX gelesen wird.
    """
    datei = Path(datei)

    if datei.stat().st_size > MAX_EINGABEDATEI_GROESSE:
        raise ValueError(
            "Die Eingabedatei ist größer als 100 MiB. "
            "Bitte prüfen Sie, ob die richtige Datei ausgewählt wurde."
        )

    endung = datei.suffix.lower()

    if endung == ".csv":
        return lese_csv(datei)

    if endung == ".xlsx":
        return lese_xlsx(datei)

    if endung == ".ods":
        return lese_ods(datei)

    raise ValueError(
        f"Nicht unterstütztes Dateiformat: {endung}"
    )

def hole_antworten(zeile):
    """
    Liest Antwort A bis Antwort L dynamisch ein.
    """
    antworten = {}

    for buchstabe in ANTWORT_BUCHSTABEN:
        spaltenname = f"Antwort {buchstabe.upper()}"

        if spaltenname not in zeile:
            continue

        antwort = textwert(
            zeile.get(spaltenname, "")
        )

        if antwort:
            antworten[buchstabe] = antwort

    return antworten


def hole_richtige_antworten(zeile):
    """
    Verarbeitet z. B.:
    a,b
    a, b
    A,B
    A, B
    """
    rohwert = textwert(
        zeile.get("Richtige Antwort", "")
    )

    if not rohwert:
        raise ValueError(
            "Keine richtige Antwort angegeben."
        )

    richtige = [
        teil.strip().lower()
        for teil in rohwert.split(",")
        if teil.strip()
    ]


    if len(richtige) != len(set(richtige)):

        raise ValueError(
            "In 'Richtige Antwort' darf dieselbe "
            "Antwort nicht mehrfach angegeben werden."
        )

    return richtige


def pruefe_antworten(
    antworten,
    richtige,
    frage_id=""
):
    """
    Prüft, ob alle als richtig angegebenen Antworten
    tatsächlich vorhanden sind.
    """

    praefix = (
        f"Frage_ID {frage_id}: "
        if frage_id
        else ""
    )

    for buchstabe in richtige:

        if buchstabe not in ANTWORT_BUCHSTABEN:

            raise ValueError(
                f"{praefix}"
                f"Ungültige Antwortbezeichnung: {buchstabe}"
            )

        if buchstabe not in antworten:

            raise ValueError(
                f"{praefix}"
                f"Antwort {buchstabe.upper()} ist als richtig angegeben, "
                "aber nicht befüllt."
            )


        
def gemeinsame_fragefelder(zeile):
    """
    Liest die allgemeinen Felder einer Frage ein,
    die von mehreren Fragetypen verwendet werden.
    """


    audio = textwert(
        zeile.get(
            "Audio",
            ""
        )
    )

    video = textwert(
        zeile.get(
            "Video",
            ""
        )
    )

    if audio and video:
        raise ValueError(
            "Audio und Video können nicht gleichzeitig "
            "in derselben Frage verwendet werden. "
            "Bitte verwenden Sie entweder eine Audiodatei "
            "oder eine Videodatei."
        )



    return {
        "einleitung": textwert(
            zeile.get(
                "Einleitungstext",
                ""
            )
        ),

        "bild": textwert(
            zeile.get(
                "Bild",
                ""
            )
        ),


        "audio": audio,

        "video": video,




        "frage": textwert(
            zeile.get(
                "Frage",
                ""
            )
        ),

        "autor": textwert(
            zeile.get(
                "Autor",
                ""
            )
        ),

        "rueckmeldung": textwert(
            zeile.get(
                "Rueckmeldung",
                ""
            )
        )
    }        
        
        

def bereite_sc_vor(zeile):
    """
    Bereitet eine SingleChoice-Frage auf.
    """
    antworten = hole_antworten(zeile)
    

    if len(antworten) < 2:
        raise ValueError(
            "SingleChoice benötigt mindestens zwei Antworten."
        )



    erwartete_buchstaben = list(
        ANTWORT_BUCHSTABEN[
            :len(antworten)
        ]
    )

    if list(antworten.keys()) != erwartete_buchstaben:
        raise ValueError(
            "Bei SingleChoice müssen die Antworten "
            "ohne Lücken ab Antwort A ausgefüllt sein."
        )


    if len(set(antworten.values())) != len(antworten):
        raise ValueError(
            "Bei SingleChoice darf derselbe Antworttext "
            "nicht mehrfach verwendet werden."
        )


   
    richtige = hole_richtige_antworten(zeile)


    pruefe_antworten(
        antworten,
        richtige,
        textwert(
            zeile.get(
                "Frage_ID",
                ""
            )
        )
    )
    
    if len(richtige) != 1:
        raise ValueError(
            "SingleChoice benötigt genau eine richtige Antwort."
        )

    richtig = richtige[0]



    answer = list(
        antworten.values()
    )

    richtige_position = (
        list(
            antworten.keys()
        ).index(richtig)
        + 1
    )



    gemeinsam = gemeinsame_fragefelder(
        zeile
    )

    return {
        "typ": "SC",
        "titel": titel_mit_punkten(zeile),



        "einleitung": gemeinsam["einleitung"],
        "bild": gemeinsam["bild"],
        "audio": gemeinsam["audio"],
        "video": gemeinsam["video"],
        "frage": gemeinsam["frage"],



        "antworten": answer,
        "richtige_position": richtige_position,
        "anzahl_richtige": 1,
        "anzahl_falsche": len(answer) - 1,


        "punkte": punktewert(
            zeile.get(
                "Punkte",
                ""
            )
        ),


        "autor": gemeinsam["autor"],
        "rueckmeldung": gemeinsam["rueckmeldung"]


    }



def bereite_mc_vor(zeile):
    """
    Bereitet eine MultipleChoice-Frage auf.
    """
    antworten = hole_antworten(zeile)
    
    if len(antworten) < 2:
        raise ValueError(
            "MultipleChoice benötigt mindestens zwei Antworten."
        )
    

    erwartete_buchstaben = list(
        ANTWORT_BUCHSTABEN[
            :len(antworten)
        ]
    )

    if list(antworten.keys()) != erwartete_buchstaben:
        raise ValueError(
            "Bei MultipleChoice müssen die Antworten "
            "ohne Lücken ab Antwort A ausgefüllt sein."
        )


    if len(set(antworten.values())) != len(antworten):
        raise ValueError(
            "Bei MultipleChoice darf derselbe Antworttext "
            "nicht mehrfach verwendet werden."
        )
    
    richtige = hole_richtige_antworten(zeile)

    pruefe_antworten(
        antworten,
        richtige,
        textwert(
            zeile.get(
                "Frage_ID",
                ""
            )
        )
    )
    





    answer = list(
        antworten.values()
    )

    wahrheitswerte = [
        buchstabe in richtige
        for buchstabe in antworten.keys()
    ]



    gesamtpunkte = punktewert(
        zeile.get(
            "Punkte",
            ""
        )
    )


    mc_bewertungsvariante = textwert(
        zeile.get(
            "MC_Bewertungsvariante",
            ""
        )
    ).lower()

    erlaubte_bewertungsvarianten = (
        "all-or-nothing",
        "nur-richtige",
        "richtig-minus-falsch",
        "jede-aussage"
    )

    if (
        mc_bewertungsvariante
        not in erlaubte_bewertungsvarianten
    ):
        raise ValueError(
            "Ungültige MC_Bewertungsvariante: "
            f"{mc_bewertungsvariante}. "
            "Erlaubt sind: "
            "all-or-nothing, "
            "nur-richtige, "
            "richtig-minus-falsch, "
            "jede-aussage."
        )



    if (
        mc_bewertungsvariante
        == "richtig-minus-falsch"
        and all(wahrheitswerte)
    ):
        raise ValueError(
            "Die Bewertungsvariante "
            "'richtig-minus-falsch' benötigt "
            "mindestens eine falsche Antwort."
        )


    gemeinsam = gemeinsame_fragefelder(
        zeile
    )


    return {
        "typ": "MC",
        "titel": titel_mit_punkten(zeile),

        "einleitung": gemeinsam["einleitung"],
        "bild": gemeinsam["bild"],
        "audio": gemeinsam["audio"],
        "video": gemeinsam["video"],
        "frage": gemeinsam["frage"],



        "antworten": answer,
        "wahrheitswerte": wahrheitswerte,
        "anzahl_richtige": sum(
            wahrheitswerte
        ),
        "anzahl_falsche": (
            len(wahrheitswerte)
            - sum(wahrheitswerte)
        ),
        "punkte": gesamtpunkte,


        "mc_bewertungsvariante": (
            mc_bewertungsvariante
        ),


        "autor": gemeinsam["autor"],
        "rueckmeldung": gemeinsam["rueckmeldung"]


    }

def bereite_kprim_vor(zeile):
    """
    Bereitet eine KPrim-Frage auf.
    """

    antworten = hole_antworten(zeile)

    #
    # KPrim benötigt genau vier Aussagen:
    # Antwort A bis Antwort D.
    #
    erwartete_antworten = (
        "a",
        "b",
        "c",
        "d"
    )

    for buchstabe in erwartete_antworten:

        if buchstabe not in antworten:
            raise ValueError(
                "KPrim benötigt genau vier ausgefüllte "
                "Antworten A bis D."
            )

    #
    # Weitere Antwortfelder dürfen bei KPrim
    # nicht befüllt sein.
    #
    weitere_antworten = [
        buchstabe
        for buchstabe in antworten
        if buchstabe not in erwartete_antworten
    ]

    if weitere_antworten:
        raise ValueError(
            "KPrim unterstützt genau vier Aussagen "
            "in Antwort A bis Antwort D."
        )

    #
    # Die Reihenfolge A, B, C, D muss erhalten bleiben.
    #
    answer = [
        antworten["a"],
        antworten["b"],
        antworten["c"],
        antworten["d"]
    ]


    if len(set(answer)) != len(answer):
        raise ValueError(
            "Bei KPrim darf derselbe Aussage-Text "
            "nicht mehrfach verwendet werden."
        )

    #
    # Richtige Aussagen auslesen.
    #
    # Ein leeres Feld bedeutet:
    # Keine der vier Aussagen ist richtig.
    #
    rohwert = textwert(
        zeile.get("Richtige Antwort", "")
    )

    if rohwert:

        richtige = [
            teil.strip().lower()
            for teil in rohwert.split(",")
            if teil.strip()
        ]

    else:

        richtige = []
        
        
    if len(richtige) != len(set(richtige)):
        raise ValueError(
            "Bei KPrim darf dieselbe Antwort in "
            "'Richtige Antwort' nicht mehrfach angegeben werden."
        )        
        

    for buchstabe in richtige:

        if buchstabe not in erwartete_antworten:
            raise ValueError(
                "Bei KPrim dürfen in 'Richtige Antwort' "
                "nur a, b, c und d angegeben werden."
            )

    #
    # Für jede These merken:
    # 1 = richtig
    # 0 = falsch
    #
    wahrheitswerte = [
        1 if buchstabe in richtige else 0
        for buchstabe in erwartete_antworten
    ]

    gemeinsam = gemeinsame_fragefelder(
        zeile
    )


    return {
        "typ": "KPRIM",
        "titel": titel_mit_punkten(zeile),


        "einleitung": gemeinsam["einleitung"],
        "bild": gemeinsam["bild"],
        "audio": gemeinsam["audio"],
        "video": gemeinsam["video"],
        "frage": gemeinsam["frage"],


        "antworten": answer,
        "wahrheitswerte": wahrheitswerte,

        "punkte": punktewert(
            zeile.get(
                "Punkte",
                ""
            )
        ),


        "autor": gemeinsam["autor"],
        "rueckmeldung": gemeinsam["rueckmeldung"]


    }




def bereite_anordnung_vertikal_vor(zeile):
    """
    Bereitet eine vertikale Anordnungsfrage vor.

    Die Antworten A, B, C, ... werden bereits
    in der richtigen Reihenfolge angegeben.
    """

    antworten = hole_antworten(
        zeile
    )

    if len(antworten) < 2:
        raise ValueError(
            "Eine vertikale Anordnungsfrage "
            "benötigt mindestens zwei Antworten."
        )

    #
    # Bei einer Anordnungsfrage dürfen innerhalb
    # der Antworten keine Lücken vorkommen.
    #
    # Also z. B. A, B, C, D - nicht A, B, D.
    #
    erwartete_buchstaben = list(
        ANTWORT_BUCHSTABEN[
            :len(antworten)
        ]
    )

    vorhandene_buchstaben = list(
        antworten.keys()
    )

    if vorhandene_buchstaben != erwartete_buchstaben:
        raise ValueError(
            "Bei einer vertikalen Anordnungsfrage "
            "müssen die Antworten ohne Lücken "
            "ab Antwort A ausgefüllt sein."
        )


    if len(set(antworten.values())) != len(antworten):
        raise ValueError(
            "Bei einer vertikalen Anordnungsfrage darf "
            "derselbe Antworttext nicht mehrfach verwendet werden. "
            "Ein LMS kann identische Antworttexte technisch unterschiedlich "
            "behandeln und dadurch ggf. falsch bzw. nicht nachvollziehbar bewerten."

        )


    #
    # Für die Anordnung brauchen wir nur die
    # Antworttexte in ihrer Reihenfolge.
    #
    answer = list(
        antworten.values()
    )


    punkte = punktewert(
        zeile.get(
            "Punkte",
            ""
        )
    )

    gemeinsam = gemeinsame_fragefelder(
        zeile
    )

    return {
        "typ": "VERT_ANORD",

        "titel": titel_mit_punkten(
            zeile
        ),

        "einleitung": gemeinsam["einleitung"],
        "bild": gemeinsam["bild"],
        "audio": gemeinsam["audio"],
        "video": gemeinsam["video"],
        "frage": gemeinsam["frage"],


        "antworten": answer,

        "punkte": punkte,

        "autor": gemeinsam["autor"],
        "rueckmeldung": gemeinsam["rueckmeldung"]


    }

def bereite_anordnung_horizontal_vor(zeile):
    """
    Bereitet eine horizontale Anordnungsfrage vor.

    Die Antworten A, B, C, ... werden bereits
    in der richtigen Reihenfolge angegeben.
    """

    antworten = hole_antworten(
        zeile
    )

    if len(antworten) < 2:
        raise ValueError(
            "Eine horizontale Anordnungsfrage "
            "benötigt mindestens zwei Antworten."
        )

    #
    # Die Antworten müssen ohne Lücken
    # ab Antwort A ausgefüllt sein.
    #
    erwartete_buchstaben = list(
        ANTWORT_BUCHSTABEN[
            :len(antworten)
        ]
    )

    vorhandene_buchstaben = list(
        antworten.keys()
    )

    if vorhandene_buchstaben != erwartete_buchstaben:
        raise ValueError(
            "Bei einer horizontalen Anordnungsfrage "
            "müssen die Antworten ohne Lücken "
            "ab Antwort A ausgefüllt sein."
        )

    if len(set(antworten.values())) != len(antworten):
        raise ValueError(
            "Bei einer horizontalen Anordnungsfrage darf "
            "derselbe Antworttext nicht mehrfach verwendet werden. "
            "Ein LMS kann identische Antworttexte technisch unterschiedlich "
            "behandeln und dadurch ggf. falsch bzw. nicht nachvollziehbar bewerten."

        )


    #
    # "::" ist bei horizontalen Anordnungsfragen
    # das technische Trennzeichen von ILIAS.
    #
    for buchstabe, antwort in antworten.items():

        if "::" in antwort:
            raise ValueError(
                "Bei einer horizontalen Anordnungsfrage "
                f"enthält Antwort {buchstabe.upper()} "
                "das reservierte Trennzeichen '::'. "
                "Verwenden Sie stattdessen ': :'."
            )
    #
    # Antworttexte in ihrer richtigen Reihenfolge.
    #
    answer = list(
        antworten.values()
    )


    punkte = punktewert(
        zeile.get(
            "Punkte",
            ""
        )
    )

    gemeinsam = gemeinsame_fragefelder(
        zeile
    )


    return {
        "typ": "HOR_ANORD",

        "titel": titel_mit_punkten(
            zeile
        ),


        "einleitung": gemeinsam["einleitung"],
        "bild": gemeinsam["bild"],
        "audio": gemeinsam["audio"],
        "video": gemeinsam["video"],
        "frage": gemeinsam["frage"],


        "antworten": answer,

        "punkte": punkte,



        "autor": gemeinsam["autor"],
        "rueckmeldung": gemeinsam["rueckmeldung"]

    }


def bereite_freitext_vor(zeile):
    """
    Bereitet eine Freitextfrage vor.

    Freitextfragen werden manuell bewertet.
    Es erfolgt keine automatische Textbewertung.
    """

    frage_text = textwert(
        zeile.get(
            "Frage",
            ""
        )
    )

    if not frage_text:
        raise ValueError(
            "Eine Freitextfrage benötigt einen Fragetext."
        )


    punkte = punktewert(
        zeile.get(
            "Punkte",
            ""
        )
    )

    gemeinsam = gemeinsame_fragefelder(
        zeile
    )


    return {
        "typ": "FREITEXT",

        "titel": titel_mit_punkten(
            zeile
        ),



        "einleitung": gemeinsam["einleitung"],
        "bild": gemeinsam["bild"],
        "audio": gemeinsam["audio"],
        "video": gemeinsam["video"],
        "frage": gemeinsam["frage"],


        "punkte": punkte,


        "autor": gemeinsam["autor"],
        "rueckmeldung": gemeinsam["rueckmeldung"]


    }


def bereite_textluecke_vor(
    zeile,
    nummer
):
    """
    Bereitet eine einzelne Textlücke vor.
    """

    richtige_antwort = textwert(
        zeile.get(
            "Richtige Antwort",
            ""
        )
    )

    if not richtige_antwort:

        raise ValueError(
            "Eine Textlücke benötigt eine "
            "richtige Antwort."
        )



    #
    # Antwort A bis L können als weitere
    # richtige Antworten verwendet werden.
    #
    antworten = hole_antworten(
        zeile
    )


    erwartete_buchstaben = list(
        ANTWORT_BUCHSTABEN[
            :len(antworten)
        ]
    )

    if list(antworten.keys()) != erwartete_buchstaben:

        raise ValueError(
            "Bei einer Textlücke müssen zusätzliche richtige Antworten "
            "ohne Lücken ab Antwort A ausgefüllt sein. "
            "Die Hauptantwort steht in der Spalte 'Richtige Antwort'."
        )



    richtige_antworten = [
        richtige_antwort
    ] + list(
        antworten.values()
    )

    if len(set(richtige_antworten)) != len(richtige_antworten):

        raise ValueError(
            "Bei einer Textlücke darf dieselbe richtige "
            "Antwort nicht mehrfach angegeben werden."
        )


    punkte = punktewert(
        zeile.get(
            "Punkte",
            ""
        )
    )


    return {
        "lueckentyp": "TEXT",

        "nummer": nummer,

        "frage": textwert(
            zeile.get(
                "Frage",
                ""
            )
        ),

        "text_nach_luecke": textwert(
            zeile.get(
                "Text_nach_Lueckentextfrage",
                ""
            )
        ),


        "richtige_antwort": richtige_antwort,

        "richtige_antworten": richtige_antworten,


        "punkte": punkte
    }
    
    
    
def bereite_auswahlluecke_vor(
    zeile,
    nummer
):
    """
    Bereitet eine einzelne Auswahllücke vor.
    """

    antworten = hole_antworten(
        zeile
    )

    if len(antworten) < 2:

        raise ValueError(
            "Eine Auswahllücke benötigt "
            "mindestens zwei Antwortoptionen."
        )

    erwartete_buchstaben = list(
        ANTWORT_BUCHSTABEN[
            :len(antworten)
        ]
    )

    if list(antworten.keys()) != erwartete_buchstaben:

        raise ValueError(
            "Bei einer Auswahllücke müssen die Antwortoptionen "
            "ohne Lücken ab Antwort A ausgefüllt sein."
        )


    if len(set(antworten.values())) != len(antworten):

        raise ValueError(
            "Bei einer Auswahllücke darf derselbe Antworttext "
            "nicht mehrfach verwendet werden. "
            "Ein LMS kann identische Antworttexte technisch unterschiedlich "
            "behandeln und dadurch ggf. falsch bzw. nicht nachvollziehbar bewerten."
        )


    richtige = hole_richtige_antworten(
        zeile
    )

    pruefe_antworten(
        antworten,
        richtige,
        textwert(
            zeile.get(
                "Frage_ID",
                ""
            )
        )
    )
    if len(richtige) != 1:

        raise ValueError(
            "Eine Auswahllücke benötigt genau "
            "eine richtige Antwort."
        )

    richtiger_buchstabe = richtige[0]

    richtige_antwort = (
        antworten[
            richtiger_buchstabe
        ]
    )


    punkte = punktewert(
        zeile.get(
            "Punkte",
            ""
        )
    )


    return {
        "lueckentyp": "AUSWAHL",

        "nummer": nummer,

        "frage": textwert(
            zeile.get(
                "Frage",
                ""
            )
        ),

        "text_nach_luecke": textwert(
            zeile.get(
                "Text_nach_Lueckentextfrage",
                ""
            )
        ),

        "antworten": list(
            antworten.values()
        ),

        "richtige_antwort": richtige_antwort,

        "punkte": punkte
    }



def bereite_zahlluecke_vor(
    zeile,
    nummer
):
    """
    Bereitet eine einzelne Zahllücke vor.

    Richtige Antwort = Sollwert
    Antwort A        = Toleranz / Abstand
    """

    sollwert = dezimalwert(
        zeile.get(
            "Richtige Antwort",
            ""
        ),
        "Richtige Antwort"
    )


    toleranz_rohwert = textwert(
        zeile.get(
            "Antwort A",
            ""
        )
    )

    if toleranz_rohwert:

        toleranz = dezimalwert(
            toleranz_rohwert,
            "Antwort A"
        )

    else:

        toleranz = Decimal("0")

    if toleranz < 0:

        raise ValueError(
            "Der Abstand einer Zahllücke "
            "darf nicht negativ sein."
        )




    #
    # Antwort B bis L dürfen bei Zahllücken
    # nicht verwendet werden.
    #
    for buchstabe in ANTWORT_BUCHSTABEN[1:]:

        spaltenname = (
            f"Antwort {buchstabe.upper()}"
        )

        if textwert(
            zeile.get(
                spaltenname,
                ""
            )
        ):

            raise ValueError(
                "Bei einer Zahllücke wird nur "
                "Antwort A als Intervallabstand "
                "verwendet. Antwort B bis L "
                "müssen leer bleiben."
            )


    punkte = punktewert(
        zeile.get(
            "Punkte",
            ""
        )
    )


    untergrenze = (
        sollwert
        - toleranz
    )

    obergrenze = (
        sollwert
        + toleranz
    )

    return {
        "lueckentyp": "ZAHL",

        "nummer": nummer,

        "frage": textwert(
            zeile.get(
                "Frage",
                ""
            )
        ),

        "text_nach_luecke": textwert(
            zeile.get(
                "Text_nach_Lueckentextfrage",
                ""
            )
        ),

        "richtige_antwort": format_dezimal(
            sollwert
        ),

        "untergrenze": format_dezimal(
            untergrenze
        ),

        "obergrenze": format_dezimal(
            obergrenze
        ),

        "punkte": punkte
    }




def hole_einzelnen_blockwert(
    zeilen,
    spaltenname,
    frage_id
):
    """
    Liest eine Eigenschaft aus mehreren Zeilen
    derselben Lückentextfrage.

    Leere Zellen werden ignoriert.
    Identische Wiederholungen sind erlaubt.
    Unterschiedliche Werte führen zum Abbruch.
    """

    werte = []

    for zeile in zeilen:

        wert = textwert(
            zeile.get(
                spaltenname,
                ""
            )
        )

        if (
            wert
            and wert not in werte
        ):

            werte.append(
                wert
            )

    if len(werte) > 1:

        raise ValueError(
            f"Frage_ID {frage_id}: "
            "Bei einer Lückentextfrage "
            f"darf '{spaltenname}' nur einmal "
            "bzw. nur mit identischem Inhalt "
            "angegeben werden. "
            f"Gefunden wurden: {werte}"
        )



    if werte:

        return werte[0]

    return ""    





def bereite_lueckentextfrage_vor(
    zeilen
):
    """
    Fasst mehrere Tabellenzeilen mit derselben
    Frage_ID zu einer einzigen ILIAS-
    Lückentextfrage zusammen.
    """

    if not zeilen:

        raise ValueError(
            "Leere Lückentextfrage."
        )

    frage_id = textwert(
        zeilen[0].get(
            "Frage_ID",
            ""
        )
    )

    if not frage_id:

        raise ValueError(
            "Eine Lückentextfrage benötigt "
            "eine Frage_ID."
        )

    ueberschrift = hole_einzelnen_blockwert(
        zeilen,
        "Ueberschrift",
        frage_id
    )

    einleitung = hole_einzelnen_blockwert(
        zeilen,
        "Einleitungstext",
        frage_id
    )
    
    bild = hole_einzelnen_blockwert(
        zeilen,
        "Bild",
        frage_id
    )    
    
    audio = hole_einzelnen_blockwert(
        zeilen,
        "Audio",
        frage_id
    )    
    


    video = hole_einzelnen_blockwert(
        zeilen,
        "Video",
        frage_id
    )



    if audio and video:

        raise ValueError(
            f"Frage_ID {frage_id}: "
            "Audio und Video können nicht gleichzeitig "
            "in derselben Lückentextfrage verwendet werden. "
            "Bitte verwenden Sie entweder eine Audiodatei "
            "oder eine Videodatei."
        )



    autor = hole_einzelnen_blockwert(
        zeilen,
        "Autor",
        frage_id
    )

    rueckmeldung = hole_einzelnen_blockwert(
        zeilen,
        "Rueckmeldung",
        frage_id
    )

    punktanzeige = hole_einzelnen_blockwert(
        zeilen,
        "Punktanzeige",
        frage_id
    )

    luecken = []

    for nummer, zeile in enumerate(
        zeilen,
        start=1
    ):

        fragetyp = textwert(
            zeile.get(
                "Fragetyp",
                ""
            )
        ).upper()

        if fragetyp == "TEXTLUECKE":

            luecke = bereite_textluecke_vor(
                zeile,
                nummer
            )

        elif fragetyp == "AUSWAHLLUECKE":

            luecke = bereite_auswahlluecke_vor(
                zeile,
                nummer
            )

        elif fragetyp == "ZAHLLUECKE":

            luecke = bereite_zahlluecke_vor(
                zeile,
                nummer
            )

        else:

            raise ValueError(
                f"Frage_ID {frage_id}: "
                f"Unbekannter Lückentyp "
                f"'{fragetyp}'."
            )

        luecken.append(
            luecke
        )

    gesamtpunkte = sum(
        luecke["punkte"]
        for luecke in luecken
    )

    #
    # titel_mit_punkten() kann weiterverwendet
    # werden. Dafür erzeugen wir eine kleine
    # temporäre Zeile mit der Gesamtpunktzahl.
    #
    titelzeile = {
        "Ueberschrift": ueberschrift,
        "Punktanzeige": punktanzeige,
        "Punkte": gesamtpunkte
    }

    return {
        "typ": "LUECKENTEXT",

        "frage_id": frage_id,

        "titel": titel_mit_punkten(
            titelzeile
        ),

        #
        # Bei Lückentextfragen wird der
        # Einleitungstext innerhalb des QTI
        # ausgegeben, nicht über das QPL.
        #
        "einleitung": einleitung,
        
        "bild": bild,
        
        "audio": audio,

        "video": video,
        
        "punktanzeige": punktanzeige.lower(),        

        "autor": autor,

        "rueckmeldung": rueckmeldung,

        "gesamtpunkte": gesamtpunkte,

        "luecken": luecken
    }




def bereite_alle_fragen_vor(
    zeilen
):
    """
    Bereitet alle Tabellenzeilen vor.

    Einzeilige Fragetypen:
        eine Zeile = eine Frage

    Lückentext:
        mehrere Zeilen mit gleicher Frage_ID
        = eine gemeinsame Frage

    Fehler werden gesammelt und erst nach der
    Prüfung aller Fragen gemeinsam ausgegeben.
    """

    lueckentypen = {
        "TEXTLUECKE",
        "AUSWAHLLUECKE",
        "ZAHLLUECKE"
    }

    fragen = []

    fehler = []

    bereits_verarbeitet = set()

    verwendete_einzelfrage_ids = set()

    for zeile in zeilen:

        fragetyp = textwert(
            zeile.get(
                "Fragetyp",
                ""
            )
        ).upper()

        frage_id = textwert(
            zeile.get(
                "Frage_ID",
                ""
            )
        )

        #
        # Einzeilige Fragetypen:
        # eine Tabellenzeile = eine Frage
        #
        if fragetyp not in lueckentypen:
        
            if not frage_id:

                fehler.append(
                    "Eine Frage benötigt eine Frage_ID."
                )

                continue

            if frage_id in verwendete_einzelfrage_ids:

                fehler.append(
                    f"Frage_ID {frage_id}: "
                    "Diese Frage_ID wurde mehrfach "
                    "für eigenständige Fragen verwendet."
                )

                continue

            verwendete_einzelfrage_ids.add(
                frage_id
            )        

            try:

                fragen.append(
                    bereite_frage_vor(
                        zeile
                    )
                )

            except Exception as e:

                meldung = str(e)

                #
                # Frage_ID ergänzen, wenn die
                # eigentliche Fehlermeldung sie
                # noch nicht enthält.
                #
                if (
                    frage_id
                    and "Frage_ID" not in meldung
                ):

                    meldung = (
                        f"Frage_ID {frage_id}: "
                        f"{meldung}"
                    )

                fehler.append(
                    meldung
                )

            continue

        #
        # Lückentextfragen benötigen zwingend
        # eine Frage_ID.
        #
        if not frage_id:

            fehler.append(
                "Eine Lückentextzeile benötigt "
                "eine Frage_ID."
            )

            continue

        #
        # Jede Lückentextfrage nur einmal als
        # gesamter Block verarbeiten.
        #
        if frage_id in bereits_verarbeitet:
            continue

        block = []

        block_fehler = False

        for andere_zeile in zeilen:

            andere_id = textwert(
                andere_zeile.get(
                    "Frage_ID",
                    ""
                )
            )

            anderer_typ = textwert(
                andere_zeile.get(
                    "Fragetyp",
                    ""
                )
            ).upper()

            if andere_id != frage_id:
                continue

            if anderer_typ not in lueckentypen:

                fehler.append(
                    f"Frage_ID {frage_id}: "
                    "Die Frage_ID wird gleichzeitig "
                    "für eine Lückentextfrage und "
                    "einen anderen Fragetyp verwendet."
                )

                block_fehler = True

                break

            block.append(
                andere_zeile
            )

        #
        # Die Frage_ID gilt unabhängig vom
        # Ergebnis als verarbeitet.
        #
        bereits_verarbeitet.add(
            frage_id
        )

        if block_fehler:
            continue

        try:

            fragen.append(
                bereite_lueckentextfrage_vor(
                    block
                )
            )

        except Exception as e:

            meldung = str(e)

            #
            # Viele Blockfehler enthalten die
            # Frage_ID bereits. Bei Fehlern einer
            # einzelnen Lücke ergänzen wir sie hier.
            #
            if "Frage_ID" not in meldung:

                meldung = (
                    f"Frage_ID {frage_id}: "
                    f"{meldung}"
                )

            fehler.append(
                meldung
            )

    #
    # Erst nach Prüfung aller Fragen abbrechen.
    #
    if fehler:

        anzahl = len(
            fehler
        )

        if anzahl == 1:

            kopf = (
                "1 Fehler gefunden:"
            )

        else:

            kopf = (
                f"{anzahl} Fehler gefunden:"
            )

        fehlertext = "\n".join(
            f"- {meldung}"
            for meldung in fehler
        )

        raise ValueError(
            f"{kopf}\n\n"
            f"{fehlertext}"
        )

    return fragen





def bereite_frage_vor(zeile):
    """
    Verteilt anhand des Fragetyps auf die
    passende Aufbereitungsfunktion.
    """
    
    fragetyp = textwert(
        zeile.get("Fragetyp", "")
    ).upper()

    if fragetyp == "SC":
        return bereite_sc_vor(zeile)

    if fragetyp == "MC":
        return bereite_mc_vor(zeile)

    if fragetyp == "KPRIM":
        return bereite_kprim_vor(zeile)

    if fragetyp == "VERT_ANORD":
        return bereite_anordnung_vertikal_vor(
            zeile
        )

    if fragetyp == "HOR_ANORD":
        return bereite_anordnung_horizontal_vor(
            zeile
        )        

    if fragetyp == "FREITEXT":
        return bereite_freitext_vor(
            zeile
        )

    raise ValueError(
        f"Unbekannter Fragetyp: {fragetyp}"
    )


def xml_text(parent, tag, text=None, attrib=None):
    """
    Kleine Hilfsfunktion für XML-Elemente mit Text.
    """
    elem = ET.SubElement(
        parent,
        tag,
        attrib or {}
    )

    if text is not None:
        elem.text = str(text)

    return elem


def metadata(qt_metadata, label, value):
    """
    Erzeugt ein QTI-Metadatenfeld.
    """
    field = ET.SubElement(
        qt_metadata,
        "qtimetadatafield"
    )

    xml_text(
        field,
        "fieldlabel",
        label
    )

    xml_text(
        field,
        "fieldentry",
        value
    )


def erzeuge_fragematerial(parent, frage):
    """
    Erzeugt den Fragetext und optional eine Abbildung.
    
    Der Einleitungstext wird bei einzeiligen Fragetypen
    separat über die QPL-Datei erzeugt.

    Die sichtbare Reihenfolge lautet letztlich:

    Einleitungstext
    Bild
    Fragetext
    Antwortoptionen
    """

    material = ET.SubElement(
        parent,
        "material"
    )

    if frage.get("bild_mob_id"):

        frage_mit_bild = (
            f'<p><img src="'
            f'{frage["bild_mob_id"]}'
            f'" /></p>'
            f'{frage["frage"]}'
        )

    else:

        frage_mit_bild = (
            frage["frage"]
        )

    xml_text(
        material,
        "mattext",
        frage_mit_bild,
        {
            "texttype": "text/xhtml"
        }
    )

    #
    # Bei einer Abbildung benötigt ILIAS zusätzlich
    # den Verweis auf die Bilddatei im Media Object.
    #
    if frage.get("bild_mob_id"):

        ET.SubElement(
            material,
            "matimage",
            {
                "label": frage["bild_mob_id"],
                "uri": frage["bild_qti_pfad"]
            }
        )

    return material


def erzeuge_sc_item(root, frage, frage_id):
    """
    Erzeugt eine SingleChoice-Frage.
    """

    item = ET.SubElement(
        root,
        "item",
        {
            "ident": frage_id,
            "title": frage["titel"]
        }
    )

    itemmetadata = ET.SubElement(
        item,
        "itemmetadata"
    )

    qt_metadata = ET.SubElement(
        itemmetadata,
        "qtimetadata"
    )

    metadata(
        qt_metadata,
        "QUESTIONTYPE",
        "SINGLE CHOICE QUESTION"
    )

    metadata(
        qt_metadata,
        "AUTHOR",
        frage["autor"]
    )

    metadata(
        qt_metadata,
        "feedback_setting",
        "2"
    )


    presentation = ET.SubElement(
        item,
        "presentation"
    )

    erzeuge_fragematerial(
        presentation,
        frage
    )





    response = ET.SubElement(
        presentation,
        "response_lid",
        {
            "ident": "MCSR",
            "rcardinality": "Single"
        }
    )

    render_choice = ET.SubElement(
        response,
        "render_choice",
        {
            "shuffle": "Yes"
        }
    )

    for i, antwort in enumerate(
        frage["antworten"],
        start=1
    ):

        response_label = ET.SubElement(
            render_choice,
            "response_label",
            {
                "ident": str(i)
            }
        )

        material = ET.SubElement(
            response_label,
            "material"
        )

        xml_text(
            material,
            "mattext",
            antwort,
            {
                "texttype": "text/xhtml"
            }
        )

    resprocessing = ET.SubElement(
        item,
        "resprocessing"
    )

    outcomes = ET.SubElement(
        resprocessing,
        "outcomes"
    )

    ET.SubElement(
        outcomes,
        "decvar"
    )

    #
    # Bewertung und individuelle Feedbackkanäle
    #
    for i, _antwort in enumerate(
        frage["antworten"],
        start=1
    ):

        respcondition = ET.SubElement(
            resprocessing,
            "respcondition",
            {
                "continue": "Yes"
            }
        )

        conditionvar = ET.SubElement(
            respcondition,
            "conditionvar"
        )

        xml_text(
            conditionvar,
            "varequal",
            str(i),
            {
                "respident": "MCSR"
            }
        )



        punkte = (
            frage["punkte"]
            if i == frage["richtige_position"]
            else 0
        )


        xml_text(
            respcondition,
            "setvar",
            format_punkte(punkte),
            {
                "action": "Add"
            }
        )

        ET.SubElement(
            respcondition,
            "displayfeedback",
            {
                "feedbacktype": "Response",
                "linkrefid": f"response_{i}"
            }
        )

    #
    # Generelle Rückmeldung bei falscher Lösung
    #
    if frage["rueckmeldung"]:

        respcondition = ET.SubElement(
            resprocessing,
            "respcondition",
            {
                "continue": "Yes"
            }
        )

        conditionvar = ET.SubElement(
            respcondition,
            "conditionvar"
        )

        not_elem = ET.SubElement(
            conditionvar,
            "not"
        )




        xml_text(
            not_elem,
            "varequal",
            str(
                frage["richtige_position"]
            ),
            {
                "respident": "MCSR"
            }
        )


        ET.SubElement(
            respcondition,
            "displayfeedback",
            {
                "feedbacktype": "Response",
                "linkrefid": "response_onenotcorrect"
            }
        )

    #
    # Leere individuelle Feedbackblöcke
    #
    for i in range(
        1,
        len(frage["antworten"]) + 1
    ):

        itemfeedback = ET.SubElement(
            item,
            "itemfeedback",
            {
                "ident": f"response_{i}",
                "view": "All"
            }
        )

        flow_mat = ET.SubElement(
            itemfeedback,
            "flow_mat"
        )

        material = ET.SubElement(
            flow_mat,
            "material"
        )

        xml_text(
            material,
            "mattext",
            "",
            {
                "texttype": "text/xhtml"
            }
        )

    #
    # Allgemeine Rückmeldung bei falscher Lösung
    #
    if frage["rueckmeldung"]:

        itemfeedback = ET.SubElement(
            item,
            "itemfeedback",
            {
                "ident": "response_onenotcorrect",
                "view": "All"
            }
        )

        flow_mat = ET.SubElement(
            itemfeedback,
            "flow_mat"
        )

        material = ET.SubElement(
            flow_mat,
            "material"
        )

        xml_text(
            material,
            "mattext",
            frage["rueckmeldung"],
            {
                "texttype": "text/xhtml"
            }
        )


def erzeuge_mc_item(root, frage, frage_id):
    """
    Erzeugt eine MultipleChoice-Frage.
    """

    item = ET.SubElement(
        root,
        "item",
        {
            "ident": frage_id,
            "title": frage["titel"]
        }
    )

    itemmetadata = ET.SubElement(
        item,
        "itemmetadata"
    )

    qt_metadata = ET.SubElement(
        itemmetadata,
        "qtimetadata"
    )

    metadata(
        qt_metadata,
        "QUESTIONTYPE",
        "assMultipleChoice"
    )

    metadata(
        qt_metadata,
        "AUTHOR",
        frage["autor"]
    )

    metadata(
        qt_metadata,
        "feedback_setting",
        "1"
    )

    presentation = ET.SubElement(
        item,
        "presentation"
    )

    erzeuge_fragematerial(
        presentation,
        frage
    )


    response = ET.SubElement(
        presentation,
        "response_lid",
        {
            "ident": "MCMR",
            "rcardinality": "Multiple"
        }
    )



    if (
        frage["mc_bewertungsvariante"]
        == "nur-richtige"
    ):

        render_choice = ET.SubElement(
            response,
            "render_choice",
            {
                "shuffle": "Yes",
                "minnumber": "0",
                "maxnumber": str(
                    frage["anzahl_richtige"]
                )
            }
        )

    else:

        render_choice = ET.SubElement(
            response,
            "render_choice",
            {
                "shuffle": "Yes",
                "minnumber": "0"
            }
        )



    
        
    for i, antwort in enumerate(
        frage["antworten"],
        start=1
    ):

        response_label = ET.SubElement(
            render_choice,
            "response_label",
            {
                "ident": str(i)
            }
        )

        material = ET.SubElement(
            response_label,
            "material"
        )

        xml_text(
            material,
            "mattext",
            antwort,
            {
                "texttype": "text/plain"
            }
        )

    resprocessing = ET.SubElement(
        item,
        "resprocessing"
    )

    outcomes = ET.SubElement(
        resprocessing,
        "outcomes"
    )

    ET.SubElement(
        outcomes,
        "decvar",
        {
            "varname": "SCORE",
            "vartype": "Decimal",
            "minvalue": "0",
            "maxvalue": format_punkte(
                frage["punkte"]
            )
        }
    )


    gesamtpunkte = Decimal(
        str(frage["punkte"])
    )

    anzahl_richtige = (
        frage["anzahl_richtige"]
    )

    anzahl_falsche = (
        frage["anzahl_falsche"]
    )

    anzahl_antworten = (
        anzahl_richtige
        + anzahl_falsche
    )


    #
    # Bewertungsvariante
    #
    bewertungsvariante = (
        frage["mc_bewertungsvariante"]
    )


    if (
        bewertungsvariante
        == "all-or-nothing"
    ):

        #
        # All-or-Nothing
        #
        # Genau eine richtige Antwort trägt technisch
        # die gesamte Punktzahl.
        #
        # Alle weiteren richtigen Antworten tragen
        # technisch 0 Punkte.
        #
        # Jeder Fehler erhält einen ausreichend großen
        # negativen Strafwert, sodass nur die vollständig
        # richtige Kombination die Gesamtpunktzahl erreicht.
        #
        strafwert = -(
            2
            * gesamtpunkte
        )

        erste_richtige_antwort = True

        for i in range(
            1,
            anzahl_antworten + 1
        ):

            ist_richtig = (
                frage["wahrheitswerte"][i - 1]
            )

            #
            # Antwort wurde angekreuzt
            #
            respcondition = ET.SubElement(
                resprocessing,
                "respcondition",
                {
                    "continue": "Yes"
                }
            )

            conditionvar = ET.SubElement(
                respcondition,
                "conditionvar"
            )

            xml_text(
                conditionvar,
                "varequal",
                str(i),
                {
                    "respident": "MCMR"
                }
            )

            if ist_richtig:

                if erste_richtige_antwort:

                    punkte = gesamtpunkte

                    erste_richtige_antwort = False

                else:

                    punkte = 0

            else:

                punkte = strafwert

            xml_text(
                respcondition,
                "setvar",
                format_punkte(punkte),
                {
                    "action": "Add"
                }
            )

            ET.SubElement(
                respcondition,
                "displayfeedback",
                {
                    "feedbacktype": "Response",
                    "linkrefid": f"response_{i}"
                }
            )

            #
            # Antwort wurde NICHT angekreuzt
            #
            respcondition = ET.SubElement(
                resprocessing,
                "respcondition",
                {
                    "continue": "Yes"
                }
            )

            conditionvar = ET.SubElement(
                respcondition,
                "conditionvar"
            )

            not_elem = ET.SubElement(
                conditionvar,
                "not"
            )

            xml_text(
                not_elem,
                "varequal",
                str(i),
                {
                    "respident": "MCMR"
                }
            )

            if ist_richtig:

                punkte = strafwert

            else:

                punkte = 0

            xml_text(
                respcondition,
                "setvar",
                format_punkte(punkte),
                {
                    "action": "Add"
                }
            )




    elif (
        bewertungsvariante
        == "nur-richtige"
    ):

        #
        # Nur richtige Kreuze werden bewertet.
        #
        # Richtig angekreuzt:
        # + Gesamtpunkte / Anzahl richtige Antworten
        #
        # Alles andere:
        # 0 Punkte
        #
        

        punkte_richtige_antworten = verteile_punkte(
            gesamtpunkte,
            anzahl_richtige,
            nachkommastellen=3
        )

        index_richtige_antwort = 0

        for i in range(
            1,
            anzahl_antworten + 1
        ):


            ist_richtig = (
                frage["wahrheitswerte"][i - 1]
            )


            #
            # Antwort wurde angekreuzt
            #
            respcondition = ET.SubElement(
                resprocessing,
                "respcondition",
                {
                    "continue": "Yes"
                }
            )

            conditionvar = ET.SubElement(
                respcondition,
                "conditionvar"
            )

            xml_text(
                conditionvar,
                "varequal",
                str(i),
                {
                    "respident": "MCMR"
                }
            )



            if ist_richtig:
                punkte = punkte_richtige_antworten[
                    index_richtige_antwort
                ]

                index_richtige_antwort += 1

            else:
                punkte = Decimal("0")


            xml_text(
                respcondition,
                "setvar",
                format_punkte(punkte),
                {
                    "action": "Add"
                }
            )

            ET.SubElement(
                respcondition,
                "displayfeedback",
                {
                    "feedbacktype": "Response",
                    "linkrefid": f"response_{i}"
                }
            )

    elif (
        bewertungsvariante
        == "richtig-minus-falsch"
    ):

        #
        # Richtige Kreuze bringen anteilig Punkte.
        # Falsche Kreuze ziehen anteilig Punkte ab.
        # Nicht angekreuzte Antworten sind neutral.
        #
        # Summe aller richtigen Kreuze:
        # maximal +Gesamtpunkte
        #
        # Summe aller falschen Kreuze:
        # maximal -Gesamtpunkte
        #
        punkte_richtige_antwort = (
            gesamtpunkte
            / anzahl_richtige
        )

        punkte_falsche_antwort = -(
            gesamtpunkte
            / anzahl_falsche
        )

        for i in range(
            1,
            anzahl_antworten + 1
        ):



            ist_richtig = (
                frage["wahrheitswerte"][i - 1]
            )

            #
            # Antwort wurde angekreuzt
            #
            respcondition = ET.SubElement(
                resprocessing,
                "respcondition",
                {
                    "continue": "Yes"
                }
            )

            conditionvar = ET.SubElement(
                respcondition,
                "conditionvar"
            )

            xml_text(
                conditionvar,
                "varequal",
                str(i),
                {
                    "respident": "MCMR"
                }
            )

            if ist_richtig:
                punkte = (
                    punkte_richtige_antwort
                )
            else:
                punkte = (
                    punkte_falsche_antwort
                )

            xml_text(
                respcondition,
                "setvar",
                format_punkte(punkte),
                {
                    "action": "Add"
                }
            )

            ET.SubElement(
                respcondition,
                "displayfeedback",
                {
                    "feedbacktype": "Response",
                    "linkrefid": f"response_{i}"
                }
            )

    elif (
        bewertungsvariante
        == "jede-aussage"
    ):

        #
        # Jede Antwortoption stellt eine eigene
        # Ja-/Nein-Entscheidung dar.
        #
        # richtige Entscheidung:
        # +Teilpunkt
        #
        # falsche Entscheidung:
        # -Teilpunkt
        #
        # Teilpunkt =
        # Gesamtpunkte / Anzahl Antwortoptionen
        #
        teilpunkt = (
            gesamtpunkte
            / anzahl_antworten
        )

        for i in range(
            1,
            anzahl_antworten + 1
        ):


            ist_richtig = (
                frage["wahrheitswerte"][i - 1]
            )



            #
            # Antwort wurde angekreuzt
            #
            respcondition = ET.SubElement(
                resprocessing,
                "respcondition",
                {
                    "continue": "Yes"
                }
            )

            conditionvar = ET.SubElement(
                respcondition,
                "conditionvar"
            )

            xml_text(
                conditionvar,
                "varequal",
                str(i),
                {
                    "respident": "MCMR"
                }
            )

            if ist_richtig:
                punkte = teilpunkt
            else:
                punkte = -teilpunkt

            xml_text(
                respcondition,
                "setvar",
                format_punkte(punkte),
                {
                    "action": "Add"
                }
            )

            ET.SubElement(
                respcondition,
                "displayfeedback",
                {
                    "feedbacktype": "Response",
                    "linkrefid": f"response_{i}"
                }
            )

            #
            # Antwort wurde NICHT angekreuzt
            #
            respcondition = ET.SubElement(
                resprocessing,
                "respcondition",
                {
                    "continue": "Yes"
                }
            )

            conditionvar = ET.SubElement(
                respcondition,
                "conditionvar"
            )

            not_elem = ET.SubElement(
                conditionvar,
                "not"
            )

            xml_text(
                not_elem,
                "varequal",
                str(i),
                {
                    "respident": "MCMR"
                }
            )

            if ist_richtig:
                punkte = -teilpunkt
            else:
                punkte = teilpunkt

            xml_text(
                respcondition,
                "setvar",
                format_punkte(punkte),
                {
                    "action": "Add"
                }
            )

    else:

        raise ValueError(
            "Unbekannte MC_Bewertungsvariante: "
            f"{bewertungsvariante}"
        )





    #
    # Generelle Rückmeldung bei falscher Lösung
    #
    if frage["rueckmeldung"]:

        respcondition = ET.SubElement(
            resprocessing,
            "respcondition",
            {
                "continue": "Yes"
            }
        )

        conditionvar = ET.SubElement(
            respcondition,
            "conditionvar"
        )

        #
        # Falsch ist die Gesamtlösung, wenn
        #
        # - mindestens eine richtige Antwort fehlt
        #   ODER
        # - mindestens eine falsche Antwort gewählt wurde.
        #
        or_elem = ET.SubElement(
            conditionvar,
            "or"
        )

        for i in range(
            1,
            anzahl_antworten + 1
        ):


            ist_richtig = (
                frage["wahrheitswerte"][i - 1]
            )


            if ist_richtig:

                not_elem = ET.SubElement(
                    or_elem,
                    "not"
                )

                xml_text(
                    not_elem,
                    "varequal",
                    str(i),
                    {
                        "respident": "MCMR"
                    }
                )

            else:

                xml_text(
                    or_elem,
                    "varequal",
                    str(i),
                    {
                        "respident": "MCMR"
                    }
                )

        ET.SubElement(
            respcondition,
            "displayfeedback",
            {
                "feedbacktype": "Response",
                "linkrefid": "response_onenotcorrect"
            }
        ) 
 
 
 
 
    #
    # Leere individuelle Feedbackblöcke
    #
    for i in range(
        1,
        anzahl_antworten + 1
    ):

        itemfeedback = ET.SubElement(
            item,
            "itemfeedback",
            {
                "ident": f"response_{i}",
                "view": "All"
            }
        )

        flow_mat = ET.SubElement(
            itemfeedback,
            "flow_mat"
        )

        material = ET.SubElement(
            flow_mat,
            "material"
        )

        xml_text(
            material,
            "mattext",
            "",
            {
                "texttype": "text/plain"
            }
        )

    #
    # Allgemeine Rückmeldung bei falscher Lösung
    #
    if frage["rueckmeldung"]:

        itemfeedback = ET.SubElement(
            item,
            "itemfeedback",
            {
                "ident": "response_onenotcorrect",
                "view": "All"
            }
        )

        flow_mat = ET.SubElement(
            itemfeedback,
            "flow_mat"
        )

        material = ET.SubElement(
            flow_mat,
            "material"
        )

        xml_text(
            material,
            "mattext",
            frage["rueckmeldung"],
            {
                "texttype": "text/plain"
            }
        )


def erzeuge_kprim_item(root, frage, frage_id):
    """
    Erzeugt eine KPrim-Frage mit klassischer
    Halbpunktbewertung.

    Die Spalte "Rueckmeldung" wird als generelle
    Rückmeldung bei nicht vollständig richtiger
    Lösung verwendet.
    """

    item = ET.SubElement(
        root,
        "item",
        {
            "ident": frage_id,
            "title": frage["titel"],
            "maxattempts": "0"
        }
    )

    xml_text(
        item,
        "qticomment",
        ""
    )

    xml_text(
        item,
        "duration",
        "P0Y0M0DT0H1M0S"
    )

    itemmetadata = ET.SubElement(
        item,
        "itemmetadata"
    )

    qt_metadata = ET.SubElement(
        itemmetadata,
        "qtimetadata"
    )


    metadata(
        qt_metadata,
        "QUESTIONTYPE",
        "KPRIM CHOICE QUESTION"
    )

    metadata(
        qt_metadata,
        "AUTHOR",
        frage["autor"]
    )

    metadata(
        qt_metadata,
        "additional_cont_edit_mode",
        "default"
    )

    metadata(
        qt_metadata,
        "externalId",
        frage_id
    )

    metadata(
        qt_metadata,
        "answer_type",
        "singleLine"
    )

    metadata(
        qt_metadata,
        "thumb_size",
        "150"
    )

    metadata(
        qt_metadata,
        "option_label_setting",
        "right_wrong"
    )

    metadata(
        qt_metadata,
        "custom_true_option_label",
        ""
    )

    metadata(
        qt_metadata,
        "custom_false_option_label",
        ""
    )

    metadata(
        qt_metadata,
        "feedback_setting",
        "1"
    )


    presentation = ET.SubElement(
        item,
        "presentation",
        {
            "label": frage["titel"]
        }
    )

    flow = ET.SubElement(
        presentation,
        "flow"
    )

    erzeuge_fragematerial(
        flow,
        frage
    )

    response = ET.SubElement(
        flow,
        "response_lid",
        {
            "ident": "MCMR",
            "rcardinality": "Multiple"
        }
    )

    render_choice = ET.SubElement(
        response,
        "render_choice",
        {
            "shuffle": "Yes"
        }
    )

    #
    # Genau vier Aussagen mit IDs 0 bis 3.
    #
    for i, antwort in enumerate(
        frage["antworten"]
    ):

        response_label = ET.SubElement(
            render_choice,
            "response_label",
            {
                "ident": str(i)
            }
        )

        material = ET.SubElement(
            response_label,
            "material"
        )

        xml_text(
            material,
            "mattext",
            antwort,
            {
                "texttype": "text/plain"
            }
        )

    resprocessing = ET.SubElement(
        item,
        "resprocessing"
    )

    outcomes = ET.SubElement(
        resprocessing,
        "outcomes"
    )

    gesamtpunkte = Decimal(
        str(frage["punkte"])
    )

    halbe_punkte = (
        gesamtpunkte
        / 2
    )

    #
    # Halbpunktbewertung:
    # minvalue = halbe Punktzahl
    # maxvalue = volle Punktzahl
    #
    ET.SubElement(
        outcomes,
        "decvar",
        {
            "varname": "SCORE",
            "vartype": "Decimal",
            "defaultval": "0",
            "minvalue": format_punkte(
                halbe_punkte
            ),
            "maxvalue": format_punkte(
                gesamtpunkte
            )
        }
    )

    #
    # Einzelbedingungen für die vier Aussagen.
    #
    for i, wahrheitswert in enumerate(
        frage["wahrheitswerte"]
    ):

        respcondition = ET.SubElement(
            resprocessing,
            "respcondition",
            {
                "continue": "Yes"
            }
        )

        conditionvar = ET.SubElement(
            respcondition,
            "conditionvar"
        )

        xml_text(
            conditionvar,
            "varequal",
            str(wahrheitswert),
            {
                "respident": str(i)
            }
        )

        ET.SubElement(
            respcondition,
            "displayfeedback",
            {
                "feedbacktype": "Response",
                "linkrefid": f"response_{i}"
            }
        )

    #
    # Alle vier Aussagen richtig beurteilt:
    # volle Punktzahl.
    #
    respcondition = ET.SubElement(
        resprocessing,
        "respcondition",
        {
            "continue": "Yes"
        }
    )

    conditionvar = ET.SubElement(
        respcondition,
        "conditionvar"
    )

    and_elem = ET.SubElement(
        conditionvar,
        "and"
    )

    for i, wahrheitswert in enumerate(
        frage["wahrheitswerte"]
    ):

        xml_text(
            and_elem,
            "varequal",
            str(wahrheitswert),
            {
                "respident": str(i)
            }
        )

    xml_text(
        respcondition,
        "setvar",
        format_punkte(
            gesamtpunkte
        ),
        {
            "action": "Add"
        }
    )

    #
    # Mindestens eine Aussage falsch beurteilt:
    # zunächst 0 Punkte.
    #
    respcondition = ET.SubElement(
        resprocessing,
        "respcondition",
        {
            "continue": "Yes"
        }
    )

    conditionvar = ET.SubElement(
        respcondition,
        "conditionvar"
    )

    or_elem = ET.SubElement(
        conditionvar,
        "or"
    )

    for i, wahrheitswert in enumerate(
        frage["wahrheitswerte"]
    ):

        not_elem = ET.SubElement(
            or_elem,
            "not"
        )

        xml_text(
            not_elem,
            "varequal",
            str(wahrheitswert),
            {
                "respident": str(i)
            }
        )

    xml_text(
        respcondition,
        "setvar",
        "0",
        {
            "action": "Add"
        }
    )

    #
    # Falls eine generelle Rückmeldung vorhanden ist,
    # wird sie bei nicht vollständig richtiger Lösung
    # angezeigt.
    #
    if frage["rueckmeldung"]:

        ET.SubElement(
            respcondition,
            "displayfeedback",
            {
                "feedbacktype": "Response",
                "linkrefid": "response_onenotcorrect"
            }
        )

    #
    # Vier leere individuelle Feedbackblöcke.
    #
    for i in range(4):

        itemfeedback = ET.SubElement(
            item,
            "itemfeedback",
            {
                "ident": f"response_{i}",
                "view": "All"
            }
        )

        flow_mat = ET.SubElement(
            itemfeedback,
            "flow_mat"
        )

        material = ET.SubElement(
            flow_mat,
            "material"
        )

        xml_text(
            material,
            "mattext",
            "",
            {
                "texttype": "text/plain"
            }
        )

    #
    # Generelle Rückmeldung bei nicht vollständig
    # richtiger Lösung.
    #
    if frage["rueckmeldung"]:

        itemfeedback = ET.SubElement(
            item,
            "itemfeedback",
            {
                "ident": "response_onenotcorrect",
                "view": "All"
            }
        )

        flow_mat = ET.SubElement(
            itemfeedback,
            "flow_mat"
        )

        material = ET.SubElement(
            flow_mat,
            "material"
        )

        xml_text(
            material,
            "mattext",
            frage["rueckmeldung"],
            {
                "texttype": "text/plain"
            }
        )



def erzeuge_anordnung_vertikal_item(
    root,
    frage,
    frage_id
):
    """
    Erzeugt eine vertikale Anordnungsfrage.

    Die Antworten werden in der Tabelle bereits
    in der richtigen Reihenfolge angegeben.
    """

    item = ET.SubElement(
        root,
        "item",
        {
            "ident": frage_id,
            "title": frage["titel"],
            "maxattempts": "0"
        }
    )

    xml_text(
        item,
        "qticomment",
        ""
    )

    xml_text(
        item,
        "duration",
        "P0Y0M0DT0H1M0S"
    )

    #
    # Metadaten
    #
    itemmetadata = ET.SubElement(
        item,
        "itemmetadata"
    )

    qt_metadata = ET.SubElement(
        itemmetadata,
        "qtimetadata"
    )


    metadata(
        qt_metadata,
        "QUESTIONTYPE",
        "ORDERING QUESTION"
    )

    metadata(
        qt_metadata,
        "AUTHOR",
        frage["autor"]
    )

    metadata(
        qt_metadata,
        "additional_cont_edit_mode",
        "default"
    )

    metadata(
        qt_metadata,
        "externalId",
        frage_id
    )

    metadata(
        qt_metadata,
        "thumb_geometry",
        "100"
    )

    metadata(
        qt_metadata,
        "element_height",
        ""
    )

    metadata(
        qt_metadata,
        "points",
        format_punkte(
            frage["punkte"]
        )
    )

    #
    # Fragetext
    #
    presentation = ET.SubElement(
        item,
        "presentation",
        {
            "label": frage["titel"]
        }
    )

    flow = ET.SubElement(
        presentation,
        "flow"
    )

    erzeuge_fragematerial(
        flow,
        frage
    )
    
    
    #
    # Geordnete Antwortliste
    #
    response = ET.SubElement(
        flow,
        "response_lid",
        {
            "ident": "OQT",
            "rcardinality": "Ordered",
            "output": "javascript"
        }
    )

    render_choice = ET.SubElement(
        response,
        "render_choice",
        {
            "shuffle": "Yes"
        }
    )

    #
    # Antworten stehen bereits in der
    # richtigen Reihenfolge.
    #
    for i, antwort in enumerate(
        frage["antworten"]
    ):

        technische_id = (
            f"{10000 + i}_{i}_{i}_0"
        )

        response_label = ET.SubElement(
            render_choice,
            "response_label",
            {
                "ident": technische_id
            }
        )

        material_aussen = ET.SubElement(
            response_label,
            "material"
        )

        material_innen = ET.SubElement(
            material_aussen,
            "material"
        )

        xml_text(
            material_innen,
            "mattext",
            antwort,
            {
                "texttype": "text/plain"
            }
        )

        material_depth = ET.SubElement(
            response_label,
            "material"
        )

        xml_text(
            material_depth,
            "mattext",
            "0",
            {
                "label": "answerdepth"
            }
        )

    #
    # Bewertung
    #
    resprocessing = ET.SubElement(
        item,
        "resprocessing"
    )

    outcomes = ET.SubElement(
        resprocessing,
        "outcomes"
    )

    ET.SubElement(
        outcomes,
        "decvar"
    )

    anzahl = len(
        frage["antworten"]
    )

    punkte_je_position = (
        Decimal(str(frage["punkte"]))
        / Decimal(anzahl)
    )

    #
    # Bewertung der einzelnen Positionen.
    #
    for i in range(
        anzahl
    ):

        respcondition = ET.SubElement(
            resprocessing,
            "respcondition",
            {
                "continue": "Yes"
            }
        )

        conditionvar = ET.SubElement(
            respcondition,
            "conditionvar"
        )

        xml_text(
            conditionvar,
            "varequal",
            str(i),
            {
                "respident": "OQT",
                "index": str(i)
            }
        )

        xml_text(
            respcondition,
            "setvar",
            format_punkte(
                punkte_je_position
            ),
            {
                "action": "Add"
            }
        )

        ET.SubElement(
            respcondition,
            "displayfeedback",
            {
                "feedbacktype": "Response",
                "linkrefid": f"link_{i}"
            }
        )

    #
    # Generelle Rückmeldung bei mindestens
    # einer falschen Position.
    #
    if frage["rueckmeldung"]:

        respcondition = ET.SubElement(
            resprocessing,
            "respcondition",
            {
                "continue": "Yes"
            }
        )

        conditionvar = ET.SubElement(
            respcondition,
            "conditionvar"
        )

        or_elem = ET.SubElement(
            conditionvar,
            "or"
        )

        #
        # Die Gesamtlösung ist falsch,
        # sobald mindestens eine Antwort
        # nicht an ihrer richtigen Position steht.
        #
        for i in range(
            anzahl
        ):

            not_elem = ET.SubElement(
                or_elem,
                "not"
            )

            xml_text(
                not_elem,
                "varequal",
                str(i),
                {
                    "respident": "OQT",
                    "index": str(i)
                }
            )

        ET.SubElement(
            respcondition,
            "displayfeedback",
            {
                "feedbacktype": "Response",
                "linkrefid": "response_onenotcorrect"
            }
        )

    #
    # Leere individuelle Feedbackkanäle.
    #
    for i in range(
        anzahl
    ):

        itemfeedback = ET.SubElement(
            item,
            "itemfeedback",
            {
                "ident": f"link_{i}",
                "view": "All"
            }
        )

        flow_mat = ET.SubElement(
            itemfeedback,
            "flow_mat"
        )

        material = ET.SubElement(
            flow_mat,
            "material"
        )

        xml_text(
            material,
            "mattext",
            ""
        )

    #
    # Generelle Rückmeldung bei falscher
    # Reihenfolge.
    #
    if frage["rueckmeldung"]:

        itemfeedback = ET.SubElement(
            item,
            "itemfeedback",
            {
                "ident": "response_onenotcorrect",
                "view": "All"
            }
        )

        flow_mat = ET.SubElement(
            itemfeedback,
            "flow_mat"
        )

        material = ET.SubElement(
            flow_mat,
            "material"
        )

        xml_text(
            material,
            "mattext",
            frage["rueckmeldung"],
            {
                "texttype": "text/plain"
            }
        )



def erzeuge_anordnung_horizontal_item(
    root,
    frage,
    frage_id
):
    """
    Erzeugt eine horizontale Anordnungsfrage.
    """

    item = ET.SubElement(
        root,
        "item",
        {
            "ident": frage_id,
            "title": frage["titel"],
            "maxattempts": "0"
        }
    )

    xml_text(
        item,
        "qticomment",
        ""
    )

    #
    # Metadaten
    #
    itemmetadata = ET.SubElement(
        item,
        "itemmetadata"
    )

    qt_metadata = ET.SubElement(
        itemmetadata,
        "qtimetadata"
    )


    metadata(
        qt_metadata,
        "QUESTIONTYPE",
        "assOrderingHorizontal"
    )

    metadata(
        qt_metadata,
        "AUTHOR",
        frage["autor"]
    )

    metadata(
        qt_metadata,
        "additional_cont_edit_mode",
        "default"
    )

    metadata(
        qt_metadata,
        "externalId",
        frage_id
    )

    metadata(
        qt_metadata,
        "points",
        format_punkte(
            frage["punkte"]
        )
    )

    #
    # ILIAS speichert die richtige Reihenfolge
    # als einen Text mit "::" als Trennzeichen.
    #
    ordertext = " :: ".join(
        frage["antworten"]
    )

    metadata(
        qt_metadata,
        "ordertext",
        ordertext
    )

    #
    # Textgröße der horizontalen Elemente.
    #
    metadata(
        qt_metadata,
        "textsize",
        "100"
    )

    metadata(
        qt_metadata,
        "separator",
        "::"
    )

    #
    # Fragetext
    #
    presentation = ET.SubElement(
        item,
        "presentation",
        {
            "label": frage["titel"]
        }
    )

    flow = ET.SubElement(
        presentation,
        "flow"
    )

    erzeuge_fragematerial(
        flow,
        frage
    )


    #
    # Für die eigentliche horizontale Anordnung
    # benötigt ILIAS keine response_label-Blöcke.
    # Die Elemente stammen aus "ordertext".
    #

    #
    # Generelle Rückmeldung bei falscher Reihenfolge.
    #
    if frage["rueckmeldung"]:

        resprocessing = ET.SubElement(
            item,
            "resprocessing"
        )

        outcomes = ET.SubElement(
            resprocessing,
            "outcomes"
        )

        ET.SubElement(
            outcomes,
            "decvar"
        )

        respcondition = ET.SubElement(
            resprocessing,
            "respcondition",
            {
                "continue": "Yes"
            }
        )

        conditionvar = ET.SubElement(
            respcondition,
            "conditionvar"
        )

        not_elem = ET.SubElement(
            conditionvar,
            "not"
        )

        xml_text(
            not_elem,
            "varequal",
            format_punkte(
                frage["punkte"]
            ),
            {
                "respident": "points"
            }
        )

        ET.SubElement(
            respcondition,
            "displayfeedback",
            {
                "feedbacktype": "Response",
                "linkrefid": "response_onenotcorrect"
            }
        )

        itemfeedback = ET.SubElement(
            item,
            "itemfeedback",
            {
                "ident": "response_onenotcorrect",
                "view": "All"
            }
        )

        flow_mat = ET.SubElement(
            itemfeedback,
            "flow_mat"
        )

        material = ET.SubElement(
            flow_mat,
            "material"
        )

        xml_text(
            material,
            "mattext",
            frage["rueckmeldung"],
            {
                "texttype": "text/plain"
            }
        )


def erzeuge_freitext_item(root, frage, frage_id):
    """
    Erzeugt eine manuell zu bewertende Freitextfrage.
    """

    item = ET.SubElement(
        root,
        "item",
        {
            "ident": frage_id,
            "title": frage["titel"],
            "maxattempts": "0"
        }
    )

    #
    # Leerer QTI-Kommentar für die Freitextfrage.
    #
    ET.SubElement(
        item,
        "qticomment"
    )

    #
    # Technische Standarddauer für Freitextfragen.
    #
    xml_text(
        item,
        "duration",
        "P0Y0M0DT0H1M0S"
    )

    #
    # Metadaten
    #
    itemmetadata = ET.SubElement(
        item,
        "itemmetadata"
    )

    qt_metadata = ET.SubElement(
        itemmetadata,
        "qtimetadata"
    )


    metadata(
        qt_metadata,
        "QUESTIONTYPE",
        "TEXT QUESTION"
    )

    metadata(
        qt_metadata,
        "AUTHOR",
        frage["autor"]
    )

    metadata(
        qt_metadata,
        "additional_cont_edit_mode",
        "default"
    )

    metadata(
        qt_metadata,
        "externalId",
        frage_id
    )

    #
    # Technische Einstellungen der Freitextfrage.
    #
    # Die automatische Textbewertung ist deaktiviert.
    #
    metadata(
        qt_metadata,
        "textrating",
        "ci"
    )

    metadata(
        qt_metadata,
        "matchcondition",
        ""
    )

    metadata(
        qt_metadata,
        "termscoring",
        "YTowOnt9"
    )

    metadata(
        qt_metadata,
        "termrelation",
        "non"
    )

    metadata(
        qt_metadata,
        "specificfeedback",
        "non"
    )

    #
    # Darstellung
    #
    presentation = ET.SubElement(
        item,
        "presentation",
        {
            "label": frage["titel"]
        }
    )

    flow = ET.SubElement(
        presentation,
        "flow"
    )


    erzeuge_fragematerial(
        flow,
        frage
    )

    #
    # Großes Freitext-Eingabefeld
    #
    response_str = ET.SubElement(
        flow,
        "response_str",
        {
            "ident": "TEXT",
            "rcardinality": "Ordered"
        }
    )

    render_fib = ET.SubElement(
        response_str,
        "render_fib",
        {
            "fibtype": "String",
            "prompt": "Box"
        }
    )

    ET.SubElement(
        render_fib,
        "response_label",
        {
            "ident": "A"
        }
    )

    #
    # Manuelle Bewertung
    #
    resprocessing = ET.SubElement(
        item,
        "resprocessing",
        {
            "scoremodel": "HumanRater"
        }
    )

    outcomes = ET.SubElement(
        resprocessing,
        "outcomes"
    )

    ET.SubElement(
        outcomes,
        "decvar",
        {
            "varname": "WritingScore",
            "vartype": "Integer",
            "minvalue": "0",
            "maxvalue": format_punkte(
                frage["punkte"]
            )
        }
    )

    #
    # Rückmeldung bei nicht voller Punktzahl.
    #
    if frage["rueckmeldung"]:

        respcondition = ET.SubElement(
            resprocessing,
            "respcondition",
            {
                "continue": "Yes"
            }
        )

        conditionvar = ET.SubElement(
            respcondition,
            "conditionvar"
        )

        not_elem = ET.SubElement(
            conditionvar,
            "not"
        )

        xml_text(
            not_elem,
            "varequal",
            format_punkte(
                frage["punkte"]
            ),
            {
                "respident": "points"
            }
        )

        ET.SubElement(
            respcondition,
            "displayfeedback",
            {
                "feedbacktype": "Response",
                "linkrefid": "response_onenotcorrect"
            }
        )

    #
    # Kennzeichnung, dass die Frage
    # durch einen Tutor bewertet wird.
    #
    respcondition = ET.SubElement(
        resprocessing,
        "respcondition"
    )

    conditionvar = ET.SubElement(
        respcondition,
        "conditionvar"
    )

    xml_text(
        conditionvar,
        "other",
        "tutor_rated"
    )

    #
    # Generelle Rückmeldung bei nicht voller Punktzahl.
    #
    if frage["rueckmeldung"]:

        itemfeedback = ET.SubElement(
            item,
            "itemfeedback",
            {
                "ident": "response_onenotcorrect",
                "view": "All"
            }
        )

        flow_mat = ET.SubElement(
            itemfeedback,
            "flow_mat"
        )

        material = ET.SubElement(
            flow_mat,
            "material"
        )

        xml_text(
            material,
            "mattext",
            frage["rueckmeldung"],
            {
                "texttype": "text/xhtml"
            }
        )



def erzeuge_textluecke(
    flow,
    luecke,
    gap_id
):
    """
    Erzeugt die Eingabe einer Textlücke.
    """

    response = ET.SubElement(
        flow,
        "response_str",
        {
            "ident": gap_id,
            "rcardinality": "Single"
        }
    )

    ET.SubElement(
        response,
        "render_fib",
        {
            "fibtype": "String",
            "prompt": "Box",
            "columns": "5",
            "maxchars": "0"
        }
    )




def erzeuge_auswahlluecke(
    flow,
    luecke,
    gap_id
):
    """
    Erzeugt eine Auswahllücke.
    """

    response = ET.SubElement(
        flow,
        "response_str",
        {
            "ident": gap_id,
            "rcardinality": "Single"
        }
    )

    render_choice = ET.SubElement(
        response,
        "render_choice",
        {
            "shuffle": "Yes"
        }
    )

    for i, antwort in enumerate(
        luecke["antworten"]
    ):

        response_label = ET.SubElement(
            render_choice,
            "response_label",
            {
                "ident": str(i)
            }
        )

        material = ET.SubElement(
            response_label,
            "material"
        )

        xml_text(
            material,
            "mattext",
            antwort
        )
        
def erzeuge_zahlluecke(
    flow,
    luecke,
    gap_id
):
    """
    Erzeugt eine numerische Zahllücke
    mit symmetrischem Lösungsintervall.
    """

    response = ET.SubElement(
        flow,
        "response_num",
        {
            "ident": gap_id,
            "numtype": "Decimal",
            "rcardinality": "Single"
        }
    )

    ET.SubElement(
        response,
        "render_fib",
        {
            "fibtype": "Decimal",
            "prompt": "Box",
            "columns": "1",
            "maxchars": "0",
            "minnumber": luecke[
                "untergrenze"
            ],
            "maxnumber": luecke[
                "obergrenze"
            ]
        }
    )

def erzeuge_lueckentext_item(
    root,
    frage,
    frage_id
):
    """
    Erzeugt eine ILIAS-Lückentextfrage
    aus mehreren Lücken.
    """

    item = ET.SubElement(
        root,
        "item",
        {
            "ident": frage_id,
            "title": frage["titel"],
            "maxattempts": "0"
        }
    )

    xml_text(
        item,
        "qticomment",
        ""
    )

    xml_text(
        item,
        "duration",
        "P0Y0M0DT0H1M0S"
    )

    #
    # Metadaten für den ILIAS-Import.
    #
    itemmetadata = ET.SubElement(
        item,
        "itemmetadata"
    )

    qt_metadata = ET.SubElement(
        itemmetadata,
        "qtimetadata"
    )

    metadata(
        qt_metadata,
        "QUESTIONTYPE",
        "CLOZE QUESTION"
    )

    metadata(
        qt_metadata,
        "AUTHOR",
        frage["autor"]
    )

    metadata(
        qt_metadata,
        "additional_cont_edit_mode",
        "default"
    )

    metadata(
        qt_metadata,
        "externalId",
        frage_id
    )

    #
    # cs = Unterscheidung zwischen
    # Groß- und Kleinschreibung.
    # ci = keine Unterscheidung
    # l1 = Levenshtein-Abstand 1
    # l2 = Levenshtein-Abstand 2
    # l3 = Levenshtein-Abstand 3
    # l4 = Levenshtein-Abstand 4
    # l5 = Levenshtein-Abstand 5
    #
    metadata(
        qt_metadata,
        "textgaprating",
        "cs"
    )

    metadata(
        qt_metadata,
        "fixedTextLength",
        ""
    )

    metadata(
        qt_metadata,
        "identicalScoring",
        "1"
    )

    metadata(
        qt_metadata,
        "combinations",
        "W10="
    )

    presentation = ET.SubElement(
        item,
        "presentation",
        {
            "label": frage["titel"]
        }
    )

    flow = ET.SubElement(
        presentation,
        "flow"
    )



    #
    # Der Einleitungstext ist bei einer
    # Lückentextfrage Bestandteil der
    # eigentlichen QTI-Frage.
    #
    # ILIAS benötigt an dieser Stelle auch dann
    # einen Inhalt, wenn kein sichtbarer
    # Einleitungstext verwendet wird.
    #
    # &#8203; erzeugt ein unsichtbares
    # Zero-Width Space und verhindert dadurch
    # eine zusätzliche Leerzeile.
    #


    einleitung = (
        frage["einleitung"]
        if frage["einleitung"]
        else "&#8203;"
    )

    material = ET.SubElement(
        flow,
        "material"
    )

    xml_text(
        material,
        "mattext",
        (
            f'<p>{einleitung}</p>'
        ),
        {
            "texttype": "text/xhtml"
        }
    )




    #
    # Optionales Bild der gesamten
    # Lückentextfrage.
    #
    if frage.get("bild_mob_id"):

        material = ET.SubElement(
            flow,
            "material"
        )

        xml_text(
            material,
            "mattext",
            (
                f'<p><img src="'
                f'{frage["bild_mob_id"]}'
                f'" /></p>'
            ),
            {
                "texttype": "text/xhtml"
            }
        )

        ET.SubElement(
            material,
            "matimage",
            {
                "label": frage["bild_mob_id"],
                "uri": frage["bild_qti_pfad"]
            }
        )


    #
    # Einzelne Lücken
    #
    for index, luecke in enumerate(
        frage["luecken"]
    ):

        gap_id = (
            f"gap_{index}"
        )


        nummerierung = roemische_zahl(
            index + 1
        )

        punktetext = (
            "1 Punkt"
            if luecke["punkte"] == 1
            else (
                f"{format_punkte(luecke['punkte'])} "
                "Punkte"
            )
        )

        hat_mehrere_luecken = (
            len(frage["luecken"]) > 1
        )

        if frage["punktanzeige"] == "ja":

            if hat_mehrere_luecken:

                praefix = (
                    f"({nummerierung}) "
                    f"({punktetext}) "
                )

            else:

                praefix = (
                    f"({punktetext}) "
                )

        else:

            if hat_mehrere_luecken:

                praefix = (
                    f"({nummerierung}) "
                )

            else:

                praefix = ""












        #
        # Jede Teilaufgabe beginnt mit einem
        # eigenen Absatz:
        #
        # (i) (2 Punkte) Text vor der Lücke
        #
        material = ET.SubElement(
            flow,
            "material"
        )

        xml_text(
            material,
            "mattext",
            (
                f"<p>{praefix}"
                f"{luecke['frage']} "            ),
            {
                "texttype": "text/xhtml"
            }
        )

        if luecke["lueckentyp"] == "TEXT":

            erzeuge_textluecke(
                flow,
                luecke,
                gap_id
            )

        elif luecke["lueckentyp"] == "AUSWAHL":

            erzeuge_auswahlluecke(
                flow,
                luecke,
                gap_id
            )

        elif luecke["lueckentyp"] == "ZAHL":

            erzeuge_zahlluecke(
                flow,
                luecke,
                gap_id
            )

        else:

            raise ValueError(
                "Unbekannter interner "
                "Lückentyp: "
                f"{luecke['lueckentyp']}"
            )

        #
        # Text hinter der Lücke beendet
        # den Absatz.
        #
        material = ET.SubElement(
            flow,
            "material"
        )

        xml_text(
            material,
            "mattext",
            (
                f" {luecke['text_nach_luecke']}"
                "</p>"
            ),
            {
                "texttype": "text/xhtml"
            }
        )

    #
    # Bewertung
    #
    resprocessing = ET.SubElement(
        item,
        "resprocessing"
    )

    outcomes = ET.SubElement(
        resprocessing,
        "outcomes"
    )

    ET.SubElement(
        outcomes,
        "decvar"
    )

    #
    # Jede Lücke wird einzeln bewertet.
    #
    for index, luecke in enumerate(
        frage["luecken"]
    ):

        gap_id = (
            f"gap_{index}"
        )

        respcondition = ET.SubElement(
            resprocessing,
            "respcondition",
            {
                "continue": "Yes"
            }
        )

        conditionvar = ET.SubElement(
            respcondition,
            "conditionvar"
        )



        if luecke["lueckentyp"] == "TEXT":

            for richtige_antwort in luecke["richtige_antworten"]:

                respcondition = ET.SubElement(
                    resprocessing,
                    "respcondition",
                    {
                        "continue": "Yes"
                    }
                )

                conditionvar = ET.SubElement(
                    respcondition,
                    "conditionvar"
                )

                xml_text(
                    conditionvar,
                    "varequal",
                    richtige_antwort,
                    {
                        "respident": gap_id
                    }
                )

                xml_text(
                    respcondition,
                    "setvar",
                    format_punkte(
                        luecke["punkte"]
                    ),
                    {
                        "action": "Add"
                    }
                )

        else:

            respcondition = ET.SubElement(
                resprocessing,
                "respcondition",
                {
                    "continue": "Yes"
                }
            )

            conditionvar = ET.SubElement(
                respcondition,
                "conditionvar"
            )

            xml_text(
                conditionvar,
                "varequal",
                luecke["richtige_antwort"],
                {
                    "respident": gap_id
                }
            )

            xml_text(
                respcondition,
                "setvar",
                format_punkte(
                    luecke["punkte"]
                ),
                {
                    "action": "Add"
                }
            )


        #
        # Technischer individueller Feedbackkanal.
        #
        ET.SubElement(
            respcondition,
            "displayfeedback",
            {
                "feedbacktype": "Response",
                "linkrefid": (
                    f"{index}_Response_0"
                )
            }
        )

    #
    # Allgemeine Rückmeldung, wenn mindestens
    # eine Lücke nicht richtig beantwortet wurde.
    #
    if frage["rueckmeldung"]:

        respcondition = ET.SubElement(
            resprocessing,
            "respcondition",
            {
                "continue": "Yes"
            }
        )

        conditionvar = ET.SubElement(
            respcondition,
            "conditionvar"
        )

        not_elem = ET.SubElement(
            conditionvar,
            "not"
        )

        and_elem = ET.SubElement(
            not_elem,
            "and"
        )



        for index, luecke in enumerate(
            frage["luecken"]
        ):

            gap_id = (
                f"gap_{index}"
            )

            if (
                luecke["lueckentyp"] == "TEXT"
                and len(luecke["richtige_antworten"]) > 1
            ):

                or_elem = ET.SubElement(
                    and_elem,
                    "or"
                )

                for richtige_antwort in luecke["richtige_antworten"]:

                    xml_text(
                        or_elem,
                        "varequal",
                        richtige_antwort,
                        {
                            "respident": gap_id
                        }
                    )

            else:

                xml_text(
                    and_elem,
                    "varequal",
                    luecke["richtige_antwort"],
                    {
                        "respident": gap_id
                    }
                )


        ET.SubElement(
            respcondition,
            "displayfeedback",
            {
                "feedbacktype": "Response",
                "linkrefid": (
                    "response_onenotcorrect"
                )
            }
        )

    #
    # Die individuellen Feedbackblöcke bleiben
    # aus Kompatibilitätsgründen vorhanden,
    # enthalten aber keinen Text.
    #
    for index, _luecke in enumerate(
        frage["luecken"]
    ):

        itemfeedback = ET.SubElement(
            item,
            "itemfeedback",
            {
                "ident": str(index),
                "view": "All"
            }
        )

        flow_mat = ET.SubElement(
            itemfeedback,
            "flow_mat"
        )

        material = ET.SubElement(
            flow_mat,
            "material"
        )

        xml_text(
            material,
            "mattext",
            "",
            {
                "texttype": "text/xhtml"
            }
        )

    #
    # Eine einzige inhaltliche Rückmeldung
    # für die gesamte Frage.
    #
    if frage["rueckmeldung"]:

        itemfeedback = ET.SubElement(
            item,
            "itemfeedback",
            {
                "ident": "response_onenotcorrect",
                "view": "All"
            }
        )

        flow_mat = ET.SubElement(
            itemfeedback,
            "flow_mat"
        )

        material = ET.SubElement(
            flow_mat,
            "material"
        )

        xml_text(
            material,
            "mattext",
            frage["rueckmeldung"],
            {
                "texttype": "text/xhtml"
            }
        )    









def bereite_bilder_vor(
    fragen,
    basisverzeichnis,
    zielordner,
    pool_id
):
    """
    Bereitet Bilddateien für den ILIAS-Export vor.

    Relative Bildpfade werden relativ zum
    Verzeichnis der eingelesenen Quelldatei
    aufgelöst.

    Jede Frage kann höchstens eine Abbildung
    enthalten.
    """

    basisverzeichnis = Path(
        basisverzeichnis
    )

    for nummer, frage in enumerate(
        fragen,
        start=1
    ):

        bild_angabe = textwert(
            frage.get("bild", "")
        )

        if not bild_angabe:
            continue




        bild_datei = Path(
            bild_angabe
        )

        #
        # Bilddateien dürfen ausschließlich im
        # selben Verzeichnis wie die Eingabedatei liegen.
        #
        # Pfadangaben wie "../bild.png",
        # Unterverzeichnisse oder absolute Pfade
        # werden nicht akzeptiert.
        #
        if (
            bild_datei.is_absolute()
            or bild_datei.parent != Path(".")
        ):

            raise ValueError(
                "Bilddateien müssen im selben Verzeichnis "
                "wie die Eingabedatei liegen. "
                f"Ungültige Pfadangabe: {bild_angabe}"
            )

        basisverzeichnis_aufgeloest = basisverzeichnis.resolve()
        bild_datei = (
            basisverzeichnis_aufgeloest
            / bild_datei.name
        ).resolve()

        # Auch symbolische Links dürfen nicht aus dem Verzeichnis
        # der Eingabedatei herausführen.
        if bild_datei.parent != basisverzeichnis_aufgeloest:
            raise ValueError(
                "Bilddateien müssen tatsächlich im selben Verzeichnis "
                "wie die Eingabedatei liegen. "
                f"Ungültige Bilddatei: {bild_angabe}"
            )

        if not bild_datei.is_file():

            raise FileNotFoundError(
                "Bilddatei nicht gefunden: "
                f"{bild_datei}"
            )









        #
        # Jede Frage erhält ein eigenes
        # ILIAS-Media-Object.
        #
        mob_id = (
            f"il_0_mob_"
            f"{pool_id}{nummer:03d}"
        )

        mob_ordner = (
            zielordner
            / "objects"
            / mob_id
        )

        mob_ordner.mkdir(
            parents=True,
            exist_ok=True
        )

        zielbild = (
            mob_ordner
            / bild_datei.name
        )

        shutil.copy2(
            bild_datei,
            zielbild
        )

        #
        # Diese Angaben benötigt später
        # die QTI-Datei.
        #
        frage["bild_mob_id"] = mob_id

        frage["bild_qti_pfad"] = (
            f"objects/"
            f"{mob_id}/"
            f"{bild_datei.name}"
        )



def bereite_audios_vor(
    fragen,
    basisverzeichnis,
    zielordner,
    pool_id
):
    """
    Bereitet MP3-Dateien für den ILIAS-Export vor.

    Audiodateien müssen im selben Verzeichnis
    wie die Eingabedatei liegen.
    """

    basisverzeichnis = Path(
        basisverzeichnis
    )

    for nummer, frage in enumerate(
        fragen,
        start=1
    ):

        audio_angabe = textwert(
            frage.get("audio", "")
        )

        if not audio_angabe:
            continue

        audio_datei = Path(
            audio_angabe
        )

        if (
            audio_datei.is_absolute()
            or audio_datei.parent != Path(".")
        ):
            raise ValueError(
                "Audiodateien müssen im selben Verzeichnis "
                "wie die Eingabedatei liegen. "
                f"Ungültige Pfadangabe: {audio_angabe}"
            )

        if audio_datei.suffix.lower() != ".mp3":
            raise ValueError(
                "Derzeit werden als Audiodateien nur MP3-Dateien unterstützt. "
                f"Ungültige Datei: {audio_angabe}"
            )

        basisverzeichnis_aufgeloest = basisverzeichnis.resolve()

        audio_datei = (
            basisverzeichnis_aufgeloest
            / audio_datei.name
        ).resolve()

        if audio_datei.parent != basisverzeichnis_aufgeloest:
            raise ValueError(
                "Audiodateien müssen tatsächlich im selben Verzeichnis "
                "wie die Eingabedatei liegen. "
                f"Ungültige Audiodatei: {audio_angabe}"
            )

        if not audio_datei.is_file():
            raise FileNotFoundError(
                "Audiodatei nicht gefunden: "
                f"{audio_datei}"
            )

        mob_id = (
            f"il_0_mob_"
            f"{pool_id}a{nummer:03d}"
        )

        mob_ordner = (
            zielordner
            / "objects"
            / mob_id
        )

        mob_ordner.mkdir(
            parents=True,
            exist_ok=True
        )

        zielaudio = (
            mob_ordner
            / audio_datei.name
        )

        shutil.copy2(
            audio_datei,
            zielaudio
        )

        frage["audio_mob_id"] = mob_id
        frage["audio_dateiname"] = audio_datei.name




def bereite_videos_vor(
    fragen,
    basisverzeichnis,
    zielordner,
    pool_id
):
    """
    Bereitet MP4- und WebM-Dateien für den ILIAS-Export vor.

    Videodateien müssen im selben Verzeichnis
    wie die Eingabedatei liegen.
    """

    video_formate = {
        ".mp4": "video/mp4",
        ".webm": "video/webm"
    }

    basisverzeichnis = Path(
        basisverzeichnis
    )

    for nummer, frage in enumerate(
        fragen,
        start=1
    ):

        video_angabe = textwert(
            frage.get("video", "")
        )

        if not video_angabe:
            continue

        video_datei = Path(
            video_angabe
        )

        if (
            video_datei.is_absolute()
            or video_datei.parent != Path(".")
        ):
            raise ValueError(
                "Videodateien müssen im selben Verzeichnis "
                "wie die Eingabedatei liegen. "
                f"Ungültige Pfadangabe: {video_angabe}"
            )

        endung = video_datei.suffix.lower()

        if endung not in video_formate:
            raise ValueError(
                "Als Videodateien werden derzeit nur "
                "MP4- und WebM-Dateien unterstützt. "
                f"Ungültige Datei: {video_angabe}"
            )

        basisverzeichnis_aufgeloest = (
            basisverzeichnis.resolve()
        )

        video_datei = (
            basisverzeichnis_aufgeloest
            / video_datei.name
        ).resolve()

        if video_datei.parent != basisverzeichnis_aufgeloest:
            raise ValueError(
                "Videodateien müssen tatsächlich im selben "
                "Verzeichnis wie die Eingabedatei liegen. "
                f"Ungültige Videodatei: {video_angabe}"
            )

        if not video_datei.is_file():
            raise FileNotFoundError(
                "Videodatei nicht gefunden: "
                f"{video_datei}"
            )

        mob_id = (
            f"il_0_mob_"
            f"{pool_id}v{nummer:03d}"
        )

        mob_ordner = (
            zielordner
            / "objects"
            / mob_id
        )

        mob_ordner.mkdir(
            parents=True,
            exist_ok=True
        )

        zielvideo = (
            mob_ordner
            / video_datei.name
        )

        shutil.copy2(
            video_datei,
            zielvideo
        )

        frage["video_mob_id"] = mob_id

        frage["video_dateiname"] = (
            video_datei.name
        )

        frage["video_mime"] = (
            video_formate[endung]
        )










def erzeuge_qti(fragen, basisname, zielordner):
    """
    Erzeugt die QTI-Datei.
    """
    root = ET.Element(
        "questestinterop"
    )

    for nummer, frage in enumerate(
        fragen,
        start=1
    ):

        frage_id = (
            f"{basisname}"
            f"{nummer:03d}"
        )

        if frage["typ"] == "SC":
            erzeuge_sc_item(
                root,
                frage,
                frage_id
            )

        elif frage["typ"] == "MC":
            erzeuge_mc_item(
                root,
                frage,
                frage_id
            )
            
        elif frage["typ"] == "KPRIM":
            erzeuge_kprim_item(
                root,
                frage,
                frage_id
            )    

        elif frage["typ"] == "VERT_ANORD":

            erzeuge_anordnung_vertikal_item(
                root,
                frage,
                frage_id
            )
            
        elif frage["typ"] == "HOR_ANORD":

            erzeuge_anordnung_horizontal_item(
                root,
                frage,
                frage_id
            )            

        elif frage["typ"] == "FREITEXT":

            erzeuge_freitext_item(
                root,
                frage,
                frage_id
            )


        elif frage["typ"] == "LUECKENTEXT":

            erzeuge_lueckentext_item(
                root,
                frage,
                frage_id
            )

        else:
            raise NotImplementedError(
                f"XML für {frage['typ']} ist nicht implementiert."
            )

    tree = ET.ElementTree(root)


    datei = (
        zielordner
        / f"{basisname}_qti.xml"
    )

    xml_inhalt = ET.tostring(
        root,
        encoding="unicode"
    )

    with open(
        datei,
        "w",
        encoding="utf-8"
    ) as f:

        f.write(
            '<?xml version="1.0" encoding="utf-8"?>'
        )

        f.write(
            '<!DOCTYPE questestinterop SYSTEM '
            '"ims_qtiasiv1p2p1.dtd">'
        )

        f.write(
            xml_inhalt
        )
    return datei


def erzeuge_qpl(fragen, basisname, zielordner, pooltitel="QuestionSheet2LMS"):
    """
    Erzeugt die QPL-Datei für den ILIAS-Fragenpoolimport.
    """


    root = ET.Element(
        "ContentObject",
        {
            "Type": "Questionpool_Test"
        }
    )

    metadata_element = ET.SubElement(
        root,
        "MetaData"
    )

    general = ET.SubElement(
        metadata_element,
        "General",
        {
            "Structure": "Hierarchical"
        }
    )

    xml_text(
        general,
        "Title",
        pooltitel,
        {
            "Language": "de"
        }
    )

    ET.SubElement(
        general,
        "Language",
        {
            "Language": "de"
        }
    )

    xml_text(
        general,
        "Description",
        "Erzeugt mit QuestionSheet2LMS",
        {
            "Language": "de"
        }
    )

    xml_text(
        general,
        "Keyword",
        "",
        {
            "Language": "de"
        }
    )

    settings = ET.SubElement(
        root,
        "Settings"
    )

    xml_text(
        settings,
        "SkillService",
        "0"
    )


    #
    # Fragen
    #

    for nummer, frage in enumerate(
        fragen,
        start=1
    ):

        #
        # Dieselbe Frage-ID wie in der QTI-Datei.
        #
        frage_id = (
            f"{basisname}"
            f"{nummer:03d}"
        )

        page_object = ET.SubElement(
            root,
            "PageObject"
        )

        #
        # Einleitungstext bleibt erhalten.
        #
        if (
            frage["einleitung"]
            and frage["typ"] != "LUECKENTEXT"
        ):
            page_content = ET.SubElement(
                page_object,
                "PageContent"
            )

            #
            # Der Einleitungstext wird von ILIAS beim Import
            # ein zweites Mal als PageObject-XML verarbeitet.
            #
            # Deshalb muss "&" hier eine zusätzliche
            # Escape-Stufe erhalten.
            #
            einleitung = frage["einleitung"].replace(
                "&",
                "&amp;"
            )

            xml_text(
                page_content,
                "Paragraph",
                einleitung,
                {
                    "Language": "de",
                    "Characteristic": "Standard"
                }
            )





        if frage.get("audio_mob_id"):

            page_content = ET.SubElement(
                page_object,
                "PageContent"
            )

            media_object = ET.SubElement(
                page_content,
                "MediaObject"
            )

            ET.SubElement(
                media_object,
                "MediaAlias",
                {
                    "OriginId": frage["audio_mob_id"],
                    "Class": "MediaContainer"
                }
            )

            media_alias_item = ET.SubElement(
                media_object,
                "MediaAliasItem",
                {
                    "Purpose": "Standard"
                }
            )

            ET.SubElement(
                media_alias_item,
                "Layout",
                {
                    "HorizontalAlign": "Left"
                }
            )











        if frage.get("video_mob_id"):

            page_content = ET.SubElement(
                page_object,
                "PageContent"
            )

            media_object = ET.SubElement(
                page_content,
                "MediaObject"
            )

            ET.SubElement(
                media_object,
                "MediaAlias",
                {
                    "OriginId": frage["video_mob_id"],
                    "Class": "MediaContainer"
                }
            )

            media_alias_item = ET.SubElement(
                media_object,
                "MediaAliasItem",
                {
                    "Purpose": "Standard"
                }
            )

            ET.SubElement(
                media_alias_item,
                "Layout",
                {
                    "HorizontalAlign": "Left"
                }
            )















        page_content = ET.SubElement(
            page_object,
            "PageContent"
        )

        ET.SubElement(
            page_content,
            "Question",
            {
                "QRef": frage_id
            }
        )


    #
    # Audio-MediaObjects
    #
    for frage in fragen:

        if not frage.get("audio_mob_id"):
            continue

        media_object = ET.SubElement(
            root,
            "MediaObject"
        )

        metadata_element = ET.SubElement(
            media_object,
            "MetaData"
        )

        general = ET.SubElement(
            metadata_element,
            "General",
            {
                "Structure": "Hierarchical"
            }
        )

        identifier = ET.SubElement(
            general,
            "Identifier",
            {
                "Catalog": "ILIAS",
                "Entry": frage["audio_mob_id"]
            }
        )

        xml_text(
           general,
            "Title",
            pooltitel,
            {
                "Language": "de"
            }
        )

        ET.SubElement(
            general,
            "Language",
            {
                "Language": "de"
            }
        )

        xml_text(
            general,
            "Description",
            "",
            {
                "Language": "de"
            }
        )

        xml_text(
            general,
            "Keyword",
            "",
            {
                "Language": "de"
            }
        )

        media_item = ET.SubElement(
            media_object,
            "MediaItem",
            {
                "Purpose": "Standard"
            }
        )

        xml_text(
            media_item,
            "Location",
            frage["audio_dateiname"],
            {
                "Type": "LocalFile"
            }
        )

        xml_text(
            media_item,
            "Format",
            "audio/mpeg"
        )

        ET.SubElement(
            media_item,
            "Layout",
            {
                "HorizontalAlign": "Left"
            }
        )







    #
    # Video-MediaObjects
    #
    for frage in fragen:

        if not frage.get("video_mob_id"):
            continue

        media_object = ET.SubElement(
            root,
            "MediaObject"
        )

        metadata_element = ET.SubElement(
            media_object,
            "MetaData"
        )

        general = ET.SubElement(
            metadata_element,
            "General",
            {
                "Structure": "Hierarchical"
            }
        )

        ET.SubElement(
            general,
            "Identifier",
            {
                "Catalog": "ILIAS",
                "Entry": frage["video_mob_id"]
            }
        )

        xml_text(
            general,
            "Title",
            pooltitel,
            {
                "Language": "de"
            }
        )

        ET.SubElement(
            general,
            "Language",
            {
                "Language": "de"
            }
        )

        xml_text(
            general,
            "Description",
            "",
            {
                "Language": "de"
            }
        )

        xml_text(
            general,
            "Keyword",
            "",
            {
                "Language": "de"
            }
        )

        media_item = ET.SubElement(
            media_object,
            "MediaItem",
            {
                "Purpose": "Standard"
            }
        )

        xml_text(
            media_item,
            "Location",
            frage["video_dateiname"],
            {
                "Type": "LocalFile"
            }
        )

        xml_text(
            media_item,
            "Format",
            frage["video_mime"]
        )

        ET.SubElement(
            media_item,
            "Layout",
            {
                "HorizontalAlign": "Left"
            }
        )










    #
    # XML erzeugen
    #
    xml_inhalt = ET.tostring(
        root,
        encoding="unicode"
    )

    datei = (
        zielordner
        / f"{basisname}_qpl.xml"
    )

    with open(
        datei,
        "w",
        encoding="utf-8"
    ) as f:

        f.write(
            '<?xml version="1.0" encoding="utf-8"?>'
        )

        f.write(
            '<!DOCTYPE Test SYSTEM '
            '"http://www.ilias.uni-koeln.de/download/dtd/ilias_co.dtd">'
        )

        f.write(
            '<!--Export of ILIAS Test Questionpool '
            'generated by QuestionSheet2LMS.-->'
        )

        f.write(
            xml_inhalt
        )

    return datei



def erzeuge_pool_id():
    """
    Erzeugt eine einfache sechsstellige ID.
    """
    return datetime.now().strftime(
        "%H%M%S"
    )



def loesche_exportarbeitsordner(zielordner, ausgabeverzeichnis):
    """
    Löscht ausschließlich von QuestionSheet2LMS erzeugte Arbeitsordner
    direkt innerhalb des vorgesehenen Ausgabeverzeichnisses.
    """
    ziel = Path(zielordner).resolve()
    basis = Path(ausgabeverzeichnis).resolve()
    name = ziel.name

    ist_poolordner = (
        name.startswith("il_1001_qst_")
        and name.endswith("_qpl")
    )
    ist_testordner = (
        "__0__tst_" in name
    )

    if ziel.parent != basis or not (ist_poolordner or ist_testordner):
        raise RuntimeError(
            "Sicherheitsabbruch: Unerwarteter temporärer Exportpfad. "
            f"Ordner wurde nicht gelöscht: {ziel}"
        )

    if ziel.exists():
        shutil.rmtree(ziel)


def erzeuge_zip(fragen, ausgabeverzeichnis, pooltitel="QuestionSheet2LMS"):
    """Erzeugt einen klassischen ILIAS-Fragenpool-Export."""
    pool_id = erzeuge_pool_id()
    basisname = f"il_1001_qst_{pool_id}"
    ordnername = f"{basisname}_qpl"

    ausgabeverzeichnis = Path(ausgabeverzeichnis)
    zielordner = ausgabeverzeichnis / ordnername
    zip_datei = ausgabeverzeichnis / f"{ordnername}.zip"
    zielordner.mkdir(parents=True, exist_ok=True)

    try:
        bereite_bilder_vor(
            fragen,
            ausgabeverzeichnis,
            zielordner,
            pool_id
        )

        bereite_audios_vor(
            fragen,
            ausgabeverzeichnis,
            zielordner,
            pool_id
        )




        bereite_videos_vor(
            fragen,
            ausgabeverzeichnis,
            zielordner,
            pool_id
        )




        erzeuge_qpl(
            fragen,
            basisname,
            zielordner,
            pooltitel
        )

        erzeuge_qti(
            fragen,
            basisname,
            zielordner
        )

        with ZipFile(
            zip_datei,
            "w",
            ZIP_DEFLATED
        ) as z:

            for datei in zielordner.rglob("*"):
                if datei.is_file():
                    arcname = (
                        Path(ordnername)
                        / datei.relative_to(zielordner)
                    )
                    z.write(
                        datei,
                        str(arcname).replace("\\", "/")
                    )

        return zip_datei

    finally:
        loesche_exportarbeitsordner(
            zielordner,
            ausgabeverzeichnis
        )




# -----------------------------------------------------------------------------
# Test-Export
# -----------------------------------------------------------------------------


def test_metadaten_11(autor):
    """
    Metadaten fuer den standardisierten QuestionSheet2LMS-Test.

    Grundlage ist der unter ILIAS 11 exportierte Test mit den
    bewaehrten Pruefungseinstellungen. Pruefungsindividuelle
    Angaben (Passwort, zeitliche Verfuegbarkeit, Online-Status)
    werden bewusst nicht vorbelegt.

    Pruefungsbedingungen werden erst aktiviert, sobald ein
    verbindlicher Text hinterlegt wird.
    """
    return [
        ("ILIAS_VERSION", "11.3 2026-08-11"),
        ("anonymity", "0"),
        ("question_set_type", "FIXED_QUEST_SET"),
        ("sequence_settings", "0"),
        ("author", autor),
        ("reset_processing_time", "0"),
        ("count_system", "0"),
        ("score_cutting", "0"),
        ("password", ""),
        ("ip_range_from", ""),
        ("ip_range_to", ""),
        ("pass_scoring", "0"),
        ("pass_deletion_allowed", "0"),
        ("nr_of_tries", "1"),
        ("block_after_passed", "0"),
        ("pass_waiting", ""),
        ("kiosk", "7"),
        ("redirection_mode", "0"),
        ("redirection_url", ""),
        ("use_previous_answers", "0"),
        ("question_list_enabled", "0"),
        ("title_output", "0"),
        ("results_presentation", "3"),
        ("examid_in_test_pass", "1"),
        ("examid_in_test_res", "0"),
        ("usr_pass_overview_mode", "0"),
        ("score_reporting", "0"),
        ("show_solution_list_comparison", "0"),
        ("instant_verification", "0"),
        ("answer_feedback", "0"),
        ("instant_feedback_specific", "0"),
        ("answer_feedback_points", "0"),
        ("follow_qst_answer_fixation", "0"),
        ("instant_feedback_answer_fixation", "0"),
        ("force_instant_feedback", "0"),
        ("highscore_enabled", "0"),
        ("highscore_anon", "0"),
        ("highscore_achieved_ts", "0"),
        ("highscore_score", "0"),
        ("highscore_percentage", "0"),
        ("highscore_wtime", "0"),
        ("highscore_own_table", "0"),
        ("highscore_top_table", "0"),
        ("highscore_top_num", "10"),
        ("suspend_test_allowed", "0"),
        ("show_marker", "0"),
        ("fixed_participants", "0"),
        # Einleitung/Pruefungsbedingungen werden im Test angezeigt.
        ("show_introduction", "1"),
        ("exam_conditions", "1"),
        ("show_concluding_remarks", "0"),
        ("exportsettings", "0"),
        ("shuffle_questions", "0"),
        ("processing_time", ""),
        ("enable_examview", "1"),
        ("skill_service", "0"),
        ("show_grading_status", "1"),
        ("show_grading_mark", "1"),
        ("hide_info_tab", "0"),
        ("autosave", "1"),
        ("autosave_ival", "30000"),
        # ILIAS 11 exportiert diese beiden Felder doppelt.
        ("instant_feedback_specific", "0"),
        ("instant_feedback_answer_fixation", "0"),
        ("enable_processing_time", "0"),
        ("mark_step_0", "<short>nicht bestanden</short><official>nicht bestanden</official><percentage>0.00</percentage><passed>0</passed>"),
        ("mark_step_1", "<short>bestanden</short><official>bestanden</official><percentage>50.00</percentage><passed>1</passed>"),
    ]


PRUEFUNGSBEDINGUNGEN_TEXT = """Prüfungsregeln und Belehrung

Bitte lesen Sie sich die folgenden Regeln genau durch. Mit dem Beginn der Klausur stimmen Sie diesen Bedingungen und Regeln zu.

Berechtigte zur Teilnahme an der Prüfung

Sie dürfen diesen Test nur starten, wenn Sie für diese Prüfung angemeldet sind.

Hilfsmittel

Alle Hilfsmittel, die nicht ausdrücklich erlaubt sind, sind verboten. Beispielsweise ist kein Zugriff auf das Internet, keine Nutzung der Anwendungen auf den Prüfungsrechnern und sind keine Kommunikationsmittel erlaubt, auch wenn dies technisch möglich (nicht gesperrt) sein sollte.

Ausweisen am Bearbeitungsplatz

An der Prüfung teilnehmende Personen müssen sich zu jeder Zeit vor und während der Prüfung - unabhängig von einem elektronischen Gerät - ausweisen können. Dieses geschieht durch einen amtlichen Lichtbildausweis. Dieser Ausweis hat während der Bearbeitungszeit für die Aufsicht einsehbar an dem Bearbeitungsplatz zu liegen. Das Ausweisen mit einem Smartphone ist ausdrücklich nicht erlaubt, da dies zu einem Täuschungsversuch führt (siehe andere Regeln).

Abschluss der Prüfung

Die Klausur muss ausdrücklich durch "Test beenden" (blaues oder grünes Feld im linken oberen Bereich der Seite) beendet werden. Anderenfalls ist die vollständige Speicherung der Antworten nicht gesichert.

Dauer der Prüfung

Die Prüfung beginnt mit der Möglichkeit, sich in dem jeweiligen Test-Element durch Eingabe des Test-Passworts anzumelden. Die Prüfungsdauer wird Ihnen zu Beginn von der Aufsicht genannt. Die Restzeit ist auf dem Beamer-Bild ablesbar. Die Aufsicht oder der Prüfer teilt das Ende der Prüfung mit. Bitte bleiben Sie solange an Ihrem Platz ruhig sitzen.

Prüfungsfähigkeit

Mit dem Start eines Tests stimmen Sie den Prüfungsregeln zu und bestätigen Sie Ihre Prüfungsfähigkeit.

Erklärung und Beachtung der gültigen Studien- und Prüfungsordnung

Es gelten die aktuellen Prüfungsregeln sowie die Studien- und Prüfungsordnungen.

Eigenständigkeitserklärung

Mit der Beendigung eines Tests bestätigen Sie die eigenständige Bearbeitung der Prüfungsleistung (siehe folgende Belehrung).

Belehrung zur Eigenständigkeit

„Prüfungsleistungen müssen selbstständig und ohne unzulässige fremde Hilfe erbracht worden sein, ansonsten liegt ein Täuschungsversuch vor. Es dürfen keine anderen als die angegebenen Quellen und Hilfsmittel benutzt werden, bei Hausaufgaben, Referaten und Projektarbeiten müssen Zitate kenntlich gemacht werden.“

Rechtsfolgen einer Täuschung

„Bei einem Täuschungsversuch wird die betreffende Prüfungsleistung mit ‚nicht ausreichend‘ (5,0) bewertet. Wer vorsätzlich gegen eine die Täuschung über Prüfungsleistungen betreffende Regelung einer Hochschulprüfungsordnung verstößt, handelt ordnungswidrig. Die Ordnungswidrigkeit kann mit einer Geldbuße von bis zu 50.000,00 € geahndet werden. Im Falle eines mehrfachen oder sonstigen schwerwiegenden Täuschungsversuches kann der Prüfling von der Erbringung weiterer Prüfungs- und Studienleistungen ausgeschlossen und zudem exmatrikuliert werden. (§ 11 Abs. 3 RahmenPO, § 63 Abs. 5 HG)“"""


def erzeuge_objectives_base64(text):
    """
    Erzeugt den von ILIAS 11 erwarteten Base64-kodierten
    PageObject-Block fuer Einleitung/Pruefungsbedingungen.

    Doppelte Leerzeilen im Ausgangstext werden als echte
    sichtbare Absatzabstaende ausgegeben. Ueberschriften
    werden fett dargestellt.
    """
    page = ET.Element("PageObject")
    page_content = ET.SubElement(
        page,
        "PageContent",
        {"PCID": "qs2lms_exam_conditions_0001"}
    )
    paragraph = ET.SubElement(
        page_content,
        "Paragraph",
        {
            "Language": "de",
            "Characteristic": "Standard"
        }
    )

    # Der Text ist als abwechselnde Bloecke
    # "Ueberschrift / Absatz" aufgebaut.
    bloecke = [
        block.strip()
        for block in text.split("\n\n")
        if block.strip()
    ]

    for index, block in enumerate(bloecke):
        ist_ueberschrift = (index % 2 == 0)

        if ist_ueberschrift:
            strong = ET.SubElement(paragraph, "Strong")
            strong.text = block
            strong.tail = ""
        else:
            # Text nach dem vorherigen Element als XML-Tail
            # einfuegen, damit ILIAS ihn innerhalb desselben
            # Paragraph-Elements korrekt rendert.
            if len(paragraph):
                letztes = paragraph[-1]
                letztes.tail = (letztes.tail or "") + block
            else:
                paragraph.text = (paragraph.text or "") + block

        # Nach jedem Block einen Zeilenumbruch; zwischen zwei
        # Bloecken einen zweiten, sodass echte Absatzabstaende
        # entstehen. Dies entspricht der Struktur nativer
        # ILIAS-11-Exporte.
        br1 = ET.SubElement(paragraph, "br")
        if index < len(bloecke) - 1:
            ET.SubElement(paragraph, "br")

    xml_bytes = ET.tostring(page, encoding="utf-8")
    return base64.b64encode(xml_bytes).decode("ascii")


def gesamtpunkte_fragen(fragen):
    """Berechnet die Gesamtpunktzahl aller vorbereiteten Fragen."""
    gesamt = Decimal("0")

    for frage in fragen:
        if frage.get("typ") == "LUECKENTEXT":
            gesamt += Decimal(str(frage.get("gesamtpunkte", 0)))
        else:
            gesamt += Decimal(str(frage.get("punkte", 0)))

    return gesamt


def erzeuge_test_qti_54(fragen, testtitel, autor, zielordner, objekt_id):
    """Erzeugt eine Test-QTI und verwendet die vorhandenen Fragegeneratoren."""
    root = ET.Element("questestinterop")
    assessment = ET.SubElement(root, "assessment", {"ident": f"il_0_tst_{objekt_id}", "title": testtitel})
    ET.SubElement(assessment, "qticomment")
    qt_metadata = ET.SubElement(assessment, "qtimetadata")
    for label, value in test_metadaten_11(autor):
        metadata(qt_metadata, label, value)

    objectives = ET.SubElement(assessment, "objectives")
    material = ET.SubElement(objectives, "material")
    xml_text(
        material,
        "mattext",
        erzeuge_objectives_base64(PRUEFUNGSBEDINGUNGEN_TEXT),
        {"texttype": "text/xml"}
    )

    # ILIAS 11 exportiert an dieser Stelle einen leeren
    # assessmentcontrol-Block. Fuer die Kompatibilitaet
    # wird dieser ebenfalls erzeugt.
    ET.SubElement(assessment, "assessmentcontrol")

    frage_ids = []
    for nummer, frage in enumerate(fragen, start=1):
        frage_id = f"il_0_qst_{nummer}"
        frage_ids.append(frage_id)
        if frage["typ"] == "SC":
            erzeuge_sc_item(assessment, frage, frage_id)

        elif frage["typ"] == "MC":
            erzeuge_mc_item(assessment, frage, frage_id)

        elif frage["typ"] == "KPRIM":
            erzeuge_kprim_item(assessment, frage, frage_id)

        elif frage["typ"] == "VERT_ANORD":
            erzeuge_anordnung_vertikal_item(
                assessment,
                frage,
                frage_id
            )

        elif frage["typ"] == "HOR_ANORD":
            erzeuge_anordnung_horizontal_item(
                assessment,
                frage,
                frage_id
            )

        elif frage["typ"] == "FREITEXT":
            erzeuge_freitext_item(
                assessment,
                frage,
                frage_id
            )

        elif frage["typ"] == "LUECKENTEXT":
            erzeuge_lueckentext_item(
                assessment,
                frage,
                frage_id
            )

        else:
            raise NotImplementedError(
                f"XML fuer {frage['typ']} ist nicht implementiert."
            )

    datei = zielordner / "qti.xml"
    with open(datei, "w", encoding="utf-8") as f:
        f.write('<?xml version="1.0" encoding="utf-8"?>')
        f.write('<!--Generated by QuestionSheet2LMS-->')
        f.write(ET.tostring(root, encoding="unicode"))
    return datei, frage_ids



def erzeuge_testobjekt_xml_54(
    fragen,
    frage_ids,
    testtitel,
    autor,
    zielordner,
    objekt_id
):
    """
    Erzeugt die Teststruktur, Einleitungstexte,
    eingebettete Audio- und Videodateien und Reihenfolge
    der Fragen.
    """

    root = ET.Element(
        "ContentObject",
        {
            "Type": "Test"
        }
    )

    metadata_elem = ET.SubElement(
        root,
        "MetaData"
    )

    general = ET.SubElement(
        metadata_elem,
        "General",
        {
            "Structure": "Hierarchical"
        }
    )

    ET.SubElement(
        general,
        "Identifier",
        {
            "Catalog": "ILIAS",
            "Entry": f"il_0_tst_{objekt_id}"
        }
    )

    xml_text(
        general,
        "Title",
        testtitel,
        {
            "Language": "de"
        }
    )

    ET.SubElement(
        general,
        "Language",
        {
            "Language": "de"
        }
    )

    xml_text(
        general,
        "Description",
        "",
        {
            "Language": "de"
        }
    )

    ET.SubElement(
        general,
        "Keyword",
        {
            "Language": "de"
        }
    )

    lifecycle = ET.SubElement(
        metadata_elem,
        "Lifecycle",
        {
            "Status": "Draft"
        }
    )

    ET.SubElement(
        lifecycle,
        "Version",
        {
            "Language": "en"
        }
    )

    contribute = ET.SubElement(
        lifecycle,
        "Contribute",
        {
            "Role": "Author"
        }
    )

    xml_text(
        contribute,
        "Entity",
        autor
    )

    ET.SubElement(
        contribute,
        "Date"
    )

    #
    # Seitenstruktur der einzelnen Fragen.
    #
    for nummer, (frage, frage_id) in enumerate(
        zip(
            fragen,
            frage_ids
        ),
        start=1
    ):

        page_object = ET.SubElement(
            root,
            "PageObject"
        )

        #
        # Einleitungstext bei allen einzeiligen
        # Fragetypen.
        #
        # Bei Lueckentextfragen ist der
        # Einleitungstext bereits Bestandteil
        # des eigentlichen QTI-Items.
        #
        if (
            frage.get("einleitung")
            and frage.get("typ") != "LUECKENTEXT"
        ):

            page_content_einleitung = ET.SubElement(
                page_object,
                "PageContent"
            )

            einleitung = frage["einleitung"].replace(
                "&",
                "&amp;"
            )

            xml_text(
                page_content_einleitung,
                "Paragraph",
                einleitung,
                {
                    "Language": "de",
                    "Characteristic": "Standard"
                }
            )

        #
        # Optionales Audio.
        #
        # Entspricht der Struktur eines nativen
        # ILIAS-9-Testexports:
        #
        # PageContent
        #   MediaObject
        #     MediaAlias
        #     MediaAliasItem
        #
        if frage.get("audio_mob_id"):

            page_content_audio = ET.SubElement(
                page_object,
                "PageContent",
                {
                    "PCID": (
                        f"qs2lmsaudio"
                        f"{nummer:021d}"
                    )
                }
            )

            media_object = ET.SubElement(
                page_content_audio,
                "MediaObject"
            )

            ET.SubElement(
                media_object,
                "MediaAlias",
                {
                    "OriginId": frage[
                        "audio_mob_id"
                    ]
                }
            )

            media_alias_item = ET.SubElement(
                media_object,
                "MediaAliasItem",
                {
                    "Purpose": "Standard"
                }
            )

            ET.SubElement(
                media_alias_item,
                "Layout",
                {
                    "HorizontalAlign": "Left"
                }
            )

 

 
        if frage.get("video_mob_id"):

            page_content = ET.SubElement(
                page_object,
                "PageContent"
            )

            media_object = ET.SubElement(
                page_content,
                "MediaObject"
            )

            ET.SubElement(
                media_object,
                "MediaAlias",
                {
                    "OriginId": frage["video_mob_id"],
                    "Class": "MediaContainer"
                }
            )

            media_alias_item = ET.SubElement(
                media_object,
                "MediaAliasItem",
                {
                    "Purpose": "Standard"
                }
            )

            ET.SubElement(
                media_alias_item,
                "Layout",
                {
                    "HorizontalAlign": "Left"
                }
            )





        #
        # Verweis auf die eigentliche Frage.
        #
        page_content = ET.SubElement(
            page_object,
            "PageContent",
            {
                "PCID": (
                    f"qs2lms"
                    f"{nummer:026d}"
                )
            }
        )

        ET.SubElement(
            page_content,
            "Question",
            {
                "QRef": frage_id
            }
        )

    #
    # Vollstaendige Audio-MediaObjects.
    #
    # Im nativen ILIAS-9-Testexport stehen
    # diese nach den PageObjects und vor den
    # QuestionSkillAssignments.
    #
    for frage in fragen:

        if not frage.get("audio_mob_id"):
            continue

        media_object = ET.SubElement(
            root,
            "MediaObject"
        )

        metadata_element = ET.SubElement(
            media_object,
            "MetaData"
        )

        general = ET.SubElement(
            metadata_element,
            "General",
            {
                "Structure": "Hierarchical"
            }
        )

        ET.SubElement(
            general,
            "Identifier",
            {
                "Catalog": "ILIAS",
                "Entry": frage[
                    "audio_mob_id"
                ]
            }
        )

        xml_text(
            general,
            "Title",
            "QuestionSheet2LMS",
            {
                "Language": "de"
            }
        )

        ET.SubElement(
            general,
            "Language",
            {
                "Language": "de"
            }
        )

        xml_text(
            general,
            "Description",
            "",
            {
                "Language": "de"
            }
        )

        xml_text(
            general,
            "Keyword",
            "",
            {
                "Language": "de"
            }
        )

        media_item = ET.SubElement(
            media_object,
            "MediaItem",
            {
                "Purpose": "Standard"
            }
        )

        xml_text(
            media_item,
            "Location",
            frage["audio_dateiname"],
            {
                "Type": "LocalFile"
            }
        )

        xml_text(
            media_item,
            "Format",
            "audio/mpeg"
        )

        ET.SubElement(
            media_item,
            "Layout",
            {
                "HorizontalAlign": "Left"
            }
        )





    #
    # Vollstaendige Video-MediaObjects.
    #
    for frage in fragen:

        if not frage.get("video_mob_id"):
            continue

        media_object = ET.SubElement(
            root,
            "MediaObject"
        )

        metadata_element = ET.SubElement(
            media_object,
            "MetaData"
        )

        general = ET.SubElement(
            metadata_element,
            "General",
            {
                "Structure": "Hierarchical"
            }
        )

        ET.SubElement(
            general,
            "Identifier",
            {
                "Catalog": "ILIAS",
                "Entry": frage[
                    "video_mob_id"
                ]
            }
        )

        xml_text(
            general,
            "Title",
            "QuestionSheet2LMS",
            {
                "Language": "de"
            }
        )

        ET.SubElement(
            general,
            "Language",
            {
                "Language": "de"
            }
        )

        xml_text(
            general,
            "Description",
            "",
            {
                "Language": "de"
            }
        )

        xml_text(
            general,
            "Keyword",
            "",
            {
                "Language": "de"
            }
        )

        media_item = ET.SubElement(
            media_object,
            "MediaItem",
            {
                "Purpose": "Standard"
            }
        )

        xml_text(
            media_item,
            "Location",
            frage["video_dateiname"],
            {
                "Type": "LocalFile"
            }
        )

        xml_text(
            media_item,
            "Format",
            frage["video_mime"]
        )

        ET.SubElement(
            media_item,
            "Layout",
            {
                "HorizontalAlign": "Left"
            }
        )









    #
    # Bestehende Skill-Struktur unveraendert.
    #
    assignments = ET.SubElement(
        root,
        "QuestionSkillAssignments"
    )

    for nummer in range(
        1,
        len(frage_ids) + 1
    ):

        ET.SubElement(
            assignments,
            "TriggerQuestion",
            {
                "Id": str(nummer)
            }
        )

    ET.SubElement(
        root,
        "SkillsLevelThresholds"
    )

    datei = (
        zielordner
        / "test.xml"
    )

    with open(
        datei,
        "w",
        encoding="utf-8"
    ) as f:

        f.write(
            '<?xml version="1.0" encoding="utf-8"?>'
        )

        f.write(
            '<!DOCTYPE Test SYSTEM '
            '"http://www.ilias.uni-koeln.de/'
            'download/dtd/ilias_co.dtd">'
        )

        f.write(
            '<!--Generated by QuestionSheet2LMS-->'
        )

        f.write(
            ET.tostring(
                root,
                encoding="unicode"
            )
        )

    return datei



def erzeuge_test_exportdateien_54(
    testtitel,
    zielordner,
    objekt_id,
    hat_medien=False
):
    """
    Erzeugt Manifest und die ILIAS-Exportdateien.

    Bei enthaltenen Audio- und Videodateien werden zusaetzlich
    die von ILIAS 9 exportierten Komponenten

    Services/MediaObjects
    Modules/File

    erzeugt.
    """

    test_set = (
        zielordner
        / "Modules"
        / "Test"
        / "set_1"
    )

    object_set = (
        zielordner
        / "Services"
        / "Object"
        / "set_1"
    )

    test_set.mkdir(
        parents=True,
        exist_ok=True
    )

    object_set.mkdir(
        parents=True,
        exist_ok=True
    )

    if hat_medien:

        media_set = (
            zielordner
            / "Services"
            / "MediaObjects"
            / "set_1"
        )

        file_set = (
            zielordner
            / "Modules"
            / "File"
            / "set_1"
        )

        media_set.mkdir(
            parents=True,
            exist_ok=True
        )

        file_set.mkdir(
            parents=True,
            exist_ok=True
        )

    #
    # Manifest.
    #
    manifest = ET.Element(
        "Manifest",
        {
            "MainEntity": "tst",
            "Title": testtitel,
            "TargetRelease": "5.4.0",
            "InstallationId": "0",
            "InstallationUrl": (
                "http://127.0.0.1:8080"
            )
        }
    )

    if hat_medien:

        ET.SubElement(
            manifest,
            "ExportFile",
            {
                "Component": (
                    "Services/MediaObjects"
                ),
                "Path": (
                    "Services/MediaObjects/"
                    "set_1/export.xml"
                )
            }
        )

        ET.SubElement(
            manifest,
            "ExportFile",
            {
                "Component": "Modules/File",
                "Path": (
                    "Modules/File/"
                    "set_1/export.xml"
                )
            }
        )

    ET.SubElement(
        manifest,
        "ExportFile",
        {
            "Component": "Modules/Test",
            "Path": (
                "Modules/Test/"
                "set_1/export.xml"
            )
        }
    )

    ET.SubElement(
        manifest,
        "ExportFile",
        {
            "Component": "Services/Object",
            "Path": (
                "Services/Object/"
                "set_1/export.xml"
            )
        }
    )

    with open(
        zielordner / "manifest.xml",
        "w",
        encoding="utf-8"
    ) as f:

        f.write(
            '<?xml version="1.0" encoding="utf-8"?>'
        )

        f.write(
            '<!--Generated by QuestionSheet2LMS-->'
        )

        f.write(
            ET.tostring(
                manifest,
                encoding="unicode"
            )
        )

    #
    # Modules/Test
    #
    test_export = """<?xml version="1.0" encoding="utf-8"?>
<exp:Export InstallationId="0" InstallationUrl="http://127.0.0.1:8080" Entity="tst" SchemaVersion="4.1.0" TargetRelease="5.4.0" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:exp="http://www.ilias.de/Services/Export/exp/4_1" xsi:schemaLocation="http://www.ilias.de/Services/Export/exp/4_1 http://127.0.0.1:8080/xml/ilias_export_4_1.xsd http://www.ilias.de/Modules/Test/htlm/4_1 http://127.0.0.1:8080/xml/ilias_tst_4_1.xsd" xmlns="http://www.ilias.de/Modules/Test/htlm/4_1"><exp:ExportItem Id="{objekt_id}"></exp:ExportItem></exp:Export>"""

    #
    # Services/Object
    #
    object_export = """<?xml version="1.0" encoding="utf-8"?>
<exp:Export InstallationId="0" InstallationUrl="http://127.0.0.1:8080" Entity="common" SchemaVersion="5.4.0" TargetRelease="5.4.0" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:exp="http://www.ilias.de/Services/Export/exp/4_1" xsi:schemaLocation="http://www.ilias.de/Services/Export/exp/4_1 http://127.0.0.1:8080/xml/ilias_export_4_1.xsd http://www.ilias.de/Services/Object/obj/5_4 http://127.0.0.1:8080/xml/ilias_obj_5_4.xsd http://www.ilias.de/Services/DataSet/ds/4_3 http://127.0.0.1:8080/xml/ilias_ds_4_3.xsd" xmlns="http://www.ilias.de/Services/Object/obj/5_4" xmlns:ds="http://www.ilias.de/Services/DataSet/ds/4_3"><exp:ExportItem Id="{objekt_id}"><ds:DataSet InstallationId="0" InstallationUrl="http://127.0.0.1:8080" TopEntity="common"><ds:Rec Entity="common"><Common><ObjId>{objekt_id}</ObjId></Common></ds:Rec></ds:DataSet></exp:ExportItem></exp:Export>"""

    (
        test_set
        / "export.xml"
    ).write_text(
        test_export.format(
            objekt_id=objekt_id
        ),
        encoding="utf-8"
    )

    (
        object_set
        / "export.xml"
    ).write_text(
        object_export.format(
            objekt_id=objekt_id
        ),
        encoding="utf-8"
    )

    #
    # Die beiden folgenden Exportdateien sind
    # im nativen ILIAS-9-Testexport mit Medien
    # vorhanden, enthalten dort aber keine
    # einzelnen Datensaetze.
    #
    if hat_medien:

        media_export = """<?xml version="1.0" encoding="utf-8"?>
<exp:Export InstallationId="0" InstallationUrl="http://127.0.0.1:8080" Entity="mob" SchemaVersion="5.1.0" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:exp="http://www.ilias.de/Services/Export/exp/4_1" xsi:schemaLocation="http://www.ilias.de/Services/Export/exp/4_1 http://127.0.0.1:8080/xml/ilias_export_4_1.xsd https://www.ilias.de/Services/MediaObjects/mob/5_1 http://127.0.0.1:8080/xml/ilias_mob_5_1.xsd http://www.ilias.de/Services/DataSet/ds/4_3 http://127.0.0.1:8080/xml/ilias_ds_4_3.xsd" xmlns="https://www.ilias.de/Services/MediaObjects/mob/5_1" xmlns:ds="http://www.ilias.de/Services/DataSet/ds/4_3"></exp:Export>"""

        file_export = """<?xml version="1.0" encoding="utf-8"?>
<exp:Export InstallationId="0" InstallationUrl="http://127.0.0.1:8080" Entity="file" SchemaVersion="4.1.0" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:exp="http://www.ilias.de/Services/Export/exp/4_1" xsi:schemaLocation="http://www.ilias.de/Services/Export/exp/4_1 http://127.0.0.1:8080/xml/ilias_export_4_1.xsd http://www.ilias.de/Modules/File/file/4_1 http://127.0.0.1:8080/xml/ilias_file_4_1.xsd" xmlns="http://www.ilias.de/Modules/File/file/4_1"></exp:Export>"""

        (
            media_set
            / "export.xml"
        ).write_text(
            media_export,
            encoding="utf-8"
        )

        (
            file_set
            / "export.xml"
        ).write_text(
            file_export,
            encoding="utf-8"
        )




def erzeuge_test_zip_54(
    fragen,
    ausgabeverzeichnis,
    testtitel
):
    """
    Erzeugt einen kompletten ILIAS-Testexport.

    Bilder, MP3-Audiodateien sowie MP4- und WebM-Videodateien
    werden als lokale Media Objects eingebettet.

    """

    export_id = str(
        int(
            datetime.now().timestamp()
        )
    )

    objekt_id = 313

    ordnername = (
        f"{export_id}"
        f"__0__tst_{objekt_id}"
    )

    ausgabeverzeichnis = Path(
        ausgabeverzeichnis
    )

    zielordner = (
        ausgabeverzeichnis
        / ordnername
    )

    zip_datei = (
        ausgabeverzeichnis
        / f"{ordnername}.zip"
    )

    zielordner.mkdir(
        parents=True,
        exist_ok=True
    )

    autor = next(
        (
            textwert(
                frage.get(
                    "autor",
                    ""
                )
            )
            for frage in fragen
            if textwert(
                frage.get(
                    "autor",
                    ""
                )
            )
        ),
        "QuestionSheet2LMS"
    )







    hat_audio = any(
        textwert(
            frage.get(
                "audio",
                ""
            )
        )
        for frage in fragen
    )

    hat_video = any(
        textwert(
            frage.get(
                "video",
                ""
            )
        )
        for frage in fragen
    )

    hat_medien = (
        hat_audio
        or hat_video
    )








    try:

        #
        # Bilder wie bisher vorbereiten.
        #
        bereite_bilder_vor(
            fragen,
            ausgabeverzeichnis,
            zielordner,
            export_id
        )

        #
        # Audiodateien analog zum Fragenpool
        # vorbereiten.
        #
        bereite_audios_vor(
            fragen,
            ausgabeverzeichnis,
            zielordner,
            export_id
        )


        #
        # Videodateien analog zum Fragenpool
        # vorbereiten.
        #

       
        bereite_videos_vor(
            fragen,
            ausgabeverzeichnis,
            zielordner,
            export_id
        )


        #
        # QTI mit den eigentlichen Fragen.
        #
        qti_datei, frage_ids = (
            erzeuge_test_qti_54(
                fragen,
                testtitel,
                autor,
                zielordner,
                objekt_id
            )
        )

        #
        # Testobjekt mit Einleitungen,
        # MediaAliases und MediaObjects.
        #
        test_datei = (
            erzeuge_testobjekt_xml_54(
                fragen,
                frage_ids,
                testtitel,
                autor,
                zielordner,
                objekt_id
            )
        )

        qti_ziel = (
            zielordner
            / (
                f"{export_id}"
                f"__0__qti_{objekt_id}.xml"
            )
        )

        test_ziel = (
            zielordner
            / (
                f"{export_id}"
                f"__0__tst_{objekt_id}.xml"
            )
        )

        qti_datei.rename(
            qti_ziel
        )

        test_datei.rename(
            test_ziel
        )

        erzeuge_test_exportdateien_54(
            testtitel,
            zielordner,
            objekt_id,
            hat_medien=hat_medien
        )

        #
        # ZIP-Struktur.
        #
        with ZipFile(
            zip_datei,
            "w",
            ZIP_DEFLATED
        ) as z:

            verzeichnisse = [
                Path(ordnername),
                Path(ordnername)
                / "objects",
                Path(ordnername)
                / "Modules",
                Path(ordnername)
                / "Modules"
                / "Test",
                Path(ordnername)
                / "Modules"
                / "Test"
                / "set_1",
                Path(ordnername)
                / "Services",
                Path(ordnername)
                / "Services"
                / "Object",
                Path(ordnername)
                / "Services"
                / "Object"
                / "set_1",
            ]

            if hat_medien:

                verzeichnisse.extend(
                    [
                        Path(ordnername)
                        / "Services"
                        / "MediaObjects",

                        Path(ordnername)
                        / "Services"
                        / "MediaObjects"
                        / "set_1",

                        Path(ordnername)
                        / "Modules"
                        / "File",

                        Path(ordnername)
                        / "Modules"
                        / "File"
                        / "set_1",
                    ]
                )

            for verzeichnis in verzeichnisse:

                z.writestr(
                    (
                        str(verzeichnis)
                        .replace("\\", "/")
                        + "/"
                    ),
                    ""
                )

            for datei in zielordner.rglob(
                "*"
            ):

                if datei.is_file():

                    arcname = (
                        Path(ordnername)
                        / datei.relative_to(
                            zielordner
                        )
                    )

                    z.write(
                        datei,
                        str(arcname).replace(
                            "\\",
                            "/"
                        )
                    )

        return zip_datei

    finally:

        loesche_exportarbeitsordner(
            zielordner,
            ausgabeverzeichnis
        )




def main():
    """
    Programmeinstieg.
    """

    if len(sys.argv) != 2:

        print(
            "Aufruf:"
        )

        print(
            "python3 QuestionSheet2LMS.py Aufgaben.xlsx"
        )

        print(
            "oder"
        )

        print(
            "python3 QuestionSheet2LMS.py Aufgaben.csv"
        )

        print(
            "python3 QuestionSheet2LMS.py Aufgaben.ods"
        )


        sys.exit(1)

    datei = Path(
        sys.argv[1]
    )

    if not datei.exists():

        print(
            f"Datei nicht gefunden: {datei}"
        )

        sys.exit(1)

    try:
        zeilen = lese_datei(
            datei
        )



        print(
            f"{len(zeilen)} Tabellenzeilen eingelesen."
        )


        fragen = bereite_alle_fragen_vor(
            zeilen
        )



        for zeile in zeilen:

            if (
                textwert(
                    zeile.get(
                        "Fragetyp",
                        ""
                    )
                ).upper() == "KPRIM"
                and not textwert(
                    zeile.get(
                        "Richtige Antwort",
                        ""
                    )
                )
            ):

                print(
                    "Hinweis: "
                    f"Bei der KPrim-Frage mit Frage_ID "
                    f"{textwert(zeile.get('Frage_ID', ''))} "
                    "sind alle vier Aussagen als falsch gekennzeichnet."
                )


















        print(
            f"Daraus wurden {len(fragen)} Fragen erkannt."
        )

        print()

        for nummer, frage in enumerate(
            fragen,
            start=1
        ):

            print(
                f"{nummer:03d}: "
                f"{frage['typ']} - "
                f"{frage['titel']}"
            )


        print(
            "Alle Fragen erfolgreich vorbereitet."
        )

        gesamtpunkte = gesamtpunkte_fragen(
            fragen
        )

        print(
            f"Gesamtpunktzahl: {format_punkte(gesamtpunkte)} Punkte."
        )

        zip_datei = erzeuge_zip(
            fragen,
            datei.parent
        )

        print(
            f"Fragenpool-ZIP erstellt: {zip_datei}"
        )

        test_zip = erzeuge_test_zip_54(
            fragen,
            datei.parent,
            datei.stem
        )

        print(
            f"Test-ZIP erstellt: {test_zip}"
        )

    except Exception as e:

        print()

        print(
            f"Abbruch: {e}"
        )

        sys.exit(1)


if __name__ == "__main__":
    main()


# -----------------------------------------------------------------------------
# Oeffentliche API fuer die grafische Benutzeroberflaeche
# -----------------------------------------------------------------------------

def bereite_export_vor(datei, logger=None):
    """Liest und validiert eine Quelldatei fuer die GUI.

    Rueckgabe:
        (tabellenzeilen, fragen, gesamtpunkte)
    """
    datei = Path(datei)
    zeilen = lese_datei(datei)
    fragen = bereite_alle_fragen_vor(zeilen)
    gesamtpunkte = gesamtpunkte_fragen(fragen)

    if logger is not None:
        logger(f"{len(zeilen)} Tabellenzeilen eingelesen.")
        logger(f"Daraus wurden {len(fragen)} Fragen erkannt.")
        logger(f"Gesamtpunktzahl: {format_punkte(gesamtpunkte)} Punkte.")

    return zeilen, fragen, gesamtpunkte


def exportiere_datei(
    datei,
    export_pool=True,
    export_test=False,
    logger=None
):
    """Erzeugt wahlweise Fragenpool, Test oder beide Exportdateien.

    Diese Funktion bildet die stabile Schnittstelle fuer die GUI.
    """
    datei = Path(datei)
    zeilen, fragen, gesamtpunkte = bereite_export_vor(
        datei,
        logger=logger
    )

    dateien = []

    if export_pool:
        pool_zip = erzeuge_zip(
            fragen,
            datei.parent,
            datei.stem
        )
        dateien.append(pool_zip)
        if logger is not None:
            logger(f"Fragenpool-ZIP erstellt: {pool_zip.name}")

    if export_test:
        test_zip = erzeuge_test_zip_54(
            fragen,
            datei.parent,
            datei.stem
        )
        dateien.append(test_zip)
        if logger is not None:
            logger(f"Test-ZIP erstellt: {test_zip.name}")

    return {
        "tabellenzeilen": len(zeilen),
        "fragen": len(fragen),
        "gesamtpunkte": gesamtpunkte,
        "dateien": dateien,
    }
