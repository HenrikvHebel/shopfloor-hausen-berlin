"""Create the empty shopfloor workbook (Excel tables + dropdowns) from schema.json.

Usage:  python tools/make_workbook.py [output.xlsx]
Default output: Shopfloor_Hausen_Berlin.xlsx in the current directory.
"""
import json
import os
import sys

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.table import Table, TableStyleInfo

HERE = os.path.dirname(os.path.abspath(__file__))
SCHEMA_PATH = os.path.join(HERE, "..", "schema.json")
MAX_ROWS = 5000  # rows pre-formatted / covered by dropdowns

# column key -> list name in schema["lists"]
DROPDOWNS = {
    ("shipments", "richtung"): "richtung",
    ("shipments", "qs_ergebnis"): "qs_ergebnis",
    ("shipments", "express"): "jn",
    ("orders", "express"): "jn",
    ("orders", "abgeschlossen"): "jn",
    ("problems", "kategorie"): "kategorie",
    ("problems", "standort"): "standort",
    ("problems", "dringlichkeit"): "dringlichkeit",
    ("problems", "status"): "status",
}

# lines starting with "## " are written bold (without the marker)
HELP = [
    "Shopfloor Hausen ↔ Berlin – Datenablage der Web-App",
    "",
    "## So läuft das Shopfloor-Meeting (ca. 10–15 Minuten, Mo–Fr bzw. jeden 2. Tag)",
    "",
    "## Vor dem Meeting – jeder Standort pflegt seine Daten selbst, am besten sofort:",
    "• Hausen: neue OrgaMax-Bestellungen unter „Bestellungen“ erfassen (BestellNr, Positionen mit S-Nr, Wunschtermin).",
    "• Berlin: jedes versendete Paket unter „Versand“ melden – auch Express. Die angezeigte LieferID und die BestellNr auf Lieferschein und Paket schreiben.",
    "• Hausen: jedes eingegangene Paket unter „Wareneingang“ erfassen: gezählte Menge, QS-Ergebnis, Menge i.O., QS geprüft von.",
    "",
    "## Im Meeting – ein Moderator, Bildschirm teilen:",
    "• Der Moderator öffnet die App (Reiter „Meeting“) und teilt den Bildschirm in Teams. Berlin schaut mit und spricht.",
    "• Nur der Moderator klickt. Die Bildschirme der anderen laufen nicht mit (Schritt und Timer sind je Browser getrennt).",
    "• Pro Schritt ca. 2 Minuten. Der Schritt-Timer wird rot, wenn die Zeit überschritten ist.",
    "• Was länger als 2 Minuten Diskussion braucht, wird als Problemkarte („+ Problem“) mit Verantwortlich und Termin erfasst – nicht im Meeting gelöst.",
    "",
    "## Die 6 Schritte:",
    "1. Kennzahlen – rote Kacheln kurz ansprechen (Liefertreue, Mengentreue, QS-Quote, Bestellbezug, Eingänge ohne QS, überfällige Positionen, offene Probleme).",
    "2. Offene Bestellungen – überfällige stehen oben. Farben: hellrot = in 3 Tagen fällig, kräftig rot = überfällig, gelb = teilgeliefert. Berlin nennt einen Termin → „Termin/Bearbeiten“ → „Bestätigter Termin“ eintragen. Lager Hausen prüfen (gelb = unter Mindestbestand, rot = 0).",
    "3. Unterwegs – was ist versendet, aber noch nicht angekommen? Express beachten; lange Laufzeiten klären.",
    "4. Eingänge & QS – Abweichungen besprechen: ohne Bestellbezug, Menge anders als versendet, QS offen oder n.i.O. Bei „QS offen“: wer prüft bis wann?",
    "5. Probleme – offene Karten nach Dringlichkeit durchgehen, Status setzen („in Arbeit“ / „erledigt“), neue Karten anlegen.",
    "6. Abschluss – Teilnehmer beider Standorte eintragen, Vereinbarungen notieren und „Meeting speichern“ (nur einmal, nur der Moderator).",
    "",
    "Der Kennzahlen-Zeitraum beginnt beim zuletzt gespeicherten Meeting – deshalb jedes Meeting speichern.",
    "",
    "## Hinweise zur Datei",
    "Die Web-App liest und schreibt die Tabellen auf den Blättern Bestellungen, Lieferungen, Probleme und Meetings.",
    "Tabellen- und Spaltennamen nicht umbenennen. Zusätzliche Spalten rechts sind erlaubt (die App lässt sie unverändert).",
    "Eine leere Zeile direkt unter der Überschrift ist normal (Excel-Tabellen brauchen mindestens eine Datenzeile).",
    "",
    "Bestellungen: eine Zeile je Bestellposition aus OrgaMax (BestellNr + Pos). LB-Nr = Laufbandnummer(n), Format 1234/26, mehrere mit Komma: 1234/26, 1235/26.",
    "Lieferungen: eine Zeile je Paket. Berlin füllt Versand, Hausen füllt Eingang und QS.",
    "Probleme: Problemkarten mit Verantwortlich, Maßnahme, Termin, Status.",
    "Meetings: Protokoll jedes Shopfloor-Meetings mit Kennzahlen.",
]


def main():
    out = sys.argv[1] if len(sys.argv) > 1 else "Shopfloor_Hausen_Berlin.xlsx"
    with open(SCHEMA_PATH, encoding="utf-8") as f:
        schema = json.load(f)
    lists = schema["lists"]

    wb = Workbook()
    info = wb.active
    info.title = "Info"
    for i, line in enumerate(HELP, start=1):
        bold = line.startswith("## ")
        c = info.cell(row=i, column=1, value=line[3:] if bold else line)
        c.alignment = Alignment(wrap_text=True, vertical="top")
        if bold:
            c.font = Font(bold=True)
    info["A1"].font = Font(bold=True, size=14)
    info.column_dimensions["A"].width = 120

    # hidden sheet with dropdown values
    ls = wb.create_sheet("Listen")
    list_ref = {}
    for ci, (name, values) in enumerate(lists.items(), start=1):
        col = get_column_letter(ci)
        ls.cell(row=1, column=ci, value=name).font = Font(bold=True)
        for ri, v in enumerate(values, start=2):
            ls.cell(row=ri, column=ci, value=v)
        list_ref[name] = "Listen!$%s$2:$%s$%d" % (col, col, len(values) + 1)
    ls.sheet_state = "hidden"

    head_fill = PatternFill("solid", fgColor="1B1D21")
    for key in ("orders", "shipments", "problems", "meetings"):
        t = schema[key]
        ws = wb.create_sheet(t["sheet"])
        cols = t["columns"]
        for ci, col in enumerate(cols, start=1):
            ckey, header, ctype = col[:3]
            c = ws.cell(row=1, column=ci, value=header)
            c.font = Font(bold=True, color="FFFFFF")
            c.fill = head_fill
            c.alignment = Alignment(wrap_text=True, vertical="top")
            letter = get_column_letter(ci)
            ws.column_dimensions[letter].width = max(12, min(40, len(header) + 4))
            fmt = {"date": "DD.MM.YYYY", "text": "@"}.get(ctype)
            if fmt:
                for r in range(2, MAX_ROWS + 2):
                    ws.cell(row=r, column=ci).number_format = fmt
            lname = DROPDOWNS.get((key, ckey))
            if lname:
                dv = DataValidation(type="list", formula1="=" + list_ref[lname], allow_blank=True)
                dv.add("%s2:%s%d" % (letter, letter, MAX_ROWS + 1))
                ws.add_data_validation(dv)
        last = get_column_letter(len(cols))
        tab = Table(displayName=t["table"], ref="A1:%s2" % last)
        tab.tableStyleInfo = TableStyleInfo(name="TableStyleLight9", showRowStripes=True)
        ws.add_table(tab)
        ws.freeze_panes = "A2"

    wb.save(out)
    print("geschrieben:", os.path.abspath(out))


if __name__ == "__main__":
    main()
