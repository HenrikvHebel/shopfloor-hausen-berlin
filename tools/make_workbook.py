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

HELP = [
    "Shopfloor Hausen ↔ Berlin – Datenablage der Web-App",
    "",
    "Die Web-App liest und schreibt die Tabellen auf den Blättern Bestellungen, Lieferungen, Probleme und Meetings.",
    "Tabellen- und Spaltennamen nicht umbenennen. Zusätzliche Spalten rechts sind erlaubt (die App lässt sie unverändert).",
    "Eine leere Zeile direkt unter der Überschrift ist normal (Excel-Tabellen brauchen mindestens eine Datenzeile).",
    "",
    "Bestellungen: eine Zeile je Bestellposition aus OrgaMax (BestellNr + Pos).",
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
        info.cell(row=i, column=1, value=line)
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
        for ci, (ckey, header, ctype) in enumerate(cols, start=1):
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
