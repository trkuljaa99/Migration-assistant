"""Demo files in samples/: a Customer template in the Migration Cockpit's XML Spreadsheet 2003 layout (made up for
the demo - in a real project use the template downloaded from the system) and a messy legacy customer extract."""
import random
from pathlib import Path

from openpyxl import Workbook

OUT = Path(__file__).parent / "samples"

BASIC = [("Key", "*Customer Number", "KUNNR", "Text(10)"), ("Key", "*Business Partner Grouping", "BU_GROUP", "Text(4)"),
         ("General Data", "*Name 1", "NAME1", "Text(40)"), ("General Data", "Name 2", "NAME2", "Text(40)"),
         ("General Data", "Search Term 1", "SORT1", "Text(20)"), ("Address", "Street", "STREET", "Text(60)"),
         ("Address", "House Number", "HOUSE_NUM1", "Text(10)"), ("Address", "*Postal Code", "POST_CODE1", "Text(10)"),
         ("Address", "*City", "CITY1", "Text(40)"), ("Address", "*Country/Region", "COUNTRY", "Text(3)"),
         ("Address", "Region", "REGION", "Text(3)"), ("Address", "*Language Key", "LANGU", "Text(2)"),
         ("Communication", "Telephone", "TEL_NUMBER", "Text(30)"), ("Communication", "E-Mail Address", "SMTP_ADDR", "Text(241)"),
         ("Tax", "VAT Registration Number", "STCEG", "Text(20)")]
SALES = [("Key", "*Customer Number", "KUNNR", "Text(10)"), ("Sales Area", "*Sales Organization", "VKORG", "Text(4)"),
         ("Sales Area", "*Distribution Channel", "VTWEG", "Text(2)"), ("Sales Area", "*Division", "SPART", "Text(2)"),
         ("Sales", "*Currency", "WAERS", "Text(5)"), ("Sales", "Terms of Payment", "ZTERM", "Text(4)"),
         ("Sales", "Incoterms", "INCO1", "Text(3)"), ("Sales", "Customer Group", "KDGRP", "Text(2)")]


def _row(cells, style=None):
    s = f' ss:StyleID="{style}"' if style else ""
    return "<Row>" + "".join(f'<Cell{s}><Data ss:Type="String">{c}</Data></Cell>' for c in cells) + "</Row>"


def _sheet(name, title, fields):
    rows = [_row([title], "t"), "<Row/>", _row(["Enter the data from row 8 on. Fields marked * are mandatory."]),
            _row([g for g, *_ in fields], "g"), _row([l for _, l, *_ in fields], "h"),
            _row([t for *_, t, _ in fields], "h"), _row([ty for *_, ty in fields])]
    return (f'<Worksheet ss:Name="{name}"><Table ss:ExpandedColumnCount="{len(fields)}" ss:ExpandedRowCount="7">'
            + "".join(rows) + "</Table></Worksheet>")


def template():
    field_list = [("Sheet Name", "Group Name", "Field Description", "Importance", "Type", "Length")]
    for sheet, fields in (("Basic Data", BASIC), ("Sales Data", SALES)):
        for g, l, t, ty in fields:
            field_list.append((sheet, g, l.strip("*"), "Mandatory" if l.startswith("*") else "Optional",
                               ty.split("(")[0], ty.split("(")[1].strip(")")))
    xml = ('<?xml version="1.0" encoding="UTF-8"?>\n<?mso-application progid="Excel.Sheet"?>\n'
           '<Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet" '
           'xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel" '
           'xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet" xmlns:html="http://www.w3.org/TR/REC-html40">'
           '<Styles><Style ss:ID="t"><Font ss:Bold="1" ss:Size="14"/></Style><Style ss:ID="g"><Font ss:Bold="1"/>'
           '</Style><Style ss:ID="h"><Font ss:Bold="1"/><Interior ss:Color="#DCE6F1" ss:Pattern="Solid"/></Style>'
           '</Styles>'
           '<Worksheet ss:Name="Introduction"><Table>' + _row(["Migration Template - Customer (demo)"], "t")
           + _row(["Fill the data sheets, then upload the file in Migrate Your Data."]) + "</Table></Worksheet>"
           '<Worksheet ss:Name="Field List"><Table>' + "".join(_row(r) for r in field_list) + "</Table></Worksheet>"
           + _sheet("Basic Data", "Customer - Basic Data", BASIC) + _sheet("Sales Data", "Customer - Sales Data", SALES)
           + "</Workbook>")
    (OUT / "Customer_migration_template_DEMO.xml").write_text(xml, encoding="utf-8")


def source():
    random.seed(7)
    wb = Workbook()
    ws = wb.active
    ws.title = "Kundenstamm"
    ws.append(["Export Altsystem ERP – Kunden, Stand 01.11.2026"])
    ws.append([])
    ws.append(["Kunden-Nr", "Firma", "Zusatz", "Adresse", "PLZ", "Ort", "Land", "Sprache", "Telefon", "E-Mail",
               "USt-IdNr", "Währung", "Zahlungsbedingung", "Angelegt am", "Vertriebsweg", "Bemerkung"])
    cities = [("80331", "München", "Deutschland"), ("10115", "Berlin", "Germany"), ("1010", "Wien", "Österreich"),
              ("8001", "Zürich", "Schweiz"), ("20095", "Hamburg", "DE"), ("50667", "Köln", "Deutschland"),
              ("4020", "Linz", "AT"), ("60311", "Frankfurt am Main", "Deutschland")]
    names = ["Müller Maschinenbau GmbH", "Schmidt & Söhne KG", "Alpenholz AG", "Brauerei Weber", "Nordlicht Logistik GmbH",
             "Keller Elektro e.K.", "Donau Pharma GmbH", "Fischer Feinkost", "Huber Bau AG", "Wagner IT Services GmbH",
             "Becker Textil", "Zimmermann Metall GmbH", "Hofmann Medizintechnik", "Schulz Druck GmbH",
             "Lange Agrar eG", "Krause Möbelwerk", "Richter Optik", "Wolf Energie AG", "Neumann Spedition",
             "Schwarz Chemie GmbH", "Braun Automotive", "Zürcher Uhren AG"]
    terms = ["30 Tage netto", "14 Tage 2% Skonto, 30 Tage netto", "sofort ohne Abzug", "60 Tage netto"]
    rows = []
    for i, name in enumerate(names):
        plz, city, land = cities[i % len(cities)]
        currency = "CHF" if land == "Schweiz" else random.choice(["EUR", "Euro", "EUR"])
        rows.append([f"{10000 + i}", name, random.choice(["", "Niederlassung Süd", ""]),
                     f"{random.choice(['Hauptstraße', 'Industriestr.', 'Am Markt', 'Bahnhofweg'])} {random.randint(1, 120)}",
                     plz, city, land, random.choice(["Deutsch", "Deutsch", "Englisch"]),
                     f"+49 {random.randint(30, 99)} {random.randint(100000, 999999)}",
                     f"info@{name.split()[0].lower().replace('ü', 'ue').replace('ö', 'oe')}.example",
                     f"DE{random.randint(100000000, 999999999)}" if land in ("Deutschland", "Germany", "DE") else "",
                     currency, random.choice(terms),
                     f"{random.randint(1, 28):02d}.{random.randint(1, 12):02d}.20{random.randint(10, 25)}",
                     random.choice(["Direktvertrieb", "Großhandel"]), ""])
    rows[4][5] = ""  # a missing city
    rows[9][4] = ""  # a missing postal code
    rows[12][1] = "  hofmann medizintechnik  "
    rows += [list(rows[2]), list(rows[7]), [*rows[15][:15], "Dublette aus Altsystem"]]  # duplicates
    for r in rows:
        ws.append(r)
    ws.append([])
    wb.save(OUT / "Legacy_customers_DEMO.xlsx")


if __name__ == "__main__":
    OUT.mkdir(exist_ok=True)
    template()
    source()
    print("samples written")
