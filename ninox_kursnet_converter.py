#!/usr/bin/env python3
"""
Ninox-KursNet Converter
=======================
Konvertiert einen Ninox-Export (KursNet.xlsx) in eine KursNet-kompatible
open-Q-cat XML-Datei (Format V1.1).

WICHTIG – 1:n-Prinzip (§ 3 Abs. 5 KURSNET-Nutzungsbedingungen):
  Einem Angebot können mehrere Standorte/Termine zugeordnet werden.
  Zeilen mit identischer PRODUCT ID werden zu EINEM SERVICE zusammengefasst.
  Jede Zeile liefert ein eigenes MODULE_COURSE-Element (Ort + Termin).

Verwendung:
    python3 ninox_kursnet_converter.py KursNet.xlsx output.xml [Optionen]

Optionen:
    --catalog-id        Katalog-ID (Standard: Zeitstempel)
    --catalog-name      Katalogname
    --catalog-version   Katalogversion (Standard: 1.0)
    --language          Sprachcode ISO 639-2 (Standard: deu)
    --currency          Währungscode ISO 4217 (Standard: EUR)

Abhängigkeiten:
    pip install openpyxl
"""

import sys
import argparse
from datetime import datetime
from collections import OrderedDict
import xml.dom.minidom
from xml.etree.ElementTree import Element, SubElement, tostring

try:
    import openpyxl
except ImportError:
    sys.exit("Fehler: 'openpyxl' nicht installiert. Bitte 'pip install openpyxl' ausführen.")


# ---------------------------------------------------------------------------
# Länderkürzel-Mapping (Klarname → ISO-Kürzel laut open-Q-cat-Schema)
# ---------------------------------------------------------------------------

COUNTRY_CODES: dict[str, str] = {
    "Deutschland":          "D",
    "Germany":              "D",
    "Österreich":           "A",
    "Austria":              "A",
    "Schweiz":              "CH",
    "Switzerland":          "CH",
    "Frankreich":           "F",
    "France":               "F",
    "Niederlande":          "NL",
    "Netherlands":          "NL",
    "Belgien":              "B",
    "Belgium":              "B",
    "Polen":                "PL",
    "Poland":               "PL",
    "Tschechien":           "CZ",
    "Czech Republic":       "CZ",
    "Ungarn":               "H",
    "Hungary":              "H",
    "Italien":              "I",
    "Italy":                "I",
    "Spanien":              "E",
    "Spain":                "E",
    "Luxemburg":            "L",
    "Luxembourg":           "L",
    "Dänemark":             "DK",
    "Denmark":              "DK",
    "Schweden":             "S",
    "Sweden":               "S",
    "Norwegen":             "N",
    "Norway":               "N",
    "Finnland":             "FIN",
    "Finland":              "FIN",
    "USA":                  "USA",
    "Vereinigte Staaten":   "USA",
    "Großbritannien":       "GB",
    "United Kingdom":       "GB",
    "Russland":             "RUS",
    "Russia":               "RUS",
    "China":                "CHN",
    "Japan":                "J",
    "Türkei":               "TR",
    "Turkey":               "TR",
}


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------

def parse_args():
    parser = argparse.ArgumentParser(
        description="Ninox KursNet-Export (xlsx) → open-Q-cat XML (V1.1)",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=__doc__,
    )
    parser.add_argument("input",  help="Eingabe-Excel-Datei (z. B. KursNet.xlsx)")
    parser.add_argument("output", help="Ausgabe-XML-Datei (z. B. kursnet_upload.xml)")
    parser.add_argument("--catalog-id",      default="",    help="Katalog-ID")
    parser.add_argument("--catalog-name",    default="",    help="Katalogname")
    parser.add_argument("--catalog-version", default="1.0", help="Katalogversion (Standard: 1.0)")
    parser.add_argument("--language",        default="deu", help="Sprachcode ISO 639-2 (Standard: deu)")
    parser.add_argument("--currency",        default="EUR", help="Währungscode (Standard: EUR)")
    return parser.parse_args()


# ---------------------------------------------------------------------------
# Excel lesen
# ---------------------------------------------------------------------------

def read_excel(filepath: str) -> list[dict]:
    """Liest die Excel-Datei und gibt eine Liste von Zeilen-Dicts zurück."""
    wb = openpyxl.load_workbook(filepath, read_only=True, data_only=True)
    ws = wb.active
    rows = list(ws.iter_rows(values_only=True))
    if not rows:
        return []
    headers = rows[0]
    result = []
    for row in rows[1:]:
        if any(v is not None for v in row):
            result.append(dict(zip(headers, row)))
    wb.close()
    return result


def group_by_product(rows: list[dict]) -> tuple[list, dict]:
    """
    Gruppiert Zeilen nach PRODUCT ID (1:n-Prinzip).
    Gibt (geordnete Liste der PIDs, Dict PID→Zeilen) zurück.
    """
    groups: dict = OrderedDict()
    for row in rows:
        pid = row.get("PRODUCT ID")
        if pid is not None:
            groups.setdefault(pid, []).append(row)
    return list(groups.keys()), groups


# ---------------------------------------------------------------------------
# Hilfsfunktionen
# ---------------------------------------------------------------------------

def add_text(parent: Element, tag: str, text, **attrib) -> Element | None:
    """Fügt ein Kind-Element mit Text ein; überspringt leere/None-Werte."""
    if text is None:
        return None
    s = str(text).strip()
    if not s:
        return None
    el = SubElement(parent, tag, **attrib)
    el.text = s
    return el


def fmt_datetime(value) -> str | None:
    """
    Formatiert einen Datumswert als dateTime (YYYY-MM-DDTHH:MM:SS).
    Das Schema erwartet xs:dateTime, nicht xs:date.
    """
    if value is None:
        return None
    if isinstance(value, datetime):
        return value.strftime("%Y-%m-%dT00:00:00")
    s = str(value).strip()
    if not s:
        return None
    # Reines Datum (YYYY-MM-DD) um Zeitanteil ergänzen
    if len(s) == 10 and "T" not in s:
        return s + "T00:00:00"
    return s


def to_bool(value) -> str | None:
    """Wandelt ja/nein in true/false um."""
    if value is None:
        return None
    v = str(value).strip().lower()
    if v in ("ja", "yes", "true", "1"):
        return "true"
    if v in ("nein", "no", "false", "0"):
        return "false"
    return None


def int_str(v) -> str | None:
    """Gibt einen Wert als ganzzahligen String zurück oder None."""
    if v is None:
        return None
    return str(int(v)) if isinstance(v, float) else str(v)


def country_code(name) -> str | None:
    """Konvertiert einen Ländernamen in das ISO-Kürzel des Schemas."""
    if not name:
        return None
    s = str(name).strip()
    return COUNTRY_CODES.get(s, s)  # Fallback: Originalwert


def truncate(value, max_len: int) -> str | None:
    """Kürzt einen String auf max_len Zeichen."""
    if value is None:
        return None
    s = str(value).strip()
    return s[:max_len] if s else None


# ---------------------------------------------------------------------------
# XML-Bausteine
# ---------------------------------------------------------------------------

def build_location(parent: Element, row: dict) -> Element:
    """Erstellt ein LOCATION-Element aus den Ortsdaten einer Zeile."""
    loc = SubElement(parent, "LOCATION")
    add_text(loc, "NAME",    row.get("ORT_NAME"))
    # STREET: Schema-Maximum 30 Zeichen
    add_text(loc, "STREET",  truncate(row.get("ORT_STRASSE"), 30))
    plz = row.get("ORT_PLZ")
    add_text(loc, "ZIP",     str(int(plz)) if isinstance(plz, float) else str(plz) if plz else None)
    add_text(loc, "CITY",    row.get("ORT_STADT"))
    add_text(loc, "COUNTRY", country_code(row.get("ORT_LAND")))
    return loc


def build_module_course(education: Element, row: dict, order: int) -> Element:
    """
    Erstellt ein MODULE_COURSE-Element für einen einzelnen Standort/Termin.
    Jede Excel-Zeile eines Produkts erzeugt genau ein MODULE_COURSE.

    Korrekte Elementreihenfolge laut Schema (kein type-Attribut, kein
    DURATION-Wrapper – START_DATE/END_DATE direkt in MODULE_COURSE).
    """
    mc = SubElement(education, "MODULE_COURSE")

    # Datumsangaben direkt (kein DURATION-Wrapper, Schema: dateTime)
    add_text(mc, "START_DATE", fmt_datetime(row.get("START_DATUM")))
    add_text(mc, "END_DATE",   fmt_datetime(row.get("END_DATUM")))

    flex = to_bool(row.get("FLEXIBLE_ANMELDUNG"))
    add_text(mc, "FLEXIBLE_START", flex)

    # Ort
    build_location(mc, row)

    # Teilnehmerzahl
    max_p = row.get("MAX_TEILNEHMER")
    min_p = row.get("MIN_TEILNEHMER")
    add_text(mc, "MAX_PARTICIPANTS", str(int(max_p)) if max_p is not None else None)
    add_text(mc, "MIN_PARTICIPANTS", str(int(min_p)) if min_p is not None else None)

    add_text(mc, "MODULE_ORDER", str(order))

    seg = row.get("SEGMENT_TYPE_KLASSE")
    add_text(mc, "SEGMENT_TYPE", int_str(seg))

    return mc


def build_contact(parent: Element, row: dict) -> Element | None:
    """Erstellt ein CONTACT-Element, falls Kontaktdaten vorhanden sind."""
    name  = row.get("KONTAKT_NAME")
    email = row.get("KONTAKT_EMAIL")
    phone = row.get("KONTAKT_TEL")

    if not any([name, email, phone]):
        return None

    contact = SubElement(parent, "CONTACT")

    if name:
        parts = str(name).strip().split(" ", 1)
        if len(parts) == 2:
            add_text(contact, "FIRST_NAME", parts[0])
            add_text(contact, "LAST_NAME",  parts[1])
        else:
            add_text(contact, "LAST_NAME", parts[0])

    if email:
        emails_el = SubElement(contact, "EMAILS")
        add_text(emails_el, "EMAIL", str(email))

    add_text(contact, "PHONE", str(phone) if phone is not None else None)
    return contact


def build_education(sd: Element, first_row: dict, rows: list[dict]) -> None:
    """
    Erstellt das EDUCATION-Element mit allen MODULE_COURSE-Einträgen
    (je eines pro Standort-Zeile).
    """
    edu = SubElement(sd, "EDUCATION")

    add_text(edu, "EDUCATION_TYPE",   int_str(first_row.get("EDUCATION_TYPE_KLASSE")))
    add_text(edu, "EXECUTION_FORM",   int_str(first_row.get("DURCHFUEHRUNGSFORM_KLASSE")))
    add_text(edu, "INSTITUTION",      int_str(first_row.get("INSTITUTION_KLASSE")))
    add_text(edu, "INSTRUCTION_FORM", int_str(first_row.get("UNTERRICHTSFORM_KLASSE")))
    add_text(edu, "LECTURE_PERIOD",   int_str(first_row.get("DAUER_KLASSE")))
    add_text(edu, "MEASURE_NUMBER",   first_row.get("MASSNAHMEN_NR"))

    # Ein MODULE_COURSE je Standort (1:n-Prinzip)
    for order, row in enumerate(rows, start=1):
        build_module_course(edu, row, order)


def build_service_classification(service: Element, row: dict) -> None:
    """Erstellt SERVICE_CLASSIFICATION aus KURSSYSTEMATIK (FNAME max. 60 Zeichen)."""
    systematik = row.get("KURSSYSTEMATIK")
    if not systematik:
        return
    sc   = SubElement(service, "SERVICE_CLASSIFICATION")
    feat = SubElement(sc, "FEATURE")
    # Schema: FNAME max. 60 Zeichen
    add_text(feat, "FNAME",  truncate(str(systematik), 60))
    add_text(feat, "FORDER", "1")


def build_service(new_el: Element, product_id, rows: list[dict]) -> None:
    """
    Erstellt ein vollständiges SERVICE-Element für ein Produkt mit
    allen zugehörigen Standort-Zeilen.

    Elementreihenfolge laut Schema:
      SERVICE > PRODUCT_ID, SERVICE_CLASSIFICATION, SERVICE_DETAILS
    SERVICE_DETAILS > TITLE zuerst, dann ANNOUNCEMENT, CONTACT, …
    """
    first = rows[0]
    modus = str(first.get("MODUS") or "new").strip()

    service = SubElement(new_el, "SERVICE", mode=modus)

    # PRODUCT_ID muss vor COURSE_TYPE stehen (Schema-Anforderung)
    add_text(service, "PRODUCT_ID", str(product_id))

    # COURSE_TYPE nur einfügen wenn numerisch (Schema erwartet Integer)
    course_type_raw = first.get("MODUL_TYP")
    if course_type_raw is not None:
        try:
            add_text(service, "COURSE_TYPE", str(int(str(course_type_raw).strip())))
        except (ValueError, TypeError):
            pass  # Textwert wie "COURSE" überspringen

    build_service_classification(service, first)

    # --- SERVICE_DETAILS ---
    sd = SubElement(service, "SERVICE_DETAILS")

    # TITLE muss als erstes Element in SERVICE_DETAILS stehen (Schema)
    add_text(sd, "TITLE", first.get("TITEL"))

    # Gesamtzeitraum (min. Startdatum / max. Enddatum über alle Standorte)
    all_starts = [fmt_datetime(r.get("START_DATUM")) for r in rows if r.get("START_DATUM")]
    all_ends   = [fmt_datetime(r.get("END_DATUM"))   for r in rows if r.get("END_DATUM")]
    overall_start = min(all_starts) if all_starts else None
    overall_end   = max(all_ends)   if all_ends   else None
    remarks_first = first.get("TERMIN_HINWEISE")

    if overall_start or overall_end or remarks_first:
        ann = SubElement(sd, "ANNOUNCEMENT")
        add_text(ann, "START_DATE",   overall_start)
        add_text(ann, "END_DATE",     overall_end)
        add_text(ann, "DATE_REMARKS", remarks_first)

    build_contact(sd, first)

    add_text(sd, "DESCRIPTION_LONG", first.get("BESCHREIBUNG"))

    build_education(sd, first, rows)

    # Schlagwörter (kommagetrennt → mehrere KEYWORD-Elemente)
    schlagworte = first.get("SCHLAGWORTE")
    if schlagworte:
        for kw in str(schlagworte).split(","):
            kw = kw.strip()
            if kw:
                add_text(sd, "KEYWORD", kw)

    seg = first.get("SEGMENT_TYPE_KLASSE")
    add_text(sd, "SEGMENT", int_str(seg))

    if overall_start or overall_end:
        sdate = SubElement(sd, "SERVICE_DATE")
        add_text(sdate, "START_DATE", overall_start)
        add_text(sdate, "END_DATE",   overall_end)

    add_text(sd, "SUPPLIER_ALT_PID", first.get("ALT_PRODUKT_ID"))


# ---------------------------------------------------------------------------
# Vollständiges XML-Dokument
# ---------------------------------------------------------------------------

def build_xml(order: list, groups: dict, args) -> Element:
    """
    Baut das komplette OPENQCAT-XML-Dokument.

    Wurzel-Elementreihenfolge laut Schema: HEADER, NEW, DELETE
    (HEADER muss als erstes Kind von OPENQCAT erscheinen).
    """
    root = Element("OPENQCAT")

    # --- HEADER zuerst (Schema-Anforderung) ---
    header = SubElement(root, "HEADER")

    catalog = SubElement(header, "CATALOG")
    # LANGUAGE muss als erstes Kind in CATALOG stehen (Schema)
    add_text(catalog, "LANGUAGE",        args.language)
    catalog_id = args.catalog_id or datetime.now().strftime("%Y%m%d%H%M%S")
    add_text(catalog, "CATALOG_ID",      catalog_id)
    add_text(catalog, "CATALOG_NAME",    args.catalog_name)
    add_text(catalog, "CATALOG_VERSION", args.catalog_version)
    add_text(catalog, "CURRENCY",        args.currency)
    add_text(catalog, "GENERATION_DATE", datetime.now().strftime("%Y-%m-%dT%H:%M:%S"))

    supplier = SubElement(header, "SUPPLIER")
    SubElement(supplier, "SUPPLIER_ID", type="supplier_specific").text = "245884"
    add_text(supplier, "SUPPLIER_NAME", "STARTUP PROFI einfach. clever. gründen.")

    sup_addr = SubElement(supplier, "ADDRESS")
    add_text(sup_addr, "NAME",    "STARTUP PROFI einfach. clever.")
    add_text(sup_addr, "STREET",  "Waldhofer Str. 102")
    add_text(sup_addr, "ZIP",     "69123")
    add_text(sup_addr, "CITY",    "Heidelberg")
    add_text(sup_addr, "COUNTRY", "D")          # ISO-Kürzel, nicht "Deutschland"
    add_text(sup_addr, "PHONE",   "+49.6221.3218416")
    sup_addr_emails = SubElement(sup_addr, "EMAILS")
    add_text(sup_addr_emails, "EMAIL", "info@startup-profi.de")

    sup_contact = SubElement(supplier, "CONTACT")
    SubElement(sup_contact, "CONTACT_ROLE", type="2").text = "Gesamtansprechpartner"
    add_text(sup_contact, "SALUTATION",  "m")
    add_text(sup_contact, "FIRST_NAME",  "Patrick")
    add_text(sup_contact, "LAST_NAME",   "Schaefer")
    add_text(sup_contact, "PHONE",       "+49.6221.3218416")
    sup_con_emails = SubElement(sup_contact, "EMAILS")
    add_text(sup_con_emails, "EMAIL", "info@startup-profi.de")
    SubElement(sup_contact, "CONTACT_REMARKS")

    add_text(supplier, "KEYWORD", "STARTUP PROFI einfach. clever. gründen.")
    ext_info = SubElement(supplier, "EXTENDED_INFO", input_type="2")
    SubElement(ext_info, "ORGANIZATIONAL_FORM", type="2").text = "Private Bildungseinrichtung"

    # --- DELETE und NEW nach HEADER ---
    delete_el = SubElement(root, "DELETE")
    new_el    = SubElement(root, "NEW")

    for pid in order:
        rows  = groups[pid]
        modus = str(rows[0].get("MODUS") or "new").strip().lower()

        if modus == "delete":
            svc = SubElement(delete_el, "SERVICE")
            add_text(svc, "PRODUCT_ID", str(pid))
        else:
            build_service(new_el, pid, rows)

    return root


# ---------------------------------------------------------------------------
# Ausgabe
# ---------------------------------------------------------------------------

def prettify(element: Element) -> str:
    """Gibt schön eingerücktes XML zurück."""
    rough  = tostring(element, encoding="unicode")
    dom    = xml.dom.minidom.parseString(rough)
    pretty = dom.toprettyxml(indent="    ", encoding=None)
    lines  = pretty.split("\n")
    if lines[0].startswith("<?xml"):
        lines = lines[1:]
    return '<?xml version="1.0" encoding="UTF-8"?>\n' + "\n".join(lines)


# ---------------------------------------------------------------------------
# Einstiegspunkt
# ---------------------------------------------------------------------------

def main():
    args = parse_args()

    print(f"Lese '{args.input}' ...")
    rows = read_excel(args.input)
    if not rows:
        sys.exit("Fehler: Keine Daten in der Excel-Datei gefunden.")
    print(f"  {len(rows)} Datenzeilen gelesen.")

    order, groups = group_by_product(rows)
    total_locations = sum(len(v) for v in groups.values())
    print(f"  {len(order)} eindeutige Produkte "
          f"(1:n-Prinzip: {total_locations} Standortzeilen → {len(order)} SERVICE-Elemente).")

    print("Erstelle XML ...")
    root = build_xml(order, groups, args)

    xml_str = prettify(root)

    with open(args.output, "w", encoding="utf-8") as fh:
        fh.write(xml_str)

    print(f"XML-Datei geschrieben: '{args.output}'")
    print("Fertig.")


if __name__ == "__main__":
    main()
