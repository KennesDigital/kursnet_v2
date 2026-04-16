#!/usr/bin/env python3
"""
Ninox-KursNet Converter
=======================
Konvertiert einen Ninox-Export (KursNet.xlsx) in eine KursNet-kompatible
open-Q-cat XML-Datei (Format V1.1).

WICHTIG – 1:n-Prinzip:
  Für jedes Produkt wird GENAU EIN Angebot-SERVICE (EDUCATION type="true")
  erzeugt, gefolgt von je einem Veranstaltungs-SERVICE (EDUCATION type="false")
  pro Standort-Zeile.

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
# Konstanten
# ---------------------------------------------------------------------------

COUNTRY_CODES: dict = {
    "Deutschland":        "D",
    "Germany":            "D",
    "Österreich":         "A",
    "Austria":            "A",
    "Schweiz":            "CH",
    "Switzerland":        "CH",
    "Frankreich":         "F",
    "France":             "F",
    "Niederlande":        "NL",
    "Netherlands":        "NL",
    "Belgien":            "B",
    "Belgium":            "B",
    "Polen":              "PL",
    "Poland":             "PL",
    "Tschechien":         "CZ",
    "Czech Republic":     "CZ",
    "Ungarn":             "H",
    "Hungary":            "H",
    "Italien":            "I",
    "Italy":              "I",
    "Spanien":            "E",
    "Spain":              "E",
    "Luxemburg":          "L",
    "Luxembourg":         "L",
    "Dänemark":           "DK",
    "Denmark":            "DK",
    "Schweden":           "S",
    "Sweden":             "S",
    "Norwegen":           "N",
    "Norway":             "N",
    "Finnland":           "FIN",
    "Finland":            "FIN",
    "USA":                "USA",
    "Vereinigte Staaten": "USA",
    "Großbritannien":     "GB",
    "United Kingdom":     "GB",
    "Russland":           "RUS",
    "Russia":             "RUS",
    "China":              "CHN",
    "Japan":              "J",
    "Türkei":             "TR",
    "Turkey":             "TR",
}

# Fallback für COURSE_TYPE wenn MODUL_TYP kein gültiger Integer
DEFAULT_COURSE_TYPE = 1

# ---------------------------------------------------------------------------
# Hilfsfunktionen
# ---------------------------------------------------------------------------

def add_text(parent: Element, tag: str, text, **attrib):
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
    """Formatiert als xs:dateTime (YYYY-MM-DDTHH:MM:SS)."""
    if value is None:
        return None
    if isinstance(value, datetime):
        return value.strftime("%Y-%m-%dT00:00:00")
    s = str(value).strip()
    if not s:
        return None
    if len(s) >= 10:
        return s[:10] + "T00:00:00"
    return None


def fmt_date(value) -> str | None:
    """Formatiert als xs:date (YYYY-MM-DD)."""
    if value is None:
        return None
    if isinstance(value, datetime):
        return value.strftime("%Y-%m-%d")
    s = str(value).strip()
    if not s:
        return None
    if len(s) >= 10:
        return s[:10]
    return None


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


def to_int_str(value, default: str = None) -> str | None:
    """Gibt Wert als Integer-String zurück oder default."""
    if value is None:
        return default
    try:
        return str(int(float(str(value).strip())))
    except (ValueError, TypeError):
        return default


def to_course_type(value) -> str:
    """Konvertiert MODUL_TYP in Integer-String für COURSE_TYPE (1-5)."""
    v = to_int_str(value)
    if v in ("1", "2", "3", "4", "5"):
        return v
    return str(DEFAULT_COURSE_TYPE)


def country_code(name) -> str | None:
    if not name:
        return None
    s = str(name).strip()
    return COUNTRY_CODES.get(s, s)


def truncate(value, max_len: int) -> str | None:
    if value is None:
        return None
    s = str(value).strip()
    if not s:
        return None
    return s[:max_len]


def fmt_zip(plz) -> str | None:
    """Konvertiert PLZ (ggf. als Float) in String."""
    if plz is None:
        return None
    if isinstance(plz, float):
        return str(int(plz))
    s = str(plz).strip()
    return s if s else None


# ---------------------------------------------------------------------------
# HEADER
# ---------------------------------------------------------------------------

def build_catalog(header: Element, args) -> None:
    """
    typeCATALOG-Reihenfolge:
      LANGUAGE → CATALOG_ID → CATALOG_VERSION → CATALOG_NAME → GENERATION_DATE → CURRENCY
    """
    catalog = SubElement(header, "CATALOG")
    # 1. LANGUAGE (Pflicht, muss ERSTE sein)
    SubElement(catalog, "LANGUAGE").text = args.language
    # 2. CATALOG_ID (Pflicht)
    SubElement(catalog, "CATALOG_ID").text = args.catalog_id or datetime.now().strftime("%Y%m%d%H%M%S")
    # 3. CATALOG_VERSION (Pflicht)
    SubElement(catalog, "CATALOG_VERSION").text = args.catalog_version
    # 4. CATALOG_NAME (optional)
    if args.catalog_name:
        SubElement(catalog, "CATALOG_NAME").text = args.catalog_name
    # 5. CURRENCY (optional)
    if args.currency:
        SubElement(catalog, "CURRENCY").text = args.currency


def build_supplier(header: Element) -> None:
    """
    typeSUPPLIER-Reihenfolge:
      SUPPLIER_ID → SUPPLIER_NAME → ADDRESS → CONTACT → KEYWORD → EXTENDED_INFO

    typeADDRESS-Reihenfolge:
      NAME → STREET → ZIP → CITY → COUNTRY (weitere optional)

    typeCONTACT-Reihenfolge:
      CONTACT_ROLE → SALUTATION → FIRST_NAME → LAST_NAME → PHONE → EMAILS

    SUPPLIER.EXTENDED_INFO:
      Pflichtattribut input_type; enthält ORGANIZATIONAL_FORM
    """
    supplier = SubElement(header, "SUPPLIER")

    # SUPPLIER_ID
    SubElement(supplier, "SUPPLIER_ID", type="supplier_specific").text = "245884"

    # SUPPLIER_NAME
    SubElement(supplier, "SUPPLIER_NAME").text = "STARTUP PROFI einfach. clever. gründen."

    # ADDRESS (NAME muss ERSTES Kind sein)
    addr = SubElement(supplier, "ADDRESS")
    SubElement(addr, "NAME").text   = "STARTUP PROFI einfach."   # max 30
    SubElement(addr, "STREET").text = "Waldhofer Str. 102"
    SubElement(addr, "ZIP").text    = "69123"
    SubElement(addr, "CITY").text   = "Heidelberg"
    SubElement(addr, "COUNTRY").text = "D"

    # CONTACT
    contact = SubElement(supplier, "CONTACT")
    SubElement(contact, "CONTACT_ROLE", type="2").text = "Gesamtansprechpartner"
    SubElement(contact, "SALUTATION").text  = "m"
    SubElement(contact, "FIRST_NAME").text  = "Patrick"
    SubElement(contact, "LAST_NAME").text   = "Schaefer"
    SubElement(contact, "PHONE").text       = "+49.6221.3218416"
    emails_el = SubElement(contact, "EMAILS")
    SubElement(emails_el, "EMAIL").text     = "info@startup-profi.de"

    # KEYWORD
    SubElement(supplier, "KEYWORD").text = "STARTUP PROFI einfach. clever. gründen."

    # EXTENDED_INFO (Pflicht, input_type Pflichtattribut)
    ext = SubElement(supplier, "EXTENDED_INFO", input_type="2")
    SubElement(ext, "ORGANIZATIONAL_FORM", type="2").text = "Privat"


# ---------------------------------------------------------------------------
# Adresse / Ort
# ---------------------------------------------------------------------------

def build_address_for_location(loc: Element, row: dict) -> None:
    """
    Füllt ein LOCATION-Element (typeADDRESS) in XSD-Reihenfolge:
      NAME → STREET → ZIP → CITY → COUNTRY
    Alle weiteren optionalen Felder werden weggelassen (nicht in den Daten).
    """
    add_text(loc, "NAME",    truncate(row.get("ORT_NAME"), 30))
    add_text(loc, "STREET",  truncate(row.get("ORT_STRASSE"), 30))
    add_text(loc, "ZIP",     fmt_zip(row.get("ORT_PLZ")))
    add_text(loc, "CITY",    row.get("ORT_STADT"))
    add_text(loc, "COUNTRY", country_code(row.get("ORT_LAND")))


# ---------------------------------------------------------------------------
# MODULE_COURSE
# ---------------------------------------------------------------------------

def build_module_course(education: Element, row: dict, dauer_klasse: str,
                        seg_type: str, is_angebot: bool) -> None:
    """
    typeMODULE_COURSE-Reihenfolge:
      MARKETINGTEXT → INSTRUCTOR → MIN_PARTICIPANTS → MAX_PARTICIPANTS
      → LOCATION → DURATION → SERVICE_REFERENCE → MODULE_ORDER
      → METHOD → MEDIA → MIME_INFO → INSTRUCTION_REMARKS
      → FLEXIBLE_START → EXTENDED_INFO(SEGMENT_TYPE)

    DURATION (typePERIOD_WITH_ATTRIBUTE):
      - Pflichtattribut 'type' = DAUER_KLASSE (Integer)
      - Für Veranstaltungen: leer (keine START_DATE/END_DATE-Kinder)
      - Für Angebote: ebenfalls leer (Datumsbereich steht in ANNOUNCEMENT)
    """
    mc = SubElement(education, "MODULE_COURSE")

    # 1. MIN_PARTICIPANTS / MAX_PARTICIPANTS
    add_text(mc, "MIN_PARTICIPANTS", to_int_str(row.get("MIN_TEILNEHMER")))
    add_text(mc, "MAX_PARTICIPANTS", to_int_str(row.get("MAX_TEILNEHMER")))

    # 2. LOCATION
    loc = SubElement(mc, "LOCATION")
    build_address_for_location(loc, row)

    # 3. DURATION (Pflichtattribut type; leer = keine Datumskinder)
    SubElement(mc, "DURATION", type=dauer_klasse)

    # 4. MODULE_ORDER
    SubElement(mc, "MODULE_ORDER").text = "1"

    # 5. INSTRUCTION_REMARKS (aus TERMIN_HINWEISE, nur bei Veranstaltung sinnvoll)
    if not is_angebot:
        add_text(mc, "INSTRUCTION_REMARKS",
                 truncate(str(row.get("TERMIN_HINWEISE") or ""), 2000))

    # 6. FLEXIBLE_START
    add_text(mc, "FLEXIBLE_START", to_bool(row.get("FLEXIBLE_ANMELDUNG")))

    # 7. EXTENDED_INFO → SEGMENT_TYPE (Pflicht innerhalb EXTENDED_INFO)
    mc_ext = SubElement(mc, "EXTENDED_INFO")
    SubElement(mc_ext, "SEGMENT_TYPE", type=seg_type)


# ---------------------------------------------------------------------------
# EDUCATION
# ---------------------------------------------------------------------------

def build_education_block(parent: Element, row: dict, is_angebot: bool,
                          angebot_pid: str = None) -> None:
    """
    Erzeugt SERVICE_MODULE > EDUCATION mit:
      - type="true"  für Angebote
      - type="false" für Veranstaltungen

    typeEDUCATION-Reihenfolge:
      COURSE_ID (opt) → ... → EXTENDED_INFO → MODULE_COURSE(s)

    EDUCATION.EXTENDED_INFO-Reihenfolge:
      INSTITUTION → INSTRUCTION_FORM → EXECUTION_FORM (opt)
      → EDUCATION_TYPE → MEASURE_NUMBER (opt)
    """
    sm  = SubElement(parent, "SERVICE_MODULE")
    edu_type = "true" if is_angebot else "false"
    edu = SubElement(sm, "EDUCATION", type=edu_type)

    # COURSE_ID bei Veranstaltungen = PRODUCT_ID des Angebots (Verknüpfung)
    if not is_angebot and angebot_pid:
        SubElement(edu, "COURSE_ID").text = str(angebot_pid)

    # EXTENDED_INFO (enthält Pflichtfelder INSTITUTION, INSTRUCTION_FORM, EDUCATION_TYPE)
    ext = SubElement(edu, "EXTENDED_INFO")

    inst = to_int_str(row.get("INSTITUTION_KLASSE"), "115")
    SubElement(ext, "INSTITUTION", type=inst)

    uf = to_int_str(row.get("UNTERRICHTSFORM_KLASSE"), "8")
    SubElement(ext, "INSTRUCTION_FORM", type=uf)

    df = to_int_str(row.get("DURCHFUEHRUNGSFORM_KLASSE"))
    if df:
        SubElement(ext, "EXECUTION_FORM", type=df)

    et = to_int_str(row.get("EDUCATION_TYPE_KLASSE"), "130")
    SubElement(ext, "EDUCATION_TYPE", type=et)

    if is_angebot:
        mn = row.get("MASSNAHMEN_NR")
        add_text(ext, "MEASURE_NUMBER", truncate(str(mn or ""), 30) if mn else None)

    # MODULE_COURSE
    dauer = to_int_str(row.get("DAUER_KLASSE"), "8")
    seg   = to_int_str(row.get("SEGMENT_TYPE_KLASSE"), "0")
    build_module_course(edu, row, dauer, seg, is_angebot)


# ---------------------------------------------------------------------------
# CONTACT in SERVICE_DETAILS
# ---------------------------------------------------------------------------

def build_contact_for_sd(parent: Element, row: dict) -> None:
    """
    typeCONTACT-Reihenfolge:
      CONTACT_ROLE → SALUTATION → FIRST_NAME → LAST_NAME → PHONE → EMAILS
    """
    name  = row.get("KONTAKT_NAME")
    email = row.get("KONTAKT_EMAIL")
    phone = row.get("KONTAKT_TEL")

    if not any([name, email, phone]):
        return

    contact = SubElement(parent, "CONTACT")

    # 1. CONTACT_ROLE (kommt VOR persönlichen Feldern per Schema)
    SubElement(contact, "CONTACT_ROLE", type="1").text = "Ansprechpartner"

    # 2. Persönliche Felder
    if name:
        parts = str(name).strip().split(" ", 1)
        if len(parts) == 2:
            add_text(contact, "FIRST_NAME", truncate(parts[0], 30))
            add_text(contact, "LAST_NAME",  truncate(parts[1], 30))
        else:
            add_text(contact, "LAST_NAME", truncate(parts[0], 30))

    # 3. Telefon
    add_text(contact, "PHONE", str(phone) if phone is not None else None)

    # 4. EMAILS (Container, nicht direktes EMAIL-Kind von CONTACT)
    if email:
        emails_el = SubElement(contact, "EMAILS")
        SubElement(emails_el, "EMAIL").text = str(email).strip()


# ---------------------------------------------------------------------------
# SERVICE_CLASSIFICATION
# ---------------------------------------------------------------------------

def build_service_classification(service: Element, row: dict) -> None:
    """
    typeSERVICE_CLASSIFICATION → typeFEATURE:
      FNAME (Pflicht, max 60) → FVALUE (Pflicht, mehrfach erlaubt)

    Mehrere Kurssystematik-Codes (semikolongetrennt) → ein FEATURE mit
    mehreren FVALUE-Kindern (FNAME muss je SERVICE_CLASSIFICATION eindeutig sein).
    """
    systematik = str(row.get("KURSSYSTEMATIK") or "").strip()
    if not systematik:
        return

    codes = [c.strip() for c in systematik.split(";") if c.strip()]
    if not codes:
        return

    sc   = SubElement(service, "SERVICE_CLASSIFICATION")
    feat = SubElement(sc, "FEATURE")
    SubElement(feat, "FNAME").text = "Kurssystematik"
    for code in codes:
        SubElement(feat, "FVALUE").text = code


# ---------------------------------------------------------------------------
# Angebot-SERVICE
# ---------------------------------------------------------------------------

def build_angebot(catalog_el: Element, product_id, rows: list[dict]) -> None:
    """
    Erzeugt einen Angebot-SERVICE (EDUCATION type="true", mode="new") mit:
      - Vollem Inhalt (Titel, Beschreibung, Kontakt, Schlagworte)
      - 1 MODULE_COURSE (erster Standort als Hauptstandort)
      - ANNOUNCEMENT mit dem Gesamtzeitraum aller Veranstaltungen
      - SERVICE_CLASSIFICATION NACH SERVICE_DETAILS

    typeSERVICE-Reihenfolge:
      PRODUCT_ID → COURSE_TYPE → SERVICE_DETAILS → SERVICE_CLASSIFICATION
    """
    first = rows[0]

    service = SubElement(catalog_el, "SERVICE", mode="new")
    SubElement(service, "PRODUCT_ID").text  = str(product_id)
    SubElement(service, "COURSE_TYPE").text = to_course_type(first.get("MODUL_TYP"))

    # --- SERVICE_DETAILS ---
    # typeSERVICE_DETAILS-Reihenfolge:
    #   TITLE → DESCRIPTION_LONG → CONTACT → KEYWORD → SERVICE_MODULE → ANNOUNCEMENT
    sd = SubElement(service, "SERVICE_DETAILS")

    # 1. TITLE (Pflicht, max 255)
    SubElement(sd, "TITLE").text = truncate(str(first.get("TITEL") or ""), 255) or ""

    # 2. DESCRIPTION_LONG (optional, max 30000)
    add_text(sd, "DESCRIPTION_LONG", first.get("BESCHREIBUNG"))

    # 3. CONTACT (optional)
    build_contact_for_sd(sd, first)

    # 4. KEYWORD(s) aus SCHLAGWORTE (semikolongetrennt, max 255 je Wort)
    schlagworte = str(first.get("SCHLAGWORTE") or "").strip()
    for kw in schlagworte.split(";"):
        kw = kw.strip()
        if kw:
            add_text(sd, "KEYWORD", truncate(kw, 255))

    # 5. SERVICE_MODULE → EDUCATION (type="true") → MODULE_COURSE
    build_education_block(sd, first, is_angebot=True)

    # 6. ANNOUNCEMENT (Sichtbarkeitsfenster: min/max über alle Zeilen)
    #    typePERIOD_DATE: START_DATE/END_DATE sind xs:date (nicht xs:dateTime!)
    start_dates = [fmt_date(r.get("START_DATUM")) for r in rows
                   if r.get("START_DATUM") and fmt_date(r.get("START_DATUM"))]
    end_dates   = [fmt_date(r.get("END_DATUM"))   for r in rows
                   if r.get("END_DATUM")   and fmt_date(r.get("END_DATUM"))]

    if start_dates or end_dates:
        ann = SubElement(sd, "ANNOUNCEMENT")
        if start_dates:
            SubElement(ann, "START_DATE").text = min(start_dates)
        if end_dates:
            SubElement(ann, "END_DATE").text   = max(end_dates)

    # --- SERVICE_CLASSIFICATION (muss NACH SERVICE_DETAILS stehen) ---
    build_service_classification(service, first)


# ---------------------------------------------------------------------------
# Veranstaltungs-SERVICE
# ---------------------------------------------------------------------------

def build_veranstaltung(catalog_el: Element, angebot_pid, index: int,
                        row: dict) -> None:
    """
    Erzeugt einen Veranstaltungs-SERVICE (EDUCATION type="false", KEIN mode-Attribut):
      - PRODUCT_ID = {angebot_pid}-V{index}
      - TITLE ist leer (<TITLE/>)
      - EDUCATION verweist via COURSE_ID auf das Angebot
      - 1 MODULE_COURSE mit Standort und DURATION (leer, nur type-Attribut)
      - ANNOUNCEMENT mit dem spezifischen Veranstaltungsdatum

    typeSERVICE-Reihenfolge:
      PRODUCT_ID → COURSE_TYPE → SERVICE_DETAILS
    """
    v_pid = f"{angebot_pid}-V{index}"

    # Kein mode-Attribut bei Veranstaltungen
    service = SubElement(catalog_el, "SERVICE")
    SubElement(service, "PRODUCT_ID").text  = v_pid
    SubElement(service, "COURSE_TYPE").text = to_course_type(row.get("MODUL_TYP"))

    # --- SERVICE_DETAILS ---
    sd = SubElement(service, "SERVICE_DETAILS")

    # TITLE ist leer (Pflichtfeld, aber bewusst leer bei Veranstaltungen)
    SubElement(sd, "TITLE")

    # SERVICE_MODULE → EDUCATION (type="false") → MODULE_COURSE
    build_education_block(sd, row, is_angebot=False, angebot_pid=str(angebot_pid))

    # ANNOUNCEMENT mit dem konkreten Veranstaltungsdatum (xs:date)
    start = fmt_date(row.get("START_DATUM"))
    end   = fmt_date(row.get("END_DATUM"))
    if start or end:
        ann = SubElement(sd, "ANNOUNCEMENT")
        if start:
            SubElement(ann, "START_DATE").text = start
        if end:
            SubElement(ann, "END_DATE").text   = end


# ---------------------------------------------------------------------------
# Haupt-Aufbau
# ---------------------------------------------------------------------------

def build_xml(rows: list[dict], args) -> Element:
    """
    Erzeugt den kompletten XML-Baum.
    - Produkte mit MODUS="delete" → UPDATE_CATALOG/DELETE/SERVICE
    - Alle anderen → Angebot + Veranstaltungen (NEW_CATALOG oder UPDATE_CATALOG/NEW)
    OPENQCAT erlaubt nur XOR: entweder NEW_CATALOG oder UPDATE_CATALOG.
    """
    # Gruppierung nach PRODUCT ID
    groups: dict = OrderedDict()
    for row in rows:
        pid = row.get("PRODUCT ID")
        if pid is not None:
            groups.setdefault(pid, []).append(row)

    delete_pids = []
    new_groups: dict = OrderedDict()
    for pid, grp in groups.items():
        modus = str(grp[0].get("MODUS") or "new").strip().lower()
        if modus == "delete":
            delete_pids.append(pid)
        else:
            new_groups[pid] = grp

    root = Element("OPENQCAT", version="1.1")
    header = SubElement(root, "HEADER")
    build_catalog(header, args)
    build_supplier(header)

    if delete_pids:
        # UPDATE_CATALOG mit DELETE und optional NEW
        uc = SubElement(root, "UPDATE_CATALOG", seq_number="1")
        del_el = SubElement(uc, "DELETE")
        for pid in delete_pids:
            svc = SubElement(del_el, "SERVICE")
            SubElement(svc, "PRODUCT_ID").text = str(pid)

        if new_groups:
            new_el = SubElement(uc, "NEW")
            for pid, grp in new_groups.items():
                build_angebot(new_el, pid, grp)
                for i, row in enumerate(grp, start=1):
                    build_veranstaltung(new_el, pid, i, row)
    else:
        # Nur neue/aktualisierte Produkte → NEW_CATALOG
        nc = SubElement(root, "NEW_CATALOG")
        for pid, grp in new_groups.items():
            build_angebot(nc, pid, grp)
            for i, row in enumerate(grp, start=1):
                build_veranstaltung(nc, pid, i, row)

    return root


# ---------------------------------------------------------------------------
# Excel lesen
# ---------------------------------------------------------------------------

def read_excel(filepath: str) -> list[dict]:
    wb = openpyxl.load_workbook(filepath, read_only=True, data_only=True)
    ws = wb.active
    rows = list(ws.iter_rows(values_only=True))
    if not rows:
        return []
    headers = [str(h).strip() if h is not None else "" for h in rows[0]]
    result = []
    for row in rows[1:]:
        if any(v is not None for v in row):
            result.append(dict(zip(headers, row)))
    wb.close()
    return result


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
    parser.add_argument("--catalog-version", default="1.0", help="Katalogversion")
    parser.add_argument("--language",        default="deu", help="Sprachcode ISO 639-2")
    parser.add_argument("--currency",        default="EUR", help="Währungscode ISO 4217")
    return parser.parse_args()


def main():
    args = parse_args()
    print(f"Lese: {args.input}")
    rows = read_excel(args.input)
    if not rows:
        sys.exit("Fehler: Keine Daten in der Excel-Datei gefunden.")
    print(f"  {len(rows)} Zeilen gelesen.")

    root = build_xml(rows, args)

    raw  = tostring(root, encoding="unicode")
    dom  = xml.dom.minidom.parseString(raw)
    pretty = dom.toprettyxml(indent="  ", encoding="UTF-8").decode("UTF-8")

    with open(args.output, "w", encoding="UTF-8") as f:
        f.write(pretty)

    print(f"Geschrieben: {args.output}")


if __name__ == "__main__":
    main()
