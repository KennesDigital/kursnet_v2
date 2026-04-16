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
    --language          Sprachcode ISO 639-2 (Standard: deu)
    --currency          Währungscode ISO 4217 (Standard: EUR)

Abhängigkeiten:
    pip install openpyxl
"""

import sys
from datetime import datetime
from collections import OrderedDict
import argparse
import xml.dom.minidom
from xml.etree.ElementTree import Element, SubElement, tostring

try:
    import openpyxl
except ImportError:
    sys.exit("Fehler: 'openpyxl' nicht installiert. Bitte 'pip install openpyxl' ausführen.")


# ---------------------------------------------------------------------------
# Konstanten
# ---------------------------------------------------------------------------

SUPPLIER_ID      = "245884"
MIME_SOURCE_URL  = "https://link.startup-profi.de/start"
CATALOG_ID       = "KATALOG_2026"
CATALOG_VERSION  = "001.001"
CATALOG_NAME     = "STARTUP PROFI Kurskatalog 2026"

# MODUL_TYP → COURSE_TYPE Mapping
MODUL_TYP_MAP: dict = {
    "COURSE":       "3",
    "SEMINAR":      "3",
    "TRAINING":     "3",
    "COURSE_UNIT":  "1",
    "PROGRAM":      "2",
    "WBT":          "4",
    "CBT":          "5",
}
DEFAULT_COURSE_TYPE = "3"

# KURSSYSTEMATIK Beschreibung → FNAME Kurzcode
KURSSYSTEMATIK_CODES: dict = {
    "§ 45 Abs. 1 Nr. 4 SGB III: Heranführung an eine selbständige Tätigkeit": "SA 04",
    "§ 45 Abs. 1 Nr. 1 SGB III: Heranführung an den Ausbildungs- und Arbeitsmarkt": "EC 01",
    "§ 45 Abs. 1 Nr. 2 SGB III: Feststellung, Verringerung oder Beseitigung von Vermittlungshemmnissen": "EC 02",
    "§ 45 Abs. 1 Nr. 3 SGB III: Vermittlung in eine versicherungspflichtige Beschäftigung": "EC 03",
    "§ 45 Abs. 1 Nr. 5 SGB III: Stabilisierung einer Beschäftigungsaufnahme": "EC 05",
}


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


def fmt_date_tz(value) -> str | None:
    """Formatiert als xs:date mit Zeitzone (YYYY-MM-DD+01:00)."""
    if value is None:
        return None
    if isinstance(value, datetime):
        return value.strftime("%Y-%m-%d") + "+01:00"
    s = str(value).strip()
    if not s:
        return None
    if len(s) >= 10:
        return s[:10] + "+01:00"
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
    """Konvertiert MODUL_TYP in COURSE_TYPE String."""
    if value is None:
        return DEFAULT_COURSE_TYPE
    s = str(value).strip()
    if s in MODUL_TYP_MAP:
        return MODUL_TYP_MAP[s]
    v = to_int_str(value)
    if v in ("1", "2", "3", "4", "5"):
        return v
    return DEFAULT_COURSE_TYPE


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


def get_fname_for_kurssystematik(description: str) -> str:
    """Gibt FNAME Kurzcode für eine KURSSYSTEMATIK-Beschreibung zurück."""
    return KURSSYSTEMATIK_CODES.get(description.strip(), description.strip()[:60])


# ---------------------------------------------------------------------------
# HEADER
# ---------------------------------------------------------------------------

def build_catalog(header: Element, args) -> None:
    """
    typeCATALOG-Reihenfolge:
      LANGUAGE → CATALOG_ID → CATALOG_VERSION → CATALOG_NAME → GENERATION_DATE → CURRENCY
    """
    catalog = SubElement(header, "CATALOG")
    SubElement(catalog, "LANGUAGE").text       = args.language
    SubElement(catalog, "CATALOG_ID").text     = CATALOG_ID
    SubElement(catalog, "CATALOG_VERSION").text = CATALOG_VERSION
    SubElement(catalog, "CATALOG_NAME").text   = CATALOG_NAME
    SubElement(catalog, "GENERATION_DATE").text = datetime.now().strftime("%Y-%m-%dT%H:%M:%S")
    SubElement(catalog, "CURRENCY").text       = args.currency


def build_supplier(header: Element) -> None:
    """
    typeSUPPLIER-Reihenfolge:
      SUPPLIER_ID → SUPPLIER_NAME → ADDRESS → CONTACT → KEYWORD → EXTENDED_INFO

    typeADDRESS-Reihenfolge:
      NAME → STREET → ZIP → CITY → COUNTRY → PHONE → EMAILS

    typeCONTACT-Reihenfolge:
      CONTACT_ROLE → SALUTATION → FIRST_NAME → LAST_NAME → PHONE → EMAILS → CONTACT_REMARKS
    """
    supplier = SubElement(header, "SUPPLIER")

    SubElement(supplier, "SUPPLIER_ID", type="supplier_specific").text = SUPPLIER_ID
    SubElement(supplier, "SUPPLIER_NAME").text = "STARTUP PROFI einfach. clever. gründen."

    addr = SubElement(supplier, "ADDRESS")
    SubElement(addr, "NAME").text    = "STARTUP PROFI einfach. clever."
    SubElement(addr, "STREET").text  = "Waldhofer Str. 102"
    SubElement(addr, "ZIP").text     = "69123"
    SubElement(addr, "CITY").text    = "Heidelberg"
    SubElement(addr, "COUNTRY").text = "Deutschland"
    SubElement(addr, "PHONE").text   = "+49.6221.3218416"
    addr_emails = SubElement(addr, "EMAILS")
    SubElement(addr_emails, "EMAIL").text = "info@startup-profi.de"

    contact = SubElement(supplier, "CONTACT")
    SubElement(contact, "CONTACT_ROLE", type="2").text = "Gesamtansprechpartner"
    SubElement(contact, "SALUTATION").text  = "m"
    SubElement(contact, "FIRST_NAME").text  = "Patrick"
    SubElement(contact, "LAST_NAME").text   = "Schaefer"
    SubElement(contact, "PHONE").text       = "+49.6221.3218416"
    con_emails = SubElement(contact, "EMAILS")
    SubElement(con_emails, "EMAIL").text    = "info@startup-profi.de"
    SubElement(contact, "CONTACT_REMARKS")

    SubElement(supplier, "KEYWORD").text = "STARTUP PROFI einfach. clever. gründen."

    ext = SubElement(supplier, "EXTENDED_INFO", input_type="2")
    SubElement(ext, "ORGANIZATIONAL_FORM", type="2").text = "Private Bildungseinrichtung"


# ---------------------------------------------------------------------------
# Adresse / Ort
# ---------------------------------------------------------------------------

def build_address_for_location(loc: Element, row: dict) -> None:
    """
    Füllt ein LOCATION-Element (typeADDRESS) in XSD-Reihenfolge:
      NAME → STREET → ZIP → CITY → COUNTRY
    COUNTRY wird direkt aus ORT_LAND übernommen (kein KFZ-Code).
    """
    add_text(loc, "NAME",    truncate(row.get("ORT_NAME"), 30))
    add_text(loc, "STREET",  truncate(row.get("ORT_STRASSE"), 30))
    add_text(loc, "ZIP",     fmt_zip(row.get("ORT_PLZ")))
    add_text(loc, "CITY",    row.get("ORT_STADT"))
    land = str(row.get("ORT_LAND") or "").strip()
    if land:
        SubElement(loc, "COUNTRY").text = land


# ---------------------------------------------------------------------------
# MODULE_COURSE
# ---------------------------------------------------------------------------

def build_module_course(education: Element, row: dict,
                        dauer_klasse: str, seg_type: str) -> None:
    """
    typeMODULE_COURSE-Reihenfolge:
      MIN_PARTICIPANTS → MAX_PARTICIPANTS → LOCATION → DURATION
      → FLEXIBLE_START → EXTENDED_INFO(SEGMENT_TYPE)

    Kein MODULE_ORDER (nicht im Referenzformat).
    """
    mc = SubElement(education, "MODULE_COURSE")

    add_text(mc, "MIN_PARTICIPANTS", to_int_str(row.get("MIN_TEILNEHMER")))
    add_text(mc, "MAX_PARTICIPANTS", to_int_str(row.get("MAX_TEILNEHMER")))

    loc = SubElement(mc, "LOCATION")
    build_address_for_location(loc, row)

    SubElement(mc, "DURATION", type=dauer_klasse)

    add_text(mc, "FLEXIBLE_START", to_bool(row.get("FLEXIBLE_ANMELDUNG")))

    mc_ext = SubElement(mc, "EXTENDED_INFO")
    SubElement(mc_ext, "SEGMENT_TYPE", type=seg_type)


# ---------------------------------------------------------------------------
# EDUCATION
# ---------------------------------------------------------------------------

def build_education_block(parent: Element, row: dict, is_angebot: bool,
                          product_id: str = None, angebot_pid: str = None) -> None:
    """
    Erzeugt SERVICE_MODULE > EDUCATION mit:
      - type="true"  für Angebote
      - type="false" für Veranstaltungen

    typeEDUCATION-Reihenfolge (Angebot):
      COURSE_ID → DEGREE → SUBSIDY → MIME_INFO → CERTIFICATE → EXTENDED_INFO → MODULE_COURSE

    typeEDUCATION-Reihenfolge (Veranstaltung):
      COURSE_ID → MIME_INFO → CERTIFICATE → EXTENDED_INFO → MODULE_COURSE

    EDUCATION.EXTENDED_INFO:
      Text-Inhalt = Attributwert (z.B. <INSTITUTION type="115">115</INSTITUTION>)
      MEASURE_NUMBER in beiden (Angebot und Veranstaltung).
    """
    sm  = SubElement(parent, "SERVICE_MODULE")
    edu = SubElement(sm, "EDUCATION", type="true" if is_angebot else "false")

    # COURSE_ID: bei Angebot = eigene PRODUCT_ID; bei Veranstaltung = Angebot-PID
    course_id = product_id if is_angebot else angebot_pid
    if course_id:
        SubElement(edu, "COURSE_ID").text = str(course_id)

    # DEGREE + SUBSIDY nur beim Angebot
    if is_angebot:
        degree = SubElement(edu, "DEGREE", type="0")
        SubElement(degree, "DEGREE_TITLE").text = "Keine Angabe zur Abschlussbezeichnung"
        deg_exam = SubElement(degree, "DEGREE_EXAM",
                              type="Abschlussart des Bildungsangebots")
        SubElement(deg_exam, "EXAMINER").text = "Keine Angabe"
        SubElement(degree, "DEGREE_ADD_QUALIFICATION").text = "Keine Angabe"
        SubElement(degree, "DEGREE_ENTITLED").text = "Keine Angabe"
        SubElement(edu, "SUBSIDY")

    # MIME_INFO (beide)
    mime_info = SubElement(edu, "MIME_INFO")
    mime_el   = SubElement(mime_info, "MIME_ELEMENT")
    SubElement(mime_el, "MIME_SOURCE").text = MIME_SOURCE_URL

    # CERTIFICATE (beide)
    cert = SubElement(edu, "CERTIFICATE")
    SubElement(cert, "CERTIFICATE_STATUS").text = "0"
    SubElement(cert, "CERTIFIER_NUMBER").text   = "25"
    SubElement(cert, "CERT_VALIDITY")

    # EXTENDED_INFO (Text-Inhalt = Attributwert)
    ext = SubElement(edu, "EXTENDED_INFO")

    inst = to_int_str(row.get("INSTITUTION_KLASSE"), "115")
    el = SubElement(ext, "INSTITUTION", type=inst)
    el.text = inst

    uf = to_int_str(row.get("UNTERRICHTSFORM_KLASSE"), "8")
    el = SubElement(ext, "INSTRUCTION_FORM", type=uf)
    el.text = uf

    df = to_int_str(row.get("DURCHFUEHRUNGSFORM_KLASSE"))
    if df:
        el = SubElement(ext, "EXECUTION_FORM", type=df)
        el.text = df

    et = to_int_str(row.get("EDUCATION_TYPE_KLASSE"), "130")
    el = SubElement(ext, "EDUCATION_TYPE", type=et)
    el.text = et

    # MEASURE_NUMBER in beiden (Angebot und Veranstaltung)
    mn = row.get("MASSNAHMEN_NR")
    add_text(ext, "MEASURE_NUMBER", truncate(str(mn or ""), 30) if mn else None)

    # MODULE_COURSE
    dauer = to_int_str(row.get("DAUER_KLASSE"), "8")
    seg   = to_int_str(row.get("SEGMENT_TYPE_KLASSE"), "0")
    build_module_course(edu, row, dauer, seg)


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
    SubElement(contact, "CONTACT_ROLE", type="1").text = "Ansprechpartner"

    if name:
        parts = str(name).strip().split(" ", 1)
        if len(parts) == 2:
            add_text(contact, "FIRST_NAME", truncate(parts[0], 30))
            add_text(contact, "LAST_NAME",  truncate(parts[1], 30))
        else:
            add_text(contact, "LAST_NAME", truncate(parts[0], 30))

    add_text(contact, "PHONE", str(phone) if phone is not None else None)

    if email:
        emails_el = SubElement(contact, "EMAILS")
        SubElement(emails_el, "EMAIL").text = str(email).strip()


# ---------------------------------------------------------------------------
# SERVICE_DATE und ANNOUNCEMENT
# ---------------------------------------------------------------------------

def build_service_date(sd: Element, row: dict) -> None:
    """SERVICE_DATE mit xs:dateTime-Datumsformat und DATE_REMARKS."""
    start = fmt_datetime(row.get("START_DATUM"))
    end   = fmt_datetime(row.get("END_DATUM"))
    if not start and not end:
        return
    sdate = SubElement(sd, "SERVICE_DATE")
    if start:
        SubElement(sdate, "START_DATE").text = start
    if end:
        SubElement(sdate, "END_DATE").text = end
    add_text(sdate, "DATE_REMARKS",
             truncate(str(row.get("TERMIN_HINWEISE") or ""), 2000))


def build_announcement(sd: Element, rows: list) -> None:
    """ANNOUNCEMENT mit xs:date + Zeitzone (+01:00)."""
    start_dates = [fmt_date_tz(r.get("START_DATUM")) for r in rows
                   if r.get("START_DATUM")]
    end_dates   = [fmt_date_tz(r.get("END_DATUM"))   for r in rows
                   if r.get("END_DATUM")]
    start_dates = [d for d in start_dates if d]
    end_dates   = [d for d in end_dates   if d]

    if not start_dates and not end_dates:
        return
    ann = SubElement(sd, "ANNOUNCEMENT")
    if start_dates:
        SubElement(ann, "START_DATE").text = min(start_dates)
    if end_dates:
        SubElement(ann, "END_DATE").text   = max(end_dates)


# ---------------------------------------------------------------------------
# SERVICE_CLASSIFICATION
# ---------------------------------------------------------------------------

def build_service_classification(service: Element, row: dict) -> None:
    """
    SERVICE_CLASSIFICATION mit REFERENCE_CLASSIFICATION_SYSTEM_NAME.
    FNAME = Kurzcode aus KURSSYSTEMATIK_CODES; FVALUE = Beschreibung.
    Mehrere Einträge (semikolongetrennt) → mehrere FVALUE.
    """
    systematik = str(row.get("KURSSYSTEMATIK") or "").strip()
    if not systematik:
        return

    codes = [c.strip() for c in systematik.split(";") if c.strip()]
    if not codes:
        return

    sc = SubElement(service, "SERVICE_CLASSIFICATION")
    SubElement(sc, "REFERENCE_CLASSIFICATION_SYSTEM_NAME").text = "Kurssystematik"
    feat = SubElement(sc, "FEATURE")
    SubElement(feat, "FNAME").text = get_fname_for_kurssystematik(codes[0])
    for code in codes:
        SubElement(feat, "FVALUE").text = code


# ---------------------------------------------------------------------------
# Angebot-SERVICE
# ---------------------------------------------------------------------------

def build_angebot(catalog_el: Element, product_id, rows: list[dict]) -> None:
    """
    Erzeugt einen Angebot-SERVICE (EDUCATION type="true", mode="new").

    typeSERVICE-Reihenfolge:
      PRODUCT_ID → COURSE_TYPE → SUPPLIER_ID_REF → SERVICE_DETAILS
      → SERVICE_CLASSIFICATION → SERVICE_PRICE_DETAILS → MIME_INFO
    """
    first = rows[0]
    pid   = str(product_id)

    service = SubElement(catalog_el, "SERVICE", mode="new")
    SubElement(service, "PRODUCT_ID").text   = pid
    SubElement(service, "COURSE_TYPE").text  = to_course_type(first.get("MODUL_TYP"))
    SubElement(service, "SUPPLIER_ID_REF",
               type="supplier_specific").text = SUPPLIER_ID

    # --- SERVICE_DETAILS ---
    # Reihenfolge (XSD typeSERVICE_DETAILS):
    #   TITLE → DESCRIPTION_LONG → SUPPLIER_ALT_PID → CONTACT → SERVICE_DATE
    #   → KEYWORD(s) → TERMS_AND_CONDITIONS → SERVICE_MODULE → ANNOUNCEMENT
    sd = SubElement(service, "SERVICE_DETAILS")

    SubElement(sd, "TITLE").text = truncate(str(first.get("TITEL") or ""), 255) or ""

    add_text(sd, "DESCRIPTION_LONG", first.get("BESCHREIBUNG"))

    add_text(sd, "SUPPLIER_ALT_PID", truncate(first.get("ALT_PRODUKT_ID"), 30))

    build_contact_for_sd(sd, first)

    build_service_date(sd, first)

    schlagworte = str(first.get("SCHLAGWORTE") or "").strip()
    for kw in schlagworte.split(","):
        kw = kw.strip()
        if kw:
            add_text(sd, "KEYWORD", truncate(kw, 255))

    SubElement(sd, "TERMS_AND_CONDITIONS")

    build_education_block(sd, first, is_angebot=True, product_id=pid)

    build_announcement(sd, rows)

    # --- SERVICE_CLASSIFICATION (nach SERVICE_DETAILS) ---
    build_service_classification(service, first)

    # --- SERVICE_PRICE_DETAILS ---
    SubElement(service, "SERVICE_PRICE_DETAILS")

    # --- MIME_INFO auf SERVICE-Ebene (nur Angebot) ---
    mime_info = SubElement(service, "MIME_INFO")
    mime_el   = SubElement(mime_info, "MIME_ELEMENT")
    SubElement(mime_el, "MIME_SOURCE").text = MIME_SOURCE_URL


# ---------------------------------------------------------------------------
# Veranstaltungs-SERVICE
# ---------------------------------------------------------------------------

def build_veranstaltung(catalog_el: Element, angebot_pid, index: int,
                        row: dict) -> None:
    """
    Erzeugt einen Veranstaltungs-SERVICE (EDUCATION type="false", kein mode-Attribut).

    PRODUCT_ID = {angebot_pid}-V{index:03d} (dreistellig, nullaufgefüllt)

    typeSERVICE-Reihenfolge:
      PRODUCT_ID → COURSE_TYPE → SUPPLIER_ID_REF → SERVICE_DETAILS → SERVICE_PRICE_DETAILS
    """
    v_pid = f"{angebot_pid}-V{index:03d}"

    service = SubElement(catalog_el, "SERVICE")
    SubElement(service, "PRODUCT_ID").text   = v_pid
    SubElement(service, "COURSE_TYPE").text  = to_course_type(row.get("MODUL_TYP"))
    SubElement(service, "SUPPLIER_ID_REF",
               type="supplier_specific").text = SUPPLIER_ID

    # --- SERVICE_DETAILS ---
    # Reihenfolge: TITLE → SUPPLIER_ALT_PID → SERVICE_DATE
    #              → TERMS_AND_CONDITIONS → SERVICE_MODULE → ANNOUNCEMENT
    sd = SubElement(service, "SERVICE_DETAILS")

    SubElement(sd, "TITLE")  # leer bei Veranstaltungen

    add_text(sd, "SUPPLIER_ALT_PID", truncate(row.get("ALT_PRODUKT_ID"), 30))

    build_service_date(sd, row)

    SubElement(sd, "TERMS_AND_CONDITIONS")

    build_education_block(sd, row, is_angebot=False,
                          angebot_pid=str(angebot_pid))

    build_announcement(sd, [row])

    # --- SERVICE_PRICE_DETAILS ---
    SubElement(service, "SERVICE_PRICE_DETAILS")


# ---------------------------------------------------------------------------
# Haupt-Aufbau
# ---------------------------------------------------------------------------

def build_xml(rows: list[dict], args) -> Element:
    """
    Erzeugt den kompletten XML-Baum.

    Immer UPDATE_CATALOG (seq_number="1"):
      - DELETE-Block bei Produkten mit MODUS="delete"
      - NEW-Block für alle anderen Produkte (Angebot + Veranstaltungen)
    """
    # Gruppierung nach PRODUCT ID
    groups: dict = OrderedDict()
    for row in rows:
        pid = row.get("PRODUCT ID")
        if pid is not None:
            groups.setdefault(pid, []).append(row)

    delete_pids  = []
    new_groups: dict = OrderedDict()
    for pid, grp in groups.items():
        modus = str(grp[0].get("MODUS") or "new").strip().lower()
        if modus == "delete":
            delete_pids.append(pid)
        else:
            new_groups[pid] = grp

    # Root-Element mit Namespaces
    root = Element(
        "OPENQCAT",
        version="1.1",
        **{
            "xmlns:xsi": "http://www.w3.org/2001/XMLSchema-instance",
            "xsi:noNamespaceSchemaLocation": "openQ-cat.V1.1.xsd",
        },
    )

    # HEADER
    header = SubElement(root, "HEADER")
    SubElement(header, "GENERATOR_INFO").text = "KURSNET XML Generator (Python)"
    build_catalog(header, args)
    build_supplier(header)
    for agreement_id in ("NB 7.0", "DSE 5.0"):
        agr = SubElement(header, "AGREEMENT")
        SubElement(agr, "AGREEMENT_ID").text = agreement_id

    # UPDATE_CATALOG (immer)
    uc = SubElement(root, "UPDATE_CATALOG", seq_number="1")

    if delete_pids:
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
    parser.add_argument("--language", default="deu", help="Sprachcode ISO 639-2")
    parser.add_argument("--currency", default="EUR", help="Währungscode ISO 4217")
    return parser.parse_args()


def main():
    args = parse_args()
    print(f"Lese: {args.input}")
    rows = read_excel(args.input)
    if not rows:
        sys.exit("Fehler: Keine Daten in der Excel-Datei gefunden.")
    print(f"  {len(rows)} Zeilen gelesen.")

    root = build_xml(rows, args)

    raw    = tostring(root, encoding="unicode")
    dom    = xml.dom.minidom.parseString(raw)
    pretty = dom.toprettyxml(indent="  ", encoding="UTF-8").decode("UTF-8")

    with open(args.output, "w", encoding="UTF-8") as f:
        f.write(pretty)

    print(f"Geschrieben: {args.output}")


if __name__ == "__main__":
    main()
