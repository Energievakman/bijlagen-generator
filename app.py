from flask import Flask, request, send_file, jsonify, abort
from werkzeug.middleware.proxy_fix import ProxyFix
from werkzeug.exceptions import HTTPException
import os
import uuid
import re
import json
import mimetypes
import subprocess
import threading
import traceback
from pathlib import Path
from datetime import datetime
from xml.sax.saxutils import escape as xml_escape

import requests
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import mm
from reportlab.platypus import (
    SimpleDocTemplate,
    Paragraph,
    Spacer,
    Table,
    TableStyle,
    Image,
)

try:
    import cloudinary.uploader
except Exception:
    cloudinary = None


app = Flask(__name__)
app.wsgi_app = ProxyFix(app.wsgi_app, x_proto=1, x_host=1)
APP_VERSION = "v9_opdrachtbevestiging_echte_branch_zonder_bijlage4_tabel_20260627"


@app.errorhandler(HTTPException)
def handle_http_exception(e):
    return jsonify({
        "status": "error",
        "error": e.name,
        "detail": e.description,
    }), e.code


@app.errorhandler(Exception)
def handle_exception(e):
    app.logger.exception("Onverwachte fout")
    return jsonify({
        "status": "error",
        "error": "server_error",
        "detail": str(e),
    }), 500

BASE_DIR = Path(__file__).resolve().parent
TMP_DIR = Path(os.environ.get("TMP_DIR", "/tmp/softr_documenten"))
TMP_DIR.mkdir(parents=True, exist_ok=True)

# Optioneel: oude Uniec3 generator laten werken op dezelfde Render-service.
GENERATOR = BASE_DIR / "generator.py"
TEMPLATE = BASE_DIR / "template.uniec3"

SOFTR_API_BASE = os.environ.get(
    "SOFTR_API_BASE",
    "https://tables-api.softr.io/api/v1",
).rstrip("/")


DOCUMENT_CONFIG = {
    "bijlage4": {
        "title": "Bijlage 4",
        "subtitle": "Beschikbaar gestelde informatie opdrachtgever",
        "source_table_env": "SOFTR_TABLE_BIJLAGE4_ID",
        "target_field_env": "SOFTR_FIELD_DOSSIER_BIJLAGE4_PDF",
        "filename_prefix": "Bijlage_4",
    },
    "opdrachtbevestiging": {
        "title": "Opdrachtbevestiging",
        "subtitle": "Opdrachtbevestiging energielabel",
        "source_table_env": "SOFTR_TABLE_OPDRACHTBEVESTIGING_ID",
        "target_field_env": "SOFTR_FIELD_DOSSIER_OPDRACHTBEVESTIGING_PDF",
        "filename_prefix": "Opdrachtbevestiging",
    },
}

DEFAULT_INFO_ROWS = [
    "Bouwkundige tekeningen",
    "Gevelaanzichten",
    "Doorsnedes",
    "Plattegronden",
    "Constructietekeningen",
    "Installatieontwerp van het gebouw",
    "Installatietekeningen voor verwarming",
    "Installatietekeningen voor tapwater",
    "Installatietekeningen voor koeling",
    "Inregeling verwarming conform protocol",
    "Verzamellijsten opwekkers verwarming",
    "Verzamellijsten kozijnen en beglazingen",
]

INTERNAL_KEYS = {
    "secret",
    "document_type",
    "source_record_id",
    "dossier_record_id",
    "target_record_id",
    "fields",
    "items",
    "html",
}

# Velden die technisch zijn of ergens anders in het document gebruikt worden.
# Deze wil je niet als losse informatie-regel in de tabel tonen.
SKIP_INFO_FIELD_NAMES = {
    "name",
    "id",
    "id (dossiers)",
    "dossier id",
    "dossiers id",
    "id_dossiers",
    "adres",
    "address",
    "adres (vol)",
    "adres vol",
    "volledig adres",
    "adviseur",
    "naam adviseur",
    "opdrachtgever",
    "klant",
    "naam opdrachtgever",
    "datum",
    "gegenereerd op",
    "paraaf",
    "signature",
    "signature_url",
    "signature_path",
    "handtekening",
}


# -----------------------------------------------------------------------------
# Helpers algemeen
# -----------------------------------------------------------------------------


def env(name: str, default: str = "") -> str:
    return os.environ.get(name, default).strip()


def truthy_env(name: str, default: str = "false") -> bool:
    return env(name, default).lower() in {"1", "true", "yes", "ja", "on"}


def safe_filename(value: str) -> str:
    value = (value or "document").strip().replace(",", "")
    value = re.sub(r"[^A-Za-z0-9_\-\. ]+", "", value)
    value = re.sub(r"\s+", "_", value)
    return value[:100] or "document"


def now_nl() -> str:
    return datetime.now().strftime("%d-%m-%Y")


def clean_value(value):
    """Maak waarden netjes leesbaar voor de PDF."""
    if value is None:
        return "-"
    if isinstance(value, bool):
        return "aanwezig" if value else "niet aanwezig"
    if isinstance(value, str):
        stripped = value.strip()
        if not stripped:
            return "-"
        if stripped.lower() == "true":
            return "aanwezig"
        if stripped.lower() == "false":
            return "niet aanwezig"
        return stripped
    if isinstance(value, (int, float)):
        return str(value)
    if isinstance(value, list):
        if not value:
            return "-"
        parts = []
        for item in value:
            if isinstance(item, dict):
                parts.append(str(item.get("name") or item.get("url") or item.get("fileUrl") or item.get("id") or item))
            else:
                parts.append(str(item))
        return ", ".join(parts) if parts else "-"
    if isinstance(value, dict):
        return str(value.get("name") or value.get("url") or value.get("fileUrl") or value.get("id") or value)
    return str(value)


def pick(data: dict, *keys, default=""):
    """Pak eerste niet-lege waarde uit dictionary, case-insensitive."""
    if not isinstance(data, dict):
        return default

    # exacte match eerst
    for key in keys:
        if key in data and clean_value(data.get(key)) != "-":
            return data.get(key)

    # daarna case-insensitive match
    lower_map = {str(k).lower(): k for k in data.keys()}
    for key in keys:
        found = lower_map.get(str(key).lower())
        if found is not None and clean_value(data.get(found)) != "-":
            return data.get(found)
    return default


def pick_raw(data: dict, *keys, default=None):
    """Pak eerste niet-lege waarde zonder clean_value toe te passen.

Handig voor file fields zoals Paraaf, omdat we het originele object/array met URL
willen houden in plaats van alleen een stringrepresentatie.
    """
    if not isinstance(data, dict):
        return default

    for key in keys:
        if key in data:
            value = data.get(key)
            if clean_value(value) != "-":
                return value

    lower_map = {str(k).lower(): k for k in data.keys()}
    for key in keys:
        found = lower_map.get(str(key).lower())
        if found is not None:
            value = data.get(found)
            if clean_value(value) != "-":
                return value
    return default


def should_skip_info_field(key: str) -> bool:
    key_norm = str(key or "").strip().lower()
    if not key_norm:
        return True
    if key_norm in SKIP_INFO_FIELD_NAMES or key_norm in INTERNAL_KEYS:
        return True
    # Technische ID velden uit linked records niet in het rapport tonen.
    if key_norm.startswith("id (") or key_norm.endswith(" record id") or key_norm.endswith("_record_id"):
        return True
    return False


def extract_file_url(value) -> str:
    """Haal een bruikbare URL uit een Softr file/image field."""
    if value is None:
        return ""
    if isinstance(value, str):
        return value.strip()
    if isinstance(value, list):
        for item in value:
            url = extract_file_url(item)
            if url:
                return url
        return ""
    if isinstance(value, dict):
        for key in ("url", "fileUrl", "downloadUrl", "signedUrl", "src", "href"):
            url = value.get(key)
            if isinstance(url, str) and url.strip():
                return url.strip()
        for key in ("file", "data", "value", "attachment"):
            nested = value.get(key)
            url = extract_file_url(nested)
            if url:
                return url
    return ""


def first_record_id(value):
    """Ondersteunt tekst, linked-record arrays en objecten uit Softr."""
    if value is None:
        return ""
    if isinstance(value, str):
        return value.strip()
    if isinstance(value, dict):
        return str(value.get("id") or value.get("recordId") or value.get("value") or "").strip()
    if isinstance(value, list) and value:
        return first_record_id(value[0])
    return ""


def get_json_payload():
    return request.get_json(silent=True) or {}


def require_secret(payload: dict):
    shared_secret = env("APP_SHARED_SECRET")
    if not shared_secret:
        return

    provided = (
        request.headers.get("X-Api-Key")
        or request.headers.get("X-App-Secret")
        or payload.get("secret")
        or request.args.get("secret")
    )
    if provided != shared_secret:
        abort(401, description="Ongeldige of ontbrekende secret")


# -----------------------------------------------------------------------------
# Softr Database API
# -----------------------------------------------------------------------------


def softr_headers():
    api_key = env("SOFTR_API_KEY")
    if not api_key:
        raise RuntimeError("SOFTR_API_KEY ontbreekt in Render environment variables")
    return {
        "Softr-Api-Key": api_key,
        "Content-Type": "application/json",
    }


def database_id():
    db_id = env("SOFTR_DATABASE_ID")
    if not db_id:
        raise RuntimeError("SOFTR_DATABASE_ID ontbreekt in Render environment variables")
    return db_id


def softr_get_record(table_id: str, record_id: str, field_names: bool = True) -> dict:
    if not table_id:
        raise RuntimeError("Softr source table ID ontbreekt")
    if not record_id:
        raise RuntimeError("source_record_id ontbreekt")

    url = f"{SOFTR_API_BASE}/databases/{database_id()}/tables/{table_id}/records/{record_id}"
    params = {"fieldNames": "true"} if field_names else None
    r = requests.get(url, headers=softr_headers(), params=params, timeout=30)
    if not r.ok:
        raise RuntimeError(f"Softr get record mislukt: {r.status_code} {r.text[:1000]}")
    return r.json().get("data", {})


def make_file_field_value(file_url: str, filename: str, mode: str):
    """Verschillende Softr setups accepteren soms een andere vorm voor file fields."""
    if mode == "url":
        return file_url
    if mode == "array_url":
        return [file_url]
    if mode == "object":
        return {"url": file_url, "name": filename}
    if mode == "array_object":
        return [{"url": file_url, "name": filename}]
    raise ValueError(f"Onbekende SOFTR_FILE_FIELD_VALUE_MODE: {mode}")


def softr_patch_dossier_file(dossier_record_id: str, field_id: str, file_url: str, filename: str):
    if not dossier_record_id:
        raise RuntimeError("dossier_record_id ontbreekt")
    if not field_id:
        raise RuntimeError("Doelveld in Dossiers ontbreekt. Vul het juiste field ID in Render in.")

    dossiers_table_id = env("SOFTR_TABLE_DOSSIERS_ID")
    if not dossiers_table_id:
        raise RuntimeError("SOFTR_TABLE_DOSSIERS_ID ontbreekt in Render environment variables")

    url = f"{SOFTR_API_BASE}/databases/{database_id()}/tables/{dossiers_table_id}/records/{dossier_record_id}"

    configured_mode = env("SOFTR_FILE_FIELD_VALUE_MODE")
    modes = [configured_mode] if configured_mode else ["url", "array_object", "array_url", "object"]
    timeout = int(env("SOFTR_REQUEST_TIMEOUT", "180"))
    accept_gateway_timeout = truthy_env("SOFTR_ACCEPT_GATEWAY_TIMEOUT_AS_SUCCESS", "true")

    errors = []
    for mode in modes:
        body = {"fields": {field_id: make_file_field_value(file_url, filename, mode)}}
        try:
            r = requests.patch(url, headers=softr_headers(), json=body, timeout=timeout)
        except requests.exceptions.ReadTimeout as e:
            app.logger.warning(
                "Softr PATCH read-timeout na %s sec. PDF-url is mogelijk alsnog verwerkt. mode=%s dossier=%s field=%s file=%s",
                timeout, mode, dossier_record_id, field_id, file_url,
            )
            if accept_gateway_timeout:
                return {
                    "mode": mode,
                    "assumed_success": True,
                    "reason": "read_timeout",
                    "detail": str(e),
                }
            errors.append({"mode": mode, "status": "read_timeout", "body": str(e)[:1000]})
            continue

        if r.ok:
            try:
                response_json = r.json()
            except Exception:
                response_json = {"raw": r.text[:1000]}
            return {
                "mode": mode,
                "response": response_json,
            }

        # Softr geeft bij file fields soms 504 terug terwijl het bestand even later wél in het record staat.
        # Daarom behandelen we 502/503/504 optioneel als 'waarschijnlijk gelukt', zodat Softr niet rood faalt.
        if r.status_code in {502, 503, 504} and accept_gateway_timeout:
            app.logger.warning(
                "Softr PATCH gaf %s, maar wordt als soft-success behandeld. mode=%s dossier=%s field=%s body=%s",
                r.status_code, mode, dossier_record_id, field_id, r.text[:500],
            )
            return {
                "mode": mode,
                "assumed_success": True,
                "reason": f"softr_{r.status_code}",
                "body": r.text[:1000],
            }

        errors.append({"mode": mode, "status": r.status_code, "body": r.text[:1000]})

    raise RuntimeError("Softr update dossier mislukt: " + json.dumps(errors, ensure_ascii=False))


def softr_patch_dossier_file_background(dossier_record_id: str, field_id: str, file_url: str, filename: str, document_type: str):
    """Voert de update naar het Dossiers-record op de achtergrond uit.

    Hierdoor krijgt Softr meteen een 200-response en blijft de knop niet minutenlang draaien.
    """
    try:
        result = softr_patch_dossier_file(dossier_record_id, field_id, file_url, filename)
        app.logger.info(
            "Achtergrond-update klaar: document_type=%s dossier=%s field=%s result=%s",
            document_type, dossier_record_id, field_id, result,
        )
    except Exception as e:
        app.logger.error(
            "Achtergrond-update mislukt: document_type=%s dossier=%s field=%s error=%s\n%s",
            document_type, dossier_record_id, field_id, e, traceback.format_exc(),
        )

# -----------------------------------------------------------------------------
# PDF generatie
# -----------------------------------------------------------------------------


def build_items(payload: dict, source_fields: dict) -> list:
    """Bouw de informatie-rijen voor Bijlage 4 / opdrachtbevestiging."""
    items = payload.get("items")

    if isinstance(items, list) and items:
        result = []
        for item in items:
            if isinstance(item, dict):
                label = item.get("label") or item.get("name") or item.get("field") or "-"
                if should_skip_info_field(label):
                    continue
                value = item.get("value") if "value" in item else item.get("status", "-")
                result.append((str(label), clean_value(value)))
            elif isinstance(item, str):
                if should_skip_info_field(item):
                    continue
                result.append((item, clean_value(source_fields.get(item))))
        return result

    if isinstance(items, dict) and items:
        return [
            (str(k), clean_value(v))
            for k, v in items.items()
            if not should_skip_info_field(k)
        ]

    # Vaste basisregels bovenaan.
    default_rows = []
    for label in DEFAULT_INFO_ROWS:
        default_rows.append((label, clean_value(pick(source_fields, label, default="-"))))

    # Extra velden die niet technisch zijn en niet al in de vaste regels zitten.
    existing = {label.lower() for label, _ in default_rows}
    extra_rows = []
    for key, value in source_fields.items():
        key_str = str(key)
        key_norm = key_str.strip().lower()
        if key_norm in existing or should_skip_info_field(key_str):
            continue
        if clean_value(value) == "-":
            continue
        extra_rows.append((key_str, clean_value(value)))

    return default_rows + extra_rows[:20]


def get_public_base_url():
    configured = env("PUBLIC_BASE_URL")
    if configured:
        return configured.rstrip("/")
    return request.url_root.rstrip("/")


def signature_to_local_path(signature_value) -> str:
    """Download of resolve een paraaf/handtekening naar een lokaal bestand."""
    signature_ref = extract_file_url(signature_value)
    if not signature_ref:
        signature_ref = clean_value(signature_value)
    if not signature_ref or signature_ref == "-":
        return ""

    try:
        if signature_ref.startswith("http://") or signature_ref.startswith("https://"):
            resp = requests.get(signature_ref, timeout=20)
            resp.raise_for_status()
            content_type = (resp.headers.get("content-type") or "").lower()
            ext = ".png"
            if "jpeg" in content_type or "jpg" in content_type:
                ext = ".jpg"
            elif "webp" in content_type:
                ext = ".webp"
            sig_path = TMP_DIR / f"signature_{uuid.uuid4().hex}{ext}"
            sig_path.write_bytes(resp.content)
            return str(sig_path)

        sig_path = Path(signature_ref)
        if sig_path.exists():
            return str(sig_path)
    except Exception as e:
        app.logger.warning("Paraaf kon niet worden geladen: %s", e)
    return ""


def make_signature_image(signature_value):
    """Maak een ReportLab Image object voor de paraaf-cel."""
    sig_path = signature_to_local_path(signature_value)
    if not sig_path:
        return ""
    try:
        return Image(sig_path, width=38 * mm, height=18 * mm, kind="proportional")
    except Exception as e:
        app.logger.warning("Paraaf kon niet in PDF worden geplaatst: %s", e)
        return ""



OPDRACHTBEVESTIGING_DEFAULT_VOORWAARDEN = [
    "De opdrachtgever is eigenaar van het betreffende object of geeft de opdracht namens de eigenaar.",
    "De opnamegegevens, aanwezige tekeningen, meetstaten, plattegronden en installatieschema’s worden toegevoegd aan het digitale dossier. Ook worden er foto’s gemaakt van en in het object. Deze gegevens worden 15 jaar bewaard.",
    "Als opdrachtgever heeft u het recht om het volledige projectdossier op te vragen bij het bedrijf dat de werkzaamheden uitvoert.",
    "De objectkenmerken die worden opgenomen in het monitoringsbestand worden geregistreerd in de landelijke energielabeldatabase van de RVO: www.ep-online.nl.",
    "U stemt ermee in dat de opnamegegevens kunnen worden doorgegeven aan de Rijksoverheid en aan de certificerende instelling (CI) van het bedrijf dat de werkzaamheden uitvoert.",
    "Wanneer er een audit of kwaliteitscontrole plaatsvindt, dient u het bedrijf dat de werkzaamheden heeft uitgevoerd en de certificerende instelling opnieuw toegang te verlenen tot het object. Als toegang wordt geweigerd, kan het energielabel worden ingetrokken. U wordt hierover geïnformeerd.",
    "Het actuele procescertificaat van De Energievakman is te vinden in het Centraal Register Techniek.",
]


def pdf_escape(value) -> str:
    """Escape tekst voor ReportLab Paragraph."""
    return xml_escape(clean_value(value)).replace("\n", "<br/>")


def p(value, style):
    return Paragraph(pdf_escape(value), style)


def split_voorwaarden(value) -> list:
    """Maak nette bullets van het veld Voorwaarden.

    Als het record geen voorwaarden bevat, gebruiken we de standaard voorwaarden.
    Ondersteunt tekst met bullettekens, newlines of één lange alinea.
    """
    raw = clean_value(value)
    if raw == "-":
        return OPDRACHTBEVESTIGING_DEFAULT_VOORWAARDEN

    normalized = raw.replace("\r", "\n").strip()
    parts = re.split(r"\s*•\s*", normalized)
    parts = [part.strip(" \n\t-") for part in parts if part.strip(" \n\t-")]

    if len(parts) <= 1:
        line_parts = [part.strip(" \n\t-") for part in normalized.split("\n") if part.strip(" \n\t-")]
        if len(line_parts) > 1:
            parts = line_parts

    return parts if parts else OPDRACHTBEVESTIGING_DEFAULT_VOORWAARDEN


def make_opdrachtbevestiging_table(source_fields: dict, styles):
    voorwaarden_raw = pick(source_fields, "Voorwaarden", "voorwaarden", default="")
    voorwaarden = split_voorwaarden(voorwaarden_raw)

    bullet_style = ParagraphStyle(
        name="BulletSmallEV",
        parent=styles["Small"],
        leftIndent=7,
        firstLineIndent=-7,
        leading=11,
        spaceAfter=4,
    )
    bullet_flowables = [
        Paragraph("• " + pdf_escape(item), bullet_style)
        for item in voorwaarden
    ]

    uw_adviseur = clean_value(
        pick(source_fields, "Uw adviseur", "Adviseur", "Naam adviseur", default=env("DEFAULT_ADVISOR", "Otto Boender"))
    )
    opnamedatum = clean_value(
        pick(source_fields, "Opnamedatum", "Opname datum", "Datum opname", default="-")
    )

    table_data = [
        [Paragraph("Omschrijving", styles["SmallBold"]), Paragraph("Waarde", styles["SmallBold"])],
        [Paragraph("Voorwaarden", styles["Small"]), bullet_flowables],
        [Paragraph("Uw adviseur", styles["Small"]), Paragraph(pdf_escape(uw_adviseur), styles["Small"])],
        [Paragraph("Opnamedatum", styles["Small"]), Paragraph(pdf_escape(opnamedatum), styles["Small"])],
    ]
    table = Table(table_data, colWidths=[95 * mm, 81 * mm], repeatRows=1)
    table.setStyle(TableStyle([
        ("BOX", (0, 0), (-1, -1), 0.35, colors.HexColor("#DDDDDD")),
        ("INNERGRID", (0, 0), (-1, -1), 0.25, colors.HexColor("#EEEEEE")),
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#F2F2F2")),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("LEFTPADDING", (0, 0), (-1, -1), 8),
        ("RIGHTPADDING", (0, 0), (-1, -1), 8),
        ("TOPPADDING", (0, 0), (-1, -1), 6),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
    ]))
    return table, uw_adviseur, opnamedatum


def make_pdf(document_type: str, payload: dict, source_fields: dict, output_path: Path) -> dict:
    cfg = DOCUMENT_CONFIG[document_type]

    address = clean_value(
        payload.get("adres")
        or payload.get("address")
        or pick(source_fields, "Adres (vol)", "Volledig adres", "Adresregel", "Address", "Adres")
    )
    opdrachtgever = clean_value(
        payload.get("opdrachtgever")
        or pick(source_fields, "Opdrachtgever", "Klant", "Naam opdrachtgever")
    )

    if document_type == "opdrachtbevestiging":
        adviseur = clean_value(
            payload.get("adviseur")
            or pick(source_fields, "Uw adviseur", "Adviseur", "Naam adviseur", default=env("DEFAULT_ADVISOR", "Otto Boender"))
        )
    else:
        adviseur = clean_value(
            payload.get("adviseur")
            or pick(source_fields, "Adviseur", "Naam adviseur", default=env("DEFAULT_ADVISOR", "Otto Boender"))
        )

    datum = clean_value(
        payload.get("datum")
        or pick(source_fields, "Datum", "Gegenereerd op", default=now_nl())
    )

    doc = SimpleDocTemplate(
        str(output_path),
        pagesize=A4,
        leftMargin=16 * mm,
        rightMargin=16 * mm,
        topMargin=14 * mm,
        bottomMargin=14 * mm,
    )
    styles = getSampleStyleSheet()
    styles.add(ParagraphStyle(name="Small", fontName="Helvetica", fontSize=8.5, leading=11))
    styles.add(ParagraphStyle(name="SmallBold", fontName="Helvetica-Bold", fontSize=8.5, leading=11))
    styles.add(ParagraphStyle(name="Meta", fontName="Helvetica", fontSize=9.5, leading=13, textColor=colors.HexColor("#444444")))
    styles.add(ParagraphStyle(name="TitleEV", fontName="Helvetica-Bold", fontSize=16, leading=19))
    styles.add(ParagraphStyle(name="CenterSmall", fontName="Helvetica", fontSize=8, alignment=TA_CENTER))

    story = []
    story.append(Paragraph(cfg["title"], styles["TitleEV"]))
    story.append(Paragraph(cfg["subtitle"], styles["Meta"]))
    story.append(Spacer(1, 8 * mm))

    header_data = [
        [Paragraph("Adres", styles["SmallBold"]), Paragraph(pdf_escape(address), styles["Small"])],
        [Paragraph("Opdrachtgever", styles["SmallBold"]), Paragraph(pdf_escape(opdrachtgever), styles["Small"])],
        [Paragraph("Datum", styles["SmallBold"]), Paragraph(pdf_escape(datum), styles["Small"])],
    ]
    header = Table(header_data, colWidths=[48 * mm, 128 * mm])
    header.setStyle(TableStyle([
        ("BOX", (0, 0), (-1, -1), 0.35, colors.HexColor("#DDDDDD")),
        ("INNERGRID", (0, 0), (-1, -1), 0.25, colors.HexColor("#EEEEEE")),
        ("BACKGROUND", (0, 0), (0, -1), colors.HexColor("#F7F7F7")),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("LEFTPADDING", (0, 0), (-1, -1), 8),
        ("RIGHTPADDING", (0, 0), (-1, -1), 8),
        ("TOPPADDING", (0, 0), (-1, -1), 6),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
    ]))
    story.append(header)
    story.append(Spacer(1, 7 * mm))

    # Belangrijk: opdrachtbevestiging krijgt NIET de Bijlage-4 informatietabel.
    # Alleen Bijlage 4 gebruikt build_items() met bouwkundige tekeningen enz.
    opnamedatum = "-"
    if document_type == "opdrachtbevestiging":
        opdracht_table, adviseur_from_table, opnamedatum = make_opdrachtbevestiging_table(source_fields, styles)
        adviseur = clean_value(payload.get("adviseur") or adviseur_from_table)
        story.append(opdracht_table)
        story.append(Spacer(1, 8 * mm))
    else:
        items = build_items(payload, source_fields)
        table_data = [[Paragraph("Omschrijving", styles["SmallBold"]), Paragraph("Aanwezig / waarde", styles["SmallBold"])]]
        for label, value in items:
            table_data.append([Paragraph(pdf_escape(label), styles["Small"]), Paragraph(pdf_escape(clean_value(value)), styles["Small"])])

        info_table = Table(table_data, colWidths=[95 * mm, 81 * mm], repeatRows=1)
        info_table.setStyle(TableStyle([
            ("BOX", (0, 0), (-1, -1), 0.35, colors.HexColor("#DDDDDD")),
            ("INNERGRID", (0, 0), (-1, -1), 0.25, colors.HexColor("#EEEEEE")),
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#F2F2F2")),
            ("VALIGN", (0, 0), (-1, -1), "TOP"),
            ("LEFTPADDING", (0, 0), (-1, -1), 8),
            ("RIGHTPADDING", (0, 0), (-1, -1), 8),
            ("TOPPADDING", (0, 0), (-1, -1), 6),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
        ]))
        story.append(info_table)
        story.append(Spacer(1, 8 * mm))

    if document_type == "opdrachtbevestiging":
        signature_value = (
            payload.get("signature_url")
            or payload.get("signature_path")
            or payload.get("paraaf")
            or pick_raw(source_fields, "Akkoord", "Paraaf", "Signature", "Handtekening", "Paraaf PNG", "Paraaf png")
            or env("SIGNATURE_URL")
            or env("SIGNATURE_PATH")
        )
    else:
        signature_value = (
            payload.get("signature_url")
            or payload.get("signature_path")
            or payload.get("paraaf")
            or pick_raw(source_fields, "Paraaf", "Signature", "Handtekening", "Paraaf PNG", "Paraaf png")
            or env("SIGNATURE_URL")
            or env("SIGNATURE_PATH")
        )

    signature_image = make_signature_image(signature_value)

    signature_data = [
        [Paragraph("Adviseur", styles["SmallBold"]), Paragraph("Paraaf", styles["SmallBold"])],
        [Paragraph(pdf_escape(adviseur), styles["Small"]), signature_image],
    ]
    signature = Table(signature_data, colWidths=[80 * mm, 96 * mm], rowHeights=[8 * mm, 24 * mm])
    signature.setStyle(TableStyle([
        ("BOX", (0, 0), (-1, -1), 0.35, colors.HexColor("#DDDDDD")),
        ("LINEBELOW", (0, 0), (-1, 0), 0.25, colors.HexColor("#EEEEEE")),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("LEFTPADDING", (0, 0), (-1, -1), 8),
        ("RIGHTPADDING", (0, 0), (-1, -1), 8),
        ("TOPPADDING", (0, 0), (-1, -1), 6),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
    ]))
    story.append(signature)

    doc.build(story)

    return {
        "address": address,
        "opdrachtgever": opdrachtgever,
        "adviseur": adviseur,
        "datum": datum,
        "opnamedatum": opnamedatum,
        "signature_included": bool(signature_image),
        "version": APP_VERSION,
    }


# -----------------------------------------------------------------------------
# File hosting / upload
# -----------------------------------------------------------------------------


def upload_or_host_pdf(pdf_path: Path, filename: str) -> str:
    """
    Voorkeur: Cloudinary raw upload als CLOUDINARY_URL is ingesteld.
    Fallback: serve vanuit deze Render-app via /files/<filename>.
    """
    if env("CLOUDINARY_URL") and cloudinary is not None:
        upload = cloudinary.uploader.upload(
            str(pdf_path),
            resource_type="raw",
            public_id=f"softr-documenten/{Path(filename).stem}_{uuid.uuid4().hex[:8]}",
            use_filename=True,
            unique_filename=False,
            overwrite=True,
        )
        return upload["secure_url"]

    public_base = get_public_base_url()
    return f"{public_base}/files/{pdf_path.name}"


@app.route("/files/<path:filename>", methods=["GET"])
def serve_generated_file(filename):
    safe_name = Path(filename).name
    file_path = TMP_DIR / safe_name
    if not file_path.exists():
        return jsonify({"error": "file_not_found"}), 404
    mimetype = mimetypes.guess_type(str(file_path))[0] or "application/pdf"
    return send_file(file_path, mimetype=mimetype, as_attachment=False, download_name=safe_name)


# -----------------------------------------------------------------------------
# Nieuwe endpoints: Bijlage 4 en Opdrachtbevestiging
# -----------------------------------------------------------------------------


def generate_document(document_type: str):
    if document_type not in DOCUMENT_CONFIG:
        return jsonify({"error": "unknown_document_type", "allowed": list(DOCUMENT_CONFIG.keys())}), 400

    payload = get_json_payload()
    require_secret(payload)
    cfg = DOCUMENT_CONFIG[document_type]

    source_fields = {}
    source_record_id = payload.get("source_record_id") or request.args.get("source_record_id")
    if source_record_id:
        source_table_id = env(cfg["source_table_env"])
        source_record = softr_get_record(source_table_id, source_record_id, field_names=True)
        source_fields.update(source_record.get("fields") or {})

    # Extra velden uit Call API body winnen altijd van opgehaalde source record velden.
    if isinstance(payload.get("fields"), dict):
        source_fields.update(payload["fields"])

    # Ook top-level waarden meenemen voor makkelijke Softr mapping.
    for key, value in payload.items():
        if key not in INTERNAL_KEYS:
            source_fields[key] = value

    dossier_record_id = (
        payload.get("dossier_record_id")
        or payload.get("target_record_id")
        or first_record_id(pick(source_fields, "dossier_record_id", "Dossier record ID", "Dossier Record ID", "Dossier", "Dossier ID"))
    )

    address_for_filename = clean_value(
        payload.get("adres")
        or payload.get("address")
        or pick(source_fields, "Adres (vol)", "Volledig adres", "Adresregel", "Address", "Adres", default="document")
    )
    filename = f"{cfg['filename_prefix']}_{safe_filename(address_for_filename)}_{uuid.uuid4().hex[:8]}.pdf"
    pdf_path = TMP_DIR / filename

    meta = make_pdf(document_type, payload, source_fields, pdf_path)
    file_url = upload_or_host_pdf(pdf_path, filename)

    target_field_id = env(cfg["target_field_env"])

    # Standaard asynchroon: Softr krijgt direct succes terug, terwijl de PDF op de achtergrond
    # naar het juiste file field in Dossiers wordt geschreven.
    async_update = truthy_env("SOFTR_UPDATE_ASYNC", "true")
    if async_update:
        worker = threading.Thread(
            target=softr_patch_dossier_file_background,
            args=(dossier_record_id, target_field_id, file_url, filename, document_type),
            daemon=True,
        )
        worker.start()
        return jsonify({
            "status": "queued",
            "document_type": document_type,
            "filename": filename,
            "file_url": file_url,
            "dossier_record_id": dossier_record_id,
            "target_field_id": target_field_id,
            "softr_update": "background_started",
            "meta": meta,
        })

    update_result = softr_patch_dossier_file(dossier_record_id, target_field_id, file_url, filename)
    return jsonify({
        "status": "ok",
        "document_type": document_type,
        "filename": filename,
        "file_url": file_url,
        "dossier_record_id": dossier_record_id,
        "target_field_id": target_field_id,
        "softr_update_mode": update_result.get("mode"),
        "softr_update_assumed_success": update_result.get("assumed_success", False),
        "meta": meta,
    })


@app.route("/generate/bijlage4", methods=["POST"])
def route_generate_bijlage4():
    return generate_document("bijlage4")


@app.route("/generate/opdrachtbevestiging", methods=["POST"])
def route_generate_opdrachtbevestiging():
    return generate_document("opdrachtbevestiging")


@app.route("/generate-document", methods=["POST"])
def route_generate_document_generic():
    payload = get_json_payload()
    document_type = payload.get("document_type") or request.args.get("document_type")
    return generate_document(str(document_type or "").lower())


# -----------------------------------------------------------------------------
# Oude Uniec3 endpoint behouden, zodat bestaande Softr/Zapier koppelingen niet breken.
# -----------------------------------------------------------------------------


@app.route("/generate", methods=["GET", "POST"])
@app.route("/generate-uniec3", methods=["GET", "POST"])
def generate_uniec3():
    if request.method == "POST":
        data = get_json_payload()
        address = data.get("address")
        height = data.get("height")
        bouwjaar = data.get("bouwjaar")
        pand_id = data.get("pand_id")
        gebruiksoppervlakte = data.get("gebruiksoppervlakte") or data.get("go")
    else:
        address = request.args.get("address")
        height = request.args.get("height")
        bouwjaar = request.args.get("bouwjaar")
        pand_id = request.args.get("pand_id")
        gebruiksoppervlakte = request.args.get("gebruiksoppervlakte") or request.args.get("go")

    if not address:
        return jsonify({"error": "address ontbreekt"}), 400
    if not GENERATOR.exists() or not TEMPLATE.exists():
        return jsonify({
            "error": "uniec_generator_missing",
            "detail": "generator.py en/of template.uniec3 staan niet in deze Render repo.",
        }), 500

    out_name = safe_filename(address) + "_" + str(uuid.uuid4()) + ".uniec3"
    output_file = Path("/tmp") / out_name

    cmd = [
        "python3",
        str(GENERATOR),
        "--template", str(TEMPLATE),
        "--address", address,
        "--output", str(output_file),
    ]

    if height not in (None, ""):
        cmd += ["--height", str(height)]
    if bouwjaar not in (None, ""):
        cmd += ["--bouwjaar", str(bouwjaar)]
    if pand_id not in (None, ""):
        cmd += ["--pand-id", str(pand_id)]
    if gebruiksoppervlakte not in (None, ""):
        cmd += ["--gebruiksoppervlakte", str(gebruiksoppervlakte)]

    try:
        result = subprocess.run(cmd, check=True, capture_output=True, text=True, timeout=60)
    except subprocess.CalledProcessError as e:
        return jsonify({
            "error": "generator_failed",
            "returncode": e.returncode,
            "stdout": e.stdout[-4000:] if e.stdout else "",
            "stderr": e.stderr[-4000:] if e.stderr else "",
            "cmd": cmd,
        }), 500
    except Exception as e:
        return jsonify({"error": "server_error", "detail": str(e)}), 500

    if not output_file.exists():
        return jsonify({
            "error": "output_missing",
            "stdout": result.stdout[-4000:] if result.stdout else "",
            "stderr": result.stderr[-4000:] if result.stderr else "",
        }), 500

    download_name = safe_filename(address) + ".uniec3"
    return send_file(output_file, as_attachment=True, download_name=download_name)


@app.route("/", methods=["GET"])
def home():
    return jsonify({
        "status": "ok",
        "message": "Softr documentgenerator draait",
        "version": APP_VERSION,
        "endpoints": {
            "bijlage4": "POST /generate/bijlage4",
            "opdrachtbevestiging": "POST /generate/opdrachtbevestiging",
            "generic": "POST /generate-document met document_type",
            "uniec3_legacy": "GET/POST /generate",
        },
    })


@app.route("/health", methods=["GET"])
def health():
    return jsonify({"status": "ok", "version": APP_VERSION})


if __name__ == "__main__":
    port = int(os.environ.get("PORT", "8000"))
    app.run(host="0.0.0.0", port=port)
