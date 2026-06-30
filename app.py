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
from reportlab.pdfgen import canvas
from reportlab.graphics.barcode import qr
from reportlab.graphics import renderPDF
from reportlab.graphics.shapes import Drawing
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
APP_VERSION = "v12_bag_afschrift_layout_full_20260630"


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
    "Plattegronden",
    "Gevelaanzichten",
    "Doorsnedes",
    "Constructietekeningen",
    "Installatieontwerp van het gebouw",
    "Installatietekeningen voor verwarming",
    "Installatietekeningen voor tapwater",
    "Installatietekeningen voor koeling",
    "Inregeling verwarming conform protocol",
    "Verzamellijsten opwekkers",
    "Verzamellijsten met type kozijnen en beglazingen",
    "Aanvraag bouwvergunning",
    "Verleende bouwvergunning",
]


# Keuzevelden uit Softr kunnen soms als interne UUID binnenkomen.
# Hiermee tonen we in de PDF de leesbare dropdown-waarde in plaats van een nieuwe losse rij.
OPTION_VALUE_LABELS = {
    "45ad207d-14ac-4c46-966b-c56c852eaf1f": "Revisie",
    "e320ced2-e7a7-4cb8-bd40-fe39e77ad5ed": "Bestek",
    "64c6112a-3ec8-4cf8-ad2e-80be9a68d50b": "Toets Bbl",
}

# Deze velden horen inhoudelijk bij bestaande regels in Bijlage 4.
# Ze mogen dus niet als extra losse rij onderaan de PDF verschijnen.
INFO_FIELD_ALIASES = {
    "met plattegronden": "Plattegronden",
    "plattegronden aanwezig": "Plattegronden",
    "type plattegronden": "Plattegronden",
    "dwarsdoorsnede": "Doorsnedes",
    "dwarsdoorsneden": "Doorsnedes",
    "met doorsnedes": "Doorsnedes",
    "met dwarsdoorsnede": "Doorsnedes",
    "detailtekeningen bouwkundige constructies": "Constructietekeningen",
    "bouwkundige constructies": "Constructietekeningen",
}

INSTALLATIE_FIELD_NAMES = {
    "installaties",
    "installatietekeningen",
    "met installatietekeningen",
    "installatietekeningen aanwezig",
    "type installatietekeningen",
}

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


def display_choice_value(value):
    """Toon Softr-keuzes leesbaar en vervang bekende UUID's door labels."""
    if value is None:
        return "-"
    if isinstance(value, list):
        parts = [display_choice_value(item) for item in value]
        parts = [part for part in parts if part and part != "-"]
        return ", ".join(parts) if parts else "-"
    if isinstance(value, dict):
        raw = value.get("name") or value.get("label") or value.get("value") or value.get("id")
        return display_choice_value(raw)

    raw = clean_value(value)
    if raw == "-":
        return "-"

    parts = [part.strip() for part in str(raw).split(",") if part.strip()]
    converted = [OPTION_VALUE_LABELS.get(part, part) for part in parts]
    return ", ".join(converted) if converted else "-"


def find_source_value(source_fields: dict, *field_names):
    """Pak een waarde op basis van veldnamen, case-insensitive."""
    if not isinstance(source_fields, dict):
        return None
    wanted = {str(name).strip().lower() for name in field_names if str(name).strip()}
    for key, value in source_fields.items():
        if str(key).strip().lower() in wanted and clean_value(value) != "-":
            return value
    return None


def label_has_choice(source_fields: dict, *choices) -> bool:
    """Controleer of een multi-select veld één van de gevraagde keuzes bevat."""
    wanted = {choice.lower() for choice in choices}
    for key, value in source_fields.items():
        if str(key).strip().lower() not in INSTALLATIE_FIELD_NAMES:
            continue
        readable = display_choice_value(value).lower()
        if any(choice in readable for choice in wanted):
            return True
    return False


def bijlage4_value_for_label(label: str, source_fields: dict) -> str:
    """Bepaal de rechterkolomwaarde voor een vaste Bijlage 4-regel."""
    direct = pick(source_fields, label, default="-")
    if clean_value(direct) != "-":
        return display_choice_value(direct)

    label_norm = label.strip().lower()

    alias_names = [
        key for key, target_label in INFO_FIELD_ALIASES.items()
        if target_label.strip().lower() == label_norm
    ]
    alias_value = find_source_value(source_fields, *alias_names)
    if alias_value is not None:
        return display_choice_value(alias_value)

    if label_norm == "installatietekeningen voor verwarming" and label_has_choice(source_fields, "verwarming"):
        return "Verwarming"
    if label_norm == "installatietekeningen voor tapwater" and label_has_choice(source_fields, "tapwater"):
        return "Tapwater"
    if label_norm == "installatietekeningen voor koeling" and label_has_choice(source_fields, "koeling"):
        return "Koeling"

    return "-"


def should_skip_extra_info_field(key: str) -> bool:
    """Voorkom dubbele losse rijen voor velden die al in vaste regels verwerkt zijn."""
    key_norm = str(key or "").strip().lower()
    if should_skip_info_field(key_norm):
        return True
    if key_norm in {label.lower() for label in DEFAULT_INFO_ROWS}:
        return True
    if key_norm in INFO_FIELD_ALIASES:
        return True
    if key_norm in INSTALLATIE_FIELD_NAMES:
        return True
    return False


def build_items(payload: dict, source_fields: dict) -> list:
    """Bouw de informatie-rijen voor Bijlage 4 / opdrachtbevestiging."""
    items = payload.get("items")

    if isinstance(items, list) and items:
        item_values = {}
        extra_rows = []
        for item in items:
            if isinstance(item, dict):
                label = str(item.get("label") or item.get("name") or item.get("field") or "-")
                if should_skip_info_field(label):
                    continue
                value = item.get("value") if "value" in item else item.get("status", "-")
                target_label = INFO_FIELD_ALIASES.get(label.strip().lower(), label)
                item_values[target_label.strip().lower()] = display_choice_value(value)
            elif isinstance(item, str):
                if should_skip_info_field(item):
                    continue
                target_label = INFO_FIELD_ALIASES.get(item.strip().lower(), item)
                item_values[target_label.strip().lower()] = display_choice_value(source_fields.get(item))

        default_rows = []
        for label in DEFAULT_INFO_ROWS:
            value = item_values.get(label.lower()) or bijlage4_value_for_label(label, source_fields)
            default_rows.append((label, value))

        for label_norm, value in item_values.items():
            if label_norm not in {label.lower() for label in DEFAULT_INFO_ROWS} and value != "-":
                extra_rows.append((label_norm, value))
        return default_rows + extra_rows[:20]

    if isinstance(items, dict) and items:
        source_fields = {**source_fields, **items}

    # Vaste basisregels bovenaan. Keuzevelden worden in de juiste vaste rij geplaatst.
    default_rows = []
    for label in DEFAULT_INFO_ROWS:
        default_rows.append((label, bijlage4_value_for_label(label, source_fields)))

    # Extra velden die niet technisch zijn en niet al in de vaste regels zitten.
    extra_rows = []
    for key, value in source_fields.items():
        key_str = str(key)
        if should_skip_extra_info_field(key_str):
            continue
        readable_value = display_choice_value(value)
        if readable_value == "-":
            continue
        extra_rows.append((key_str, readable_value))

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
    table = Table(table_data, colWidths=[48 * mm, 128 * mm], repeatRows=1)
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

    if document_type == "opdrachtbevestiging":
        signature_data = [
            [Paragraph("Opdrachtgever", styles["SmallBold"]), Paragraph("Paraaf", styles["SmallBold"])],
            [Paragraph(pdf_escape(opdrachtgever), styles["Small"]), signature_image],
        ]
    else:
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
# BAG-afschrift generatie
# -----------------------------------------------------------------------------

BAG_API_BASE = "https://api.bag.kadaster.nl/lvbag/individuelebevragingen/v2"
BAG_VBO_ENDPOINT = f"{BAG_API_BASE}/verblijfsobjecten/{{id}}"


def first_non_empty(*values, default="-"):
    for value in values:
        cleaned = clean_value(value)
        if cleaned != "-":
            return cleaned
    return default


def clean_bag_id(value) -> str:
    """Accepteert Softr-strings, arrays en objecten en haalt een 16-cijferig BAG-ID eruit."""
    if value is None:
        return ""
    if isinstance(value, list):
        for item in value:
            found = clean_bag_id(item)
            if found:
                return found
        return ""
    if isinstance(value, dict):
        for key in ("id", "value", "identificatie", "adresseerbaarobject_id", "adresseerbaar_object_id"):
            found = clean_bag_id(value.get(key))
            if found:
                return found
        return ""
    text = str(value).strip()
    match = re.search(r"\d{16}", text)
    return match.group(0) if match else ""


def format_m2(value):
    cleaned = clean_value(value)
    if cleaned == "-":
        return "-"
    text = str(cleaned)
    if "m²" in text or "m2" in text.lower():
        return text
    return f"{text} m²"


def bag_get_nested(data, *path, default=""):
    current = data
    for part in path:
        if isinstance(current, dict):
            current = current.get(part)
        elif isinstance(current, list) and isinstance(part, int) and 0 <= part < len(current):
            current = current[part]
        else:
            return default
    return current if current is not None else default


def extract_bag_identificatie_from_href(href: str) -> str:
    if not href:
        return ""
    match = re.search(r"/(verblijfsobjecten|nummeraanduidingen|panden|openbare-ruimten|woonplaatsen)/([^/?#]+)", str(href))
    return match.group(2) if match else ""


def bag_headers():
    api_key = env("BAG_API_KEY") or env("KADASTER_BAG_API_KEY")
    if not api_key:
        return None
    return {
        "X-Api-Key": api_key,
        "Accept": "application/hal+json",
        "Accept-Crs": "epsg:28992",
    }


def bag_api_get(url: str) -> dict:
    headers = bag_headers()
    if not headers:
        return {}
    r = requests.get(url, headers=headers, timeout=30)
    if not r.ok:
        raise RuntimeError(f"BAG API ophalen mislukt: {r.status_code} {r.text[:1000]}")
    return r.json()


def bag_api_get_object(kind: str, object_id: str) -> dict:
    object_id = clean_bag_id(object_id)
    if not object_id:
        return {}
    return bag_api_get(f"{BAG_API_BASE}/{kind}/{object_id}")


def fetch_bag_verblijfsobject(adresseerbaar_object_id: str) -> dict:
    """Haal verblijfsobject + gekoppelde nummeraanduiding + pand(en) op."""
    vbo_id = clean_bag_id(adresseerbaar_object_id)
    if not bag_headers() or not vbo_id:
        return {}

    result = {"verblijfsobject_response": bag_api_get(BAG_VBO_ENDPOINT.format(id=vbo_id))}
    vbo_response = result["verblijfsobject_response"]

    # Nummeraanduiding via hoofdadres/nevenadres-link.
    nummeraanduiding_id = first_non_empty(
        extract_bag_identificatie_from_href(bag_get_nested(vbo_response, "_links", "heeftAlsHoofdAdres", "href", default="")),
        extract_bag_identificatie_from_href(bag_get_nested(vbo_response, "_links", "heeftAlsNevenAdres", "href", default="")),
        default="",
    )
    if nummeraanduiding_id:
        result["nummeraanduiding_response"] = bag_api_get_object("nummeraanduidingen", nummeraanduiding_id)

    # Pand(en): soms embedded, soms alleen links.
    pand_ids = []
    embedded_panden = bag_get_nested(vbo_response, "_embedded", "panden", default=[])
    if isinstance(embedded_panden, list):
        for item in embedded_panden:
            pid = clean_bag_id(bag_get_nested(item, "pand", "identificatie", default=""))
            if not pid:
                pid = clean_bag_id(extract_bag_identificatie_from_href(bag_get_nested(item, "_links", "self", "href", default="")))
            if pid:
                pand_ids.append(pid)

    links = bag_get_nested(vbo_response, "_links", "maaktDeelUitVan", default=[])
    if isinstance(links, dict):
        links = [links]
    if isinstance(links, list):
        for link in links:
            pid = clean_bag_id(extract_bag_identificatie_from_href(link.get("href") if isinstance(link, dict) else ""))
            if pid:
                pand_ids.append(pid)

    result["pand_responses"] = []
    for pid in dict.fromkeys(pand_ids):
        try:
            result["pand_responses"].append(bag_api_get_object("panden", pid))
        except Exception as e:
            app.logger.warning("Pand %s kon niet worden opgehaald: %s", pid, e)

    return result


def unwrap_bag_object(response: dict, object_key: str) -> dict:
    if not isinstance(response, dict):
        return {}
    obj = response.get(object_key)
    return obj if isinstance(obj, dict) else response


def fetch_linked_name(response: dict, rel: str, object_key: str, name_keys: tuple) -> str:
    href = bag_get_nested(response, "_links", rel, "href", default="")
    if not href or not bag_headers():
        return ""
    try:
        data = bag_api_get(href if str(href).startswith("http") else f"{BAG_API_BASE}{href}")
        obj = unwrap_bag_object(data, object_key)
        return first_non_empty(*[obj.get(k) for k in name_keys], default="")
    except Exception as e:
        app.logger.warning("Gekoppelde BAG-link %s kon niet worden opgehaald: %s", rel, e)
        return ""


def format_date_bag(value):
    cleaned = clean_value(value)
    if cleaned == "-":
        return "-"
    text = str(cleaned)[:10]
    if re.match(r"^\d{4}-\d{2}-\d{2}$", text):
        y, m, d = text.split("-")
        return f"{d}-{m}-{y}"
    return cleaned


def flatten_bag_data(payload: dict, source_fields: dict, api_data: dict) -> dict:
    """Maak één nette dictionary voor het BAG-afschrift uit Softr + BAG API."""
    vbo_response = api_data.get("verblijfsobject_response") if isinstance(api_data, dict) else {}
    # Backwards compatible met oude fetch die direct de vbo-response teruggaf.
    if not vbo_response and isinstance(api_data, dict) and ("verblijfsobject" in api_data or "_links" in api_data):
        vbo_response = api_data
    nummer_response = api_data.get("nummeraanduiding_response", {}) if isinstance(api_data, dict) else {}
    pand_responses = api_data.get("pand_responses", []) if isinstance(api_data, dict) else []

    vbo = unwrap_bag_object(vbo_response, "verblijfsobject")
    nummer = unwrap_bag_object(nummer_response, "nummeraanduiding")
    pand_objects = [unwrap_bag_object(p, "pand") for p in pand_responses if isinstance(p, dict)]
    pand0 = pand_objects[0] if pand_objects else {}

    # Als oude API-response embedded panden bevatte.
    if not pand0:
        embedded = vbo_response.get("_embedded", {}) if isinstance(vbo_response, dict) else {}
        panden = embedded.get("panden", []) if isinstance(embedded, dict) else []
        pand0 = panden[0].get("pand", {}) if panden and isinstance(panden[0], dict) else {}

    verblijfsobject_id = first_non_empty(
        payload.get("adresseerbaar_object_id"), payload.get("verblijfsobject_id"), payload.get("object_id"),
        pick(source_fields, "adresseerbaar_object_id", "adresseerbaarobject_id", "Adresseerbaar object ID", "Adresseerbaar object id", "verblijfsobject_id", "Verblijfsobject ID", "object_id", default=""),
        vbo.get("identificatie"),
        default="-",
    )
    verblijfsobject_id = clean_bag_id(verblijfsobject_id) or clean_value(verblijfsobject_id)

    nummeraanduiding_id = first_non_empty(
        nummer.get("identificatie"),
        pick(source_fields, "nummeraanduiding_id", "nummeraanduidingid", "Nummeraanduiding ID", "Nummeraanduiding", default=""),
        extract_bag_identificatie_from_href(bag_get_nested(vbo_response, "_links", "heeftAlsHoofdAdres", "href", default="")),
        default="-",
    )

    pand_ids = [clean_bag_id(p.get("identificatie")) for p in pand_objects if p.get("identificatie")]
    pand_id = first_non_empty(
        pick(source_fields, "pand_id", "Pand ID", "pand identificatie", "pand_identificatie", default=""),
        ", ".join([p for p in dict.fromkeys(pand_ids) if p]),
        default="-",
    )

    openbare_ruimte = first_non_empty(
        nummer.get("_embedded", {}).get("openbareRuimte", {}).get("naam") if isinstance(nummer.get("_embedded"), dict) else "",
        pick(source_fields, "straat", "Straat", "openbare_ruimte", "Openbare ruimte", "straatnaam", "straatnaam_verkort", default=""),
        fetch_linked_name(nummer_response, "ligtAanOpenbareRuimte", "openbareRuimte", ("naam", "verkorteNaam")),
        default="-",
    )
    woonplaats = first_non_empty(
        payload.get("woonplaats"), payload.get("plaats"),
        pick(source_fields, "woonplaats", "Woonplaats", "plaats", "Plaats", "woonplaatsnaam", default=""),
        fetch_linked_name(nummer_response, "ligtInWoonplaats", "woonplaats", ("naam",)),
        default="-",
    )
    gemeente = first_non_empty(
        pick(source_fields, "gemeente", "Gemeente", "gemeentenaam", default=""),
        default="-",
    )

    huisnummer = first_non_empty(nummer.get("huisnummer"), pick(source_fields, "huisnummer", "Huisnummer", default=""), default="-")
    huisletter = first_non_empty(nummer.get("huisletter"), pick(source_fields, "huisletter", "Huisletter", default=""), default="-")
    toevoeging = first_non_empty(nummer.get("huisnummertoevoeging"), pick(source_fields, "toevoeging", "huisnummertoevoeging", "Huisnummertoevoeging", default=""), default="-")
    postcode = first_non_empty(payload.get("postcode"), nummer.get("postcode"), pick(source_fields, "postcode", "Postcode", default=""), default="-")

    adres_parts = [openbare_ruimte, huisnummer]
    if clean_value(huisletter) != "-":
        adres_parts.append(huisletter)
    if clean_value(toevoeging) != "-":
        adres_parts.append(toevoeging)
    adres_from_parts = " ".join([clean_value(x) for x in adres_parts if clean_value(x) != "-"])
    adres = first_non_empty(payload.get("adres"), payload.get("address"), pick(source_fields, "Adres (vol)", "Volledig adres", "Adresregel", "Adres", "Address", "weergavenaam", default=""), adres_from_parts, default="-")

    gebruiksdoelen = vbo.get("gebruiksdoelen") or vbo.get("gebruiksdoel") or pick(source_fields, "gebruiksdoel", "Gebruiksdoel", default="")
    if isinstance(gebruiksdoelen, list):
        gebruiksdoelen = ", ".join(str(x) for x in gebruiksdoelen)

    begin_geldigheid_vbo = first_non_empty(
        bag_get_nested(vbo, "voorkomen", "beginGeldigheid", default=""), vbo.get("beginGeldigheid"),
        pick(source_fields, "begindatum_geldigheid", "Begindatum geldigheid", default=""),
        default="-",
    )
    begin_geldigheid_pand = first_non_empty(
        bag_get_nested(pand0, "voorkomen", "beginGeldigheid", default=""), pand0.get("beginGeldigheid"), begin_geldigheid_vbo,
        default="-",
    )

    coords = first_non_empty(
        pick(source_fields, "centroide_ll", "centroide_rd", "geometrie_ll", "geometrie_rd", default=""),
        bag_get_nested(vbo, "geometrie", "punt", "coordinates", default=""),
        default="",
    )

    return {
        "afschriftnummer": payload.get("afschriftnummer") or f"{datetime.now().strftime('%Y-%m-%d')}-{uuid.uuid4().hex[:4]}",
        "documentdatum": now_nl(),
        "tijdstip": datetime.now().strftime("%H:%M:%S"),
        "bron": "BAG API Individuele Bevragingen",
        "aanvrager": env("BAG_AANVRAGER", "Mijn Portaal B.V."),
        "adres": adres,
        "straat": openbare_ruimte,
        "openbare_ruimte": openbare_ruimte,
        "huisnummer": huisnummer,
        "huisletter": huisletter,
        "toevoeging": toevoeging,
        "postcode": postcode,
        "woonplaats": woonplaats,
        "gemeente": gemeente,
        "status_nummeraanduiding": first_non_empty(nummer.get("status"), pick(source_fields, "status_nummeraanduiding", default=""), default="-"),
        "type_nummeraanduiding": first_non_empty(nummer.get("typeAdresseerbaarObject"), "Verblijfsobjectnummer", default="-"),
        "geconstateerd_nummeraanduiding": first_non_empty(nummer.get("geconstateerd"), default="Nee"),
        "in_onderzoek_nummeraanduiding": first_non_empty(nummer.get("inOnderzoek"), default="Nee"),
        "datum_naamgeving": format_date_bag(begin_geldigheid_vbo),
        "verblijfsobject_id": verblijfsobject_id,
        "nummeraanduiding_id": clean_bag_id(nummeraanduiding_id) or clean_value(nummeraanduiding_id),
        "pand_id": pand_id,
        "bouwjaar": first_non_empty(pand0.get("oorspronkelijkBouwjaar"), pick(source_fields, "bouwjaar", "Bouwjaar", default=""), default="-"),
        "gebruiksoppervlakte": first_non_empty(payload.get("gebruiksoppervlakte"), payload.get("oppervlakte"), payload.get("go"), pick(source_fields, "gebruiksoppervlakte", "Gebruiksoppervlakte", "oppervlakte", "Oppervlakte", "go", default=""), vbo.get("oppervlakte"), default="-"),
        "gebruiksdoel": first_non_empty(gebruiksdoelen, default="-"),
        "status_verblijfsobject": first_non_empty(vbo.get("status"), pick(source_fields, "status", "Status verblijfsobject", "object_status", default=""), default="-"),
        "status_pand": first_non_empty(pand0.get("status"), pick(source_fields, "status_pand", "Status pand", default=""), default="-"),
        "pand_oppervlakte": first_non_empty(pand0.get("oppervlakte"), pick(source_fields, "pand_oppervlakte", "Pand oppervlakte", default=""), default="-"),
        "aantal_bouwlagen": first_non_empty(pand0.get("aantalBouwlagen"), pick(source_fields, "aantal_bouwlagen", "Aantal bouwlagen", default=""), default="-"),
        "aantal_verblijfsobjecten": first_non_empty(pick(source_fields, "aantal_verblijfsobjecten", "Aantal verblijfsobjecten", default=""), default="1"),
        "aantal_gebruiksdoelen": first_non_empty(str(len([x for x in str(gebruiksdoelen).split(',') if x.strip()])) if gebruiksdoelen else "", default="1"),
        "begin_geldigheid_vbo": format_date_bag(begin_geldigheid_vbo),
        "begin_geldigheid_pand": format_date_bag(begin_geldigheid_pand),
        "eind_geldigheid_vbo": "-",
        "eind_geldigheid_pand": "-",
        "coords": coords,
    }


def draw_wrapped(c, text, x, y, max_width, font="Helvetica", size=8.2, leading=11, bold=False):
    from reportlab.pdfbase.pdfmetrics import stringWidth
    text = clean_value(text)
    font_name = "Helvetica-Bold" if bold else font
    words = str(text).split()
    lines = []
    line = ""
    for word in words:
        trial = f"{line} {word}".strip()
        if stringWidth(trial, font_name, size) <= max_width or not line:
            line = trial
        else:
            lines.append(line)
            line = word
    if line:
        lines.append(line)
    c.setFont(font_name, size)
    for i, line in enumerate(lines[:3]):
        c.drawString(x, y - i * leading, line)
    return y - max(1, len(lines[:3])) * leading


def draw_section_title(c, title, x, y, w):
    blue = colors.HexColor("#00508F")
    c.setFillColor(blue)
    c.setFont("Helvetica-Bold", 9.2)
    c.drawString(x, y, title)
    c.setStrokeColor(blue)
    c.setLineWidth(0.9)
    c.line(x, y - 4, x + w, y - 4)
    c.setFillColor(colors.black)


def draw_label_value_rows(c, rows, x, y, label_w=42*mm, row_h=14, font_size=7.4, value_bold=False):
    for label, value in rows:
        c.setFillColor(colors.HexColor("#4B5563"))
        c.setFont("Helvetica", font_size)
        c.drawString(x, y, clean_value(label))
        c.setFillColor(colors.black)
        c.setFont("Helvetica-Bold" if value_bold else "Helvetica", font_size)
        c.drawString(x + label_w, y, clean_value(value))
        y -= row_h
    return y


def draw_box_table(c, title, rows, x, y, w, row_h=14, label_w=None):
    if label_w is None:
        label_w = w * 0.45
    blue = colors.HexColor("#00508F")
    border = colors.HexColor("#1F5E95")
    fill = colors.HexColor("#F3F8FC")
    c.setStrokeColor(border)
    c.setLineWidth(0.7)
    total_h = 22 + row_h * len(rows)
    c.roundRect(x, y - total_h, w, total_h, 3, stroke=1, fill=0)
    c.setFillColor(blue)
    c.setFont("Helvetica-Bold", 8.5)
    c.drawString(x + 7, y - 14, title)
    c.setStrokeColor(colors.HexColor("#D6E2ED"))
    c.line(x + 7, y - 20, x + w - 7, y - 20)
    cur_y = y - 33
    for i, (label, value) in enumerate(rows):
        if i % 2 == 1:
            c.setFillColor(fill)
            c.rect(x + 1, cur_y - 4, w - 2, row_h, stroke=0, fill=1)
        c.setFillColor(colors.HexColor("#4B5563"))
        c.setFont("Helvetica", 7.2)
        c.drawString(x + 7, cur_y, clean_value(label))
        c.setFillColor(colors.black)
        c.setFont("Helvetica-Bold" if i == 0 else "Helvetica", 7.2)
        draw_wrapped(c, value, x + label_w, cur_y, w - label_w - 8, size=7.2, leading=8.5, bold=(i == 0))
        cur_y -= row_h
    return y - total_h


def draw_placeholder_map(c, bag, x, y, w, h):
    c.setStrokeColor(colors.HexColor("#1F5E95"))
    c.setLineWidth(0.7)
    c.roundRect(x, y - h, w, h, 3, stroke=1, fill=0)
    c.setFillColor(colors.HexColor("#00508F"))
    c.setFont("Helvetica-Bold", 8.5)
    c.drawString(x + 7, y - 14, "Ligging in Nederland")
    inner_x, inner_y, inner_w, inner_h = x + 7, y - h + 7, w - 14, h - 29
    c.setFillColor(colors.HexColor("#EEF3F4"))
    c.rect(inner_x, inner_y, inner_w, inner_h, stroke=0, fill=1)
    c.setStrokeColor(colors.HexColor("#D6D6D6"))
    for i in range(6):
        c.line(inner_x, inner_y + i * inner_h / 5, inner_x + inner_w, inner_y + (i + 0.7) * inner_h / 5)
        c.line(inner_x + i * inner_w / 5, inner_y, inner_x + (i + 0.8) * inner_w / 5, inner_y + inner_h)
    # Pin
    px, py = inner_x + inner_w * 0.55, inner_y + inner_h * 0.47
    c.setFillColor(colors.HexColor("#1686C4"))
    c.circle(px, py + 8, 7, stroke=0, fill=1)
    c.setFillColor(colors.white)
    c.circle(px, py + 8, 2.5, stroke=0, fill=1)
    c.setFillColor(colors.HexColor("#1686C4"))
    pth = c.beginPath()
    pth.moveTo(px - 5, py + 4)
    pth.lineTo(px + 5, py + 4)
    pth.lineTo(px, py - 8)
    pth.close()
    c.drawPath(pth, stroke=0, fill=1)
    c.setFillColor(colors.HexColor("#666666"))
    c.setFont("Helvetica", 6.2)
    c.drawRightString(inner_x + inner_w - 3, inner_y + 3, "© OpenStreetMap contributors")


def draw_qr(c, value, x, y, size=31*mm):
    c.setStrokeColor(colors.HexColor("#8AA8C5"))
    c.roundRect(x, y - size, size, size, 3, stroke=1, fill=0)
    qr_code = qr.QrCodeWidget(value or "BAG afschrift")
    bounds = qr_code.getBounds()
    width = bounds[2] - bounds[0]
    height = bounds[3] - bounds[1]
    d = Drawing(size - 14, size - 14, transform=[(size - 14) / width, 0, 0, (size - 14) / height, 0, 0])
    d.add(qr_code)
    renderPDF.draw(d, c, x + 7, y - size + 7)


def make_bag_pdf(bag: dict, out_path: Path) -> dict:
    """Maak een BAG-afschrift in de layout van het voorbeeld."""
    c = canvas.Canvas(str(out_path), pagesize=A4)
    W, H = A4
    margin = 16 * mm
    blue = colors.HexColor("#00508F")

    # Header
    c.setFillColor(colors.HexColor("#123A63"))
    c.setFont("Helvetica-Bold", 23)
    c.drawString(margin, H - 33 * mm, "BAG AFSCHRIFT")
    c.setFillColor(colors.HexColor("#374151"))
    c.setFont("Helvetica", 11)
    c.drawString(margin, H - 43 * mm, "Basisregistratie Adressen en Gebouwen (BAG)")

    # Tekstlogo rechtsboven. Geen officieel beeldmerkbestand nodig.
    c.setFillColor(colors.HexColor("#0085B2"))
    c.setFont("Helvetica-Oblique", 9)
    c.drawString(W - 50 * mm, H - 30 * mm, "kadaster")
    c.setFillColor(colors.HexColor("#0077A8"))
    p = c.beginPath()
    p.moveTo(W - 35 * mm, H - 27 * mm)
    p.lineTo(W - 23 * mm, H - 27 * mm)
    p.lineTo(W - 44 * mm, H - 58 * mm)
    p.close()
    c.drawPath(p, stroke=0, fill=1)
    c.setFillColor(colors.HexColor("#00508F"))
    p = c.beginPath()
    p.moveTo(W - 20 * mm, H - 27 * mm)
    p.lineTo(W - 15 * mm, H - 27 * mm)
    p.lineTo(W - 31 * mm, H - 58 * mm)
    p.lineTo(W - 36 * mm, H - 58 * mm)
    p.close()
    c.drawPath(p, stroke=0, fill=1)

    # Linker adresblok
    left_x = margin
    top_y = H - 62 * mm
    draw_section_title(c, "Adres", left_x, top_y, 86 * mm)
    adres_rows = [
        ("Adres", bag.get("adres")),
        ("Postcode", bag.get("postcode")),
        ("Woonplaats", bag.get("woonplaats")),
        ("Gemeente", bag.get("gemeente")),
        ("Openbare ruimte", bag.get("openbare_ruimte")),
        ("Huisnummer", bag.get("huisnummer")),
        ("Huisletter", bag.get("huisletter")),
        ("Huisnummertoevoeging", bag.get("toevoeging")),
        ("Status", bag.get("status_nummeraanduiding")),
        ("Datum naamgeving", bag.get("datum_naamgeving")),
        ("Datum beëindiging", "-"),
    ]
    draw_label_value_rows(c, adres_rows, left_x, top_y - 16, label_w=33 * mm, row_h=13.2, font_size=7.4, value_bold=True)

    # Rechter metadata + kaart
    meta_x = W - margin - 89 * mm
    meta_y = top_y + 1
    meta_rows = [
        ("Afschriftnummer", bag.get("afschriftnummer")),
        ("Datum afschrift", bag.get("documentdatum")),
        ("Tijdstip", bag.get("tijdstip")),
        ("Bron", bag.get("bron")),
        ("Aanvrager", bag.get("aanvrager")),
    ]
    draw_label_value_rows(c, meta_rows, meta_x, meta_y, label_w=36 * mm, row_h=14, font_size=7.3, value_bold=True)
    draw_placeholder_map(c, bag, meta_x, H - 118 * mm, 89 * mm, 54 * mm)

    # Nummeraanduiding breed
    y = H - 237 * mm
    # Actually use absolute placement similar to sample.
    num_y = H - 123 * mm
    draw_box_table(c, "Nummeraanduiding", [
        ("Nummeraanduiding ID", bag.get("nummeraanduiding_id")),
        ("Type nummeraanduiding", bag.get("type_nummeraanduiding")),
        ("Status", bag.get("status_nummeraanduiding")),
        ("Geconstateerd", bag.get("geconstateerd_nummeraanduiding")),
        ("Begindatum geldigheid", bag.get("begin_geldigheid_vbo")),
        ("Einddatum geldigheid", "-"),
        ("In onderzoek", bag.get("in_onderzoek_nummeraanduiding")),
    ], margin, num_y, W - 2 * margin, row_h=12.7, label_w=55 * mm)

    # Onderste kolommen
    col_y = H - 174 * mm
    col_w = (W - 2 * margin - 6 * mm) / 2
    draw_box_table(c, "Adresseerbaar object", [
        ("Adresseerbaar object ID", bag.get("verblijfsobject_id")),
        ("Type adresseerbaar object", "Verblijfsobject"),
        ("Status", bag.get("status_verblijfsobject")),
        ("Gebruiksdoel", bag.get("gebruiksdoel")),
        ("Oppervlakte gebruiksdoel (m²)", clean_value(bag.get("gebruiksoppervlakte")).replace(" m²", "")),
        ("Vloeroppervlakte (m²)", clean_value(bag.get("gebruiksoppervlakte")).replace(" m²", "")),
        ("Aantal verblijfsobjecten", bag.get("aantal_verblijfsobjecten")),
        ("Aantal gebruiksdoelen", bag.get("aantal_gebruiksdoelen")),
        ("Pand ID", bag.get("pand_id")),
        ("Begindatum geldigheid", bag.get("begin_geldigheid_vbo")),
        ("Einddatum geldigheid", bag.get("eind_geldigheid_vbo")),
    ], margin, col_y, col_w, row_h=11.7, label_w=42 * mm)

    right_x = margin + col_w + 6 * mm
    pand_bottom = draw_box_table(c, "Pand", [
        ("Pand ID", bag.get("pand_id")),
        ("Status", bag.get("status_pand")),
        ("Bouwjaar", bag.get("bouwjaar")),
        ("Gebruiksdoel", bag.get("gebruiksdoel")),
        ("Oppervlakte (m²)", clean_value(bag.get("pand_oppervlakte")).replace(" m²", "")),
        ("Aantal bouwlagen", bag.get("aantal_bouwlagen")),
        ("Begindatum geldigheid", bag.get("begin_geldigheid_pand")),
        ("Einddatum geldigheid", bag.get("eind_geldigheid_pand")),
    ], right_x, col_y, col_w, row_h=12.3, label_w=42 * mm)

    draw_box_table(c, "Standplaats", [("Er is geen standplaats geregistreerd voor dit adres.", "")], right_x, pand_bottom - 6 * mm, col_w, row_h=15, label_w=col_w - 14)

    # Overig + QR
    overig_y = 47 * mm
    overig_w = W - 2 * margin - 63 * mm
    draw_box_table(c, "Overig", [
        ("Gerelateerde objecten", "Zie BAG Viewer of API response voor alle relaties"),
        ("Opmerkingen", "Dit afschrift is automatisch gegenereerd op basis van BAG API Individuele Bevragingen."),
    ], margin, overig_y, overig_w, row_h=18, label_w=38 * mm)
    qr_value = f"BAG afschrift | {bag.get('adres')} | VBO {bag.get('verblijfsobject_id')}"
    draw_qr(c, qr_value, W - margin - 31 * mm, overig_y, size=31 * mm)

    c.setFillColor(colors.HexColor("#4B5563"))
    c.setFont("Helvetica", 7)
    c.drawString(margin, 10 * mm, "Dit afschrift is geen officieel bewijsstuk. Raadpleeg het Kadaster voor juridische informatie.")
    c.save()
    return {"address": bag.get("adres"), "verblijfsobject_id": bag.get("verblijfsobject_id"), "version": APP_VERSION}


def generate_bag_afschrift():
    payload = get_json_payload()
    require_secret(payload)

    source_fields = {}
    source_record_id = payload.get("source_record_id") or payload.get("dossier_record_id") or request.args.get("source_record_id")
    source_table_id = payload.get("source_table_id") or env("SOFTR_TABLE_DOSSIERS_ID")
    if source_record_id and source_table_id and truthy_env("BAG_FETCH_SOFTR_RECORD", "true"):
        source_record = softr_get_record(source_table_id, source_record_id, field_names=True)
        source_fields.update(source_record.get("fields") or {})

    if isinstance(payload.get("fields"), dict):
        source_fields.update(payload["fields"])
    for key, value in payload.items():
        if key not in INTERNAL_KEYS:
            source_fields[key] = value

    adresseerbaar_object_id = first_non_empty(
        payload.get("adresseerbaar_object_id"), payload.get("verblijfsobject_id"), payload.get("object_id"),
        pick(source_fields, "adresseerbaar_object_id", "adresseerbaarobject_id", "Adresseerbaar object ID", "Adresseerbaar object id", "verblijfsobject_id", "Verblijfsobject ID", "object_id", default=""),
        default="",
    )
    adresseerbaar_object_id = clean_bag_id(adresseerbaar_object_id)

    api_data = {}
    if truthy_env("BAG_USE_API", "true") and adresseerbaar_object_id:
        api_data = fetch_bag_verblijfsobject(adresseerbaar_object_id)

    bag = flatten_bag_data(payload, source_fields, api_data)
    address_for_filename = clean_value(bag.get("adres") or bag.get("verblijfsobject_id") or "bag_afschrift")
    filename = f"BAG_afschrift_{safe_filename(address_for_filename)}_{uuid.uuid4().hex[:8]}.pdf"
    pdf_path = TMP_DIR / filename
    meta = make_bag_pdf(bag, pdf_path)
    file_url = upload_or_host_pdf(pdf_path, filename)

    dossier_record_id = (
        payload.get("dossier_record_id")
        or payload.get("target_record_id")
        or first_record_id(pick(source_fields, "dossier_record_id", "Dossier record ID", "Dossier Record ID", "Dossier", "Dossier ID"))
        or source_record_id
    )
    target_field_id = env("SOFTR_FIELD_DOSSIER_BAG_AFSCHRIFT_PDF")

    async_update = truthy_env("SOFTR_UPDATE_ASYNC", "true")
    should_patch = bool(dossier_record_id and target_field_id)
    if should_patch and async_update:
        worker = threading.Thread(
            target=softr_patch_dossier_file_background,
            args=(dossier_record_id, target_field_id, file_url, filename, "bag_afschrift"),
            daemon=True,
        )
        worker.start()
        return jsonify({
            "status": "queued",
            "document_type": "bag_afschrift",
            "filename": filename,
            "file_url": file_url,
            "dossier_record_id": dossier_record_id,
            "target_field_id": target_field_id,
            "softr_update": "background_started",
            "bag_data": bag,
            "meta": meta,
        })

    update_result = None
    if should_patch:
        update_result = softr_patch_dossier_file(dossier_record_id, target_field_id, file_url, filename)

    return jsonify({
        "status": "ok",
        "document_type": "bag_afschrift",
        "filename": filename,
        "file_url": file_url,
        "dossier_record_id": dossier_record_id,
        "target_field_id": target_field_id,
        "softr_update_mode": update_result.get("mode") if update_result else None,
        "softr_update_assumed_success": update_result.get("assumed_success", False) if update_result else False,
        "bag_data": bag,
        "meta": meta,
    })

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


@app.route("/generate/bag-afschrift", methods=["POST"])
@app.route("/generate/bag", methods=["POST"])
def route_generate_bag_afschrift():
    return generate_bag_afschrift()


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
            "bag_afschrift": "POST /generate/bag-afschrift",
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
