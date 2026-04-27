import json
import mimetypes
import os
import re
import uuid
from datetime import datetime
from io import BytesIO
from pathlib import Path
from string import Template
from typing import Dict, List, Tuple

import requests
from flask import Flask, jsonify, make_response, request, send_from_directory
from jinja2 import Environment, FileSystemLoader, meta, select_autoescape


app = Flask(__name__)

BASE_DIR = Path(__file__).resolve().parent
OUTPUT_DIR = BASE_DIR / "output"
TEMPLATE_DIR = BASE_DIR / "templates"
MONDAY_CONFIG_PATH = BASE_DIR / "monday_config.json"

DEFAULT_TEMPLATE_TYPE = "allocation_notice_gmo"

PDF_TEMPLATE_REGISTRY = {
    "allocation_notice": {
        "file": "allocation_notice.html",
        "label": "Allocation Decision Notice",
    },
    "allocation_notice_gmo": {
        "file": "allocation_notice_gmo.html",
        "label": "Allocation Decision Notice - GMO Bank",
    },
    "application_form": {
        "file": "application_form.html",
        "label": "Bond Application Form",
    },
    "application_form_period": {
        "file": "application_form_period.html",
        "label": "Bond Application Form With Period",
    },
    "condition_summary": {
        "file": "condition_summary.html",
        "label": "Bond Condition Summary",
    },
    "interest_calculation": {
        "file": "interest_calculation.html",
        "label": "Interest Notice Calculation",
    },
    "monthly_interest_notice": {
        "file": "monthly_interest_notice.html",
        "label": "Monthly Interest Notice",
    },
    "issuance_terms_long": {
        "file": "issuance_terms_long.html",
        "label": "Long-Term Issuance Terms",
    },
    "payment_receipt": {
        "file": "payment_receipt.html",
        "label": "Payment Deposit Receipt",
    },
    "terms_two_page": {
        "file": "terms_two_page.html",
        "label": "Bond Terms - Two Pages",
    },
}

TEMPLATE_TYPE_ALIASES = {
    "allocation": "allocation_notice",
    "allocation_gmo": "allocation_notice_gmo",
    "gmo": "allocation_notice_gmo",
    "application": "application_form",
    "application_period": "application_form_period",
    "summary": "condition_summary",
    "interest": "interest_calculation",
    "monthly_interest": "monthly_interest_notice",
    "receipt": "payment_receipt",
    "terms": "terms_two_page",
    "gls_bond_allocation_decision_notice": "allocation_notice",
    "gls_bond_allocation_decision_notice_gmo": "allocation_notice_gmo",
    "gls_bond_application_form": "application_form",
    "gls_bond_application_form_with_period": "application_form_period",
    "gls_bond_condition_summary_sheet": "condition_summary",
    "gls_bond_interest_notice_calculation": "interest_calculation",
    "gls_bond_interest_notice_monthly_payment": "monthly_interest_notice",
    "gls_bond_issuance_terms_long_term_421": "issuance_terms_long",
    "gls_bond_payment_deposit_receipt": "payment_receipt",
    "gls_bond_terms_two_pages": "terms_two_page",
}

CONTROL_PAYLOAD_KEYS = {
    "data",
    "document_type",
    "monday",
    "save_local_pdf_copy",
    "save_to_monday",
    "template",
    "template_type",
}

SAMPLE_FIELD_VALUES = {
    "account_holder": "グローバルロジスティックスサービス（カ",
    "account_name": "グローバルロジスティックスサービス（カ",
    "account_number": "235859999",
    "account_type": "普通預金",
    "address": "佐賀県佐賀市天神1-2-55 IK天神ビル2階 西北号室",
    "allocated_amount": "600万円",
    "allocated_unit_count": "6",
    "amount": "600万円",
    "applicant_address_line_1": "佐賀県佐賀市天神1-2-55 IK天神ビル2階 西北号室",
    "applicant_address_line_2": "〒840-0815",
    "applicant_name": "山田 太郎",
    "applicant_postal_code": "840-0815",
    "application_date": "2026年4月25日",
    "bank_name": "GMOあおぞらネット銀行",
    "bond_number": "8",
    "bond_period": "2026年4月27日から2027年4月27日まで",
    "bond_title": "第8回普通社債",
    "bond_unit_amount": "1口 金1,000,000円",
    "bond_unit_text": "1口 金1,000,000円",
    "bondholder_address": "佐賀県佐賀市天神1-2-55 IK天神ビル2階 西北号室",
    "bondholder_name": "山田 太郎",
    "branch_name": "法人営業部",
    "calculation_detail": "600万円 × 年8.0% ÷ 12ヶ月",
    "company_name": "グローバル・ロジスティックス・サービス株式会社",
    "contact_note": "本書の内容についてご不明な点がございましたら、発行会社までお問い合わせください。",
    "contact_note_line_1": "本書の内容についてご不明な点がございましたら、発行会社までお問い合わせください。",
    "contact_note_line_2": "",
    "contact_note_line_3": "",
    "contact_note_line_4": "",
    "created_date": "2026年4月25日",
    "customer_name": "山田 太郎",
    "date": "2026年4月25日",
    "deposit_amount": "6,000,000",
    "deposit_date": "2026年4月27日",
    "face_amount": "6,000,000",
    "fax": "0952-00-0000",
    "guarantor_1": "藤井 雄太郎",
    "guarantor_2": "藤井 太郎",
    "head_office_address": "佐賀県佐賀市天神1-2-55 IK天神ビル2階 西北号室",
    "head_office_zip": "840-0815",
    "income_tax": "6,126円",
    "interest_amount": "40,000円",
    "interest_date": "2026年5月27日",
    "interest_description": "利息",
    "interest_payment_date": "毎月27日",
    "interest_rate": "年8.0%",
    "issue_date": "2026年4月27日",
    "issuer_address": "佐賀県佐賀市天神1-2-55 IK天神ビル2階 西北号室",
    "issuer_company_name": "グローバル・ロジスティックス・サービス株式会社",
    "issuer_zip": "840-0815",
    "joint_guarantee_text": "当社の代表取締役は、本社債の元利金の返還債務について、当社と連帯して保証する。",
    "monthly_interest_after_tax": "31,874円",
    "net_payment_amount": "31,874円",
    "notice_date": "2026年4月25日",
    "paid_amount": "600万円",
    "payment_date": "2026年4月27日",
    "payment_deadline": "2026年4月27日",
    "principal_amount": "600万円",
    "principal_date": "2026年4月27日",
    "principal_description": "元金",
    "recipient_name": "山田 太郎",
    "redemption_date": "2027年4月27日",
    "registration_number": "T0000000000000",
    "representative_name": "藤井 雄太郎",
    "resident_tax": "2,000円",
    "tel": "0952-00-0000",
    "tokyo_branch_address": "東京都千代田区丸の内1-1-1",
    "tokyo_branch_zip": "100-0005",
    "total_amount": "40,000円",
    "total_bond_amount": "金6,000,000円",
    "transfer_account_note": "GMOあおぞらネット銀行　法人営業部　普通預金　235859999",
    "unit_count": "6",
}

TEMPLATE_FIELD_CACHE: Dict[str, List[str]] = {}
TEST_HIGHLIGHT_RE = re.compile(
    r"\s*background(?:-color)?\s*:\s*(?:#fff2cc|#f3ecc9|yellow)\s*;?",
    re.IGNORECASE,
)
FILENAME_UNSAFE_RE = re.compile(r'[\\/:*?"<>|\s]+')


def blank_if_none(value):
    return "" if value is None else value


JINJA_ENV = Environment(
    loader=FileSystemLoader(str(TEMPLATE_DIR)),
    autoescape=select_autoescape(
        enabled_extensions=("html", "xml"),
        default_for_string=True,
    ),
    finalize=blank_if_none,
)


def ensure_output_dir() -> None:
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)


@app.before_request
def handle_preflight():
    if request.method == "OPTIONS":
        response = make_response("", 204)
        response.headers["Access-Control-Allow-Origin"] = "*"
        response.headers["Access-Control-Allow-Headers"] = "Content-Type, Authorization"
        response.headers["Access-Control-Allow-Methods"] = "GET, POST, OPTIONS"
        return response
    return None


@app.after_request
def add_cors_headers(response):
    response.headers["Access-Control-Allow-Origin"] = "*"
    response.headers["Access-Control-Allow-Headers"] = "Content-Type, Authorization"
    response.headers["Access-Control-Allow-Methods"] = "GET, POST, OPTIONS"
    return response


def parse_bool(value, default: bool = False) -> bool:
    if value is None:
        return default
    if isinstance(value, bool):
        return value
    return str(value).strip().lower() in {"1", "true", "yes", "on"}


def safe_json_loads(raw_value: str, default):
    if not raw_value:
        return default
    try:
        return json.loads(raw_value)
    except json.JSONDecodeError:
        return default


def load_local_monday_config() -> Dict:
    if not MONDAY_CONFIG_PATH.exists():
        return {}

    try:
        return json.loads(MONDAY_CONFIG_PATH.read_text(encoding="utf-8"))
    except json.JSONDecodeError as exc:
        raise ValueError(f"Invalid JSON in {MONDAY_CONFIG_PATH.name}: {exc}") from exc


def resolve_runtime_config(payload: Dict) -> Dict:
    file_config = load_local_monday_config()
    config = {
        "save_local_pdf_copy": parse_bool(os.getenv("SAVE_LOCAL_PDF_COPY"), False),
    }

    if isinstance(file_config, dict):
        config["save_local_pdf_copy"] = parse_bool(
            file_config.get("save_local_pdf_copy"),
            config["save_local_pdf_copy"],
        )

    if isinstance(payload, dict) and "save_local_pdf_copy" in payload:
        config["save_local_pdf_copy"] = parse_bool(
            payload.get("save_local_pdf_copy"),
            config["save_local_pdf_copy"],
        )

    return config


def get_default_monday_config() -> Dict:
    return {
        "enabled": parse_bool(os.getenv("MONDAY_ENABLED"), False),
        "api_token": os.getenv("MONDAY_API_TOKEN", "").strip(),
        "api_url": os.getenv("MONDAY_API_URL", "https://api.monday.com/v2").strip(),
        "file_api_url": os.getenv("MONDAY_FILE_API_URL", "https://api.monday.com/v2/file").strip(),
        "api_version": os.getenv("MONDAY_API_VERSION", "").strip(),
        "board_id": os.getenv("MONDAY_BOARD_ID", "").strip(),
        "group_id": os.getenv("MONDAY_GROUP_ID", "").strip(),
        "file_column_id": os.getenv("MONDAY_FILE_COLUMN_ID", "").strip(),
        "pdf_path_column_id": os.getenv("MONDAY_PDF_PATH_COLUMN_ID", "").strip(),
        "raw_json_column_id": os.getenv("MONDAY_RAW_JSON_COLUMN_ID", "").strip(),
        "item_name_template": os.getenv(
            "MONDAY_ITEM_NAME_TEMPLATE",
            "${template_label} ${recipient_name} ${bond_number}",
        ).strip(),
        "upload_pdf": parse_bool(os.getenv("MONDAY_UPLOAD_PDF"), True),
        "save_mapped_columns": parse_bool(os.getenv("MONDAY_SAVE_MAPPED_COLUMNS"), False),
        "column_map": safe_json_loads(os.getenv("MONDAY_COLUMN_MAP_JSON", ""), {}),
        "column_values_override": safe_json_loads(
            os.getenv("MONDAY_COLUMN_VALUES_OVERRIDE_JSON", ""),
            {},
        ),
    }


def resolve_monday_config(payload: Dict) -> Dict:
    file_config = load_local_monday_config()
    request_config = payload.get("monday", {}) if isinstance(payload.get("monday"), dict) else {}

    config = get_default_monday_config()
    config.update(file_config)
    config.update(request_config)

    config["enabled"] = parse_bool(config.get("enabled"), False)
    config["upload_pdf"] = parse_bool(config.get("upload_pdf"), True)
    config["save_mapped_columns"] = parse_bool(config.get("save_mapped_columns"), False)
    config["board_id"] = str(config.get("board_id", "")).strip()
    config["group_id"] = str(config.get("group_id", "")).strip()
    config["file_column_id"] = str(config.get("file_column_id", "")).strip()
    config["pdf_path_column_id"] = str(config.get("pdf_path_column_id", "")).strip()
    config["raw_json_column_id"] = str(config.get("raw_json_column_id", "")).strip()
    config["item_name_template"] = str(
        config.get("item_name_template", "${template_label} ${recipient_name} ${bond_number}")
    ).strip()

    if not isinstance(config.get("column_map"), dict):
        raise ValueError("monday.column_map must be an object/dictionary.")
    if not isinstance(config.get("column_values_override"), dict):
        raise ValueError("monday.column_values_override must be an object/dictionary.")

    return config


def monday_headers(config: Dict) -> Dict[str, str]:
    headers = {"Authorization": config["api_token"]}
    if config.get("api_version"):
        headers["API-Version"] = config["api_version"]
    return headers


def monday_json_headers(config: Dict) -> Dict[str, str]:
    headers = monday_headers(config)
    headers["Content-Type"] = "application/json"
    return headers


def monday_request(query: str, variables: Dict, config: Dict) -> Dict:
    response = requests.post(
        config["api_url"],
        headers=monday_json_headers(config),
        json={"query": query, "variables": variables},
        timeout=45,
    )
    response.raise_for_status()
    payload = response.json()

    if payload.get("errors"):
        messages = "; ".join(error.get("message", "Unknown monday error") for error in payload["errors"])
        raise ValueError(messages)

    return payload.get("data", {})


def build_monday_item_name(data: Dict, config: Dict) -> str:
    template = Template(config["item_name_template"])
    safe_values = {key: str(value) for key, value in data.items() if not is_blank(value)}
    item_name = template.safe_substitute(safe_values).strip()

    if item_name and "${" not in item_name:
        return item_name

    fallback_parts = [
        first_value(data, "template_label", "template_type"),
        first_value(data, "recipient_name", "customer_name", "applicant_name", "bondholder_name"),
        first_value(data, "issuer_company_name", "company_name"),
    ]

    bond_number = first_value(data, "bond_number")
    if bond_number:
        fallback_parts.append(f"第{bond_number}回普通社債")
    else:
        fallback_parts.append(first_value(data, "bond_title"))

    fallback_parts.append(
        first_value(
            data,
            "notice_date",
            "created_date",
            "application_date",
            "issue_date",
            "payment_deadline",
            "redemption_date",
            "payment_date",
            "deposit_date",
            "date",
        )
    )

    return " - ".join(str(part).strip() for part in fallback_parts if not is_blank(part)) or "Generated PDF"


def normalize_monday_column_value(value):
    if isinstance(value, (dict, list)):
        return value
    if value is None:
        return ""
    return str(value)


def build_monday_column_values(data: Dict, local_pdf_path: str, config: Dict) -> Dict:
    column_values = {}

    if config.get("save_mapped_columns"):
        for data_key, column_id in config["column_map"].items():
            if not column_id:
                continue
            column_values[str(column_id)] = normalize_monday_column_value(data.get(data_key, ""))

    if config.get("pdf_path_column_id") and local_pdf_path:
        column_values[config["pdf_path_column_id"]] = local_pdf_path

    if config.get("raw_json_column_id"):
        column_values[config["raw_json_column_id"]] = json.dumps(
            {key: value for key, value in data.items() if key != "monday"},
            ensure_ascii=False,
        )

    for column_id, value in config["column_values_override"].items():
        if column_id:
            column_values[str(column_id)] = value

    return column_values


def update_monday_item_columns(item_id: str, column_values: Dict, config: Dict) -> Dict:
    if not column_values:
        return {}

    mutation = """
    mutation ChangeMultipleColumnValues($board_id: ID!, $item_id: ID!, $column_values: JSON!) {
      change_multiple_column_values(
        board_id: $board_id,
        item_id: $item_id,
        column_values: $column_values
      ) {
        id
      }
    }
    """

    variables = {
        "board_id": str(config["board_id"]),
        "item_id": str(item_id),
        "column_values": json.dumps(column_values, ensure_ascii=False),
    }
    result = monday_request(mutation, variables, config)
    return result.get("change_multiple_column_values", {})


def create_monday_item(data: Dict, config: Dict) -> Dict:
    item_name = build_monday_item_name(data, config)

    if config.get("group_id"):
        mutation = """
        mutation CreateItem($board_id: ID!, $group_id: String!, $item_name: String!) {
          create_item(
            board_id: $board_id,
            group_id: $group_id,
            item_name: $item_name
          ) {
            id
            name
          }
        }
        """
        variables = {
            "board_id": str(config["board_id"]),
            "group_id": str(config["group_id"]),
            "item_name": item_name,
        }
    else:
        mutation = """
        mutation CreateItem($board_id: ID!, $item_name: String!) {
          create_item(
            board_id: $board_id,
            item_name: $item_name
          ) {
            id
            name
          }
        }
        """
        variables = {
            "board_id": str(config["board_id"]),
            "item_name": item_name,
        }

    result = monday_request(mutation, variables, config)
    return result["create_item"]


def upload_pdf_to_monday_file_column(item_id: str, pdf_bytes: bytes, pdf_filename: str, config: Dict) -> Dict:
    mime_type = mimetypes.guess_type(pdf_filename)[0] or "application/pdf"
    query = (
        f'mutation ($file: File!) {{ '
        f'add_file_to_column (item_id: {int(item_id)}, column_id: "{config["file_column_id"]}", file: $file) '
        f'{{ id }} }}'
    )

    with BytesIO(pdf_bytes) as pdf_file:
        response = requests.post(
            config["file_api_url"],
            headers=monday_headers(config),
            data={"query": query},
            files={"variables[file]": (pdf_filename, pdf_file, mime_type)},
            timeout=90,
        )

    response.raise_for_status()
    payload = response.json()

    if payload.get("errors"):
        messages = "; ".join(error.get("message", "Unknown monday file upload error") for error in payload["errors"])
        raise ValueError(messages)

    return payload.get("data", {}).get("add_file_to_column", {})


def validate_monday_config(config: Dict) -> None:
    if not config.get("enabled"):
        raise ValueError("monday integration is disabled.")
    if not config.get("api_token"):
        raise ValueError("Missing monday api_token. Set MONDAY_API_TOKEN or monday_config.json.")
    if not config.get("board_id"):
        raise ValueError("Missing monday board_id.")


def list_monday_boards(config: Dict) -> List[Dict]:
    query = """
    query ListBoards {
      boards(limit: 50) {
        id
        name
      }
    }
    """
    result = monday_request(query, {}, config)
    return result.get("boards", [])


def fetch_monday_board_schema(config: Dict) -> Dict:
    if not config.get("board_id"):
        raise ValueError("board_id is required to fetch board schema.")

    query = """
    query BoardSchema($boardId: [ID!]) {
      boards(ids: $boardId) {
        id
        name
        columns {
          id
          title
          type
        }
        groups {
          id
          title
        }
      }
    }
    """
    result = monday_request(query, {"boardId": [config["board_id"]]}, config)
    boards = result.get("boards", [])
    if not boards:
        raise ValueError("No board found for the provided board_id.")
    return boards[0]


def upload_to_monday(pdf_bytes: bytes, pdf_filename: str, data: Dict, local_pdf_path: str = "") -> Dict:
    """
    Create a monday item from the generated PDF payload and optionally upload
    the PDF into a monday file column.

    Configuration priority:
    1. Request JSON -> "monday": {...}
    2. Local file -> ./monday_config.json
    3. Environment variables
    """
    config = resolve_monday_config(data)
    validate_monday_config(config)

    column_values = build_monday_column_values(data, local_pdf_path, config)
    created_item = create_monday_item(data, config)
    result = {
        "success": True,
        "item_id": created_item["id"],
        "item_name": created_item["name"],
        "board_id": config["board_id"],
        "pdf_uploaded": False,
        "columns_updated": False,
        "saved_locally": bool(local_pdf_path),
    }

    if column_values:
        update_result = update_monday_item_columns(created_item["id"], column_values, config)
        result["columns_updated"] = True
        result["column_update_response"] = update_result

    if config.get("upload_pdf") and config.get("file_column_id"):
        file_result = upload_pdf_to_monday_file_column(created_item["id"], pdf_bytes, pdf_filename, config)
        result["pdf_uploaded"] = True
        result["file_column_id"] = config["file_column_id"]
        result["file_upload_response"] = file_result

    return result


def normalize_template_type(raw_template_type) -> str:
    template_type = str(raw_template_type or DEFAULT_TEMPLATE_TYPE).strip()
    template_type = template_type.rsplit("/", 1)[-1].replace(".html", "")
    template_type = TEMPLATE_TYPE_ALIASES.get(template_type, template_type)

    if template_type not in PDF_TEMPLATE_REGISTRY:
        choices = ", ".join(sorted(PDF_TEMPLATE_REGISTRY.keys()))
        raise ValueError(f"Unknown template_type '{raw_template_type}'. Choose one of: {choices}.")

    return template_type


def requested_template_type(payload: Dict) -> str:
    return normalize_template_type(
        payload.get("template_type")
        or payload.get("document_type")
        or payload.get("template")
        or DEFAULT_TEMPLATE_TYPE
    )


def template_config(template_type: str) -> Dict:
    normalized_type = normalize_template_type(template_type)
    return PDF_TEMPLATE_REGISTRY[normalized_type]


def template_source(template_type: str) -> str:
    config = template_config(template_type)
    source, _, _ = JINJA_ENV.loader.get_source(JINJA_ENV, config["file"])
    return source


def template_fields(template_type: str) -> List[str]:
    normalized_type = normalize_template_type(template_type)
    if normalized_type in TEMPLATE_FIELD_CACHE:
        return TEMPLATE_FIELD_CACHE[normalized_type]

    ast = JINJA_ENV.parse(template_source(normalized_type))
    fields = sorted(meta.find_undeclared_variables(ast))
    TEMPLATE_FIELD_CACHE[normalized_type] = fields
    return fields


def template_catalog() -> List[Dict]:
    catalog = []
    for template_type, config in PDF_TEMPLATE_REGISTRY.items():
        fields = template_fields(template_type)
        catalog.append(
            {
                "type": template_type,
                "label": config["label"],
                "file": config["file"],
                "fields": fields,
                "sample_data": {
                    field: SAMPLE_FIELD_VALUES.get(field, "")
                    for field in fields
                },
            }
        )
    return catalog


def extract_document_data(payload: Dict) -> Dict:
    document_data = {}

    nested_data = payload.get("data")
    if isinstance(nested_data, dict):
        document_data.update(nested_data)

    for key, value in payload.items():
        if key not in CONTROL_PAYLOAD_KEYS:
            document_data[key] = value

    return document_data


def is_blank(value) -> bool:
    return value is None or str(value).strip() == ""


def first_value(data: Dict, *keys: str):
    for key in keys:
        value = data.get(key)
        if not is_blank(value):
            return value
    return ""


def set_if_blank(data: Dict, key: str, value) -> None:
    if is_blank(data.get(key)) and not is_blank(value):
        data[key] = value


def normalize_unit_count(value) -> str:
    text = str(value or "").strip()
    return text[:-1].strip() if text.endswith("口") else text


def filename_part(value, fallback: str = "") -> str:
    text = str(value or fallback or "").strip()
    text = FILENAME_UNSAFE_RE.sub("-", text)
    text = re.sub(r"-{2,}", "-", text).strip("-. _")
    return text[:60]


def build_pdf_filename(template_info: Dict, context: Dict) -> str:
    recipient = first_value(
        context,
        "recipient_name",
        "customer_name",
        "applicant_name",
        "bondholder_name",
    )
    bond = first_value(context, "bond_number", "bond_title")
    document_date = first_value(
        context,
        "notice_date",
        "created_date",
        "application_date",
        "issue_date",
        "payment_deadline",
        "payment_date",
        "deposit_date",
        "date",
    )
    timestamp = datetime.utcnow().strftime("%Y%m%dT%H%M%S")
    unique_id = uuid.uuid4().hex[:8]

    parts = [
        filename_part(template_info["template_type"], "document"),
        filename_part(recipient),
        filename_part(f"bond-{bond}" if bond else ""),
        filename_part(document_date),
        timestamp,
        unique_id,
    ]
    return "-".join(part for part in parts if part) + ".pdf"


def parse_bond_number(value) -> str:
    match = re.search(r"第\s*([0-9０-９]+)\s*回", str(value or ""))
    return match.group(1) if match else ""


def add_derived_aliases(data: Dict) -> Dict:
    set_if_blank(data, "recipient_name", first_value(data, "customer_name", "applicant_name", "bondholder_name"))
    set_if_blank(data, "customer_name", first_value(data, "recipient_name", "applicant_name", "bondholder_name"))
    set_if_blank(data, "applicant_name", first_value(data, "recipient_name", "customer_name"))
    set_if_blank(data, "bondholder_name", first_value(data, "recipient_name", "customer_name"))

    set_if_blank(data, "issuer_company_name", first_value(data, "company_name"))
    set_if_blank(data, "company_name", first_value(data, "issuer_company_name"))

    set_if_blank(data, "issuer_address", first_value(data, "address", "head_office_address", "bondholder_address"))
    set_if_blank(data, "address", first_value(data, "issuer_address", "head_office_address", "bondholder_address"))
    set_if_blank(data, "head_office_address", first_value(data, "issuer_address", "address"))
    set_if_blank(data, "bondholder_address", first_value(data, "address", "issuer_address"))
    set_if_blank(data, "applicant_address_line_1", first_value(data, "address", "bondholder_address"))

    set_if_blank(data, "notice_date", first_value(data, "date", "created_date", "application_date", "issue_date"))
    set_if_blank(data, "created_date", first_value(data, "date", "notice_date", "issue_date"))
    set_if_blank(data, "application_date", first_value(data, "date", "notice_date", "issue_date"))
    set_if_blank(data, "issue_date", first_value(data, "date", "notice_date", "payment_date", "payment_deadline"))
    set_if_blank(data, "date", first_value(data, "notice_date", "created_date", "application_date", "issue_date"))

    set_if_blank(data, "payment_deadline", first_value(data, "payment_date"))
    set_if_blank(data, "payment_date", first_value(data, "payment_deadline"))
    set_if_blank(data, "deposit_date", first_value(data, "payment_date", "payment_deadline"))

    set_if_blank(data, "account_holder", first_value(data, "account_name"))
    set_if_blank(data, "account_name", first_value(data, "account_holder"))

    set_if_blank(data, "allocated_amount", first_value(data, "amount", "face_amount", "paid_amount", "deposit_amount"))
    set_if_blank(data, "amount", first_value(data, "allocated_amount", "face_amount", "paid_amount", "deposit_amount"))
    set_if_blank(data, "paid_amount", first_value(data, "allocated_amount", "amount"))
    set_if_blank(data, "deposit_amount", first_value(data, "paid_amount", "amount"))
    set_if_blank(data, "principal_amount", first_value(data, "paid_amount", "amount"))

    if not is_blank(data.get("unit_count")):
        data["unit_count"] = normalize_unit_count(data["unit_count"])
    if not is_blank(data.get("allocated_unit_count")):
        data["allocated_unit_count"] = normalize_unit_count(data["allocated_unit_count"])
    set_if_blank(data, "allocated_unit_count", normalize_unit_count(first_value(data, "unit_count")))
    set_if_blank(data, "unit_count", normalize_unit_count(first_value(data, "allocated_unit_count")))

    set_if_blank(data, "bond_number", parse_bond_number(first_value(data, "bond_title")))
    if not is_blank(data.get("bond_number")):
        set_if_blank(data, "bond_title", f'第{data["bond_number"]}回普通社債')

    set_if_blank(data, "bank_name", first_value(data, "bank"))
    set_if_blank(data, "branch_name", first_value(data, "branch"))

    return data


def build_template_context(document_data: Dict, template_type: str) -> Dict:
    context = {
        key: blank_if_none(value)
        for key, value in document_data.items()
    }
    add_derived_aliases(context)

    for field in template_fields(template_type):
        context.setdefault(field, "")

    return context


def remove_test_highlights(html_content: str) -> str:
    return TEST_HIGHLIGHT_RE.sub("", html_content)


def ensure_full_html_document(html_content: str) -> str:
    if "<html" in html_content.lower():
        return html_content

    return f"""<!DOCTYPE html>
<html lang="ja">
  <head>
    <meta charset="UTF-8" />
    <style>
      html,
      body {{
        margin: 0;
        padding: 0;
        background: #ffffff;
      }}
    </style>
  </head>
  <body>
    {html_content}
  </body>
</html>
"""


def render_template_html(template_type: str, context: Dict) -> str:
    config = template_config(template_type)
    template = JINJA_ENV.get_template(config["file"])
    html_content = template.render(context)
    html_content = remove_test_highlights(html_content)
    return ensure_full_html_document(html_content)


def configure_native_library_paths() -> None:
    fallback_paths = ["/opt/homebrew/lib", "/usr/local/lib"]
    existing_paths = [
        path
        for path in os.getenv("DYLD_FALLBACK_LIBRARY_PATH", "").split(":")
        if path
    ]

    for path in fallback_paths:
        if Path(path).exists() and path not in existing_paths:
            existing_paths.append(path)

    if existing_paths:
        os.environ["DYLD_FALLBACK_LIBRARY_PATH"] = ":".join(existing_paths)


def html_to_pdf_bytes(html_content: str) -> bytes:
    configure_native_library_paths()

    try:
        from weasyprint import HTML
    except Exception as exc:
        raise RuntimeError(
            "WeasyPrint is required for HTML template PDF generation. "
            "Install Python dependencies with pip install -r requirements.txt. "
            "Native rendering libraries are also required: use brew install pango on macOS, "
            "or deploy Cloud Run with the included Dockerfile so Debian packages such as "
            "libpango-1.0-0, libpangoft2-1.0-0, libharfbuzz-subset0, and fonts-noto-cjk are installed."
        ) from exc

    return HTML(string=html_content, base_url=str(BASE_DIR)).write_pdf()


def build_pdf(document_data: Dict, template_type: str) -> Tuple[bytes, Dict, Dict]:
    normalized_type = normalize_template_type(template_type)
    config = template_config(normalized_type)
    fields = template_fields(normalized_type)
    context = build_template_context(document_data, normalized_type)
    html_content = render_template_html(normalized_type, context)
    pdf_bytes = html_to_pdf_bytes(html_content)
    empty_fields = [
        field
        for field in fields
        if is_blank(context.get(field))
    ]

    return (
        pdf_bytes,
        {
            "template_type": normalized_type,
            "template_label": config["label"],
            "template_file": config["file"],
            "template_fields": fields,
            "empty_fields": empty_fields,
        },
        context,
    )


def monday_requested_from_payload(payload: Dict, config: Dict) -> bool:
    if "save_to_monday" in payload:
        return parse_bool(payload.get("save_to_monday"), False)
    return bool(config.get("enabled", False))


@app.route("/health", methods=["GET", "OPTIONS"])
def health_check():
    if request.method == "OPTIONS":
        return make_response("", 204)

    return jsonify(
        {
            "success": True,
            "message": "GLS PDF backend is running.",
            "template_count": len(PDF_TEMPLATE_REGISTRY),
            "templates_dir_exists": TEMPLATE_DIR.exists(),
        }
    ), 200


@app.route("/templates", methods=["GET", "OPTIONS"])
def list_templates():
    if request.method == "OPTIONS":
        return make_response("", 204)

    try:
        return jsonify(
            {
                "success": True,
                "default_template_type": DEFAULT_TEMPLATE_TYPE,
                "templates": template_catalog(),
            }
        ), 200
    except Exception as exc:
        return jsonify({"success": False, "message": str(exc)}), 500


@app.route("/generate-pdf", methods=["POST", "OPTIONS"])
def generate_pdf():
    if request.method == "OPTIONS":
        return make_response("", 204)

    payload = request.get_json(silent=True)
    if not isinstance(payload, dict):
        return jsonify({"success": False, "message": "Request body must be valid JSON."}), 400

    try:
        template_type = requested_template_type(payload)
        document_data = extract_document_data(payload)
    except Exception as exc:
        return jsonify({"success": False, "message": str(exc)}), 400

    local_pdf_path = ""

    try:
        pdf_bytes, template_info, render_context = build_pdf(document_data, template_type)
    except Exception as exc:
        return (
            jsonify(
                {
                    "success": False,
                    "message": "Failed to generate PDF.",
                    "error": str(exc),
                    "template_type": template_type,
                }
            ),
            500,
        )

    file_name = build_pdf_filename(template_info, render_context)
    pdf_path = OUTPUT_DIR / file_name

    runtime_config = resolve_runtime_config(payload)
    if runtime_config.get("save_local_pdf_copy"):
        ensure_output_dir()
        pdf_path.write_bytes(pdf_bytes)
        local_pdf_path = str(pdf_path)

    monday_result = None
    monday_error = None
    monday_requested = False

    try:
        monday_config = resolve_monday_config(payload)
        monday_requested = monday_requested_from_payload(payload, monday_config)
    except Exception as exc:
        monday_config = None
        if parse_bool(payload.get("save_to_monday"), False):
            monday_error = str(exc)

    if monday_requested:
        if monday_error is None:
            monday_payload = dict(render_context)
            monday_payload.update(
                {
                    "template_type": template_info["template_type"],
                    "template_label": template_info["template_label"],
                }
            )
            if isinstance(payload.get("monday"), dict):
                monday_payload["monday"] = payload["monday"]

            try:
                monday_result = upload_to_monday(
                    pdf_bytes,
                    file_name,
                    monday_payload,
                    local_pdf_path=local_pdf_path,
                )
            except Exception as exc:
                monday_error = str(exc)

    return (
        jsonify(
            {
                "success": True,
                "message": "PDF generated successfully.",
                "file_path": local_pdf_path or None,
                "generated_filename": file_name,
                "saved_locally": bool(local_pdf_path),
                "template_type": template_info["template_type"],
                "template_label": template_info["template_label"],
                "template_file": template_info["template_file"],
                "template_fields": template_info["template_fields"],
                "empty_fields": template_info["empty_fields"],
                "monday_requested": monday_requested,
                "monday_result": monday_result,
                "monday_error": monday_error,
            }
        ),
        201,
    )


@app.route("/monday/discover", methods=["POST", "OPTIONS"])
def monday_discover():
    if request.method == "OPTIONS":
        return make_response("", 204)

    payload = request.get_json(silent=True) or {}
    if not isinstance(payload, dict):
        return jsonify({"success": False, "message": "Request body must be valid JSON."}), 400

    try:
        config = resolve_monday_config(payload)
        if not config.get("api_token"):
            return (
                jsonify(
                    {
                        "success": False,
                        "message": "Missing monday api_token. Set MONDAY_API_TOKEN or monday_config.json.",
                    }
                ),
                400,
            )

        if config.get("board_id"):
            board = fetch_monday_board_schema(config)
            return jsonify({"success": True, "mode": "board_schema", "board": board}), 200

        boards = list_monday_boards(config)
        return jsonify({"success": True, "mode": "boards", "boards": boards}), 200
    except Exception as exc:
        return jsonify({"success": False, "message": str(exc)}), 500


@app.route("/config/defaults", methods=["GET", "OPTIONS"])
def config_defaults():
    if request.method == "OPTIONS":
        return make_response("", 204)

    try:
        config = resolve_monday_config({})
        missing_fields = []
        if not config.get("enabled"):
            missing_fields.append("enabled")
        if not config.get("board_id"):
            missing_fields.append("board_id")
        if config.get("upload_pdf") and not config.get("file_column_id"):
            missing_fields.append("file_column_id")

        return (
            jsonify(
                {
                    "success": True,
                    "default_template_type": DEFAULT_TEMPLATE_TYPE,
                    "templates": template_catalog(),
                    "monday": {
                        "enabled": config.get("enabled", False),
                        "board_id": config.get("board_id", ""),
                        "group_id": config.get("group_id", ""),
                        "file_column_id": config.get("file_column_id", ""),
                        "item_name_template": config.get("item_name_template", "${customer_name}"),
                        "upload_pdf": config.get("upload_pdf", True),
                        "save_mapped_columns": config.get("save_mapped_columns", False),
                        "save_local_pdf_copy": resolve_runtime_config({}).get("save_local_pdf_copy", False),
                        "integration_ready": len(missing_fields) == 0,
                        "missing_fields": missing_fields,
                    },
                }
            ),
            200,
        )
    except Exception as exc:
        return jsonify({"success": False, "message": str(exc)}), 500


@app.route("/output/<path:filename>", methods=["GET"])
def output_file(filename: str):
    return send_from_directory(OUTPUT_DIR, filename, as_attachment=False)


if __name__ == "__main__":
    ensure_output_dir()
    app.run(
        host="0.0.0.0",
        port=int(os.getenv("PORT", "5000")),
        debug=os.getenv("FLASK_DEBUG", "false").lower() == "true",
    )
