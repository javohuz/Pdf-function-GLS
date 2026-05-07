import json
import mimetypes
import os
import re
from datetime import datetime
from io import BytesIO
from pathlib import Path
from string import Template
from typing import Any, Dict, List, Optional, Tuple

import requests
from dotenv import load_dotenv
from flask import Flask, jsonify, request
from jinja2 import Environment, FileSystemLoader, meta, select_autoescape


app = Flask(__name__)

BASE_DIR = Path(__file__).resolve().parent
TEMPLATE_DIR = BASE_DIR / "templates"
BOARD_DOCUMENT_CONFIG_PATH = Path(
    os.getenv("BOARD_DOCUMENT_CONFIG_PATH", str(BASE_DIR / "board_document_config.json"))
)

load_dotenv(BASE_DIR / ".env")

DEFAULT_TEMPLATE_TYPE = "allocation_notice_gmo"
DEFAULT_OUTPUT_FORMAT = "pdf"
DOCX_TEMPLATE_DIR_CANDIDATES = [
    BASE_DIR / "docx_templates",
    BASE_DIR / "dock_tempaltes",
]

OUTPUT_FORMAT_REGISTRY = {
    "pdf": {
        "label": "PDF",
        "extension": "pdf",
    },
    "docx": {
        "label": "Word (.docx)",
        "extension": "docx",
    },
}

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

DOCX_TEMPLATE_REGISTRY = {
    template_type: {
        "file": config["file"].replace(".html", ".docx"),
        "label": f'{config["label"]} (Word Template)',
    }
    for template_type, config in PDF_TEMPLATE_REGISTRY.items()
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

HTML_TEMPLATE_FIELD_CACHE: Dict[str, List[str]] = {}
DOCX_TEMPLATE_FIELD_CACHE: Dict[str, List[str]] = {}
TEMPLATE_FIELD_CACHE: Dict[Tuple[str, str], List[str]] = {}
PROCESSED_TRIGGER_UUIDS: set[str] = set()

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


def is_blank(value) -> bool:
    return value is None or str(value).strip() == ""


def first_value(data: Dict[str, Any], *keys: str):
    for key in keys:
        value = data.get(key)
        if not is_blank(value):
            return value
    return ""


def set_if_blank(data: Dict[str, Any], key: str, value) -> None:
    if is_blank(data.get(key)) and not is_blank(value):
        data[key] = value


def resolve_monday_api_config() -> Dict[str, Any]:
    return {
        "api_token": os.getenv("MONDAY_API_TOKEN", "").strip(),
        "api_url": os.getenv("MONDAY_API_URL", "https://api.monday.com/v2").strip(),
        "file_api_url": os.getenv("MONDAY_FILE_API_URL", "https://api.monday.com/v2/file").strip(),
        "api_version": os.getenv("MONDAY_API_VERSION", "").strip(),
    }


def validate_monday_api_config(config: Dict[str, Any]) -> None:
    if not config.get("api_token"):
        raise ValueError("Missing monday api_token. Set MONDAY_API_TOKEN in environment or .env.")


def load_board_document_config() -> Dict[str, Any]:
    if not BOARD_DOCUMENT_CONFIG_PATH.exists():
        return {"boards": {}}

    try:
        config = json.loads(BOARD_DOCUMENT_CONFIG_PATH.read_text(encoding="utf-8"))
    except json.JSONDecodeError as exc:
        raise ValueError(f"Invalid JSON in {BOARD_DOCUMENT_CONFIG_PATH.name}: {exc}") from exc

    if not isinstance(config, dict):
        raise ValueError(f"{BOARD_DOCUMENT_CONFIG_PATH.name} must contain a JSON object.")

    boards = config.get("boards", {})
    if not isinstance(boards, dict):
        raise ValueError(f"{BOARD_DOCUMENT_CONFIG_PATH.name} must contain a 'boards' object.")

    return config


def normalize_template_type(raw_template_type) -> str:
    template_type = str(raw_template_type or DEFAULT_TEMPLATE_TYPE).strip()
    template_type = template_type.rsplit("/", 1)[-1].replace(".html", "").replace(".docx", "")
    template_type = TEMPLATE_TYPE_ALIASES.get(template_type, template_type)

    if template_type not in PDF_TEMPLATE_REGISTRY:
        choices = ", ".join(sorted(PDF_TEMPLATE_REGISTRY.keys()))
        raise ValueError(f"Unknown template_type '{raw_template_type}'. Choose one of: {choices}.")

    return template_type


def normalize_output_format(raw_output_format) -> str:
    output_format = str(raw_output_format or DEFAULT_OUTPUT_FORMAT).strip().lower()
    if output_format not in OUTPUT_FORMAT_REGISTRY:
        choices = ", ".join(sorted(OUTPUT_FORMAT_REGISTRY.keys()))
        raise ValueError(f"Unknown output_format '{raw_output_format}'. Choose one of: {choices}.")
    return output_format


def resolve_board_document_setup(board_id) -> Dict[str, Any]:
    board_key = str(board_id or "").strip()
    config = load_board_document_config()
    board_config = config.get("boards", {}).get(board_key)
    if not isinstance(board_config, dict):
        raise ValueError(
            f"No board document configuration found for board_id '{board_key}' in "
            f"{BOARD_DOCUMENT_CONFIG_PATH.name}."
        )

    setup = dict(board_config)
    setup["board_id"] = board_key
    setup["enabled"] = parse_bool(setup.get("enabled"), True)
    setup["document_type"] = str(
        setup.get("document_type") or setup.get("template_type") or DEFAULT_TEMPLATE_TYPE
    ).strip()
    setup["template_type"] = normalize_template_type(setup["document_type"])
    setup["output_format"] = normalize_output_format(
        setup.get("output_format") or DEFAULT_OUTPUT_FORMAT
    )
    setup["match_mode"] = str(setup.get("match_mode") or "column_name").strip().lower()
    setup["file_name_template"] = str(setup.get("file_name_template") or "").strip()

    if setup["match_mode"] != "column_name":
        raise ValueError(
            f"Unsupported match_mode '{setup['match_mode']}' for board_id '{board_key}'."
        )

    if not isinstance(setup.get("monday"), dict):
        raise ValueError(f"board_id '{board_key}' is missing a valid 'monday' object.")

    setup.setdefault("column_name_rules", {})
    setup.setdefault("overrides", {})
    setup.setdefault("defaults", {})
    setup.setdefault("upload", {})
    setup.setdefault("status_labels", {})

    for key in ("column_name_rules", "overrides", "defaults", "upload", "status_labels"):
        if not isinstance(setup.get(key), dict):
            raise ValueError(f"board_id '{board_key}' has an invalid '{key}' object.")

    return setup


def monday_headers(config: Dict[str, Any]) -> Dict[str, str]:
    headers = {"Authorization": config["api_token"]}
    if config.get("api_version"):
        headers["API-Version"] = config["api_version"]
    return headers


def monday_json_headers(config: Dict[str, Any]) -> Dict[str, str]:
    headers = monday_headers(config)
    headers["Content-Type"] = "application/json"
    return headers


def monday_request(query: str, variables: Dict[str, Any], config: Dict[str, Any]) -> Dict[str, Any]:
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


def update_monday_item_columns(
    item_id: str,
    board_id: str,
    column_values: Dict[str, Any],
    config: Dict[str, Any],
) -> Dict[str, Any]:
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
        "board_id": str(board_id),
        "item_id": str(item_id),
        "column_values": json.dumps(column_values, ensure_ascii=False),
    }
    result = monday_request(mutation, variables, config)
    return result.get("change_multiple_column_values", {})


def clear_monday_file_column(
    item_id: str,
    board_id: str,
    file_column_id: str,
    config: Dict[str, Any],
) -> Dict[str, Any]:
    mutation = """
    mutation ClearFileColumn($board_id: ID!, $item_id: ID!, $column_id: String!, $value: JSON!) {
      change_column_value(
        board_id: $board_id,
        item_id: $item_id,
        column_id: $column_id,
        value: $value
      ) {
        id
      }
    }
    """

    variables = {
        "board_id": str(board_id),
        "item_id": str(item_id),
        "column_id": str(file_column_id),
        "value": json.dumps({"clear_all": True}),
    }
    result = monday_request(mutation, variables, config)
    return result.get("change_column_value", {})


def fetch_monday_board_schema(board_id: str, config: Dict[str, Any]) -> Dict[str, Any]:
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
      }
    }
    """
    result = monday_request(query, {"boardId": [str(board_id)]}, config)
    boards = result.get("boards", [])
    if not boards:
        raise ValueError(f"No board found for board_id '{board_id}'.")
    return boards[0]


def fetch_monday_item_row(item_id: str, board_id: str, config: Dict[str, Any]) -> Dict[str, Any]:
    query_with_titles = """
    query FetchItem($itemIds: [ID!]) {
      items(ids: $itemIds) {
        id
        name
        column_values {
          id
          text
          type
          value
          column {
            title
          }
        }
      }
    }
    """
    query_without_titles = """
    query FetchItem($itemIds: [ID!]) {
      items(ids: $itemIds) {
        id
        name
        column_values {
          id
          text
          type
          value
        }
      }
    }
    """

    items: List[Dict[str, Any]] = []
    try:
        result = monday_request(query_with_titles, {"itemIds": [str(item_id)]}, config)
        items = result.get("items", [])
    except Exception:
        result = monday_request(query_without_titles, {"itemIds": [str(item_id)]}, config)
        items = result.get("items", [])
        board = fetch_monday_board_schema(board_id, config)
        title_map = {
            str(column.get("id", "")): str(column.get("title", "")).strip()
            for column in board.get("columns", [])
            if column.get("id")
        }
        for item in items:
            for column_value in item.get("column_values", []):
                column_value["column"] = {"title": title_map.get(str(column_value.get("id", "")), "")}

    if not items:
        raise ValueError(f"No monday item found for item_id '{item_id}'.")

    item = items[0]
    normalized_columns = []
    for column_value in item.get("column_values", []):
        column_info = column_value.get("column") or {}
        normalized_columns.append(
            {
                "id": str(column_value.get("id", "")).strip(),
                "title": str(column_info.get("title", "")).strip(),
                "text": blank_if_none(column_value.get("text")),
                "type": str(column_value.get("type", "")).strip(),
                "value": column_value.get("value"),
            }
        )

    return {
        "id": str(item.get("id", "")).strip(),
        "name": str(item.get("name", "")).strip(),
        "column_values": normalized_columns,
    }


def normalize_variable_name(value: str, rules: Optional[Dict[str, Any]] = None) -> str:
    rules = rules or {}
    text = str(value or "")
    if parse_bool(rules.get("trim_spaces"), True):
        text = text.strip()
    if parse_bool(rules.get("normalize_case"), True):
        text = text.lower()
    text = text.replace("　", " ")
    text = re.sub(r"[^\w]+", "_", text, flags=re.UNICODE)
    text = re.sub(r"_+", "_", text).strip("_")
    return text


def monday_column_display_value(column_value: Dict[str, Any]) -> str:
    text = blank_if_none(column_value.get("text"))
    if not is_blank(text):
        return str(text).strip()

    raw_value = column_value.get("value")
    parsed = raw_value
    if isinstance(raw_value, str) and raw_value.strip():
        try:
            parsed = json.loads(raw_value)
        except json.JSONDecodeError:
            return raw_value.strip()

    if isinstance(parsed, dict):
        label = parsed.get("label")
        if isinstance(label, dict) and not is_blank(label.get("text")):
            return str(label.get("text")).strip()
        for key in ("text", "name", "date", "value"):
            if not is_blank(parsed.get(key)):
                return str(parsed.get(key)).strip()

    if isinstance(parsed, list):
        flattened = [str(value).strip() for value in parsed if not is_blank(value)]
        return ", ".join(flattened)

    if parsed is None:
        return ""

    return str(parsed).strip()


def monday_item_to_document_data(
    item_row: Dict[str, Any],
    board_setup: Dict[str, Any],
) -> Dict[str, Any]:
    rules = board_setup.get("column_name_rules", {})
    item_id = str(item_row.get("id", "")).strip()
    item_name = str(item_row.get("name", "")).strip()
    board_id = str(board_setup.get("board_id", "")).strip()
    data: Dict[str, Any] = {
        "item_id": item_id,
        "item_name": item_name,
        "pulse_name": item_name,
        "row_name": item_name,
        "board_id": board_id,
        "itemId": item_id,
        "itemName": item_name,
        "pulseId": item_id,
        "pulseName": item_name,
        "boardId": board_id,
    }

    by_normalized_title: Dict[str, Any] = {}
    for column_value in item_row.get("column_values", []):
        value = monday_column_display_value(column_value)
        column_id = str(column_value.get("id", "")).strip()
        title = str(column_value.get("title", "")).strip()
        if column_id:
            data[column_id] = value
            normalized_id = normalize_variable_name(column_id, rules)
            if normalized_id:
                data[normalized_id] = value
                by_normalized_title[normalized_id] = value

        if title:
            data[title] = value
        normalized_title = normalize_variable_name(title, rules)
        if not normalized_title:
            continue
        by_normalized_title[normalized_title] = value
        data[normalized_title] = value

    data.update(by_normalized_title)

    for target_key, source_key in board_setup.get("overrides", {}).items():
        normalized_target = normalize_variable_name(target_key, rules)
        normalized_source = normalize_variable_name(source_key, rules)
        if normalized_target and normalized_source and is_blank(data.get(normalized_target)):
            data[normalized_target] = by_normalized_title.get(normalized_source, "")

    for key, value in board_setup.get("defaults", {}).items():
        normalized_key = normalize_variable_name(key, rules)
        if normalized_key and is_blank(data.get(normalized_key)):
            data[normalized_key] = value

    return data


def template_config(template_type: str) -> Dict[str, Any]:
    normalized_type = normalize_template_type(template_type)
    return PDF_TEMPLATE_REGISTRY[normalized_type]


def template_source(template_type: str) -> str:
    config = template_config(template_type)
    source, _, _ = JINJA_ENV.loader.get_source(JINJA_ENV, config["file"])
    return source


def docx_template_path(template_type: str) -> Optional[Path]:
    config = DOCX_TEMPLATE_REGISTRY.get(normalize_template_type(template_type))
    if not config:
        return None

    for directory in DOCX_TEMPLATE_DIR_CANDIDATES:
        candidate = directory / config["file"]
        if candidate.exists():
            return candidate

    return None


def html_template_fields(template_type: str) -> List[str]:
    normalized_type = normalize_template_type(template_type)
    if normalized_type in HTML_TEMPLATE_FIELD_CACHE:
        return HTML_TEMPLATE_FIELD_CACHE[normalized_type]

    ast = JINJA_ENV.parse(template_source(normalized_type))
    fields = sorted(meta.find_undeclared_variables(ast))
    HTML_TEMPLATE_FIELD_CACHE[normalized_type] = fields
    return fields


def docx_template_fields(template_type: str) -> List[str]:
    normalized_type = normalize_template_type(template_type)
    if normalized_type in DOCX_TEMPLATE_FIELD_CACHE:
        return DOCX_TEMPLATE_FIELD_CACHE[normalized_type]

    template_path = docx_template_path(normalized_type)
    if not template_path:
        DOCX_TEMPLATE_FIELD_CACHE[normalized_type] = []
        return []

    try:
        from docxtpl import DocxTemplate
    except Exception:
        DOCX_TEMPLATE_FIELD_CACHE[normalized_type] = []
        return []

    template = DocxTemplate(str(template_path))
    fields = sorted(template.get_undeclared_template_variables())
    DOCX_TEMPLATE_FIELD_CACHE[normalized_type] = fields
    return fields


def template_fields(template_type: str, output_format: Optional[str] = None) -> List[str]:
    normalized_type = normalize_template_type(template_type)
    normalized_format = normalize_output_format(output_format) if output_format else "all"
    cache_key = (normalized_type, normalized_format)

    if cache_key in TEMPLATE_FIELD_CACHE:
        return TEMPLATE_FIELD_CACHE[cache_key]

    if normalized_format == "pdf":
        fields = html_template_fields(normalized_type)
    elif normalized_format == "docx":
        fields = docx_template_fields(normalized_type)
    else:
        fields = sorted(set(html_template_fields(normalized_type)) | set(docx_template_fields(normalized_type)))

    TEMPLATE_FIELD_CACHE[cache_key] = fields
    return fields


def normalize_unit_count(value) -> str:
    text = str(value or "").strip()
    return text[:-1].strip() if text.endswith("口") else text


def extract_money_number(value) -> str:
    text = str(value or "").strip()
    match = re.search(r"金\s*([0-9０-９,]+)", text)
    if not match:
        match = re.search(r"([0-9０-９,]+)", text)
    return match.group(1) if match else ""


def yen_to_man_yen(value) -> str:
    number_text = extract_money_number(value)
    normalized = number_text.replace(",", "").translate(str.maketrans("０１２３４５６７８９", "0123456789"))
    if not normalized.isdigit():
        return number_text

    amount = int(normalized)
    if amount >= 10000 and amount % 10000 == 0:
        return f"{amount // 10000:,}"
    return number_text


def parse_bond_number(value) -> str:
    match = re.search(r"第\s*([0-9０-９]+)\s*回", str(value or ""))
    return match.group(1) if match else ""


def add_derived_aliases(data: Dict[str, Any]) -> Dict[str, Any]:
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

    set_if_blank(data, "bank_name", first_value(data, "bank"))
    set_if_blank(data, "branch_name", first_value(data, "branch"))

    set_if_blank(
        data,
        "bank_info",
        "\n".join(
            part
            for part in [
                "　".join(
                    part
                    for part in [first_value(data, "bank_name"), first_value(data, "branch_name")]
                    if not is_blank(part)
                ),
                "　".join(
                    part
                    for part in [
                        first_value(data, "account_type"),
                        f'口座番号　{first_value(data, "account_number")}' if not is_blank(first_value(data, "account_number")) else "",
                    ]
                    if not is_blank(part)
                ),
                first_value(data, "account_holder", "account_name"),
            ]
            if not is_blank(part)
        ),
    )

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

    set_if_blank(data, "bond_unit_cost", yen_to_man_yen(first_value(data, "bond_unit_amount", "bond_unit_text")))
    set_if_blank(data, "monthly_interest_net", first_value(data, "monthly_interest_after_tax", "net_payment_amount"))
    set_if_blank(data, "monthly_interest_after_tax", first_value(data, "monthly_interest_net", "net_payment_amount"))

    return data


def build_template_context(document_data: Dict[str, Any], template_type: str, output_format: Optional[str] = None) -> Dict[str, Any]:
    context = {
        key: blank_if_none(value)
        for key, value in document_data.items()
    }
    add_derived_aliases(context)

    for field in template_fields(template_type, output_format):
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


def render_template_html(template_type: str, context: Dict[str, Any]) -> str:
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
            "libpango-1.0-0, libpangocairo-1.0-0, libpangoft2-1.0-0, "
            f"libharfbuzz-subset0, and fonts-noto-cjk are installed. Underlying import error: {exc!r}"
        ) from exc

    return HTML(string=html_content, base_url=str(BASE_DIR)).write_pdf()


def docx_template_to_bytes(template_type: str, context: Dict[str, Any]) -> bytes:
    template_path = docx_template_path(template_type)
    if not template_path:
        raise RuntimeError(
            f"No .docx template is configured for template_type '{template_type}'."
        )

    try:
        from docxtpl import DocxTemplate
    except Exception as exc:
        raise RuntimeError(
            "docxtpl is required for Word template generation. "
            "Install Python dependencies with pip install -r requirements.txt. "
            f"Underlying import error: {exc!r}"
        ) from exc

    template = DocxTemplate(str(template_path))
    template.render(context)
    output = BytesIO()
    template.save(output)
    return output.getvalue()


def filename_part(value, fallback: str = "") -> str:
    text = str(value or fallback or "").strip()
    text = FILENAME_UNSAFE_RE.sub("-", text)
    text = re.sub(r"-{2,}", "-", text).strip("-. _")
    return text[:60]


def build_document_filename(
    template_info: Dict[str, Any],
    context: Dict[str, Any],
    output_format: str,
    board_setup: Optional[Dict[str, Any]] = None,
) -> str:
    recipient = first_value(
        context,
        "recipient_name",
        "customer_name",
        "applicant_name",
        "bondholder_name",
        "item_name",
    )
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

    safe_values: Dict[str, str] = {
        key: filename_part(value)
        for key, value in context.items()
        if not is_blank(value)
    }
    safe_values.update(
        {
            "template_type": filename_part(template_info["template_type"], "document"),
            "template_label": filename_part(template_info["template_label"]),
            "output_format": filename_part(output_format),
            "recipient_name": filename_part(recipient),
            "payment_deadline": filename_part(first_value(context, "payment_deadline")),
            "date": filename_part(document_date),
            "timestamp": timestamp,
            "board_id": filename_part((board_setup or {}).get("board_id", "")),
            "item_id": filename_part(first_value(context, "item_id")),
            "item_name": filename_part(first_value(context, "item_name", "row_name", "pulse_name")),
        }
    )

    configured_template = str((board_setup or {}).get("file_name_template", "")).strip()
    if configured_template:
        candidate = Template(configured_template).safe_substitute(safe_values).strip()
        if candidate and "${" not in candidate:
            candidate = filename_part(candidate, "document")
            if candidate:
                extension = OUTPUT_FORMAT_REGISTRY[output_format]["extension"]
                return f"{candidate}.{extension}"

    parts = [
        filename_part(template_info["template_type"], "document"),
        filename_part(recipient),
        filename_part(document_date),
        timestamp,
    ]
    extension = OUTPUT_FORMAT_REGISTRY[output_format]["extension"]
    base_name = "-".join(part for part in parts if part) or "document"
    return f"{base_name}.{extension}"


def build_document(
    document_data: Dict[str, Any],
    template_type: str,
    output_format: str,
) -> Tuple[bytes, Dict[str, Any], Dict[str, Any]]:
    normalized_type = normalize_template_type(template_type)
    normalized_output_format = normalize_output_format(output_format)
    config = template_config(normalized_type)
    fields = template_fields(normalized_type, normalized_output_format)
    context = build_template_context(document_data, normalized_type, normalized_output_format)

    if normalized_output_format == "pdf":
        html_content = render_template_html(normalized_type, context)
        document_bytes = html_to_pdf_bytes(html_content)
    else:
        document_bytes = docx_template_to_bytes(normalized_type, context)

    empty_fields = [
        field
        for field in fields
        if is_blank(context.get(field))
    ]

    return (
        document_bytes,
        {
            "template_type": normalized_type,
            "template_label": config["label"],
            "template_file": config["file"],
            "docx_template_file": DOCX_TEMPLATE_REGISTRY.get(normalized_type, {}).get("file", ""),
            "template_fields": fields,
            "empty_fields": empty_fields,
            "output_format": normalized_output_format,
            "output_label": OUTPUT_FORMAT_REGISTRY[normalized_output_format]["label"],
        },
        context,
    )


def upload_generated_file_to_column(
    item_id: str,
    file_column_id: str,
    file_bytes: bytes,
    filename: str,
    config: Dict[str, Any],
) -> Dict[str, Any]:
    mime_type = mimetypes.guess_type(filename)[0] or "application/octet-stream"
    query = (
        f'mutation ($file: File!) {{ '
        f'add_file_to_column (item_id: {int(item_id)}, column_id: "{file_column_id}", file: $file) '
        f'{{ id }} }}'
    )

    with BytesIO(file_bytes) as generated_file:
        response = requests.post(
            config["file_api_url"],
            headers=monday_headers(config),
            data={"query": query},
            files={"variables[file]": (filename, generated_file, mime_type)},
            timeout=90,
        )

    response.raise_for_status()
    payload = response.json()

    if payload.get("errors"):
        messages = "; ".join(error.get("message", "Unknown monday file upload error") for error in payload["errors"])
        raise ValueError(messages)

    return payload.get("data", {}).get("add_file_to_column", {})


def webhook_trigger_already_processed(trigger_uuid: str) -> bool:
    trigger_key = str(trigger_uuid or "").strip()
    if not trigger_key:
        return False
    return trigger_key in PROCESSED_TRIGGER_UUIDS


def mark_webhook_trigger_processed(trigger_uuid: str) -> None:
    trigger_key = str(trigger_uuid or "").strip()
    if trigger_key:
        PROCESSED_TRIGGER_UUIDS.add(trigger_key)


def build_status_column_payload(
    board_setup: Dict[str, Any],
    status_key: str,
    message: str = "",
) -> Dict[str, Any]:
    monday_settings = board_setup.get("monday", {})
    status_column_id = str(monday_settings.get("status_column_id", "")).strip()
    result_message_column_id = str(monday_settings.get("result_message_column_id", "")).strip()
    status_labels = board_setup.get("status_labels", {})
    column_values: Dict[str, Any] = {}

    status_label = str(status_labels.get(f"{status_key}_label", "")).strip()
    status_index = status_labels.get(f"{status_key}_index")

    if status_column_id:
        if status_label:
            column_values[status_column_id] = {"label": status_label}
        elif status_index not in (None, ""):
            try:
                column_values[status_column_id] = {"index": int(status_index)}
            except (TypeError, ValueError):
                pass

    if result_message_column_id:
        column_values[result_message_column_id] = message

    return column_values


def update_webhook_item_status(
    item_id: str,
    board_id: str,
    config: Dict[str, Any],
    board_setup: Dict[str, Any],
    status_key: str,
    message: str = "",
) -> Dict[str, Any]:
    column_values = build_status_column_payload(board_setup, status_key, message=message)
    if not column_values:
        return {}
    return update_monday_item_columns(item_id, board_id, column_values, config)


@app.route("/webhooks/monday/file-generator", methods=["POST"])
def monday_file_generator_webhook():
    payload = request.get_json(silent=True) or {}
    if not isinstance(payload, dict):
        return jsonify({"success": False, "message": "Request body must be valid JSON."}), 400

    if payload.get("challenge") is not None:
        return jsonify({"challenge": payload.get("challenge")}), 200

    event = payload.get("event")
    if not isinstance(event, dict):
        return jsonify({"success": False, "message": "Webhook payload must include an event object."}), 400

    board_id = str(event.get("boardId", "")).strip()
    item_id = str(event.get("pulseId", "")).strip()
    trigger_uuid = str(event.get("triggerUuid", "")).strip()

    if not board_id or not item_id:
        return jsonify({"success": False, "message": "Webhook event is missing boardId or pulseId."}), 400

    if webhook_trigger_already_processed(trigger_uuid):
        return jsonify(
            {
                "success": True,
                "skipped": True,
                "message": "Webhook trigger was already processed.",
                "board_id": board_id,
                "item_id": item_id,
                "trigger_uuid": trigger_uuid,
            }
        ), 200

    monday_config: Optional[Dict[str, Any]] = None
    board_setup: Optional[Dict[str, Any]] = None

    try:
        monday_config = resolve_monday_api_config()
        validate_monday_api_config(monday_config)

        board_setup = resolve_board_document_setup(board_id)
        if not board_setup.get("enabled", True):
            return jsonify(
                {
                    "success": True,
                    "skipped": True,
                    "message": "Board document generation is disabled for this board.",
                    "board_id": board_id,
                    "item_id": item_id,
                }
            ), 200

        processing_status_error = ""
        try:
            update_webhook_item_status(
                item_id,
                board_id,
                monday_config,
                board_setup,
                "processing",
                message="Generating file...",
            )
        except Exception as exc:
            processing_status_error = str(exc)

        item_row = fetch_monday_item_row(item_id, board_id, monday_config)
        document_data = monday_item_to_document_data(item_row, board_setup)
        template_type = board_setup["template_type"]
        output_format = board_setup["output_format"]
        file_bytes, template_info, render_context = build_document(document_data, template_type, output_format)
        file_name = build_document_filename(template_info, render_context, output_format, board_setup=board_setup)
        upload_settings = board_setup.get("upload", {})

        monday_settings = board_setup.get("monday", {})
        file_column_id = str(monday_settings.get("file_column_id", "")).strip()
        upload_generated_file = parse_bool(upload_settings.get("upload_generated_file"), True)
        replace_existing_file = parse_bool(upload_settings.get("replace_existing_file"), True)
        file_clear_response = None
        file_upload_response = None

        if upload_generated_file:
            if not file_column_id:
                raise ValueError(
                    f"board_id '{board_id}' is configured to upload files, but monday.file_column_id is missing."
                )

            if replace_existing_file:
                file_clear_response = clear_monday_file_column(
                    item_id=item_id,
                    board_id=board_id,
                    file_column_id=file_column_id,
                    config=monday_config,
                )

            file_upload_response = upload_generated_file_to_column(
                item_id=item_id,
                file_column_id=file_column_id,
                file_bytes=file_bytes,
                filename=file_name,
                config=monday_config,
            )

        success_status_error = ""
        try:
            update_webhook_item_status(
                item_id,
                board_id,
                monday_config,
                board_setup,
                "success",
                message=f"Generated {file_name}",
            )
        except Exception as exc:
            success_status_error = str(exc)

        mark_webhook_trigger_processed(trigger_uuid)

        return jsonify(
            {
                "success": True,
                "message": "Webhook file generation completed successfully.",
                "board_id": board_id,
                "item_id": item_id,
                "trigger_uuid": trigger_uuid,
                "template_type": template_info["template_type"],
                "template_label": template_info["template_label"],
                "output_format": output_format,
                "generated_filename": file_name,
                "file_uploaded": bool(file_upload_response),
                "file_replaced": bool(file_clear_response),
                "file_clear_response": file_clear_response,
                "file_upload_response": file_upload_response,
                "empty_fields": template_info["empty_fields"],
                "processing_status_error": processing_status_error or None,
                "success_status_error": success_status_error or None,
            }
        ), 200
    except Exception as exc:
        error_message = str(exc)
        if monday_config and board_setup:
            try:
                update_webhook_item_status(
                    item_id,
                    board_id,
                    monday_config,
                    board_setup,
                    "error",
                    message=error_message,
                )
            except Exception:
                pass

        return jsonify(
            {
                "success": False,
                "message": "Webhook file generation failed.",
                "error": error_message,
                "board_id": board_id,
                "item_id": item_id,
                "trigger_uuid": trigger_uuid,
            }
        ), 500


@app.route("/health", methods=["GET"])
def health_check():
    weasyprint_available = True
    weasyprint_error = ""
    weasyprint_version = ""
    docxtpl_available = True
    docxtpl_error = ""

    try:
        import weasyprint

        weasyprint_version = getattr(weasyprint, "__version__", "")
    except Exception as exc:
        weasyprint_available = False
        weasyprint_error = repr(exc)

    try:
        import docxtpl  # noqa: F401
    except Exception as exc:
        docxtpl_available = False
        docxtpl_error = repr(exc)

    board_config = load_board_document_config()
    configured_board_ids = sorted(str(board_id) for board_id in board_config.get("boards", {}).keys())

    return jsonify(
        {
            "success": True,
            "message": "GLS monday webhook document backend is running.",
            "template_count": len(PDF_TEMPLATE_REGISTRY),
            "templates_dir_exists": TEMPLATE_DIR.exists(),
            "docx_templates_available": sorted(
                template_type
                for template_type in PDF_TEMPLATE_REGISTRY
                if docx_template_path(template_type)
            ),
            "docxtpl_available": docxtpl_available,
            "docxtpl_error": docxtpl_error,
            "board_document_config_path": str(BOARD_DOCUMENT_CONFIG_PATH),
            "configured_board_ids": configured_board_ids,
            "weasyprint_available": weasyprint_available,
            "weasyprint_version": weasyprint_version,
            "weasyprint_error": weasyprint_error,
        }
    ), 200
if __name__ == "__main__":
    app.run(
        host="0.0.0.0",
        port=int(os.getenv("PORT", "5000")),
        debug=os.getenv("FLASK_DEBUG", "false").lower() == "true",
    )
