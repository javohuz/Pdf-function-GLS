import html
import json
import mimetypes
import os
import uuid
from datetime import datetime
from io import BytesIO
from pathlib import Path
from string import Template
from typing import Dict, List, Tuple

import requests
from flask import Flask, jsonify, make_response, request, send_from_directory
from reportlab.lib.colors import Color, HexColor
from reportlab.lib.enums import TA_CENTER, TA_LEFT
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.units import mm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.cidfonts import UnicodeCIDFont
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfgen import canvas
from reportlab.platypus import Paragraph


app = Flask(__name__)

BASE_DIR = Path(__file__).resolve().parent
OUTPUT_DIR = BASE_DIR / "output"
MONDAY_CONFIG_PATH = BASE_DIR / "monday_config.json"
TESTER_HTML_PATH = BASE_DIR / "local_tester.html"

EMBEDDED_FONT_NAME = "PrototypeJapaneseFont"
FALLBACK_HEADING_FONT_NAME = "HeiseiKakuGo-W5"
FALLBACK_BODY_FONT_NAME = "HeiseiMin-W3"

TEXT_COLOR = HexColor("#1A1A1A")
MUTED_TEXT_COLOR = HexColor("#555555")
LINE_COLOR = HexColor("#B9BDC4")
SOFT_FILL_COLOR = HexColor("#F5F6F7")

REQUIRED_FIELDS = [
    "date",
    "customer_name",
    "address",
    "company_name",
    "representative_name",
    "bond_title",
    "amount",
    "unit_count",
    "payment_date",
    "bank_name",
    "branch_name",
    "account_type",
    "account_number",
    "account_name",
]

LOCAL_FONT_CANDIDATES = [
    BASE_DIR / "fonts" / "NotoSansJP-Regular.ttf",
    BASE_DIR / "fonts" / "NotoSerifJP-Regular.otf",
    BASE_DIR / "fonts" / "NotoSansCJKjp-Regular.otf",
]

FONT_INFO_CACHE: Dict[str, str] = {}


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


def html_text(value: str) -> str:
    return html.escape(str(value or "")).replace("\n", "<br/>")


def validate_payload(payload: Dict) -> List[str]:
    missing = []
    for field in REQUIRED_FIELDS:
        value = payload.get(field)
        if value is None or str(value).strip() == "":
            missing.append(field)
    return missing


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
            "${customer_name}",
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
        config.get("item_name_template", "${customer_name}")
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
    safe_values = {key: str(value) for key, value in data.items() if value is not None}
    return template.safe_substitute(safe_values)


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


def get_font_info() -> Dict[str, str]:
    """
    Register Japanese-capable fonts and cache the result.

    Preferred:
    1. Bundled font file in ./fonts or PDF_JP_FONT_PATH
    2. ReportLab built-in Japanese CID fonts

    Bundling a real font file is the best production option because the PDF
    embeds it and stays more consistent across viewers and Cloud Run instances.
    """
    global FONT_INFO_CACHE
    if FONT_INFO_CACHE:
        return FONT_INFO_CACHE

    candidate_paths: List[Path] = []
    env_font_path = os.getenv("PDF_JP_FONT_PATH", "").strip()
    if env_font_path:
        candidate_paths.append(Path(env_font_path))
    candidate_paths.extend(LOCAL_FONT_CANDIDATES)

    last_error = None
    for font_path in candidate_paths:
        if not font_path.exists():
            continue

        try:
            if EMBEDDED_FONT_NAME not in pdfmetrics.getRegisteredFontNames():
                if font_path.suffix.lower() == ".ttc":
                    pdfmetrics.registerFont(
                        TTFont(EMBEDDED_FONT_NAME, str(font_path), subfontIndex=0)
                    )
                else:
                    pdfmetrics.registerFont(TTFont(EMBEDDED_FONT_NAME, str(font_path)))

            FONT_INFO_CACHE = {
                "heading_font": EMBEDDED_FONT_NAME,
                "body_font": EMBEDDED_FONT_NAME,
                "font_mode": "embedded_font_file",
                "font_source": str(font_path),
            }
            return FONT_INFO_CACHE
        except Exception as exc:
            last_error = str(exc)

    if FALLBACK_HEADING_FONT_NAME not in pdfmetrics.getRegisteredFontNames():
        pdfmetrics.registerFont(UnicodeCIDFont(FALLBACK_HEADING_FONT_NAME))
    if FALLBACK_BODY_FONT_NAME not in pdfmetrics.getRegisteredFontNames():
        pdfmetrics.registerFont(UnicodeCIDFont(FALLBACK_BODY_FONT_NAME))

    FONT_INFO_CACHE = {
        "heading_font": FALLBACK_HEADING_FONT_NAME,
        "body_font": FALLBACK_BODY_FONT_NAME,
        "font_mode": "reportlab_cid_fallback",
        "font_source": last_error or "built-in Japanese CID fonts",
    }
    return FONT_INFO_CACHE


def make_styles(font_info: Dict[str, str]) -> Dict[str, ParagraphStyle]:
    heading_font = font_info["heading_font"]
    body_font = font_info["body_font"]

    return {
        "sender": ParagraphStyle(
            "sender",
            fontName=body_font,
            fontSize=9.7,
            leading=14.0,
            alignment=TA_LEFT,
            textColor=TEXT_COLOR,
            wordWrap="CJK",
        ),
        "body": ParagraphStyle(
            "body",
            fontName=body_font,
            fontSize=10.6,
            leading=17.5,
            alignment=TA_LEFT,
            textColor=TEXT_COLOR,
            wordWrap="CJK",
        ),
        "table_value": ParagraphStyle(
            "table_value",
            fontName=body_font,
            fontSize=10.4,
            leading=16.0,
            alignment=TA_LEFT,
            textColor=TEXT_COLOR,
            wordWrap="CJK",
        ),
        "table_label": ParagraphStyle(
            "table_label",
            fontName=heading_font,
            fontSize=9.6,
            leading=14.0,
            alignment=TA_LEFT,
            textColor=MUTED_TEXT_COLOR,
            wordWrap="CJK",
        ),
        "title_meta": ParagraphStyle(
            "title_meta",
            fontName=heading_font,
            fontSize=9.2,
            leading=13.0,
            alignment=TA_CENTER,
            textColor=MUTED_TEXT_COLOR,
            wordWrap="CJK",
        ),
        "title": ParagraphStyle(
            "title",
            fontName=heading_font,
            fontSize=16.2,
            leading=21.5,
            alignment=TA_CENTER,
            textColor=TEXT_COLOR,
            wordWrap="CJK",
        ),
    }


def draw_text(
    pdf: canvas.Canvas,
    text: str,
    x: float,
    y: float,
    font_name: str,
    font_size: float,
    align: str = "left",
    color: Color = TEXT_COLOR,
) -> float:
    text = str(text or "")
    width = pdf.stringWidth(text, font_name, font_size)

    if align == "center":
        draw_x = x - (width / 2)
    elif align == "right":
        draw_x = x - width
    else:
        draw_x = x

    pdf.saveState()
    pdf.setFillColor(color)
    pdf.setFont(font_name, font_size)
    pdf.drawString(draw_x, y, text)
    pdf.restoreState()
    return width


def draw_paragraph(
    pdf: canvas.Canvas,
    text: str,
    x: float,
    top_y: float,
    width: float,
    style: ParagraphStyle,
) -> Tuple[float, float]:
    paragraph = Paragraph(html_text(text), style)
    wrapped_width, wrapped_height = paragraph.wrap(width, 200 * mm)
    paragraph.drawOn(pdf, x, top_y - wrapped_height)
    return wrapped_width, wrapped_height


def draw_rule(
    pdf: canvas.Canvas,
    x1: float,
    y1: float,
    x2: float,
    y2: float,
    stroke_color: Color = LINE_COLOR,
    line_width: float = 0.8,
) -> None:
    pdf.saveState()
    pdf.setStrokeColor(stroke_color)
    pdf.setLineWidth(line_width)
    pdf.line(x1, y1, x2, y2)
    pdf.restoreState()


def make_detail_rows(data: Dict) -> List[Tuple[str, str]]:
    return [
        ("社債名", data["bond_title"]),
        ("割当金額", f'{data["amount"]}（{data["unit_count"]}）'),
        ("払込期日", data["payment_date"]),
        ("金融機関", f'{data["bank_name"]}　{data["branch_name"]}'),
        ("預金種別", data["account_type"]),
        ("口座番号", data["account_number"]),
        ("口座名義", data["account_name"]),
    ]


def draw_detail_table(
    pdf: canvas.Canvas,
    rows: List[Tuple[str, str]],
    x: float,
    top_y: float,
    width: float,
    styles: Dict[str, ParagraphStyle],
    font_info: Dict[str, str],
) -> float:
    heading_font = font_info["heading_font"]
    label_style = styles["table_label"]
    value_style = styles["table_value"]

    header_height = 8.5 * mm
    label_width = 30 * mm
    inner_padding_x = 3.8 * mm
    inner_padding_y = 2.7 * mm
    value_width = width - label_width - (2 * inner_padding_x)

    prepared_rows = []
    total_rows_height = 0.0
    for label, value in rows:
        label_paragraph = Paragraph(html_text(label), label_style)
        _, label_height = label_paragraph.wrap(label_width - (2 * inner_padding_x), 40 * mm)
        value_paragraph = Paragraph(html_text(value), value_style)
        _, value_height = value_paragraph.wrap(value_width, 200 * mm)
        content_height = max(label_height, value_height)
        row_height = max(10.2 * mm, content_height + (2 * inner_padding_y))
        prepared_rows.append(
            {
                "label": label,
                "value": value,
                "label_paragraph": label_paragraph,
                "label_height": label_height,
                "paragraph": value_paragraph,
                "value_height": value_height,
                "row_height": row_height,
            }
        )
        total_rows_height += row_height

    table_height = header_height + total_rows_height
    table_bottom_y = top_y - table_height

    pdf.saveState()
    pdf.setFillColor(SOFT_FILL_COLOR)
    pdf.rect(x, top_y - header_height, width, header_height, fill=1, stroke=0)
    pdf.setStrokeColor(LINE_COLOR)
    pdf.setLineWidth(0.9)
    pdf.rect(x, table_bottom_y, width, table_height, fill=0, stroke=1)
    pdf.restoreState()

    draw_text(
        pdf,
        "払込内容",
        x + (width / 2),
        top_y - (header_height / 2) - 4.0,
        heading_font,
        9.8,
        align="center",
        color=TEXT_COLOR,
    )

    current_top_y = top_y - header_height
    separator_x = x + label_width

    for index, row in enumerate(prepared_rows):
        row_bottom_y = current_top_y - row["row_height"]

        draw_rule(pdf, separator_x, current_top_y, separator_x, row_bottom_y, line_width=0.55)
        if index < len(prepared_rows) - 1:
            draw_rule(pdf, x, row_bottom_y, x + width, row_bottom_y, line_width=0.55)

        label_top_y = current_top_y - ((row["row_height"] - row["label_height"]) / 2)
        row["label_paragraph"].drawOn(
            pdf,
            x + inner_padding_x,
            label_top_y - row["label_height"],
        )

        value_top_y = current_top_y - ((row["row_height"] - row["value_height"]) / 2)
        row["paragraph"].drawOn(
            pdf,
            separator_x + inner_padding_x,
            value_top_y - row["value_height"],
        )

        current_top_y = row_bottom_y

    return table_bottom_y

def build_pdf(data: Dict) -> Tuple[bytes, Dict[str, str]]:
    font_info = get_font_info()
    styles = make_styles(font_info)
    heading_font = font_info["heading_font"]
    body_font = font_info["body_font"]

    pdf_buffer = BytesIO()
    pdf = canvas.Canvas(pdf_buffer, pagesize=A4)
    page_width, page_height = A4

    left_margin = 24 * mm
    right_margin = page_width - (24 * mm)
    content_width = page_width - (48 * mm)

    pdf.setTitle("Bond Allocation Notice Prototype")
    pdf.setAuthor("Local Flask PDF Prototype")

    draw_rule(pdf, left_margin, page_height - (19 * mm), right_margin, page_height - (19 * mm), line_width=1.0)

    draw_text(
        pdf,
        data["date"],
        right_margin,
        page_height - (29 * mm),
        body_font,
        10.0,
        align="right",
        color=MUTED_TEXT_COLOR,
    )

    customer_line_y = page_height - (49 * mm)
    draw_text(
        pdf,
        f'{data["customer_name"]}　様',
        left_margin,
        customer_line_y,
        body_font,
        12.8,
    )
    draw_rule(pdf, left_margin, customer_line_y - 2.1 * mm, left_margin + (50 * mm), customer_line_y - 2.1 * mm, line_width=0.7)

    sender_block = (
        f'{data["address"]}\n'
        f'{data["company_name"]}\n'
        f'代表取締役　{data["representative_name"]}'
    )
    draw_paragraph(
        pdf,
        sender_block,
        page_width - (86 * mm),
        page_height - (39 * mm),
        62 * mm,
        styles["sender"],
    )

    title_top_y = page_height - (74 * mm)
    draw_paragraph(
        pdf,
        data["company_name"],
        left_margin,
        title_top_y,
        content_width,
        styles["title_meta"],
    )
    draw_paragraph(
        pdf,
        f'{data["bond_title"]}　割当決定通知書',
        left_margin,
        title_top_y - (8.5 * mm),
        content_width,
        styles["title"],
    )
    draw_rule(pdf, left_margin + (28 * mm), page_height - (94 * mm), right_margin - (28 * mm), page_height - (94 * mm), line_width=0.8)

    body_text = (
        f'拝啓　平素より格別のご高配を賜り、厚く御礼申し上げます。'
        f'このたびは当社の{data["bond_title"]}へお申込みいただき、誠にありがとうございました。'
        "申込内容に基づき割当を決定いたしましたので、下記の内容をご確認のうえ、"
        "払込期日までにお手続きくださいますようお願い申し上げます。"
    )
    draw_paragraph(
        pdf,
        body_text,
        left_margin + (4 * mm),
        page_height - (104 * mm),
        content_width - (8 * mm),
        styles["body"],
    )

    draw_text(
        pdf,
        "記",
        page_width / 2,
        page_height - (136 * mm),
        heading_font,
        12,
        align="center",
    )

    table_bottom_y = draw_detail_table(
        pdf,
        make_detail_rows(data),
        left_margin,
        page_height - (144 * mm),
        content_width,
        styles,
        font_info,
    )

    note_style = ParagraphStyle(
        "note",
        fontName=body_font,
        fontSize=9.0,
        leading=13.0,
        alignment=TA_LEFT,
        textColor=MUTED_TEXT_COLOR,
        wordWrap="CJK",
    )
    note_text = (
        "なお、本通知書は払込内容のご案内を目的としたものです。"
        "ご不明な点がございましたら、発行会社までお問い合わせください。"
    )
    note_top_y = table_bottom_y - (7 * mm)
    _, note_height = draw_paragraph(
        pdf,
        note_text,
        left_margin,
        note_top_y,
        content_width,
        note_style,
    )
    note_bottom_y = note_top_y - note_height
    footer_y = max(8 * mm, min(16 * mm, note_bottom_y - (7 * mm)))

    draw_text(
        pdf,
        "以上",
        right_margin,
        footer_y,
        body_font,
        10.0,
        align="right",
    )

    pdf.showPage()
    pdf.save()
    return pdf_buffer.getvalue(), font_info


@app.route("/generate-pdf", methods=["POST", "OPTIONS"])
def generate_pdf():
    if request.method == "OPTIONS":
        return make_response("", 204)

    payload = request.get_json(silent=True)
    if not isinstance(payload, dict):
        return jsonify({"success": False, "message": "Request body must be valid JSON."}), 400

    missing_fields = validate_payload(payload)
    if missing_fields:
        return (
            jsonify(
                {
                    "success": False,
                    "message": "Missing required fields.",
                    "missing_fields": missing_fields,
                }
            ),
            400,
        )

    runtime_config = resolve_runtime_config(payload)
    timestamp = datetime.utcnow().strftime("%Y%m%dT%H%M%S")
    file_name = f"bond_notice_{timestamp}_{uuid.uuid4().hex[:8]}.pdf"
    pdf_path = OUTPUT_DIR / file_name
    pdf_bytes = b""
    local_pdf_path = ""

    try:
        pdf_bytes, font_info = build_pdf(payload)
    except Exception as exc:
        return (
            jsonify(
                {
                    "success": False,
                    "message": "Failed to generate PDF.",
                    "error": str(exc),
                }
            ),
            500,
        )

    if runtime_config.get("save_local_pdf_copy"):
        ensure_output_dir()
        pdf_path.write_bytes(pdf_bytes)
        local_pdf_path = str(pdf_path)

    monday_requested = parse_bool(payload.get("save_to_monday"), False)
    monday_result = None
    monday_error = None

    try:
        monday_config = resolve_monday_config(payload)
        monday_requested = monday_requested or monday_config.get("enabled", False)
    except Exception as exc:
        monday_config = None
        if monday_requested:
            monday_error = str(exc)

    if monday_requested:
        if monday_error is None:
            try:
                monday_result = upload_to_monday(pdf_bytes, file_name, payload, local_pdf_path=local_pdf_path)
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
                "heading_font": font_info["heading_font"],
                "body_font": font_info["body_font"],
                "font_mode": font_info["font_mode"],
                "font_source": font_info["font_source"],
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
    # This keeps the app close to a future serverless deployment:
    # - Local run: python app.py
    # - Cloud Run: gunicorn app:app
    # - Cloud Functions 2nd gen: adapt the same logic to an HTTP entrypoint
    app.run(
        host="0.0.0.0",
        port=int(os.getenv("PORT", "5000")),
        debug=os.getenv("FLASK_DEBUG", "false").lower() == "true",
    )
