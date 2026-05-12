# GLS monday Document Generator

Flask backend for monday.com webhook-driven document generation. It can generate PDF files from HTML templates and Word files from `.docx` templates, then upload the generated files back to monday Files columns.

## Current Flow

1. monday sends a webhook to `POST /webhooks/monday/file-generator`.
2. The backend reads `boardId` and `pulseId` from the webhook event.
3. The backend loads that board setup from `board_document_config.json`.
4. The backend fetches the monday row.
5. For every document configured for that board, it fills the correct template.
6. Each generated file is uploaded to its configured monday Files column.
7. The row status is updated to `Generated` or `Error`.

Generated files are kept in memory and uploaded directly. The production webhook flow does not save generated files locally.

## Main Files

- `app.py`: Flask webhook backend and generation logic.
- `board_document_config.json`: real board/document mapping.
- `board_document_config.example.json`: safe example config.
- `.env.example`: required environment variables.
- `templates/`: HTML templates used for PDF output.
- `dock_tempaltes/`: Word `.docx` templates used for DOCX output.
- `Dockerfile`: Cloud Run image with Japanese fonts and WeasyPrint native libraries.

## Environment

Create `.env` locally from `.env.example`.

Required:

- `MONDAY_API_TOKEN`

Optional:

- `MONDAY_API_URL`
- `MONDAY_FILE_API_URL`
- `MONDAY_API_VERSION`
- `BOARD_DOCUMENT_CONFIG_PATH`
- `PORT`

Do not commit or deploy `.env`. It is ignored by `.gitignore` and `.gcloudignore`.

## Board Config

`board_document_config.json` supports multiple documents for one board. Use one board ID once, then put all files for that board inside `documents`.

Important: JSON cannot keep repeated keys. If the same board ID appears multiple times, only the last one is used.

Example:

```json
{
  "boards": {
    "YOUR_BOARD_ID": {
      "enabled": true,
      "output_format": "docx",
      "file_name_template": "${template_type}-${item_name}-${bond_number}",
      "match_mode": "column_name",
      "monday": {
        "status_column_id": "YOUR_STATUS_COLUMN_ID",
        "result_message_column_id": ""
      },
      "upload": {
        "upload_generated_file": true,
        "replace_existing_file": true
      },
      "documents": [
        {
          "template_type": "allocation_notice_gmo",
          "monday": {
            "file_column_id": "FILE_COLUMN_FOR_ALLOCATION_NOTICE"
          }
        },
        {
          "template_type": "monthly_interest_notice",
          "monday": {
            "file_column_id": "FILE_COLUMN_FOR_MONTHLY_NOTICE"
          }
        }
      ]
    }
  }
}
```

Each document can override board-level settings such as `output_format`, `file_name_template`, `defaults`, `overrides`, `upload`, or `monday.file_column_id`.

## Variable Matching

The backend fetches the monday row and makes values available to templates using:

- monday column ID
- normalized monday column title
- row aliases like `item_name`, `pulse_name`, `row_name`, `item_id`, and `board_id`

Templates can use placeholders like:

```text
{{ recipient_name }}
{{ bond_number }}
{{ payment_deadline }}
```

Use `overrides` in `board_document_config.json` if a monday column title does not match the template variable name.

## Local Run

On macOS, install WeasyPrint native libraries once:

```bash
brew install pango
```

Run locally:

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
cp .env.example .env
PORT=5001 python app.py
```

Health check:

```bash
curl http://127.0.0.1:5001/health
```

Webhook test:

```bash
curl -X POST http://127.0.0.1:5001/webhooks/monday/file-generator \
  -H "Content-Type: application/json" \
  -d '{"event":{"boardId":"YOUR_BOARD_ID","pulseId":"YOUR_ITEM_ID","triggerUuid":"local-test-001"}}'
```

## Cloud Run Deploy

The included `Dockerfile` is the recommended deployment path because it installs Japanese fonts and native WeasyPrint libraries.

Deploy from source:

```bash
gcloud run deploy YOUR_SERVICE_NAME \
  --source . \
  --region YOUR_REGION \
  --allow-unauthenticated \
  --set-env-vars MONDAY_API_TOKEN=YOUR_TOKEN
```

Better production option: store `MONDAY_API_TOKEN` in Secret Manager and mount it as an environment variable.

After deploy:

1. Open the Cloud Run service URL.
2. Check `GET /health`.
3. Add monday webhook URL:

```text
https://YOUR_CLOUD_RUN_URL/webhooks/monday/file-generator
```

## Deployment Notes

- `Dockerfile` uses `gunicorn app:app` and Cloud Run `$PORT`.
- `.gcloudignore` excludes `.env`, local output files, virtualenv, and cache files.
- `board_document_config.json` is included in deployment, so keep it updated before deploy.
- File columns with `"Not found"` are treated as invalid and will fail only that document, while other documents continue.
- If `replace_existing_file` is `true`, the target monday Files column is cleared before uploading the new file.
