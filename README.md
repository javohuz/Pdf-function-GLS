# GLS monday File Generator

Small Flask backend for monday.com webhook-driven document generation.

## What It Does

The service:

1. receives a monday webhook
2. reads the `boardId` and `pulseId`
3. loads that board's document setup from `board_document_config.json`
4. fetches only that one monday row
5. fills the matching PDF or DOCX template
6. removes the old file from the monday Files column
7. uploads the new file back to the same row

It does not save generated files locally in production flow.

## Main Files

- `app.py`: webhook backend and document generation logic
- `board_document_config.json`: board-by-board document settings
- `board_document_config.example.json`: example config for multiple boards
- `.env.example`: monday API environment variables
- `templates/`: HTML templates for PDF generation
- `dock_tempaltes/`: DOCX templates for Word generation
- `Dockerfile`: Cloud Run image with required native libraries

## Config

### 1. monday API config

Create a local `.env` from `.env.example`.

Required:

- `MONDAY_API_TOKEN`

Optional:

- `MONDAY_API_URL`
- `MONDAY_FILE_API_URL`
- `MONDAY_API_VERSION`
- `PORT`

### 2. Board document config

`board_document_config.json` controls:

- which board uses which template
- output type: `pdf` or `docx`
- monday file column id
- monday status column id
- filename template
- default values
- optional overrides
- whether old files are replaced before upload

Example:

```json
{
  "boards": {
    "5095195874": {
      "enabled": true,
      "template_type": "terms_two_page",
      "output_format": "docx",
      "file_name_template": "${item_name}-${bond_number}",
      "monday": {
        "file_column_id": "file_mm2qgmt1",
        "status_column_id": "color_mm34cpyx"
      },
      "upload": {
        "upload_generated_file": true,
        "replace_existing_file": true
      }
    }
  }
}
```

## Variable Matching

The backend fetches the monday row and passes row data into the template.

Best options for template placeholders:

- monday column id, like `{{ text_mm34gq11 }}`
- monday column title if it is stable
- built-in item aliases like `{{ pulseName }}`, `{{ item_name }}`, `{{ boardId }}`

Extra monday fields are ignored if the template does not use them.

## Local Run

On macOS, install WeasyPrint native libraries once:

```bash
brew install pango
```

Then run:

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

## Webhook Test

The production endpoint is:

```text
POST /webhooks/monday/file-generator
```

Local example:

```bash
curl -X POST http://127.0.0.1:5001/webhooks/monday/file-generator \
  -H "Content-Type: application/json" \
  -d '{"event":{"boardId":5095195874,"pulseId":2873011595,"triggerUuid":"test-001"}}'
```

## Cloud Run

This project is prepared for Cloud Run.

Recommended:

- keep monday token in Secret Manager or env vars
- keep board logic in `board_document_config.json`
- use the included `Dockerfile` for WeasyPrint and Japanese font support

Deploy:

```bash
gcloud run deploy YOUR_SERVICE_NAME --source . --allow-unauthenticated
```

## Notes

- Repeated generation replaces the old file in the monday Files column when `replace_existing_file` is `true`.
- DOCX output should use real `.docx` templates in `dock_tempaltes/`.
- PDF output uses the HTML templates in `templates/`.
