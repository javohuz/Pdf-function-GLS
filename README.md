# GLS Multi-Template PDF Service

Local Flask service that:

- receives JSON
- selects a real client HTML template by `template_type`
- fills `{{ dynamic_field }}` values safely with Jinja
- renders the HTML template to PDF with WeasyPrint
- optionally saves a local preview copy
- optionally creates a monday.com item and uploads the PDF
- names generated PDF files with template type, recipient, bond/date data, timestamp, and a short unique id

## Project Files

- `app.py`: backend API, template registry, PDF generation, monday upload flow
- `templates/`: real client HTML templates
- `local_tester.html`: standalone browser UI for testing each template
- `sample_request.json`: sample API payload
- `monday_config.example.json`: safe example config for local monday setup
- `Procfile`: Cloud Run source deploy entrypoint

## Template Types

- `allocation_notice`
- `allocation_notice_gmo`
- `application_form`
- `application_form_period`
- `condition_summary`
- `interest_calculation`
- `monthly_interest_notice`
- `issuance_terms_long`
- `payment_receipt`
- `terms_two_page`

## Local Run

On macOS, install WeasyPrint's native rendering libraries once:

```bash
brew install pango
```

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
PORT=5001 python app.py
```

Then open `local_tester.html` in your browser and set:

- `Backend URL` = `http://127.0.0.1:5001`

## API

List templates:

```bash
curl http://127.0.0.1:5001/templates
```

Generate a PDF:

```bash
curl -X POST http://127.0.0.1:5001/generate-pdf \
  -H "Content-Type: application/json" \
  --data @sample_request.json
```

Payload shape:

```json
{
  "template_type": "allocation_notice_gmo",
  "save_local_pdf_copy": false,
  "save_to_monday": false,
  "data": {
    "recipient_name": "山田 太郎"
  }
}
```

Missing or empty dynamic template fields render as blank strings.

## monday Local Config

Do not commit real secrets.

Create a local file named `monday_config.json` by copying:

```bash
cp monday_config.example.json monday_config.json
```

Then fill in:

- `api_token`
- `board_id`
- `file_column_id`

If `save_to_monday` is included in a request, that explicit value controls upload for that request. If it is omitted, the backend falls back to the monday config `enabled` value.

## Cloud Run Notes

This repo is prepared for Cloud Run source deploy:

- Python version pinned in `.python-version`
- production server in `requirements.txt`
- startup command in `Procfile`
- local secret file excluded in `.gcloudignore`

For Cloud Run, prefer environment variables and Secret Manager instead of `monday_config.json`.

WeasyPrint also needs native Pango/GLib libraries in the deployment image. If source deploy does not provide them, use a Docker-based Cloud Run deploy and install the native packages in the image.

## Security Note

If a real monday token was ever pasted into local files, terminal history, or chat, rotate it before production use.
