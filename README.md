# GLS monday Webhook Document Service

Flask service that:

- receives a monday webhook
- finds the board setup from `board_document_config.json`
- fetches the changed row from monday
- matches monday column names to template variables
- renders the document as PDF or DOCX
- uploads the generated file back to the same monday item

## Project Files

- `app.py`: webhook backend, template registry, document generation, monday upload flow
- `templates/`: real client HTML templates
- `dock_tempaltes/`: DOCX templates
- `board_document_config.json`: real board-to-document routing config
- `.env.example`: environment variable template for monday API config
- `Procfile`: Cloud Run source deploy entrypoint
- `Dockerfile`: Cloud Run container image with WeasyPrint native Linux libraries

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

Create your local environment file:

```bash
cp .env.example .env
```

Then fill in:

- `MONDAY_API_TOKEN`
- `MONDAY_API_URL` if different
- `MONDAY_FILE_API_URL` if different

## API

Health check:

```bash
curl http://127.0.0.1:5001/health
```

Production webhook endpoint:

```bash
POST /webhooks/monday/file-generator
```

## Cloud Run Notes

This repo is prepared for Cloud Run deploy:

- Python version pinned in `.python-version`
- production server in `requirements.txt`
- startup command in `Procfile`
- local env file excluded in `.gcloudignore`
- Docker image installs WeasyPrint's native Linux libraries and Japanese fonts

For Cloud Run, prefer environment variables and Secret Manager instead of local files.

WeasyPrint also needs native Pango/GLib/font libraries in the deployment image. The included `Dockerfile` installs the Debian packages needed by WeasyPrint plus `fonts-noto-cjk` for Japanese output.

Deploy with Dockerfile-based Cloud Run build:

```bash
gcloud run deploy YOUR_SERVICE_NAME \
  --source . \
  --allow-unauthenticated
```

If Cloud Run was connected to GitHub before this file existed, confirm the build is using the new `Dockerfile`, then redeploy after pushing.

## Security Note

If a real monday token was ever pasted into old local files, terminal history, or chat, rotate it before production use.
