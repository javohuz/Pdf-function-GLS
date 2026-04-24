# PDF to monday Prototype

Local Flask service that:

- receives JSON
- generates a Japanese PDF notice in memory
- uploads the PDF to monday.com
- is prepared for Cloud Run deployment

## Project Files

- `app.py`: backend API
- `local_tester.html`: standalone browser UI for local/manual testing
- `sample_request.json`: sample API payload
- `monday_config.example.json`: safe example config for local monday setup
- `Procfile`: Cloud Run source deploy entrypoint

## Local Run

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
PORT=5001 python app.py
```

Then open `local_tester.html` in your browser and set:

- `Backend URL` = `http://127.0.0.1:5001`

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

## API Endpoint

`POST /generate-pdf`

Example:

```bash
curl -X POST http://127.0.0.1:5001/generate-pdf \
  -H "Content-Type: application/json" \
  --data @sample_request.json
```

## Cloud Run Notes

This repo is prepared for Cloud Run source deploy:

- Python version pinned in `.python-version`
- production server in `requirements.txt`
- startup command in `Procfile`
- local secret file excluded in `.gcloudignore`

For Cloud Run, prefer environment variables and Secret Manager instead of `monday_config.json`.

## Push To GitHub

```bash
git init -b main
git add .
git commit -m "Initial commit"
```

Then create a GitHub repo and connect it:

```bash
git remote add origin https://github.com/YOUR-USER/YOUR-REPO.git
git push -u origin main
```

## Security Note

If a real monday token was ever pasted into local files, terminal history, or chat, rotate it before production use.
