# Ride Report Tracker — Vercel + Render + PostgreSQL

This repository separates the static frontend (`frontend/`) from the Flask API (`app.py`). Render hosts the API and managed PostgreSQL; Vercel hosts the frontend. The existing Data Logger application is not modified.

## Important deployment/security notes

- `render.yaml` creates a paid Render web service and a paid managed PostgreSQL database. Review current pricing and region before deploying.
- Set `IMPORT_TOKEN` as a secret in Render. The current sample static frontend does not provide production-grade user authentication. Do **not** place the import token in public `frontend/config.js` for a public production site. For production, use an authenticated upload proxy or implement user authentication before enabling browser uploads.
- Set `FRONTEND_ORIGIN` to the exact Vercel deployment origin. The current simple CORS header is intended for a single origin.
- Keep the database URL and import token private. Never commit secrets.

## Render backend

1. Create a new Render Blueprint from this repository, or create a Python web service using this repository.
2. Use `pip install -r requirements.txt` and `gunicorn app:app --bind 0.0.0.0:$PORT --workers 2 --timeout 120`.
3. Create managed PostgreSQL and set `DATABASE_URL` to its internal connection string. Set `IMPORT_TOKEN` to a long random secret and `FRONTEND_ORIGIN` to the Vercel origin.
4. Verify `https://YOUR-API/health` returns `{"ok":true,...}`.

## Vercel frontend

1. Import this GitHub repository into Vercel.
2. Set Root Directory to `frontend` and Framework Preset to Other; no build command is needed and output directory is `.`.
3. Add `config.js` in `frontend/` using `config.js.example` as a template and set `API_BASE_URL` to the Render API URL.
4. Do not put `IMPORT_TOKEN` in a public frontend. The upload button is a development scaffold until authentication/proxying is implemented. Download and status calls can be used without a token.

## API

- `GET /health` health check
- `GET /api/status` record count
- `POST /api/import` multipart field `file` or JSON `{filename,csv}`; requires `X-Import-Token`
- `PATCH /api/reports/<sha256>` JSON field/value corrections; requires `X-Import-Token`
- `GET /download` editable `.xlsx` export

Exact-byte duplicate imports are ignored. A repeated Drive ID with different CSV content is retained and flagged as `drive_id_conflict`; it is not automatically merged. Corrections are stored separately and overlaid on workbook exports so later source imports do not overwrite those corrections.

## Tracker layout

The workbook uses the screenshot-derived 45-column A:AS layout. Metadata absent from the Data Logger CSV (for example Route, State/City, Start/End Point, Vehicle, and Z Frame Checker) remains blank for later completion.

## Testing

Run `python -m pytest` after installing requirements. Before using real ride reports, validate parsing against representative CSV exports, verify date ordering and totals, test correction persistence, and confirm your organization permits uploading the data to this cloud service.
