# OnlyOffice Dashboard

A self-hosted web dashboard for your OnlyOffice Document Server. Upload, create, and open documents, spreadsheets, and presentations — all from a clean browser UI.

---

## Features

- **File dashboard** — grid view with type badges, size, and date
- **Upload** — click the button *or* drag & drop files onto the page
- **New document** — create blank `.docx`, `.xlsx`, or `.pptx` files
- **Open in editor** — full OnlyOffice editor opens in a new tab
- **Delete files** — remove files from the server
- **Filter & search** — filter by document type, search by filename
- **JWT-secured** — signs every editor config with your OnlyOffice JWT secret

---

## Quick Start

### Option A — Docker Compose (recommended)

```bash
# 1. Copy the env example (values are already pre-filled for your setup)
cp .env.example .env

# 2. Edit docker-compose.yml if your OnlyOffice is not on port 8080
#    ONLYOFFICE_URL: http://<your-onlyoffice-host>:<port>

# 3. Build & start
docker compose up -d --build
```

Open **http://localhost:3000** in your browser.

---

### Option B — Node.js directly

```bash
npm install

# Copy and edit the env file
cp .env.example .env

node server.js
```

---

## Configuration

| Variable | Default | Description |
|---|---|---|
| `ONLYOFFICE_URL` | `http://localhost:8080` | Document Server URL **as seen by the browser** |
| `APP_URL` | `http://host.docker.internal:3000` | Dashboard URL **as seen by the OnlyOffice container** (for callbacks & file download) |
| `PORT` | `3000` | Port the dashboard listens on |
| `JWT_SECRET` | *(pre-filled)* | Must match `services.CoAuthoring.secret.*.string` in your OnlyOffice `local.json` |

> **Important:** `APP_URL` must be reachable from inside the OnlyOffice container. If both containers are on the same Docker network, use the service name (e.g. `http://dashboard:3000`). When running Docker Desktop on Windows/Mac, `host.docker.internal` resolves to the host machine.

---

## Blank document templates

New document creation requires blank template files. Create them once:

```bash
mkdir -p templates

# Easiest way — copy any existing blank docx/xlsx/pptx here and rename them:
#   templates/blank.docx
#   templates/blank.xlsx
#   templates/blank.pptx
```

You can grab minimal blank templates from any Office installation or download them from the OnlyOffice GitHub sdkjs-forms repository.

---

## File storage

Uploaded files are stored in the `uploads/` directory. Mount it as a Docker volume so files survive container restarts (already configured in `docker-compose.yml`).

---

## Security notes

- Files are served over plain HTTP — run behind a reverse proxy (nginx / Traefik) with HTTPS in production.
- The JWT secret in this repo is your **live** secret — treat it like a password and do not commit it to a public repository.
- Multer limits uploads to 100 MB per file and only accepts known Office extensions.
