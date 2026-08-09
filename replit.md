# LabVault

LabVault is a professional, AI-assisted laboratory inventory workspace for students and lab assistants. It tracks real equipment and consumables from the uploaded workbooks through request, approval, collection, issue, and return workflows.

## Run & Operate

- `python app.py` — run the Flask application (port 5000)
- `python import_inventory.py` — safely refresh inventory from the uploaded workbooks
- `pnpm run typecheck` — full typecheck across all packages
- `pnpm run build` — typecheck + build all packages
- `pnpm --filter @workspace/api-spec run codegen` — regenerate API hooks and Zod schemas from the OpenAPI spec
- `pnpm --filter @workspace/db run push` — push DB schema changes (dev only)
- Required env: `SECRET_KEY` and `MAIL_PASSWORD` are managed securely; SQLite is used by default for a self-contained deploy and SQL Server connection values are documented in `.env.example`.

## Stack

- pnpm workspaces, Node.js 24, TypeScript 5.9
- Application: Flask + Jinja2
- Database: SQLite by default, with SQL Server settings reserved in `.env.example`
- Frontend: server-rendered HTML, custom CSS, and vanilla JavaScript
- Inventory source: openpyxl import from the two attached Excel workbooks
- Email: SMTP with `labvault.lab@gmail.com` as the sender and secure approval links

## Where things live

- `app.py` — Flask routes, schema creation, inventory import, approval tokens, OTPs, email, and persistence
- `templates/` — server-rendered student, lab assistant, request, approval, returns, and landing pages
- `static/css/style.css` — LabVault light/dark visual system and responsive layout
- `static/js/app.js` — theme persistence, password visibility, modal, and rejection interactions
- `attached_assets/` — source inventory workbooks and product reference image

## Architecture decisions

- Inventory is imported idempotently using a stable source key derived from source file, stock/register number or row identity, and normalized name.
- Email links open a time-limited signed review page; only an explicit POST from that page changes approval state.
- SQLite keeps the first deployment self-contained and durable; all credentials and mail settings are environment-driven.
- OTP values are shown only once at generation time and only SHA-256 hashes are stored.

## Product

Students can register with a UID, search real lab stock, create multi-item requests, use the project assistant, track approvals, generate collection/return OTPs, and see active issues. The Lab Assistant can review metrics, issue approved requests, process returns, and add or deactivate inventory.

## User preferences

- Use `labvault.lab@gmail.com` as the official sender address. Keep its password in secure environment settings only.
- Keep the visual language professional and inspired by the uploaded LabVault reference without copying it pixel-for-pixel.

## Gotchas

- The application workflow is `LabVault: python app.py`; restart it after changing server-side code.
- SMTP delivery is intentionally skipped when mail settings are missing, while requests remain persisted and visible.
- The default admin password must be replaced with `ADMIN_PASSWORD` before a public deployment.

## Pointers

- See the `pnpm-workspace` skill for workspace structure, TypeScript setup, and package details
