# RoboMANTAP HTMX KUIS — Vercel runtime

This repository contains only the student-facing FastAPI/HTMX runtime for KUIS. It does not modify the main RoboMANTAP repository.

## Vercel setup
- Framework Preset: Other
- Root Directory: `./`
- Build/Output Command: leave unset (if the dashboard allows it)
- Environment Variables: `DATABASE_URL`, `GEMINI_API_KEYS` (or `GEMINI_API_KEY`), `STREAMLIT_URL=https://robomantap-intelligence.vercel.app/`
- Redeploy after changing variables.

`GEMINI_API_KEYS` should be a one-line JSON array, e.g. `["key-one","key-two"]` without the invisible characters shown here; use your real values and never commit them. A single `GEMINI_API_KEY` is also supported.

## Important
- The main route is served by FastAPI from `api/index.py`.
- PostgreSQL must be reachable from Vercel. Do not commit `.env` or credentials.
- Current session dictionaries are still in memory. This package is for deployment testing, not yet production-safe for concurrent/high-stakes exams until session state is persisted and load-tested.

## Smoke test
After deployment, open `/healthz`. It should return `status: ok` and only boolean flags for configured environment variables. If `/healthz` returns FastAPI 404, check the Vercel function routing/build log.
