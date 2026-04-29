# Weekly Dashboard Project Structure

Keep this entire dashboard inside one dedicated project folder:

`weekly_dashboard/`

This helps prevent overlap with other projects and keeps Drive/local storage clean.

## Recommended layout

- `app_streamlit.py`
  Main Streamlit dashboard app.
- `app.py`
  FastAPI app for the static dashboard/API flow.
- `start_dashboard.sh`
  Local startup script.
- `scripts/`
  Utility scripts for generating sample or support data.
- `data/`
  Active dashboard data files used by the app.
- `data/backups/`
  Auto-generated recovery backups only.
- `pmo_storage/`
  Project-specific PMO storage and Drive-mirrored content.
- `static/`
  HTML, CSS, and JavaScript assets for the FastAPI version.
- `requirements.txt`
  Python dependencies.

## Folder rules

- Keep all dashboard files inside `weekly_dashboard/` only.
- Do not mix files from different projects into this folder.
- Save live app data in `data/`, not in the root folder.
- Save backups in `data/backups/`, not beside live data files.
- Keep project/topic folders inside `pmo_storage/` only.
- Keep utility or one-time scripts inside `scripts/`, not in the root.
- Avoid unclear names like `final`, `copy`, `new`, or `fix_backup` for active files.
- Use one clear file name for each active app entry point.
