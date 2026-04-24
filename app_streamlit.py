from pathlib import Path

_BACKUP_APP = Path(__file__).with_name("app_streamlit.py.fix_backup")
exec(compile(_BACKUP_APP.read_text(encoding="utf-8"), str(_BACKUP_APP), "exec"))
