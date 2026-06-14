from pathlib import Path


app_path = Path(__file__).with_name("app-general.py")
exec(app_path.read_text(encoding="utf-8"), globals())
