import sys
from cx_Freeze import setup, Executable

# Зависимости, которые нужно включить
build_exe_options = {
    "packages": ["PySide6", "pptx", "qrcode", "sqlite3"],
    "include_files": [
        "app.ico",
        "ШАБЛОН.pptx",
        "справочник.db",
        "история_форм.db"
    ],
    "excludes": [],
    "include_msvcr": True
}

# Базовое имя для exe
base = None
if sys.platform == "win32":
    base = "Win32GUI"

setup(
    name="Генератор маршрутных карт",
    version="1.0",
    description="Программа для генерации маршрутных карт",
    options={"build_exe": build_exe_options},
    executables=[
        Executable(
            "form_generator_gui.py",
            base=base,
            target_name="Генератор маршрутных карт.exe",
            icon="app.ico"
        )
    ]
) 