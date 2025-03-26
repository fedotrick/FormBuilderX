import PyInstaller.__main__
import os

PyInstaller.__main__.run([
    'form_generator_gui.py',
    '--name=Генератор маршрутных карт. Версия 2.0.0',
    '--onedir',
    '--windowed',
    '--icon=app.ico',
    '--add-data=ШАБЛОН.pptx;.',
    '--add-data=справочник.db;.',
    '--add-data=история_форм.db;.',
    '--add-data=app.ico;.',
    '--clean'
]) 