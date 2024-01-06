@echo off
pyinstaller --clean --onefile --noconsole --add-data "res/icon.ico;res" --icon "res/icon.ico" main.py
pause
