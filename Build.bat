@echo off
cd /d %~dp0
::pyinstaller -F main.py
pyinstaller -F --icon=.\SogouIndex.ico main.py
copy .\dist\main.exe .\SogouIndex.exe
pause