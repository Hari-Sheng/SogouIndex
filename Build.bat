@echo off
cd /d %~dp0
pyinstaller main.py
::pyinstaller -F --icon=.\search_web.ico main.py
copy .\dist\main.exe .\SogouIndex.exe
pause