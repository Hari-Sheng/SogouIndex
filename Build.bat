@echo off
cd /d %~dp0
<<<<<<< HEAD
::pyinstaller -F main.py
pyinstaller -F --icon=.\SogouIndex.ico main.py
=======
pyinstaller main.py
::pyinstaller -F --icon=.\search_web.ico main.py
>>>>>>> origin/master
copy .\dist\main.exe .\SogouIndex.exe
pause