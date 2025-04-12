@echo off
rd /s /q ".\build"
rd /s /q ".\dist"
del "main.spec"
del "..\main.exe"
del "..\main.log"
rd /s /q ".\logs"

pip install --upgrade pyinstaller
pyinstaller --onefile --noconsole --strip  main.py
move ".\dist\main.exe" ".\"

rd /s /q ".\build"
rd /s /q ".\dist"
del "main.spec"
del "..\main.exe"
del "..\main.log"
rd /s /q ".\logs"