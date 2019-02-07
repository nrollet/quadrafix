SET MYPY=%LOCALAPPDATA%/Programs/Python/Python36-32/python.exe
SET ROOT="C:\Users\nicolas\Documents\GIT\quadrafix"

cd %ROOT%
%MYPY% -m PyInstaller -p . --onefile .\py\quadrafix.py
