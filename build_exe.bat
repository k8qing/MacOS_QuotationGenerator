@echo off
REM Double-click on Windows to build KeithsQuotationGenerator.exe  (needs Python 3.10+ installed)
py -m pip install -r requirements.txt
if errorlevel 1 goto :fail
py build_exe.py
if errorlevel 1 goto :fail
echo.
echo Finished. Your app is in the "release" folder.
pause
exit /b 0
:fail
echo Build failed - see the messages above.
pause
