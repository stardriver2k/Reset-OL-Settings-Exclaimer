:: Zurücksetzen von Outlook-Einstellungen (nur bei Problemen mit Exclaimer Cloud Signature Agent ab August 2020)
:: Autor: Alexander Karls, pegasus GmbH, Stand 10.08.20

@echo off
echo Vorhandene Outlook-Einstellungen werden geloescht. Bitte mit erhöhten Rechten starten, falls notwendig!
pause


:: Outlook stoppen
taskkill /F /IM outlook.exe

:: Office 2016
reg delete "HKEY_CURRENT_USER\Software\Microsoft\Office\15.0\Outlook\Options\Mail" /v "Send Pictures With Document" /f

:: Office 2019
reg delete "HKEY_CURRENT_USER\Software\Microsoft\Office\16.0\Outlook\Options\Mail" /v "Send Pictures With Document" /f

:: Outlook starten
:CheckOS
IF EXIST "%PROGRAMFILES(X86)%" (GOTO 64BIT) ELSE (GOTO 32BIT)

:64BIT
echo 64-bit Betriebssystem gefunden...
"C:\Program Files\Microsoft Office\Office15\OUTLOOK.EXE"
"C:\Program Files\Microsoft Office\Office16\OUTLOOK.EXE" 
GOTO END

:32BIT
echo 32-bit Betriebssystem gefunden...
"C:\Program Files (x86)\Microsoft Office\Office15\OUTLOOK.EXE" 
"C:\Program Files (x86)\Microsoft Office\Office16\OUTLOOK.EXE"
GOTO END

:END
echo Standardeinstellungen wiederhergestellt, Outlook bitte neu starten
pause