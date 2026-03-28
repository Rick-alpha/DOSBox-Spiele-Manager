@echo off
setlocal

set "PS_EXE="
if exist "%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe" set "PS_EXE=%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe"
if not defined PS_EXE if exist "%SystemRoot%\SysWOW64\WindowsPowerShell\v1.0\powershell.exe" set "PS_EXE=%SystemRoot%\SysWOW64\WindowsPowerShell\v1.0\powershell.exe"
if not defined PS_EXE set "PS_EXE=powershell.exe"

"%PS_EXE%" -NoProfile -ExecutionPolicy Bypass -File "%~dp0DOSBox-Spiele-Manager.ps1"

if errorlevel 9009 (
	where pwsh >nul 2>nul
	if not errorlevel 1 (
		pwsh -NoProfile -ExecutionPolicy Bypass -File "%~dp0DOSBox-Spiele-Manager.ps1"
	)
)

if errorlevel 1 (
	echo.
	echo Fehler beim Starten des DOSBox-Spiele-Managers.
	echo Bitte Fehlermeldung oben pruefen.
	pause
)
