@echo off
rem Builds the DOCXClass samples: docxwriter, docxtemplate, docxreader
rem   build                 with the default C compiler of hbmk2
rem   build -comp=bcc       other hbmk2 options are passed on (also HBMK_OPTIONS)
rem HBMK2  hbmk2.exe to use (default: hbmk2 from the PATH)
setlocal enabledelayedexpansion
cd /d "%~dp0"
if not defined HBMK2 set HBMK2=hbmk2
set FAILED=
for %%p in (docxwriter docxtemplate docxreader) do (
   echo.
   echo === %%p
   "%HBMK2%" %%p.hbp %*
   if errorlevel 1 set FAILED=!FAILED! %%p
)
echo.
if defined FAILED (
   echo Not built:%FAILED%
) else (
   echo Built: docxwriter docxtemplate docxreader
)
rem started with a double click: keep the window open
echo %CMDCMDLINE% | find /i "/c" >nul && pause
if defined FAILED exit /b 1
exit /b 0
