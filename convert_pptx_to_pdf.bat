@echo off
setlocal enabledelayedexpansion

echo Converting all .pptx files to PDF...

for /r %%f in (*.pptx) do (
    echo Processing: %%f
    cscript //nologo "%~dp0pptx_to_pdf.vbs" "%%f"
)

echo Done!
pause
