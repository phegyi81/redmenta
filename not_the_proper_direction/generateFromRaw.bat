@echo off
for /f "delims=|" %%f in ('dir /b *.html') do node html-parser.js "%%f"

