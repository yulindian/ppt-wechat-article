@echo off
chcp 65001 >nul
py -3 "%~dp0copy_docx_to_clipboard.py" "%~dp0公众号文章导入稿.docx" --expected-images 5
if errorlevel 1 pause
