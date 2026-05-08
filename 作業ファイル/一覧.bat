@echo off
if "%~1" == "" (
  echo 使い方: 一覧 マクロ名等
  exit /b
)
python "c:\Users\shu\Desktop\秀 20260113\作業ファイル\project\python_scripts\list_macros.py" "c:\Users\shu\Desktop\秀 20260113\%~1"