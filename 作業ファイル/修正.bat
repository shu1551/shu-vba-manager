@echo off
if "%~1" == "" (
  echo 使い方: 修正 [Excel名] [モジュール名] [ソースファイル名]
  echo 例: 修正 秀 shu001 shu001_FINAL.vba
  exit /b
)

set XLSM=%~1
if "%XLSM:~-5%" neq ".xlsm" set XLSM=%XLSM%.xlsm

python "c:\Users\shu\Desktop\秀 20260113\作業ファイル\project\python_scripts\vba_patch.py" "c:\Users\shu\Desktop\秀 20260113\%XLSM%" "%~2" "c:\Users\shu\Desktop\秀 20260113\作業ファイル\%~3"
