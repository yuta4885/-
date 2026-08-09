@echo off
REM 社員スキルマップ ツールの実行用バッチファイル
REM これをダブルクリックするだけで、必要なライブラリの確認とスクリプトの実行を行います

cd /d %~dp0

echo 必要なライブラリを確認しています…
python -m pip install --quiet -r requirements.txt

echo.
echo スクリプトを実行しています…
python main.py

echo.
echo 処理が完了しました。ウィンドウを閉じるには何かキーを押してください。
pause
