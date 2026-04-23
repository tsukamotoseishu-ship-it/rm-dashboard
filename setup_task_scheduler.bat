@echo off
REM 毎朝6時に競合価格を自動取得するWindowsタスクを登録します
schtasks /create /tn "RM_競合価格取得" ^
  /tr "C:\Users\tsukamoto.seishu\rm_system\scrape_daily.bat" ^
  /sc daily /st 06:00 ^
  /ru "%USERNAME%" ^
  /f
echo.
echo タスクを登録しました。毎朝6:00に自動実行されます。
echo 確認: タスクスケジューラ → RM_競合価格取得
pause
