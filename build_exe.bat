@echo off
chcp 65001 > nul
cd /d "%~dp0"

echo ========================================
echo  exe ビルドスクリプト
echo ========================================

REM 仮想環境があれば有効化
if exist ".venv\Scripts\activate.bat" (
    call .venv\Scripts\activate.bat
    echo 仮想環境を有効化しました
)

REM PyInstaller がなければインストール
python -c "import PyInstaller" 2>nul
if errorlevel 1 (
    echo PyInstaller をインストールしています...
    pip install pyinstaller
)

echo.
echo ビルドを開始します...
echo.

REM 1つのexeにまとめる（コンソール非表示 = GUIアプリ）
pyinstaller --onefile --windowed ^
    --name "大沼運輸倉庫_SGシステム" ^
    --hidden-import openpyxl ^
    --hidden-import openpyxl.styles ^
    --clean ^
    "ダブルクリックボタン.py"

if errorlevel 1 (
    echo.
    echo ビルドに失敗しました。
    pause
    exit /b 1
)

echo.
echo ========================================
echo  ビルド完了
echo ========================================
echo.
echo 出力先: dist\大沼運輸倉庫_SGシステム.exe
echo.
echo ※ exe と同じフォルダに以下が自動作成されます:
echo   - ここにファイルを入れてください
echo   - 完成フォルダ
echo   - ゴミ箱
echo   - tenko_settings.json （設定で保存時）
echo.
pause
