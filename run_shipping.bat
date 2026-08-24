@echo off
title 出荷表作成ツール (Shipping List Generator)
chcp 65001 >nul
cd /d "%~dp0"

echo =======================================================
echo  店舗間移動 出荷表・CSV生成ツール
echo =======================================================
echo.
echo 暫定移動明細.xlsx と 商品マスタ.csv から
echo 出荷指示Excel (WOS_Report_店舗出荷.xlsx) と
echo 取込用CSV (transfer_upload.csv) を生成します...
echo.

python src/generate_shipping_list.py

echo.
if %ERRORLEVEL% NEQ 0 (
    echo -------------------------------------------------------
    echo [エラー] 処理中にエラーが発生しました。
    echo 上記のエラーメッセージを確認してください。
    echo -------------------------------------------------------
) else (
    echo -------------------------------------------------------
    echo [完了] 出荷表とCSVの生成が正常に完了しました！
    echo.
    echo 出力ファイル:
    echo  - WOS_Report_店舗出荷.xlsx (店舗別出荷指示Excel)
    echo  - transfer_upload.csv (システム取込用CSV)
    echo -------------------------------------------------------
)
echo.
pause
