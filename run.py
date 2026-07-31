import sys
import os
import glob
import argparse
import io

# Windows コマンドプロンプトで日本語が文字化けするのを防ぐ
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8', errors='replace')

from src.data_loader import DataLoader
from src.wos_calculator import WOSCalculator
from src.item_allocator import ItemAllocator
from src.reporter import Reporter


def find_latest_csv(keyword: str, search_dir: str = ".") -> str:
    """
    指定キーワードを含むCSVファイルを検索し、更新日時が最新のものを返す。
    ダウンロードの重複で「ファイル名 (1).csv」のような連番が付いた場合にも対応。

    Args:
        keyword: ファイル名に含まれるべきキーワード（例: '在庫一覧'）
        search_dir: 検索対象ディレクトリ（デフォルトは実行カレント）

    Returns:
        最新ファイルのパス文字列。見つからない場合はNone。
    """
    pattern = os.path.join(search_dir, "*.csv")
    candidates = [
        f for f in glob.glob(pattern)
        if keyword in os.path.basename(f)
    ]
    if not candidates:
        return None
    # 更新日時が最も新しいファイルを選択
    latest = max(candidates, key=os.path.getmtime)
    return latest


def main():
    print("=======================================")
    print(" WOS アイテム集約ツール")
    print("=======================================")

    # コマンドライン引数の設定（省略時はキーワード自動検索）
    parser = argparse.ArgumentParser(description="売上・在庫データからアイテム移動リストを生成します")
    parser.add_argument("--sales", type=str, default=None, help="売上データのCSVパス（省略時は自動検索）")
    parser.add_argument("--stock", type=str, default=None, help="在庫データのCSVパス（省略時は自動検索）")
    args = parser.parse_args()

    # --- 売上データのパス解決 ---
    if args.sales:
        sales_path = args.sales
    else:
        sales_path = find_latest_csv("営業日付別売上分析")
        if sales_path:
            print(f"[自動検索] 売上データを検出しました: {os.path.basename(sales_path)}")
        else:
            print("[エラー] 売上データ（'営業日付別売上分析' を含むCSV）が見つかりません。")
            print("         --sales オプションで明示的にパスを指定してください。")
            sys.exit(1)

    # --- 在庫データのパス解決 ---
    if args.stock:
        stock_path = args.stock
    else:
        stock_path = find_latest_csv("在庫一覧")
        if stock_path:
            print(f"[自動検索] 在庫データを検出しました: {os.path.basename(stock_path)}")
        else:
            print("[エラー] 在庫データ（'在庫一覧' を含むCSV）が見つかりません。")
            print("         --stock オプションで明示的にパスを指定してください。")
            sys.exit(1)

    if not os.path.exists(sales_path):
        print(f"[エラー] 売上データが見つかりません: {sales_path}")
        sys.exit(1)

    if not os.path.exists(stock_path):
        print(f"[エラー] 在庫データが見つかりません: {stock_path}")
        sys.exit(1)

    try:
        # 1. データロード
        loader = DataLoader(sales_path, stock_path)
        sales_df = loader.load_sales_history()
        stock_df = loader.load_current_stock()

        # 2. WOS計算
        wos_df = WOSCalculator.calculate(sales_df, stock_df)

        # 3. アイテム移動推奨算出
        move_df = ItemAllocator.allocate(wos_df)

        # 4. レポート生成
        Reporter.generate_excel(wos_df, move_df, "WOS_Report.xlsx")
        Reporter.generate_html(wos_df, move_df, "WOS_Report.html")

        print("\n処理が完了しました！")
        print("- WOS_Report.xlsx (Excelレポート)")
        print("- WOS_Report.html (HTMLレポート)")

    except Exception as e:
        print(f"\n[実行エラー] 処理中にエラーが発生しました:\n{e}")
        sys.exit(1)


if __name__ == "__main__":
    main()
