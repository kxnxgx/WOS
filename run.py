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


def find_latest_xlsx(keyword1: str, keyword2: str, search_dir: str = ".") -> str:
    """
    指定キーワードを2つとも含むExcelファイルを検索し、更新日時が最新のものを返す。
    来年以降も使えるようにシーズン番号を問わず検索する。

    Args:
        keyword1: ファイル名に含まれるべきキーワード1（例: 'FRV'）
        keyword2: ファイル名に含まれるべきキーワード2（例: '発注数'）
        search_dir: 検索対象ディレクトリ（デフォルトは実行カレント）

    Returns:
        最新ファイルのパス文字列。見つからない場合はNone。
    """
    pattern = os.path.join(search_dir, "*.xlsx")
    candidates = [
        f for f in glob.glob(pattern)
        if keyword1 in os.path.basename(f) and keyword2 in os.path.basename(f)
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
    parser.add_argument("--order", type=str, default=None, help="発注数ExcelファイルのパスCSVパス（省略時は自動検索）")
    parser.add_argument("--next-order", type=str, default=None,
                        help="次シーズン発注数Excelのパス（省略時はFW発注数ファイルを自動検索）")
    parser.add_argument("--threshold", type=float, default=80.0,
                        help="消化率の優先集約閾値 %% (デフォルト: 80.0)")
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

    # --- 発注数データのパス解決（任意）---
    order_path = None
    if args.order:
        order_path = args.order
        print(f"[指定] 発注数データ: {os.path.basename(order_path)}")
    else:
        found = find_latest_xlsx("FRV", "発注数")
        if found:
            order_path = found
            print(f"[自動検索] 発注数データを検出しました: {os.path.basename(order_path)}")
        else:
            print("[警告] 発注数データ（'FRV' かつ '発注数' を含むxlsx）が見つかりません。")
    # --- 次シーズン発注数データのパス解決（任意）---
    next_order_path = None
    if args.next_order:
        next_order_path = args.next_order
        print(f"[指定] 次シーズン発注数データ: {os.path.basename(next_order_path)}")
    else:
        # FWをハードコードせず、今季と違う方のファイルを探すなどのロジックも可能ですが
        # 今回は暫定的に「FW」かつ「発注数」を検索します
        found_next = find_latest_xlsx("FW", "発注数")
        if found_next:
            next_order_path = found_next
            print(f"[自動検索] 次シーズン発注数データを検出しました: {os.path.basename(next_order_path)}")

    print(f"[設定] 優先集約の消化率閾値: {args.threshold:.0f}%")

    try:
        # 1. データロード
        loader = DataLoader(sales_path, stock_path)
        sales_df = loader.load_sales_history()
        stock_df = loader.load_current_stock()

        # 発注数データ（任意）
        order_df = None
        if order_path and os.path.exists(order_path):
            order_df = loader.load_order_data(order_path)

        # 次シーズン発注数データ（任意）
        next_order_df = None
        if next_order_path and os.path.exists(next_order_path):
            next_order_df = loader.load_order_data(next_order_path)

        # 2. WOS計算（+ 消化率・継続品判定）
        wos_df = WOSCalculator.calculate(sales_df, stock_df, order_df, next_order_df)

        # 3. アイテム移動推奨算出
        move_df = ItemAllocator.allocate(wos_df, sell_through_threshold=args.threshold)

        # 4. レポート生成
        Reporter.generate_excel(wos_df, move_df, "WOS_Report.xlsx")
        Reporter.generate_html(wos_df, move_df, "WOS_Report.html", threshold=args.threshold)

        print("\n処理が完了しました！")
        print("- WOS_Report.xlsx (Excelレポート)")
        print("- WOS_Report.html (HTMLレポート)")

    except Exception as e:
        import traceback
        print(f"\n[実行エラー] 処理中にエラーが発生しました:\n{e}")
        traceback.print_exc()
        sys.exit(1)


if __name__ == "__main__":
    main()
