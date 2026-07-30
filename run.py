import sys
import os
import argparse
from src.data_loader import DataLoader
from src.wos_calculator import WOSCalculator
from src.item_allocator import ItemAllocator
from src.reporter import Reporter

def main():
    print("=======================================")
    print(" WOS アイテム集約ツール")
    print("=======================================")
    
    # コマンドライン引数の設定
    parser = argparse.ArgumentParser(description="売上・在庫データからアイテム移動リストを生成します")
    parser.add_argument("--sales", type=str, default="営業日付別売上分析(旧).csv", help="売上データのCSVパス")
    parser.add_argument("--stock", type=str, default="在庫一覧.csv", help="在庫データのCSVパス")
    args = parser.parse_args()
    
    sales_path = args.sales
    stock_path = args.stock
    
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
