import pandas as pd

class DataLoader:
    def __init__(self, sales_csv_path: str, stock_csv_path: str):
        self.sales_csv_path = sales_csv_path
        self.stock_csv_path = stock_csv_path

    def load_sales_history(self) -> pd.DataFrame:
        """売上データを読み込む"""
        print(f"売上データを読み込んでいます: {self.sales_csv_path}")
        try:
            df = pd.read_csv(self.sales_csv_path, encoding='cp932', low_memory=False)
            
            # 列インデックスで確実に取得する
            date_col = df.columns[1]   # 営業日付
            store_col = df.columns[2]  # 店舗名
            item_col = df.columns[3]   # 商品コード
            brand_col = df.columns[4]  # ブランド名1
            name_col = df.columns[6]   # 商品名
            color_col = df.columns[8]  # ColorName
            sales_col = df.columns[10] # 数量
            
            df[date_col] = pd.to_datetime(df[date_col])
            
            # ブランドが 'FJALLRAVEN' のものだけ抽出
            df = df[df[brand_col].astype(str).str.upper() == 'FJALLRAVEN'].copy()
            
            result = df[[date_col, store_col, item_col, name_col, color_col, sales_col]].copy()
            result.columns = ['date', 'store', 'sku', 'item_name', 'color_name', 'sales_qty']
            
            # 数値変換とNA埋め
            result['sales_qty'] = pd.to_numeric(result['sales_qty'], errors='coerce').fillna(0).astype(int)
            # 商品名・ColorNameの欠損値埋め
            result['item_name'] = result['item_name'].fillna('Unknown')
            result['color_name'] = result['color_name'].fillna('Unknown')
            
            # 商品コードが '23510' で始まるものを除外 (KANKEN除外ルール)
            result = result[~result['sku'].astype(str).str.startswith('23510')].copy()
            
            # 商品名に 'Tokyo' (大文字小文字問わず) が含まれるものを除外
            result = result[~result['item_name'].astype(str).str.contains('Tokyo', case=False, na=False)].copy()
            
            return result
        except Exception as e:
            raise RuntimeError(f"売上データの読み込みに失敗しました: {e}")

    def load_current_stock(self) -> pd.DataFrame:
        """在庫データを読み込む"""
        print(f"在庫データを読み込んでいます: {self.stock_csv_path}")
        try:
            df = pd.read_csv(self.stock_csv_path, encoding='cp932', low_memory=False)
            
            # 列インデックスで確実に取得する
            store_col = df.columns[1]  # 店舗名
            item_col = df.columns[2]   # 商品コード
            name_col = df.columns[3]   # 商品名
            brand_col = df.columns[7]  # BrandName
            stock_col = df.columns[17] # 現在庫 (R列: インデックス17)
            
            # ブランドが 'FJALLRAVEN' のものだけ抽出
            df = df[df[brand_col].astype(str).str.upper() == 'FJALLRAVEN'].copy()
                    
            result = df[[store_col, item_col, name_col, stock_col]].copy()
            result.columns = ['store', 'sku', 'item_name', 'stock_qty']
            
            result['stock_qty'] = pd.to_numeric(result['stock_qty'], errors='coerce').fillna(0).astype(int)
            
            # 商品コードが '23510' で始まるものを除外 (KANKEN除外ルール)
            result = result[~result['sku'].astype(str).str.startswith('23510')].copy()
            
            # 商品名に 'Tokyo' (大文字小文字問わず) が含まれるものを除外
            result = result[~result['item_name'].astype(str).str.contains('Tokyo', case=False, na=False)].copy()
            
            # 不要になった item_name 列を削除して既存のインターフェースに合わせる
            result = result[['store', 'sku', 'stock_qty']]
            
            return result
        except Exception as e:
            raise RuntimeError(f"在庫データの読み込みに失敗しました: {e}")

    def load_order_data(self, order_xlsx_path: str) -> pd.DataFrame:
        """発注数データを読み込む（消化率計算用）"""
        print(f"発注数データを読み込んでいます: {order_xlsx_path}")
        try:
            df = pd.read_excel(order_xlsx_path, header=0)

            # 列インデックスで取得（FRV発注数ファイルの列構成に合わせる）
            # [0]=品コード(sku), [11]=計(total_order)
            sku_col = df.columns[0]
            total_col = df.columns[11]

            result = df[[sku_col, total_col]].copy()
            result.columns = ['sku', 'total_order']

            # SKUが空またはNaNの行を除外（ヘッダー行などを除去）
            result = result[result['sku'].notna()].copy()
            result['sku'] = result['sku'].astype(str).str.strip()
            result = result[result['sku'] != 'nan'].copy()

            # total_order を数値変換
            result['total_order'] = pd.to_numeric(result['total_order'], errors='coerce').fillna(0).astype(int)

            # 発注数が0の行は除外（消化率計算の分母がゼロになるため）
            result = result[result['total_order'] > 0].copy()

            # SKU の重複がある場合は合計する（念のため）
            result = result.groupby('sku', as_index=False)['total_order'].sum()

            print(f"  -> 発注数データ: {len(result)} SKU 読み込み完了")
            return result
        except Exception as e:
            raise RuntimeError(f"発注数データの読み込みに失敗しました: {e}")
