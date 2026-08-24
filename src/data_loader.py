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

    def load_bulk_stock(self, ship_alloc_csv_path: str) -> pd.DataFrame:
        """出荷予定振分CSVから自社リテールBULK在庫数を読み込む（H列: 列インデックス7）"""
        print(f"出荷予定振分データを読み込んでいます: {ship_alloc_csv_path}")
        try:
            # 4行スキップしてデータ行を読み込み
            df = pd.read_csv(ship_alloc_csv_path, header=None, skiprows=4, encoding='cp932', low_memory=False)

            # [3]=商品コード(sku), [5]=Brand, [7]=BULK(bulk_stock)
            sku_col = 3
            brand_col = 5
            bulk_col = 7

            # ブランドが FRV / FJALLRAVEN のものを抽出
            if brand_col < len(df.columns):
                df = df[df[brand_col].astype(str).str.upper().str.contains('FRV|FJALLRAVEN', na=False)].copy()

            result = df[[sku_col, bulk_col]].copy()
            result.columns = ['sku', 'bulk_stock']

            # SKUの整形
            result = result[result['sku'].notna()].copy()
            result['sku'] = result['sku'].astype(str).str.strip()
            result = result[result['sku'] != 'nan'].copy()

            # bulk_stock の数値変換
            result['bulk_stock'] = pd.to_numeric(result['bulk_stock'], errors='coerce').fillna(0).astype(int)

            # SKU ごとに合算
            result = result.groupby('sku', as_index=False)['bulk_stock'].sum()

            print(f"  -> 自社BULK在庫データ: {len(result)} SKU 読み込み完了（BULK在庫あり: {len(result[result['bulk_stock'] > 0])} SKU）")
            return result
        except Exception as e:
            raise RuntimeError(f"出荷予定振分データの読み込みに失敗しました: {e}")

    def load_item_master(self, master_csv_path: str = "商品マスタ.csv") -> pd.DataFrame:
        """商品マスタCSVから商品コード、商品名（StyleName）、カラー名（ColorName）を読み込む"""
        print(f"商品マスタデータを読み込んでいます: {master_csv_path}")
        try:
            df = pd.read_csv(master_csv_path, encoding='cp932', low_memory=False)

            sku_col = '商品コード' if '商品コード' in df.columns else df.columns[0]
            name_col = 'StyleName' if 'StyleName' in df.columns else df.columns[4]
            color_col = 'ColorName' if 'ColorName' in df.columns else df.columns[6]
            brand_col = 'Brand' if 'Brand' in df.columns else (df.columns[22] if len(df.columns) > 22 else None)

            if brand_col and brand_col in df.columns:
                # FRV / FJALLRAVEN のものを優先抽出
                frv_df = df[df[brand_col].astype(str).str.upper().str.contains('FRV|FJALLRAVEN', na=False)].copy()
                if not frv_df.empty:
                    df = frv_df

            result = df[[sku_col, name_col, color_col]].copy()
            result.columns = ['sku', 'item_name', 'color_name']

            result['sku'] = result['sku'].astype(str).str.strip()
            result['item_name'] = result['item_name'].fillna('Unknown').astype(str).str.strip()
            result['color_name'] = result['color_name'].fillna('Unknown').astype(str).str.strip()

            result = result[result['sku'] != 'nan']
            result = result.drop_duplicates(subset=['sku'])

            print(f"  -> 商品マスタデータ: {len(result)} SKU 読み込み完了")
            return result
        except Exception as e:
            print(f"  [警告] 商品マスタの読み込みに失敗しました ({e})。処理をスキップします。")
            return pd.DataFrame(columns=['sku', 'item_name', 'color_name'])
