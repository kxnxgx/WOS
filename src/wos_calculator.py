import pandas as pd
import numpy as np

class WOSCalculator:
    @staticmethod
    def calculate(sales_df: pd.DataFrame, stock_df: pd.DataFrame) -> pd.DataFrame:
        """
        売上データと在庫データから、SKU・店舗ごとのWOSを計算する
        
        WOS = 現在在庫 ÷ 最近4週平均販売数
        最近4週平均販売数 = (最新日付から過去28日間の合計販売数) ÷ 4
        """
        print("WOSを計算しています...")
        
        if sales_df.empty or stock_df.empty:
            raise ValueError("売上データまたは在庫データが空です。")
            
        # 最近4週間（28日間）のデータを抽出
        max_date = sales_df['date'].max()
        cutoff_date = max_date - pd.Timedelta(days=28)
        
        recent_sales = sales_df[sales_df['date'] > cutoff_date].copy()
        
        # 店舗×SKUごとの合計販売数を計算
        sales_agg = recent_sales.groupby(['store', 'sku'])['sales_qty'].sum().reset_index()
        
        # 4週平均を算出
        sales_agg['avg_sales_4w'] = sales_agg['sales_qty'] / 4.0
        
        # 在庫データと結合 (Left Join: 在庫があるものをベースにする)
        # 在庫がなくても売上があるケースは、現在庫0として扱う
        
        # 店舗・SKUのユニークな組み合わせを作成
        all_combinations = pd.merge(
            stock_df[['store', 'sku']], 
            sales_agg[['store', 'sku']], 
            on=['store', 'sku'], 
            how='outer'
        ).drop_duplicates()
        
        # マージ
        wos_df = pd.merge(all_combinations, stock_df, on=['store', 'sku'], how='left')
        wos_df = pd.merge(wos_df, sales_agg[['store', 'sku', 'avg_sales_4w']], on=['store', 'sku'], how='left')
        
        # 欠損値補完
        wos_df['stock_qty'] = wos_df['stock_qty'].fillna(0)
        wos_df['avg_sales_4w'] = wos_df['avg_sales_4w'].fillna(0)
        
        # WOS計算
        # avg_sales_4w が 0 の場合は NaN にする
        wos_df['wos'] = np.where(
            wos_df['avg_sales_4w'] > 0, 
            wos_df['stock_qty'] / wos_df['avg_sales_4w'], 
            np.nan
        )
        
        # 商品名とColorNameをマスタとして結合する
        if 'item_name' in sales_df.columns and 'color_name' in sales_df.columns:
            sku_master = sales_df[['sku', 'item_name', 'color_name']].drop_duplicates(subset=['sku'])
            wos_df = pd.merge(wos_df, sku_master, on='sku', how='left')
            wos_df['item_name'] = wos_df['item_name'].fillna('Unknown')
            wos_df['color_name'] = wos_df['color_name'].fillna('Unknown')
        
        return wos_df
