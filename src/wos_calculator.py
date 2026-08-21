import pandas as pd
import numpy as np

class WOSCalculator:
    @staticmethod
    def calculate(sales_df: pd.DataFrame, stock_df: pd.DataFrame,
                  order_df: pd.DataFrame = None,
                  next_order_df: pd.DataFrame = None,
                  exclude_stores: list = None) -> pd.DataFrame:
        """
        売上データと在庫データから、SKU・店舗ごとのWOSを計算する
        
        WOS = 現在在庫 ÷ 最近4週平均販売数
        最近4週平均販売数 = (最新日付から過去28日間の合計販売数) ÷ 4

        消化率（order_dfがある場合）:
        消化率 = 全期間累計売上数 ÷ 発注数（計列）× 100

        継続品（next_order_dfがある場合）:
        次シーズンの発注がある商品は is_continuation = True
        """
        print("WOSを計算しています...")
        
        if sales_df.empty or stock_df.empty:
            raise ValueError("売上データまたは在庫データが空です。")

        # 除外店舗のフィルタリング（WOS計算・移動推奨・ヒートマップ対象外）
        if exclude_stores:
            exclude_strs = [str(s).strip() for s in exclude_stores]
            print(f"WOS計算・移動推奨から除外する店舗: {exclude_strs}")
            stock_for_wos = stock_df[~stock_df['store'].astype(str).str.strip().isin(exclude_strs)].copy()
            sales_for_wos = sales_df[~sales_df['store'].astype(str).str.strip().isin(exclude_strs)].copy()
        else:
            stock_for_wos = stock_df.copy()
            sales_for_wos = sales_df.copy()
            
        # 最近4週間（28日間）のデータを抽出（除外店舗を除く）
        max_date = sales_for_wos['date'].max()
        cutoff_date = max_date - pd.Timedelta(days=28)
        
        recent_sales = sales_for_wos[sales_for_wos['date'] > cutoff_date].copy()
        
        # 店舗×SKUごとの合計販売数を計算
        sales_agg = recent_sales.groupby(['store', 'sku'])['sales_qty'].sum().reset_index()
        
        # 4週平均を算出
        sales_agg['avg_sales_4w'] = sales_agg['sales_qty'] / 4.0
        
        # 在庫データと結合 (Left Join: 在庫があるものをベースにする)
        # 在庫がなくても売上があるケースは、現在庫0として扱う
        
        # 店舗・SKUのユニークな組み合わせを作成
        all_combinations = pd.merge(
            stock_for_wos[['store', 'sku']], 
            sales_agg[['store', 'sku']], 
            on=['store', 'sku'], 
            how='outer'
        ).drop_duplicates()
        
        # マージ
        wos_df = pd.merge(all_combinations, stock_for_wos, on=['store', 'sku'], how='left')
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

        # --- 消化率の計算（order_df がある場合のみ）---
        if order_df is not None and not order_df.empty:
            print("消化率を計算しています...")
            # 全期間（期初～最新日付）の SKU ごと累計売上を集計
            cumulative_sales = sales_df.groupby('sku')['sales_qty'].sum().reset_index()
            cumulative_sales.columns = ['sku', 'cumulative_sales']

            # 発注数と結合
            st_df = pd.merge(cumulative_sales, order_df, on='sku', how='inner')

            # 消化率 = 累計売上 ÷ 発注数 × 100（小数点1位、100%上限）
            st_df['sell_through'] = (st_df['cumulative_sales'] / st_df['total_order'] * 100).round(1)
            st_df['sell_through'] = st_df['sell_through'].clip(upper=100.0)

            # wos_df に消化率・累計売上・発注数を結合
            wos_df = pd.merge(
                wos_df,
                st_df[['sku', 'sell_through', 'cumulative_sales', 'total_order']],
                on='sku',
                how='left'
            )

        # --- 継続品判定（next_order_df がある場合のみ）---
        if next_order_df is not None and not next_order_df.empty:
            print("次シーズン継続品を判定しています...")
            next_skus = next_order_df['sku'].unique()
            wos_df['is_continuation'] = wos_df['sku'].isin(next_skus)
        else:
            wos_df['is_continuation'] = False

        return wos_df

