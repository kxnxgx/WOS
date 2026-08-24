import pandas as pd
import numpy as np
import re

class WOSCalculator:
    EXCLUDE_STORE_PATTERN = r'バルク|BULK|bulk|New Way|NWA|new way|ZOZO|zozo|ユニフォーム|ﾕﾆﾌｫｰﾑ|uniform|丸井|マルイ|marui|テスト|test|^FJALLRAVEN by 3NITY$'

    @classmethod
    def is_excluded_store(cls, store_name: str, exclude_stores: list = None) -> bool:
        """除外対象の拠点・店舗かどうかを判定する"""
        s = str(store_name).strip()
        if not s:
            return True
        if exclude_stores and s in [str(x).strip() for x in exclude_stores]:
            return True
        if re.search(cls.EXCLUDE_STORE_PATTERN, s, re.IGNORECASE):
            return True
        return False

    @staticmethod
    def calculate(sales_df: pd.DataFrame, stock_df: pd.DataFrame,
                  item_master_df: pd.DataFrame = None,
                  order_df: pd.DataFrame = None,
                  next_order_df: pd.DataFrame = None,
                  bulk_df: pd.DataFrame = None,
                  exclude_stores: list = None) -> pd.DataFrame:
        """
        売上データと在庫データから、SKU・店舗ごとのWOSを計算する
        
        WOS = 現在在庫 ÷ 最近4週平均販売数
        最近4週平均販売数 = (最新日付から過去28日間の合計販売数) ÷ 4

        消化率（order_dfがある場合）:
        消化率 = 全期間累計売上数 ÷ 発注数（計列）× 100

        継続品（next_order_dfがある場合）:
        次シーズンの発注がある商品は is_continuation = True

        BULK在庫（bulk_dfがある場合）:
        出荷予定振分CSVから読み込んだ自社リテールBULK在庫数
        """
        print("WOSを計算しています...")
        
        if sales_df.empty or stock_df.empty:
            raise ValueError("売上データまたは在庫データが空です。")

        # 除外店舗のフィルタリング（NWA, ZOZO, ユニフォーム, 丸井, バルク, テスト店舗, 6142等）
        stock_for_wos = stock_df[
            ~stock_df['store'].apply(lambda s: WOSCalculator.is_excluded_store(s, exclude_stores))
        ].copy()
        
        sales_for_wos = sales_df[
            ~sales_df['store'].apply(lambda s: WOSCalculator.is_excluded_store(s, exclude_stores))
        ].copy()

        excluded_stock_count = len(stock_df) - len(stock_for_wos)
        if excluded_stock_count > 0:
            print(f"対象外店舗の在庫レコード {excluded_stock_count} 件を除外しました。")
            
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

        # --- フォールバックWOS（全期間平均週販ベース）---
        # 直近4週に売上がなくても、全期間（期初〜最新日付）に売上実績がある
        # 店舗×SKUは「全期間平均週販」を用いてWOSを補完する。
        # これにより「在庫あり・最近売れていない」店舗も移動推奨の対象になる。
        total_weeks = max((max_date - sales_for_wos['date'].min()).days / 7, 1)
        full_period_sales = (
            sales_for_wos
            .groupby(['store', 'sku'])['sales_qty']
            .sum()
            .reset_index()
            .rename(columns={'sales_qty': 'total_sales_full'})
        )
        full_period_sales['avg_sales_full'] = full_period_sales['total_sales_full'] / total_weeks

        wos_df = pd.merge(
            wos_df,
            full_period_sales[['store', 'sku', 'avg_sales_full']],
            on=['store', 'sku'],
            how='left'
        )
        wos_df['avg_sales_full'] = wos_df['avg_sales_full'].fillna(0)

        # フォールバックWOS: avg_sales_4w=0 かつ avg_sales_full が月1点以上（>=0.25/週）の場合のみ算出
        # ※ 月1点未満の超低速品は「実質的に動いていない」とみなし移動推奨対象外とする
        FALLBACK_MIN_AVG = 0.25  # 週0.25点 = 月1点相当
        wos_df['wos_fallback'] = np.where(
            (wos_df['avg_sales_4w'] == 0) & (wos_df['avg_sales_full'] >= FALLBACK_MIN_AVG),
            wos_df['stock_qty'] / wos_df['avg_sales_full'],
            np.nan
        )

        # 商品名とColorNameをマスタとして結合する
        sku_master_list = []
        if item_master_df is not None and not item_master_df.empty:
            sku_master_list.append(item_master_df[['sku', 'item_name', 'color_name']])
        if 'item_name' in sales_df.columns and 'color_name' in sales_df.columns:
            sku_master_list.append(sales_df[['sku', 'item_name', 'color_name']])

        if sku_master_list:
            combined_master = pd.concat(sku_master_list, ignore_index=True)
            # Unknownでないレコードを優先
            valid_mask = (combined_master['item_name'] != 'Unknown') & (combined_master['color_name'] != 'Unknown')
            combined_master = pd.concat([combined_master[valid_mask], combined_master[~valid_mask]], ignore_index=True)
            combined_master = combined_master.drop_duplicates(subset=['sku'], keep='first')

            wos_df = pd.merge(wos_df, combined_master, on='sku', how='left')
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

        # --- BULK在庫の結合 ---
        if bulk_df is not None and not bulk_df.empty:
            wos_df = pd.merge(wos_df, bulk_df, on='sku', how='left')
            wos_df['bulk_stock'] = wos_df['bulk_stock'].fillna(0).astype(int)
        else:
            wos_df['bulk_stock'] = 0

        return wos_df

