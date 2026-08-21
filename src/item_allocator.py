import pandas as pd
import math

class ItemAllocator:
    @staticmethod
    def allocate(wos_df: pd.DataFrame,
                 sell_through_threshold: float = 80.0) -> pd.DataFrame:
        """
        WOSデータから店舗間のアイテム移動推奨リストを作成する

        Args:
            wos_df: WOS計算済みDataFrame（sell_through列があれば消化率で優先度を判定）
            sell_through_threshold: 優先集約の消化率閾値（デフォルト80%）
        """
        print("アイテム移動推奨リストを作成しています...")
        
        move_records = []
        all_skus = wos_df['sku'].unique()
        
        for sku in all_skus:
            stores = wos_df[wos_df['sku'] == sku]
            valid = stores[stores['wos'].notna()]
            
            if valid.empty:
                continue
                
            wos_avg = valid['wos'].mean()
            
            # wos_avg が 0 の場合は移動不要
            if wos_avg == 0:
                continue
                
            shippers = valid[valid['wos'] > wos_avg].copy()  # WOSが高い = 在庫余剰 = 出荷側
            receivers = valid[valid['wos'] < wos_avg].copy() # WOSが低い = 在庫不足 = 受入れ側

            shippers['surplus'] = shippers['stock_qty'] - (wos_avg * shippers['avg_sales_4w'])
            receivers['deficit'] = (wos_avg * receivers['avg_sales_4w']) - receivers['stock_qty']
            
            # 余剰分と不足分をマッチング
            shippers = shippers.sort_values('surplus', ascending=False).to_dict('records')
            receivers = receivers.sort_values('deficit', ascending=False).to_dict('records')

            # このSKUの消化率（wos_df の sell_through 列があれば参照）
            has_sell_through = 'sell_through' in wos_df.columns
            if has_sell_through:
                sell_through_val = wos_df.loc[wos_df['sku'] == sku, 'sell_through'].iloc[0]
                if pd.notna(sell_through_val) and sell_through_val >= sell_through_threshold:
                    priority = "優先"
                else:
                    priority = "通常"
            else:
                priority = "通常"
            # このSKUの継続品フラグ
            is_cont = False
            if 'is_continuation' in wos_df.columns:
                is_cont = wos_df.loc[wos_df['sku'] == sku, 'is_continuation'].iloc[0]

            # このSKUのBULK在庫
            bulk_stock_val = 0
            if 'bulk_stock' in wos_df.columns:
                b_val = wos_df.loc[wos_df['sku'] == sku, 'bulk_stock']
                if not b_val.empty and pd.notna(b_val.iloc[0]):
                    bulk_stock_val = int(b_val.iloc[0])

            for s in shippers:
                s['current_stock'] = s['stock_qty']
            for r in receivers:
                r['current_stock'] = r['stock_qty']

            for s in shippers:
                if s['surplus'] <= 0: continue
                
                for r in receivers:
                    if r['deficit'] <= 0: continue
                    if s['surplus'] <= 0: break
                    
                    move = min(s['surplus'], r['deficit'])
                    move_qty = math.ceil(move)
                    
                    if move_qty > 0:
                        shipper_pre_wos = round(s['current_stock'] / s['avg_sales_4w'], 1) if s['avg_sales_4w'] > 0 else None
                        receiver_pre_wos = round(r['current_stock'] / r['avg_sales_4w'], 1) if r['avg_sales_4w'] > 0 else None

                        shipper_post_stock = s['current_stock'] - move_qty
                        receiver_post_stock = r['current_stock'] + move_qty
                        
                        shipper_post_wos = round(shipper_post_stock / s['avg_sales_4w'], 1) if s['avg_sales_4w'] > 0 else None
                        receiver_post_wos = round(receiver_post_stock / r['avg_sales_4w'], 1) if r['avg_sales_4w'] > 0 else None

                        reason = (
                            f"{s['store']}のWOSが{s['wos']:.1f}週、"
                            f"{r['store']}のWOSが{r['wos']:.1f}週"
                            f"（全店平均{wos_avg:.1f}週）のため、"
                            f"{s['store']}から{r['store']}へ{move_qty}点移動を推奨"
                        )
                        
                        item_name = s.get('item_name', 'Unknown')
                        color_name = s.get('color_name', 'Unknown')
                        
                        move_records.append({
                            'priority': priority,
                            'is_continuation': is_cont,
                            'sku': sku,
                            'item_name': item_name,
                            'color_name': color_name,
                            'bulk_stock': bulk_stock_val,
                            'sell_through': sell_through_val if has_sell_through else None,
                            'shipper': s['store'],
                            'shipper_stock': s['stock_qty'],
                            'shipper_pre_wos': shipper_pre_wos,
                            'shipper_post_wos': shipper_post_wos,
                            'receiver': r['store'],
                            'receiver_stock': r['stock_qty'],
                            'move_qty': move_qty,
                            'receiver_pre_wos': receiver_pre_wos,
                            'receiver_post_wos': receiver_post_wos,
                            'reason': reason
                        })
                        
                        s['current_stock'] -= move_qty
                        r['current_stock'] += move_qty
                        s['surplus'] -= move
                        r['deficit'] -= move
                        
        result = pd.DataFrame(move_records)
        return result
