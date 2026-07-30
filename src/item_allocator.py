import pandas as pd
import math

class ItemAllocator:
    @staticmethod
    def allocate(wos_df: pd.DataFrame) -> pd.DataFrame:
        """
        WOSデータから店舗間のアイテム移動推奨リストを作成する
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
            
            # TODO: implementation_plan では `WOS < 平均 → 出荷候補` と書かれていたが、
            # 物理的な意味（WOSが低い＝在庫不足）を考えると、WOSが高い方から低い方へ移動すべき。
            # ここでは WOS > 平均 を出荷側 (余剰)、WOS < 平均 を受入れ側 (不足) とする。
            
            shippers['surplus'] = shippers['stock_qty'] - (wos_avg * shippers['avg_sales_4w'])
            receivers['deficit'] = (wos_avg * receivers['avg_sales_4w']) - receivers['stock_qty']
            
            # 余剰分と不足分をマッチング
            shippers = shippers.sort_values('surplus', ascending=False).to_dict('records')
            receivers = receivers.sort_values('deficit', ascending=False).to_dict('records')
            
            for s in shippers:
                if s['surplus'] <= 0: continue
                
                for r in receivers:
                    if r['deficit'] <= 0: continue
                    if s['surplus'] <= 0: break
                    
                    move = min(s['surplus'], r['deficit'])
                    move_qty = math.ceil(move)
                    
                    if move_qty > 0:
                        reason = f"WOS {s['wos']:.1f}週 (平均{wos_avg:.1f}週)のため、{s['store']}から{r['store']}へ{move_qty}点移動を推奨"
                        
                        item_name = s.get('item_name', 'Unknown')
                        color_name = s.get('color_name', 'Unknown')
                        
                        move_records.append({
                            'sku': sku,
                            'item_name': item_name,
                            'color_name': color_name,
                            'shipper': s['store'],
                            'shipper_stock': s['stock_qty'],
                            'receiver': r['store'],
                            'receiver_stock': r['stock_qty'],
                            'move_qty': move_qty,
                            'reason': reason
                        })
                        
                        s['surplus'] -= move
                        r['deficit'] -= move
                        
        result = pd.DataFrame(move_records)
        return result
