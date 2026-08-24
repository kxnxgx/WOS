import pandas as pd
import numpy as np
import math

class ItemAllocator:
    @staticmethod
    def _get_store_priority_rank(store_name: str) -> int:
        """店舗名から受入優先順位（1〜9、小さいほど優先）を判定する"""
        s = str(store_name)
        # TOKYO NODE を先に判定（TOKYO と誤判定しないため）
        if 'NODE' in s:
            return 4
        elif 'TOKYO' in s:
            return 1
        elif 'ルクア' in s:
            return 2
        elif '名古屋' in s:
            return 3
        elif '京王新宿' in s:
            return 5
        elif '大丸心斎橋' in s:
            return 6
        elif '玉川高島屋' in s:
            return 7
        elif 'HUTTE' in s or 'ヒュッテ' in s:
            return 8
        elif 'NARITA' in s:
            return 9
        else:
            return 99

    @staticmethod
    def _fmt_wos(stock, avg_sales, is_fallback=False):
        """WOS値を計算して返す（フォールバック時はavg_salesに全期間平均を使う）"""
        if avg_sales and avg_sales > 0:
            return round(stock / avg_sales, 1)
        return None

    @staticmethod
    def allocate(wos_df: pd.DataFrame,
                 sell_through_threshold: float = 80.0) -> pd.DataFrame:
        """
        WOSデータから店舗間のアイテム移動推奨リストを作成する

        Args:
            wos_df: WOS計算済みDataFrame（sell_through列があれば消化率で優先度を判定）
            sell_through_threshold: 優先集約の消化率閾値（デフォルト80%）

        処理フロー:
            1. Phase 1 (滞留在庫集約):
               売上実績のない店舗（直近4週=0 かつ 全期間<0.25）に在庫がある場合、
               売上実績のある店舗（直近売上あり、または全期間売上あり）へ引き上げ移動を推奨。
            2. Phase 2 (通常WOS平準化):
               売上実績店舗同士で依然としてWOSの偏りがある場合、全店平均WOSを基準に平準化。
        """
        print("アイテム移動推奨リストを作成しています...")

        if wos_df.empty:
            return pd.DataFrame()

        has_fallback_col = 'wos_fallback' in wos_df.columns
        has_avg_full_col = 'avg_sales_full' in wos_df.columns
        has_sell_through = 'sell_through' in wos_df.columns
        has_continuation = 'is_continuation' in wos_df.columns
        has_bulk = 'bulk_stock' in wos_df.columns

        move_records = []
        all_skus = wos_df['sku'].unique()

        for sku in all_skus:
            stores = wos_df[wos_df['sku'] == sku].copy()
            if stores.empty:
                continue

            # このSKUのメタ情報
            if has_sell_through:
                st_series = stores['sell_through'].dropna()
                sell_through_val = st_series.iloc[0] if not st_series.empty else None
                if sell_through_val is not None and pd.notna(sell_through_val) and sell_through_val >= sell_through_threshold:
                    priority = "優先"
                else:
                    priority = "通常"
            else:
                sell_through_val = None
                priority = "通常"

            is_cont = False
            if has_continuation:
                c_series = stores['is_continuation'].dropna()
                if not c_series.empty:
                    is_cont = bool(c_series.iloc[0])

            bulk_stock_val = 0
            if has_bulk:
                b_series = stores['bulk_stock'].dropna()
                if not b_series.empty and pd.notna(b_series.iloc[0]):
                    bulk_stock_val = int(b_series.iloc[0])

            item_name = stores['item_name'].iloc[0] if 'item_name' in stores.columns else 'Unknown'
            color_name = stores['color_name'].iloc[0] if 'color_name' in stores.columns else 'Unknown'

            # 店舗別辞書リストの構築
            store_dicts = []
            for _, r in stores.iterrows():
                stk = int(r.get('stock_qty', 0)) if pd.notna(r.get('stock_qty', 0)) else 0
                a4 = float(r.get('avg_sales_4w', 0.0)) if pd.notna(r.get('avg_sales_4w', 0.0)) else 0.0
                af = float(r.get('avg_sales_full', 0.0)) if (has_avg_full_col and pd.notna(r.get('avg_sales_full', 0.0))) else 0.0

                # 有効週販（eff_avg）と有効WOS（effective_wos）
                if a4 > 0:
                    eff_avg = a4
                    eff_wos = stk / a4
                    is_fb = False
                elif af >= 0.25:
                    eff_avg = af
                    eff_wos = stk / af
                    is_fb = True
                else:
                    eff_avg = 0.0
                    eff_wos = np.nan
                    is_fb = False

                store_dicts.append({
                    'store': r['store'],
                    'sku': sku,
                    'item_name': item_name,
                    'color_name': color_name,
                    'stock_qty': stk,
                    'current_stock': stk,
                    'avg_sales_4w': a4,
                    'avg_sales_full': af,
                    'eff_avg': eff_avg,
                    'effective_wos': eff_wos,
                    'is_wos_fallback': is_fb,
                    'store_rank': ItemAllocator._get_store_priority_rank(r['store'])
                })

            # ==========================================
            # Phase 1: 滞留在庫集約（売上ゼロ店舗からの引き上げ）
            # ==========================================
            # 出荷元（滞留店）: 有効週販なし（直近4週売上0 かつ 全期間売上<0.25） かつ 在庫あり
            idle_shippers = [
                d for d in store_dicts
                if d['eff_avg'] == 0.0 and d['current_stock'] > 0
            ]

            # 受入先候補: 有効週販あり（直近4週売上>0 または 全期間売上>=0.25）
            active_receivers = [
                d for d in store_dicts
                if d['eff_avg'] > 0.0
            ]

            if idle_shippers and active_receivers:
                # 在庫が多い順、同数時はランク下位（数字が大きい）順に出荷
                idle_shippers.sort(key=lambda s: (-s['current_stock'], -s['store_rank']))

                for s in idle_shippers:
                    s_orig_stk = s['current_stock']
                    if s_orig_stk <= 0:
                        continue

                    # 出荷元以外の受入先候補
                    candidate_receivers = [r for r in active_receivers if r['store'] != s['store']]
                    if not candidate_receivers:
                        continue

                    k_receivers = len(candidate_receivers)

                    # 放出可能数量の算出
                    # 受入先が1店舗のみ: 手元に1点残す（在庫1点なら1点全量、在庫2点なら1点、在庫3点以上は残1点で移動）
                    # 受入先が2店舗以上: 全量移動可能
                    if k_receivers == 1:
                        avail_qty = 1 if s_orig_stk == 1 else (s_orig_stk - 1)
                    else:
                        avail_qty = s_orig_stk

                    # 受入先ごとの移動数量集計用
                    alloc_to_r = {id(r): 0 for r in candidate_receivers}
                    r_pre_stocks = {id(r): r['current_stock'] for r in candidate_receivers}

                    while avail_qty > 0 and s['current_stock'] > 0:
                        # 受入優先順位でソート
                        # 1. 直近4週売上あり (0) > なし (1)
                        # 2. 現在庫0 (0) > 在庫あり (1)
                        # 3. 現在の推定WOS（現在庫 / 週販）昇順
                        # 4. store_rank 昇順
                        def _rec_sort_key(r):
                            has_recent = 0 if r['avg_sales_4w'] > 0 else 1
                            is_zero = 0 if r['current_stock'] == 0 else 1
                            cur_w = (r['current_stock'] / r['eff_avg']) if r['eff_avg'] > 0 else 999.0
                            return (has_recent, is_zero, cur_w, r['store_rank'])

                        candidate_receivers.sort(key=_rec_sort_key)
                        target_r = candidate_receivers[0]

                        alloc_to_r[id(target_r)] += 1
                        target_r['current_stock'] += 1
                        s['current_stock'] -= 1
                        avail_qty -= 1

                    # 集計結果から移動レコードを生成
                    for r in candidate_receivers:
                        m_qty = alloc_to_r[id(r)]
                        if m_qty > 0:
                            r_pre_stk = r_pre_stocks[id(r)]
                            r_eff_avg = r['eff_avg']
                            r_fallback = r.get('is_wos_fallback', False)

                            receiver_pre_wos = ItemAllocator._fmt_wos(r_pre_stk, r_eff_avg)
                            receiver_post_wos = ItemAllocator._fmt_wos(r['current_stock'], r_eff_avg)

                            fallback_note = f"（{r['store']}は全期間平均週販を使用）" if r_fallback else ""
                            reason = (
                                f"{s['store']}は売上実績がなく在庫滞留（現在庫{s_orig_stk}点）、"
                                f"売上実績のある{r['store']}（現在庫{r_pre_stk}点）へ"
                                f"{m_qty}点集約移動を推奨"
                                + (f"。{fallback_note}" if fallback_note else "")
                            )

                            move_records.append({
                                'priority': priority,
                                'is_continuation': is_cont,
                                'sku': sku,
                                'item_name': item_name,
                                'color_name': color_name,
                                'bulk_stock': bulk_stock_val,
                                'sell_through': sell_through_val if has_sell_through else None,
                                'shipper': s['store'],
                                'shipper_stock': s_orig_stk,
                                'shipper_pre_wos': None,
                                'shipper_post_wos': None,
                                'move_qty': m_qty,
                                'receiver': r['store'],
                                'receiver_stock': r_pre_stk,
                                'receiver_pre_wos': receiver_pre_wos,
                                'receiver_post_wos': receiver_post_wos,
                                'is_shipper_fallback': False,
                                'is_receiver_fallback': r_fallback,
                                'reason': reason
                            })

            # ==========================================
            # Phase 2: 通常WOS平準化（売上実績店舗間）
            # ==========================================
            # Phase 1の移動後、有効なWOSを持つ店舗の effective_wos を更新して再計算
            valid_stores = [d for d in store_dicts if d['eff_avg'] > 0]
            if not valid_stores:
                continue

            for d in valid_stores:
                d['effective_wos'] = d['current_stock'] / d['eff_avg']

            wos_vals = [d['effective_wos'] for d in valid_stores]
            wos_avg = float(np.mean(wos_vals)) if wos_vals else 0.0

            if pd.isna(wos_avg) or wos_avg == 0:
                continue

            norm_shippers = [d for d in valid_stores if d['effective_wos'] > wos_avg and d['current_stock'] > 0]
            norm_receivers = [d for d in valid_stores if d['effective_wos'] < wos_avg]

            for s in norm_shippers:
                s['surplus'] = max(0.0, s['current_stock'] - (wos_avg * s['eff_avg']))
            for r in norm_receivers:
                r['deficit'] = max(0.0, (wos_avg * r['eff_avg']) - r['current_stock'])

            norm_shippers = [s for s in norm_shippers if s['surplus'] > 0]
            norm_receivers = [r for r in norm_receivers if r['deficit'] > 0]

            norm_shippers.sort(key=lambda s: -s['surplus'])
            norm_receivers.sort(key=lambda r: (-r['deficit'], r['store_rank']))

            for s in norm_shippers:
                if s['surplus'] <= 0 or s['current_stock'] <= 0:
                    continue

                for r in norm_receivers:
                    if r['store'] == s['store']:
                        continue
                    if r['deficit'] <= 0:
                        continue
                    if s['surplus'] <= 0 or s['current_stock'] <= 0:
                        break

                    move = min(s['surplus'], r['deficit'])
                    candidate_qty = math.ceil(move)
                    move_qty = min(candidate_qty, int(s['current_stock']))

                    if move_qty > 0:
                        s_fallback = s.get('is_wos_fallback', False)
                        r_fallback = r.get('is_wos_fallback', False)

                        s_eff_avg = s['eff_avg']
                        r_eff_avg = r['eff_avg']

                        shipper_pre_wos = ItemAllocator._fmt_wos(s['current_stock'], s_eff_avg)
                        receiver_pre_wos = ItemAllocator._fmt_wos(r['current_stock'], r_eff_avg)

                        shipper_post_stock = s['current_stock'] - move_qty
                        receiver_post_stock = r['current_stock'] + move_qty

                        shipper_post_wos = ItemAllocator._fmt_wos(shipper_post_stock, s_eff_avg)
                        receiver_post_wos = ItemAllocator._fmt_wos(receiver_post_stock, r_eff_avg)

                        fallback_note = ""
                        if s_fallback:
                            fallback_note += f"（{s['store']}は全期間平均週販を使用）"
                        if r_fallback:
                            fallback_note += f"（{r['store']}は全期間平均週販を使用）"

                        s_wos_display = s['effective_wos']
                        r_wos_display = r['effective_wos']

                        reason = (
                            f"{s['store']}のWOSが{s_wos_display:.1f}週{'(参考)' if s_fallback else ''}、"
                            f"{r['store']}のWOSが{r_wos_display:.1f}週{'(参考)' if r_fallback else ''}"
                            f"（全店平均{wos_avg:.1f}週）のため、"
                            f"{s['store']}から{r['store']}へ{move_qty}点移動を推奨"
                            + (f"。{fallback_note}" if fallback_note else "")
                        )

                        move_records.append({
                            'priority': priority,
                            'is_continuation': is_cont,
                            'sku': sku,
                            'item_name': item_name,
                            'color_name': color_name,
                            'bulk_stock': bulk_stock_val,
                            'sell_through': sell_through_val if has_sell_through else None,
                            'shipper': s['store'],
                            'shipper_stock': s['current_stock'],
                            'shipper_pre_wos': shipper_pre_wos,
                            'shipper_post_wos': shipper_post_wos,
                            'move_qty': move_qty,
                            'receiver': r['store'],
                            'receiver_stock': r['current_stock'],
                            'receiver_pre_wos': receiver_pre_wos,
                            'receiver_post_wos': receiver_post_wos,
                            'is_shipper_fallback': s_fallback,
                            'is_receiver_fallback': r_fallback,
                            'reason': reason
                        })

                        s['current_stock'] -= move_qty
                        r['current_stock'] += move_qty
                        s['surplus'] = max(0.0, s['surplus'] - move_qty)
                        r['deficit'] = max(0.0, r['deficit'] - move_qty)

        result = pd.DataFrame(move_records)
        return result

