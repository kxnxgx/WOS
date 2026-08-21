import pandas as pd
import numpy as np
from datetime import datetime

class Reporter:
    @staticmethod
    def _prepare_move_df(move_df: pd.DataFrame) -> pd.DataFrame:
        if move_df.empty: return move_df
        rename_dict = {
            'priority': '優先度',
            'sku': '商品コード',
            'item_name': '商品名',
            'color_name': 'カラー',
            'sell_through': '消化率(%)',
            'shipper': '出荷店舗',
            'shipper_stock': '出荷元在庫',
            'receiver': '受入店舗',
            'receiver_stock': '受入先在庫',
            'move_qty': '移動推奨数',
            'reason': '理由'
        }
        rename_dict = {k: v for k, v in rename_dict.items() if k in move_df.columns}
        return move_df.rename(columns=rename_dict)

    @staticmethod
    def _wos_color(val, vmin, vmax):
        """WOS値に応じた背景色を返す（赤=不足、白=平均、青=余剰）"""
        if pd.isna(val) or vmin == vmax:
            return '#ffffff'
        ratio = (val - vmin) / (vmax - vmin)
        if ratio < 0.5:
            t = ratio / 0.5
            r = 255
            g = int(255 * t)
            b = int(255 * t)
        else:
            t = (ratio - 0.5) / 0.5
            r = int(255 * (1 - t))
            g = int(255 * (1 - t))
            b = 255
        return f'rgb({r},{g},{b})'

    @staticmethod
    def _st_color(val):
        """消化率に応じた背景色を返す（緑=高、黄=中、赤=低）"""
        if pd.isna(val):
            return '#f0f0f0'
        if val >= 80:
            return '#d4edda'  # 緑
        elif val >= 60:
            return '#fff3cd'  # 黄
        else:
            return '#fde8e8'  # 赤

    @staticmethod
    def _build_heatmap_table(pivot_df: pd.DataFrame) -> str:
        """WOSヒートマップHTMLテーブルを生成する"""
        all_vals = pivot_df.values.flatten()
        valid_vals = all_vals[~np.isnan(all_vals.astype(float))]
        vmin = float(valid_vals.min()) if len(valid_vals) > 0 else 0
        vmax = float(valid_vals.max()) if len(valid_vals) > 0 else 1

        rows = []
        index_names = [n if n else '' for n in (pivot_df.index.names if hasattr(pivot_df.index, 'names') else [pivot_df.index.name])]
        col_headers = list(pivot_df.columns)
        header_html = ''.join(f'<th>{n}</th>' for n in index_names)
        header_html += ''.join(f'<th>{c}</th>' for c in col_headers)
        rows.append(f'<tr>{header_html}</tr>')

        for idx, row in pivot_df.iterrows():
            idx_vals = idx if isinstance(idx, tuple) else (idx,)
            cells = ''.join(f'<td class="idx-cell">{v}</td>' for v in idx_vals)
            for val in row:
                if pd.isna(val):
                    cells += '<td class="wos-cell na-cell">—</td>'
                else:
                    color = Reporter._wos_color(val, vmin, vmax)
                    cells += f'<td class="wos-cell" style="background:{color};" title="WOS: {val:.1f}週">{val:.1f}</td>'
            rows.append(f'<tr>{cells}</tr>')

        return f'<table class="heatmap-table"><thead>{rows[0]}</thead><tbody>{"".join(rows[1:])}</tbody></table>'

    @staticmethod
    def _build_move_table(df: pd.DataFrame) -> str:
        """移動推奨テーブルHTMLを生成する"""
        if df.empty:
            return '<p class="no-data">該当する移動推奨アイテムはありません</p>'
        
        disp = df.copy()
        
        # 継続品バッジを商品名に追加
        if 'is_continuation' in disp.columns:
            disp['item_name'] = disp.apply(
                lambda row: f"{row['item_name']} <span class='cont-badge'>🔄 継続品</span>" 
                            if row['is_continuation'] else row['item_name'], 
                axis=1
            )
            # is_continuation列は表示不要なので退避
            is_cont_series = disp['is_continuation']
            disp = disp.drop(columns=['is_continuation'])
        else:
            is_cont_series = pd.Series([False] * len(disp), index=disp.index)
            
        disp = Reporter._prepare_move_df(disp)
        
        # HTMLテーブルを生成
        html = disp.to_html(index=False, classes='move-table', border=0, escape=False)
        
        # 行ごとに継続品クラスを付与
        html_lines = html.splitlines()
        tr_idx = 0
        for i, line in enumerate(html_lines):
            if '<tr>' in line and '<th>' not in line:
                if tr_idx < len(is_cont_series):
                    if is_cont_series.iloc[tr_idx]:
                        html_lines[i] = line.replace('<tr>', '<tr class="continuation-row">')
                tr_idx += 1
                
        return '\n'.join(html_lines)

    @staticmethod
    def _build_summary_table(wos_df: pd.DataFrame, move_df: pd.DataFrame) -> str:
        """全社サマリーテーブルHTMLを生成する（消化率列・消化率順ソート対応）"""
        if wos_df.empty:
            return '<p>データがありません</p>'

        valid_wos = wos_df[wos_df['wos'].notna()]
        summary = valid_wos.groupby('sku').agg(
            avg_wos=('wos', 'mean'),
            min_wos=('wos', 'min'),
            max_wos=('wos', 'max'),
            store_count=('store', 'nunique'),
        ).reset_index()
        summary['avg_wos'] = summary['avg_wos'].round(1)
        summary['min_wos'] = summary['min_wos'].round(1)
        summary['max_wos'] = summary['max_wos'].round(1)

        if 'item_name' in wos_df.columns:
            master = wos_df[['sku', 'item_name']].drop_duplicates('sku')
            summary = summary.merge(master, on='sku', how='left')
        if 'color_name' in wos_df.columns:
            master_c = wos_df[['sku', 'color_name']].drop_duplicates('sku')
            summary = summary.merge(master_c, on='sku', how='left')

        # 消化率列の追加
        has_st = 'sell_through' in wos_df.columns
        if has_st:
            st_master = wos_df[['sku', 'sell_through', 'cumulative_sales', 'total_order']].drop_duplicates('sku')
            summary = summary.merge(st_master, on='sku', how='left')

        # 移動推奨件数
        if not move_df.empty:
            move_counts = move_df.groupby('sku').size().reset_index(name='move_count')
            summary = summary.merge(move_counts, on='sku', how='left')
        else:
            summary['move_count'] = 0
        summary['move_count'] = summary['move_count'].fillna(0).astype(int)

        # カラム整理
        col_order = ['sku']
        if 'item_name' in summary.columns: col_order.append('item_name')
        if 'color_name' in summary.columns: col_order.append('color_name')
        if has_st:
            col_order += ['sell_through', 'cumulative_sales', 'total_order']
        col_order += ['store_count', 'avg_wos', 'min_wos', 'max_wos', 'move_count']
        summary = summary[[c for c in col_order if c in summary.columns]]

        # 消化率がある場合は消化率降順、ない場合は平均WOS昇順
        if has_st and 'sell_through' in summary.columns:
            summary = summary.sort_values('sell_through', ascending=False)
        else:
            summary = summary.sort_values('avg_wos', ascending=True)

        col_rename = {
            'sku': '商品コード', 'item_name': '商品名', 'color_name': 'カラー',
            'sell_through': '消化率(%)', 'cumulative_sales': '累計売上数',
            'total_order': '発注数', 'store_count': '対象店舗数',
            'avg_wos': '平均WOS', 'min_wos': '最小WOS', 'max_wos': '最大WOS',
            'move_count': '移動推奨件数'
        }
        summary = summary.rename(columns={k: v for k, v in col_rename.items() if k in summary.columns})

        overall_avg_wos = summary['平均WOS'].mean()

        rows = []
        headers = list(summary.columns)
        rows.append('<tr>' + ''.join(f'<th>{h}</th>' for h in headers) + '</tr>')

        for _, row in summary.iterrows():
            cells = ''
            for col in headers:
                val = row[col]
                cell_style = ''
                if col == '消化率(%)' and pd.notna(val):
                    bg = Reporter._st_color(val)
                    cell_style = f' style="background:{bg};font-weight:bold;"'
                    val = f'{val:.1f}%'
                elif col == '移動推奨件数' and val > 0:
                    cell_style = ' style="background:#fff3cd;font-weight:bold;"'
                elif col == '平均WOS' and pd.notna(val):
                    if val < overall_avg_wos * 0.7:
                        cell_style = ' style="background:#fde8e8;"'
                    elif val > overall_avg_wos * 1.3:
                        cell_style = ' style="background:#e8f0fd;"'
                cells += f'<td{cell_style}>{val}</td>'
            rows.append(f'<tr>{cells}</tr>')

        return (
            f'<p style="color:#666;font-size:0.9em;">全商品の平均WOS: <strong>{overall_avg_wos:.1f}週</strong></p>'
            f'<table class="summary-table"><thead>{rows[0]}</thead><tbody>{"".join(rows[1:])}</tbody></table>'
        )

    @staticmethod
    def _apply_excel_formatting(writer: pd.ExcelWriter, sheet_name: str, df: pd.DataFrame):
        """Excelのシートに書式（継続品の色付けなど）を適用する"""
        from openpyxl.styles import PatternFill
        workbook = writer.book
        if sheet_name in workbook.sheetnames:
            worksheet = workbook[sheet_name]
            # 青系の薄い色
            cont_fill = PatternFill(start_color="E6F2FF", end_color="E6F2FF", fill_type="solid")
            
            # is_continuation はデータフレームのどこにあるか？
            # 準備前の df からインデックスを取得する
            if 'is_continuation' in df.columns:
                is_cont_series = df['is_continuation'].reset_index(drop=True)
                for row_idx, is_cont in enumerate(is_cont_series):
                    if is_cont:
                        # ヘッダーが1行目なのでデータは2行目から
                        for col_idx in range(1, worksheet.max_column + 1):
                            worksheet.cell(row=row_idx + 2, column=col_idx).fill = cont_fill

    @staticmethod
    def generate_excel(wos_df: pd.DataFrame, move_df: pd.DataFrame, output_path: str = "WOS_Report.xlsx"):
        """Excelレポートの出力"""
        print(f"Excelレポートを生成しています: {output_path}")

        with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
            # 優先集約リスト
            priority_df = move_df[move_df['priority'] == '優先'].copy() if not move_df.empty and 'priority' in move_df.columns else pd.DataFrame()
            normal_df = move_df[move_df['priority'] != '優先'].copy() if not move_df.empty and 'priority' in move_df.columns else move_df.copy()

            # 表示用DF作成時に is_continuation は削除されるため、元のDataFrameを保持
            def _prepare_and_drop(df):
                if df.empty: return df
                disp = df.copy()
                if 'is_continuation' in disp.columns:
                    disp['item_name'] = disp.apply(
                        lambda row: f"🔄 {row['item_name']}" if row['is_continuation'] else row['item_name'], axis=1
                    )
                    disp = disp.drop(columns=['is_continuation'])
                return Reporter._prepare_move_df(disp)

            disp_priority = _prepare_and_drop(priority_df)
            disp_normal = _prepare_and_drop(normal_df)

            if not disp_priority.empty:
                disp_priority.to_excel(writer, sheet_name='⭐優先集約リスト', index=False)
                Reporter._apply_excel_formatting(writer, '⭐優先集約リスト', priority_df)
            else:
                pd.DataFrame([{"Message": "優先集約アイテムはありません"}]).to_excel(writer, sheet_name='⭐優先集約リスト', index=False)

            if not disp_normal.empty:
                disp_normal.to_excel(writer, sheet_name='📦通常集約リスト', index=False)
                Reporter._apply_excel_formatting(writer, '📦通常集約リスト', normal_df)
            else:
                pd.DataFrame([{"Message": "通常集約アイテムはありません"}]).to_excel(writer, sheet_name='📦通常集約リスト', index=False)

            if not wos_df.empty:
                idx_cols = ['sku']
                if 'item_name' in wos_df.columns: idx_cols.append('item_name')
                if 'color_name' in wos_df.columns: idx_cols.append('color_name')
                if 'sell_through' in wos_df.columns: idx_cols.append('sell_through')
                if 'is_continuation' in wos_df.columns: idx_cols.append('is_continuation')

                pivot_wos = wos_df.pivot(index=idx_cols, columns='store', values='wos').round(1)
                pivot_wos.to_excel(writer, sheet_name='店舗別WOSサマリー')
                wos_df.to_excel(writer, sheet_name='WOS生データ', index=False)

    @staticmethod
    def generate_html(wos_df: pd.DataFrame, move_df: pd.DataFrame,
                      output_path: str = "WOS_Report.html",
                      threshold: float = 80.0):
        """HTMLレポートの出力（4タブ・消化率対応）"""
        print(f"HTMLレポートを生成しています: {output_path}")

        generated_at = datetime.now().strftime('%Y年%m月%d日 %H:%M')
        has_sell_through = not move_df.empty and 'sell_through' in move_df.columns

        # --- 優先集約 / 通常集約 の分離 ---
        if not move_df.empty and 'priority' in move_df.columns:
            priority_move = move_df[move_df['priority'] == '優先'].copy()
            normal_move = move_df[move_df['priority'] != '優先'].copy()
        else:
            priority_move = pd.DataFrame()
            normal_move = move_df.copy()

        # 消化率降順ソート
        if not priority_move.empty and 'sell_through' in priority_move.columns:
            priority_move = priority_move.sort_values('sell_through', ascending=False)
        if not normal_move.empty and 'sell_through' in normal_move.columns:
            normal_move = normal_move.sort_values('sell_through', ascending=False)

        priority_html = Reporter._build_move_table(priority_move)
        normal_html = Reporter._build_move_table(normal_move)
        summary_html = Reporter._build_summary_table(wos_df, move_df)

        # WOS ヒートマップ
        if not wos_df.empty:
            idx_cols = ['sku']
            if 'item_name' in wos_df.columns: idx_cols.append('item_name')
            if 'color_name' in wos_df.columns: idx_cols.append('color_name')
            pivot_wos = wos_df.pivot(index=idx_cols, columns='store', values='wos').round(1)
            heatmap_html = Reporter._build_heatmap_table(pivot_wos)
        else:
            heatmap_html = '<p class="no-data">データがありません</p>'

        priority_count = len(priority_move)
        normal_count = len(normal_move)
        sku_count = wos_df['sku'].nunique() if not wos_df.empty else 0
        store_count = wos_df['store'].nunique() if not wos_df.empty else 0

        st_badge = f'（消化率 ≥ {threshold:.0f}% 基準）' if has_sell_through else '（消化率データなし）'

        html_content = f"""<!DOCTYPE html>
<html lang="ja">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>WOS アイテム集約レポート</title>
    <style>
        *, *::before, *::after {{ box-sizing: border-box; }}
        body {{
            font-family: 'Segoe UI', 'Meiryo', Tahoma, Geneva, Verdana, sans-serif;
            margin: 0; padding: 20px 30px; color: #2d3748;
            background: #f7fafc;
        }}
        h1 {{
            color: #1a202c; font-size: 1.6em;
            border-bottom: 3px solid #4299e1; padding-bottom: 10px;
            margin-bottom: 6px;
        }}
        .meta {{ color: #718096; font-size: 0.85em; margin-bottom: 20px; }}
        .kpi-bar {{ display: flex; gap: 16px; margin-bottom: 24px; flex-wrap: wrap; }}
        .kpi {{
            background: #fff; border-radius: 8px; padding: 14px 22px;
            box-shadow: 0 1px 4px rgba(0,0,0,0.1); min-width: 140px;
        }}
        .kpi .label {{ font-size: 0.78em; color: #718096; margin-bottom: 4px; }}
        .kpi .value {{ font-size: 1.7em; font-weight: 700; color: #2b6cb0; }}
        .kpi.priority .value {{ color: #c05621; }}
        .logic-box {{
            background: #ebf8ff; border-left: 4px solid #4299e1;
            padding: 10px 16px; margin-bottom: 20px; border-radius: 0 6px 6px 0;
            font-size: 0.88em;
        }}
        .logic-box summary {{ font-weight: bold; cursor: pointer; color: #2b6cb0; }}
        .logic-box ul {{ margin: 8px 0 0 0; padding-left: 20px; color: #4a5568; }}
        .tabs {{ display: flex; border-bottom: 2px solid #e2e8f0; margin-bottom: 20px; }}
        .tab {{
            padding: 10px 24px; cursor: pointer; border: none;
            background: none; font-size: 0.95em; color: #718096;
            border-bottom: 3px solid transparent; margin-bottom: -2px;
            transition: color 0.15s;
        }}
        .tab:hover {{ color: #4299e1; }}
        .tab.active {{ color: #2b6cb0; font-weight: bold; border-bottom-color: #4299e1; }}
        .tab.priority-tab.active {{ color: #c05621; border-bottom-color: #ed8936; }}
        .tab-content {{ display: none; }}
        .tab-content.active {{ display: block; }}
        /* 移動推奨テーブル */
        .move-table {{
            border-collapse: collapse; width: 100%;
            box-shadow: 0 1px 6px rgba(0,0,0,0.08); border-radius: 8px; overflow: hidden;
        }}
        .move-table th {{
            background: #2b6cb0; color: #fff;
            padding: 10px 14px; text-align: left; font-weight: 600; font-size: 0.88em;
        }}
        .move-table td {{
            border-bottom: 1px solid #e2e8f0; padding: 9px 14px; font-size: 0.87em;
        }}
        .move-table tr:last-child td {{ border-bottom: none; }}
        .move-table tr:hover td {{ background: #ebf8ff; }}
        .move-table tr.continuation-row td {{ background: #e6f2ff; }}
        .move-table tr.continuation-row:hover td {{ background: #cce6ff; }}
        .cont-badge {{
            display: inline-block; background: #e1effe; color: #1e40af;
            border: 1px solid #93c5fd; border-radius: 4px;
            padding: 1px 6px; font-size: 0.82em; margin-left: 6px;
        }}
        /* ヒートマップテーブル */
        .heatmap-table {{
            border-collapse: collapse; font-size: 0.82em;
            box-shadow: 0 1px 6px rgba(0,0,0,0.08);
        }}
        .heatmap-table th {{
            background: #4a5568; color: #fff;
            padding: 8px 12px; text-align: center; white-space: nowrap;
        }}
        .heatmap-table .idx-cell {{
            background: #f7fafc; padding: 7px 12px;
            border-bottom: 1px solid #e2e8f0; white-space: nowrap;
            font-size: 0.85em; color: #4a5568;
        }}
        .heatmap-table .wos-cell {{
            text-align: center; padding: 7px 10px; min-width: 54px;
            border-bottom: 1px solid #e2e8f0;
            font-weight: 600; font-size: 0.92em;
        }}
        .heatmap-table .na-cell {{ color: #a0aec0; font-weight: normal; background: #f7fafc; }}
        .heatmap-table tr:hover .wos-cell {{ filter: brightness(0.9); }}
        /* 凡例 */
        .legend {{
            display: flex; align-items: center; gap: 8px;
            font-size: 0.82em; color: #718096; margin-bottom: 14px;
        }}
        .legend-bar {{
            width: 180px; height: 14px; border-radius: 4px;
            background: linear-gradient(to right, #ff8080, #ffffff, #8080ff);
            border: 1px solid #e2e8f0;
        }}
        /* 全社サマリーテーブル */
        .summary-table {{
            border-collapse: collapse; width: 100%;
            box-shadow: 0 1px 6px rgba(0,0,0,0.08); border-radius: 8px; overflow: hidden;
        }}
        .summary-table th {{
            background: #553c9a; color: #fff;
            padding: 10px 14px; text-align: left; font-size: 0.88em; font-weight: 600;
        }}
        .summary-table td {{
            border-bottom: 1px solid #e2e8f0; padding: 8px 14px; font-size: 0.87em;
        }}
        .summary-table tr:hover td {{ background: #faf5ff; }}
        .no-data {{ color: #a0aec0; font-style: italic; }}
        .overflow-x {{ overflow-x: auto; }}
        .st-badge {{
            display: inline-block; background: #fef3c7; color: #92400e;
            border: 1px solid #f59e0b; border-radius: 4px;
            padding: 2px 10px; font-size: 0.82em; margin-left: 8px;
        }}
    </style>
    <script>
        function showTab(tabId, el) {{
            document.querySelectorAll('.tab-content').forEach(e => e.classList.remove('active'));
            document.querySelectorAll('.tab').forEach(e => e.classList.remove('active'));
            document.getElementById(tabId).classList.add('active');
            el.classList.add('active');
        }}
    </script>
</head>
<body>
    <h1>WOS アイテム集約レポート</h1>
    <p class="meta">生成日時: {generated_at}</p>

    <div class="kpi-bar">
        <div class="kpi"><div class="label">対象SKU数</div><div class="value">{sku_count}</div></div>
        <div class="kpi"><div class="label">対象店舗数</div><div class="value">{store_count}</div></div>
        <div class="kpi priority"><div class="label">⭐ 優先集約件数</div><div class="value">{priority_count}</div></div>
        <div class="kpi"><div class="label">📦 通常集約件数</div><div class="value">{normal_count}</div></div>
    </div>

    <details class="logic-box">
        <summary>計算ロジックについて（クリックで展開）</summary>
        <ul>
            <li><strong>WOS（週数）</strong>: 現在在庫 ÷ 直近4週の平均販売数</li>
            <li><strong>消化率</strong>: 全期間累計売上数 ÷ 発注数（総計列）× 100</li>
            <li><strong>⭐ 優先集約</strong>: 消化率 ≥ {threshold:.0f}% の SKU（完売を狙う）</li>
            <li><strong>📦 通常集約</strong>: 消化率 &lt; {threshold:.0f}% または消化率データなし</li>
            <li><strong>出荷候補</strong>: SKUの全店舗平均WOSより高い店舗（在庫余剰）</li>
            <li><strong>受入候補</strong>: SKUの全店舗平均WOSより低い店舗（在庫不足）</li>
        </ul>
    </details>

    <div class="tabs">
        <button class="tab priority-tab active" onclick="showTab('tab-priority', this)">⭐ 優先集約リスト<span class="st-badge">{st_badge}</span></button>
        <button class="tab" onclick="showTab('tab-normal', this)">📦 通常集約リスト</button>
        <button class="tab" onclick="showTab('tab-summary', this)">📊 全社サマリー</button>
        <button class="tab" onclick="showTab('tab-wos', this)">🌡 WOS ヒートマップ</button>
    </div>

    <div id="tab-priority" class="tab-content active">
        <h2 style="font-size:1.1em;color:#c05621;margin-bottom:8px;">⭐ 優先集約リスト（消化率 ≥ {threshold:.0f}%）</h2>
        <p style="font-size:0.85em;color:#718096;margin-bottom:12px;">
            完売を狙える商品です。消化率の高い順に並んでいます。
        </p>
        <div class="overflow-x">{priority_html}</div>
    </div>

    <div id="tab-normal" class="tab-content">
        <h2 style="font-size:1.1em;color:#2b6cb0;margin-bottom:8px;">📦 通常集約リスト（消化率 &lt; {threshold:.0f}%）</h2>
        <p style="font-size:0.85em;color:#718096;margin-bottom:12px;">
            在庫の偏りを解消する移動推奨リストです。
        </p>
        <div class="overflow-x">{normal_html}</div>
    </div>

    <div id="tab-summary" class="tab-content">
        <h2 style="font-size:1.1em;color:#553c9a;margin-bottom:8px;">📊 全社サマリー（SKU別統計）</h2>
        <p style="font-size:0.85em;color:#718096;margin-bottom:12px;">
            消化率の高い順に並んでいます。🟢 ≥ 80% / 🟡 60〜79% / 🔴 &lt; 60%
        </p>
        <div class="overflow-x">{summary_html}</div>
    </div>

    <div id="tab-wos" class="tab-content">
        <h2 style="font-size:1.1em;color:#2b6cb0;margin-bottom:8px;">🌡 店舗別 WOS ヒートマップ</h2>
        <div class="legend">
            <span>在庫不足</span>
            <div class="legend-bar"></div>
            <span>在庫余剰</span>
            &nbsp;|&nbsp; セルにカーソルを合わせるとWOS値が表示されます
        </div>
        <div class="overflow-x">{heatmap_html}</div>
    </div>

</body>
</html>
"""
        with open(output_path, "w", encoding="utf-8") as f:
            f.write(html_content)

