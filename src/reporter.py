import pandas as pd
import json

class Reporter:
    @staticmethod
    def _prepare_move_df(move_df: pd.DataFrame) -> pd.DataFrame:
        if move_df.empty: return move_df
        rename_dict = {
            'sku': '商品コード',
            'item_name': '商品名',
            'color_name': 'カラー',
            'shipper': '出荷店舗',
            'shipper_stock': '出荷元在庫',
            'receiver': '受入店舗',
            'receiver_stock': '受入先在庫',
            'move_qty': '移動推奨数',
            'reason': '理由'
        }
        return move_df.rename(columns=rename_dict)

    @staticmethod
    def generate_excel(wos_df: pd.DataFrame, move_df: pd.DataFrame, output_path: str = "WOS_Report.xlsx"):
        """Excelレポートの出力"""
        print(f"Excelレポートを生成しています: {output_path}")
        
        disp_move_df = Reporter._prepare_move_df(move_df)
        
        with pd.ExcelWriter(output_path) as writer:
            if not disp_move_df.empty:
                disp_move_df.to_excel(writer, sheet_name='アイテム移動推奨リスト', index=False)
            else:
                pd.DataFrame([{"Message": "移動推奨アイテムはありません"}]).to_excel(writer, sheet_name='アイテム移動推奨リスト', index=False)
                
            if not wos_df.empty:
                # WOSのヒートマップ用ピボットテーブル
                # 商品名とColorNameが存在する場合はindexに含める
                idx_cols = ['sku']
                if 'item_name' in wos_df.columns: idx_cols.append('item_name')
                if 'color_name' in wos_df.columns: idx_cols.append('color_name')
                
                pivot_wos = wos_df.pivot(index=idx_cols, columns='store', values='wos')
                pivot_wos.to_excel(writer, sheet_name='店舗別WOSサマリー')
                
                wos_df.to_excel(writer, sheet_name='WOS生データ', index=False)

    @staticmethod
    def generate_html(wos_df: pd.DataFrame, move_df: pd.DataFrame, output_path: str = "WOS_Report.html"):
        """HTMLレポートの出力（自己完結型）"""
        print(f"HTMLレポートを生成しています: {output_path}")
        
        disp_move_df = Reporter._prepare_move_df(move_df)
        move_html = disp_move_df.to_html(index=False, classes='table table-striped table-hover') if not disp_move_df.empty else "<p>移動推奨アイテムはありません</p>"
        
        if not wos_df.empty:
            idx_cols = ['sku']
            if 'item_name' in wos_df.columns: idx_cols.append('item_name')
            if 'color_name' in wos_df.columns: idx_cols.append('color_name')
            
            pivot_wos = wos_df.pivot(index=idx_cols, columns='store', values='wos').round(1)
            pivot_html = pivot_wos.to_html(classes='table table-bordered table-sm')
        else:
            pivot_html = "<p>データがありません</p>"
            
        html_content = f"""<!DOCTYPE html>
<html lang="ja">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>WOS アイテム集約レポート</title>
    <style>
        body {{ font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; margin: 20px; color: #333; }}
        h1 {{ color: #2c3e50; border-bottom: 2px solid #3498db; padding-bottom: 10px; }}
        .tabs {{ display: flex; border-bottom: 1px solid #ccc; margin-bottom: 20px; }}
        .tab {{ padding: 10px 20px; cursor: pointer; border: 1px solid transparent; }}
        .tab:hover {{ background-color: #f1f1f1; }}
        .tab.active {{ background-color: #fff; border-color: #ccc #ccc transparent; border-bottom-color: #fff; margin-bottom: -1px; font-weight: bold; color: #3498db; }}
        .tab-content {{ display: none; }}
        .tab-content.active {{ display: block; }}
        table {{ border-collapse: collapse; width: 100%; margin-bottom: 20px; box-shadow: 0 0 20px rgba(0,0,0,0.1); }}
        th, td {{ border: 1px solid #ddd; padding: 12px; text-align: left; }}
        th {{ background-color: #f4f6f7; color: #2c3e50; }}
        tr:nth-child(even) {{ background-color: #f9f9f9; }}
        tr:hover {{ background-color: #f1f1f1; }}
        details {{ background: #f8f9fa; padding: 10px; border-left: 4px solid #3498db; margin-bottom: 20px; }}
        summary {{ font-weight: bold; cursor: pointer; outline: none; }}
    </style>
    <script>
        function showTab(tabId, element) {{
            document.querySelectorAll('.tab-content').forEach(el => el.classList.remove('active'));
            document.querySelectorAll('.tab').forEach(el => el.classList.remove('active'));
            document.getElementById(tabId).classList.add('active');
            element.classList.add('active');
        }}
    </script>
</head>
<body>
    <h1>WOS アイテム集約レポート</h1>
    
    <details>
        <summary>計算式・ロジックについて（クリックで展開）</summary>
        <ul>
            <li><strong>WOS（週数）</strong>: 現在在庫 ÷ 最近4週の平均販売数</li>
            <li><strong>出荷判定</strong>: 各SKUの全店舗平均WOSより高い店舗（在庫余剰）を出荷候補とします。</li>
            <li><strong>受入判定</strong>: 各SKUの全店舗平均WOSより低い店舗（在庫不足）を受入れ候補とします。</li>
            <li><strong>移動数量</strong>: 余剰分と不足分を比較し、両者がWOS平均に近づくように算出されます。</li>
        </ul>
    </details>

    <div class="tabs">
        <div class="tab active" onclick="showTab('tab-move', this)">アイテム集約リスト</div>
        <div class="tab" onclick="showTab('tab-wos', this)">店舗別WOSサマリー</div>
    </div>

    <div id="tab-move" class="tab-content active">
        <h2>アイテム移動推奨リスト</h2>
        {move_html}
    </div>

    <div id="tab-wos" class="tab-content">
        <h2>店舗別WOSサマリー (WOS値)</h2>
        <div style="overflow-x: auto;">
            {pivot_html}
        </div>
    </div>
</body>
</html>
"""
        with open(output_path, "w", encoding="utf-8") as f:
            f.write(html_content)
