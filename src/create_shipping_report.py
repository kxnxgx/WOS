"""
create_shipping_report.py
-------------------------
WOS_Report.xlsx から目星（黄色ハイライト）を読み取り、
バックアップを作成の上、店舗出荷指示に特化した「WOS_Report_店舗出荷.xlsx」を生成するスクリプト。
"""

import os
import shutil
import sys
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# 店舗名正規化マッピング
STORE_NAME_MAP = {
    'FJALLRAVEN by 3NITY TOKYO': 'TOKYO',
    'TOKYO': 'TOKYO',
    'FJALLRAVEN by 3NITY TOKYO NODE': 'TOKYO NODE',
    'TOKYO NODE': 'TOKYO NODE',
    'FJALLRAVEN by 3NITY 京王新宿': '京王新宿',
    '京王新宿': '京王新宿',
    'FJALLRAVEN by 3NITY玉川高島屋S・C': '玉川高島屋',
    '玉川高島屋': '玉川高島屋',
    'FJALLRAVEN POPUP NARITA': 'NARITA',
    'NARITA': 'NARITA',
    'FJALLRAVEN by 3NITY SAPPORO HUTTE': 'SAPPORO HUTTE',
    'SAPPORO HUTTE': 'SAPPORO HUTTE',
    'FJALLRAVEN STORE 名古屋ファッションワン': '名古屋',
    '名古屋': '名古屋',
    'FJALLRAVEN by 3NITY 大丸心斎橋': '大丸心斎橋',
    '大丸心斎橋': '大丸心斎橋',
    'FJALLRAVEN by 3NITY ルクア大阪': 'ルクア大阪',
    'ルクア大阪': 'ルクア大阪',
}

ALL_STORES = [
    'TOKYO', 'TOKYO NODE', '京王新宿', '玉川高島屋',
    'NARITA', 'SAPPORO HUTTE', '名古屋', '大丸心斎橋', 'ルクア大阪'
]

def normalize_store(name):
    if not name:
        return ''
    s = str(name).strip()
    return STORE_NAME_MAP.get(s, s)

def load_data_from_excel(filepath):
    """Excelから優先集約リストと通常集約リストを読み込み、黄色ハイライト行を特定する"""
    wb = openpyxl.load_workbook(filepath, data_only=True)
    records = []

    for sheet_name in ['⭐優先集約リスト', '📦通常集約リスト']:
        if sheet_name not in wb.sheetnames:
            continue
        ws = wb[sheet_name]
        headers = [cell.value for cell in ws[1]]
        
        for r in range(2, ws.max_row + 1):
            row_cells = ws[r]
            row_dict = {}
            for h, cell in zip(headers, row_cells):
                if h:
                    row_dict[h] = cell.value

            # 黄色ハイライトの判定（ユーザー指定の FFFFFF00 やその他明示的黄色）
            is_yellow = False
            for cell in row_cells:
                if cell.fill and cell.fill.fill_type:
                    fg = cell.fill.fgColor
                    rgb = getattr(fg, 'rgb', None)
                    if rgb == 'FFFFFF00' or (rgb and 'FFFF00' in str(rgb)):
                        is_yellow = True
                        break
            
            row_dict['is_selected'] = is_yellow
            row_dict['original_sheet'] = sheet_name
            records.append(row_dict)

    df = pd.DataFrame(records)
    
    # 店舗名の正規化
    if '出荷店舗' in df.columns:
        df['出荷店舗'] = df['出荷店舗'].apply(normalize_store)
    if '受入店舗' in df.columns:
        df['受入店舗'] = df['受入店舗'].apply(normalize_store)

    return df

def setup_styles():
    """各種セルの書式スタイルを定義"""
    font_main = Font(name='Meiryo UI', size=9)
    font_bold = Font(name='Meiryo UI', size=9, bold=True)
    font_title = Font(name='Meiryo UI', size=13, bold=True, color='1F4E79')
    font_subtitle = Font(name='Meiryo UI', size=9.5, bold=True, color='595959')
    font_header = Font(name='Meiryo UI', size=9, bold=True, color='FFFFFF')

    fill_header_navy = PatternFill(start_color='1F4E79', end_color='1F4E79', fill_type='solid') # 統合リスト用
    fill_header_blue = PatternFill(start_color='2E75B6', end_color='2E75B6', fill_type='solid') # 店舗出荷用
    fill_header_summary = PatternFill(start_color='333F48', end_color='333F48', fill_type='solid') # サマリー用

    fill_yellow_selected = PatternFill(start_color='FFF2CC', end_color='FFF2CC', fill_type='solid') # 選定行
    fill_cont_blue = PatternFill(start_color='E6F2FF', end_color='E6F2FF', fill_type='solid') # 継続品
    fill_bulk_orange = PatternFill(start_color='FFEAA7', end_color='FFEAA7', fill_type='solid') # BULK在庫
    fill_highlight_qty = PatternFill(start_color='E2EFDA', end_color='E2EFDA', fill_type='solid') # サマリー強調
    fill_total_row = PatternFill(start_color='F2F2F2', end_color='F2F2F2', fill_type='solid') # 合計行

    thin_border_side = Side(border_style='thin', color='D9D9D9')
    border_all_thin = Border(left=thin_border_side, right=thin_border_side, top=thin_border_side, bottom=thin_border_side)
    
    double_bottom_side = Side(border_style='double', color='000000')
    border_total_row = Border(left=thin_border_side, right=thin_border_side, top=thin_border_side, bottom=double_bottom_side)

    align_center = Alignment(horizontal='center', vertical='center')
    align_left = Alignment(horizontal='left', vertical='center')
    align_right = Alignment(horizontal='right', vertical='center')
    align_wrap_left = Alignment(horizontal='left', vertical='center', wrap_text=True)

    return {
        'font_main': font_main,
        'font_bold': font_bold,
        'font_title': font_title,
        'font_subtitle': font_subtitle,
        'font_header': font_header,
        'fill_header_navy': fill_header_navy,
        'fill_header_blue': fill_header_blue,
        'fill_header_summary': fill_header_summary,
        'fill_yellow_selected': fill_yellow_selected,
        'fill_cont_blue': fill_cont_blue,
        'fill_bulk_orange': fill_bulk_orange,
        'fill_highlight_qty': fill_highlight_qty,
        'fill_total_row': fill_total_row,
        'border_all_thin': border_all_thin,
        'border_total_row': border_total_row,
        'align_center': align_center,
        'align_left': align_left,
        'align_right': align_right,
        'align_wrap_left': align_wrap_left,
    }

def build_shipping_summary_sheet(wb, df_selected, styles):
    """シート1: 📋店舗別出荷サマリー の作成"""
    ws = wb.create_sheet(title='📋店舗別出荷サマリー', index=0)
    ws.views.sheetView[0].showGridLines = True

    # タイトル
    ws['B2'] = "【店舗間移動 出荷・受入サマリー（選定アイテム計40件）】"
    ws['B2'].font = styles['font_title']
    ws['B3'] = "※各店舗が出荷する点数および受入店舗の内訳一覧です。"
    ws['B3'].font = styles['font_subtitle']

    # 集計マトリクスの作成
    summary_df = df_selected.groupby(['出荷店舗', '受入店舗'])['移動推奨数'].sum().unstack(fill_value=0)
    summary_df = summary_df.reindex(index=ALL_STORES, columns=ALL_STORES, fill_value=0)
    
    start_row = 5
    start_col = 2  # B列から

    # テーブルヘッダー
    ws.cell(row=start_row, column=start_col, value="出荷元店舗 ＼ 受入先店舗").font = styles['font_header']
    ws.cell(row=start_row, column=start_col).fill = styles['fill_header_summary']
    ws.cell(row=start_row, column=start_col).alignment = styles['align_center']
    ws.cell(row=start_row, column=start_col).border = styles['border_all_thin']

    for c_idx, recv_store in enumerate(ALL_STORES, start=start_col + 1):
        cell = ws.cell(row=start_row, column=c_idx, value=recv_store)
        cell.font = styles['font_header']
        cell.fill = styles['fill_header_summary']
        cell.alignment = styles['align_center']
        cell.border = styles['border_all_thin']

    # 出荷合計列ヘッダー
    total_col_idx = start_col + len(ALL_STORES) + 1
    cell = ws.cell(row=start_row, column=total_col_idx, value="出荷合計点数")
    cell.font = styles['font_header']
    cell.fill = styles['fill_header_navy']
    cell.alignment = styles['align_center']
    cell.border = styles['border_all_thin']

    # アイテム件数列ヘッダー
    item_col_idx = total_col_idx + 1
    cell = ws.cell(row=start_row, column=item_col_idx, value="選定アイテム数")
    cell.font = styles['font_header']
    cell.fill = styles['fill_header_navy']
    cell.alignment = styles['align_center']
    cell.border = styles['border_all_thin']

    # データ行
    for r_idx, ship_store in enumerate(ALL_STORES, start=start_row + 1):
        store_cell = ws.cell(row=r_idx, column=start_col, value=ship_store)
        store_cell.font = styles['font_bold']
        store_cell.fill = styles['fill_total_row']
        store_cell.alignment = styles['align_center']
        store_cell.border = styles['border_all_thin']

        row_total_qty = 0
        for c_idx, recv_store in enumerate(ALL_STORES, start=start_col + 1):
            qty = int(summary_df.loc[ship_store, recv_store])
            row_total_qty += qty
            cell = ws.cell(row=r_idx, column=c_idx, value=qty if qty > 0 else "-")
            cell.font = styles['font_bold'] if qty > 0 else styles['font_main']
            cell.alignment = styles['align_center']
            cell.border = styles['border_all_thin']
            if qty > 0:
                cell.fill = styles['fill_highlight_qty']

        # 出荷合計点数
        cell_total = ws.cell(row=r_idx, column=total_col_idx, value=row_total_qty)
        cell_total.font = styles['font_bold']
        cell_total.alignment = styles['align_center']
        cell_total.border = styles['border_all_thin']
        if row_total_qty > 0:
            cell_total.fill = styles['fill_yellow_selected']

        # 選定アイテム数
        item_count = len(df_selected[df_selected['出荷店舗'] == ship_store])
        cell_item = ws.cell(row=r_idx, column=item_col_idx, value=f"{item_count}件" if item_count > 0 else "-")
        cell_item.font = styles['font_bold'] if item_count > 0 else styles['font_main']
        cell_item.alignment = styles['align_center']
        cell_item.border = styles['border_all_thin']

    # 最下部：受入合計行
    bottom_row = start_row + len(ALL_STORES) + 1
    total_label_cell = ws.cell(row=bottom_row, column=start_col, value="受入合計点数")
    total_label_cell.fill = styles['fill_header_navy']
    total_label_cell.font = styles['font_header']
    total_label_cell.alignment = styles['align_center']
    total_label_cell.border = styles['border_total_row']

    grand_total_qty = 0
    for c_idx, recv_store in enumerate(ALL_STORES, start=start_col + 1):
        col_total_qty = int(summary_df[recv_store].sum())
        grand_total_qty += col_total_qty
        cell = ws.cell(row=bottom_row, column=c_idx, value=col_total_qty if col_total_qty > 0 else "-")
        cell.font = styles['font_bold']
        cell.alignment = styles['align_center']
        cell.border = styles['border_total_row']
        if col_total_qty > 0:
            cell.fill = styles['fill_yellow_selected']

    # 総合計点数
    grand_total_cell = ws.cell(row=bottom_row, column=total_col_idx, value=grand_total_qty)
    grand_total_cell.font = Font(name='Meiryo UI', size=11, bold=True, color='C00000')
    grand_total_cell.alignment = styles['align_center']
    grand_total_cell.fill = styles['fill_yellow_selected']
    grand_total_cell.border = styles['border_total_row']

    # 総アイテム数
    total_items_cell = ws.cell(row=bottom_row, column=item_col_idx, value=f"{len(df_selected)}件")
    total_items_cell.font = Font(name='Meiryo UI', size=10, bold=True, color='C00000')
    total_items_cell.alignment = styles['align_center']
    total_items_cell.fill = styles['fill_yellow_selected']
    total_items_cell.border = styles['border_total_row']

    # 列幅調整
    ws.column_dimensions['A'].width = 3
    ws.column_dimensions['B'].width = 24
    for c_idx in range(start_col + 1, total_col_idx + 2):
        col_letter = get_column_letter(c_idx)
        ws.column_dimensions[col_letter].width = 16

def build_merged_summary_sheet(wb, df_all, styles):
    """シート2: 📋統合集約リスト の作成"""
    ws = wb.create_sheet(title='📋統合集約リスト', index=1)
    ws.views.sheetView[0].showGridLines = True

    # ソート：優先度（優先➔通常）＋ 消化率降順（NaN末尾）
    df_sorted = df_all.copy()
    df_sorted['priority_rank'] = df_sorted['優先度'].apply(lambda x: 0 if x == '優先' else 1)
    df_sorted['st_num'] = pd.to_numeric(df_sorted['消化率(%)'], errors='coerce')
    df_sorted = df_sorted.sort_values(by=['priority_rank', 'st_num'], ascending=[True, False]).drop(columns=['priority_rank', 'st_num'])

    columns = [
        ('選定対象', 10, 'center'),
        ('優先度', 10, 'center'),
        ('商品コード', 16, 'center'),
        ('商品名', 28, 'left'),
        ('カラー', 20, 'left'),
        ('BULK在庫', 12, 'right'),
        ('消化率(%)', 12, 'right'),
        ('出荷店舗', 15, 'center'),
        ('出荷元在庫', 12, 'right'),
        ('出荷前WOS', 12, 'right'),
        ('出荷後WOS', 12, 'right'),
        ('移動推奨数', 12, 'right'),
        ('受入店舗', 15, 'center'),
        ('受入先在庫', 12, 'right'),
        ('受入前WOS', 12, 'right'),
        ('受入後WOS', 12, 'right'),
        ('理由', 50, 'left')
    ]

    # ヘッダー書き込み
    for col_idx, (col_name, col_width, align) in enumerate(columns, start=1):
        cell = ws.cell(row=1, column=col_idx, value=col_name)
        cell.font = styles['font_header']
        cell.fill = styles['fill_header_navy']
        cell.alignment = styles['align_center']
        cell.border = styles['border_all_thin']
        col_letter = get_column_letter(col_idx)
        ws.column_dimensions[col_letter].width = col_width

    # データ行書き込み
    for row_idx, (_, row_data) in enumerate(df_sorted.iterrows(), start=2):
        is_sel = bool(row_data.get('is_selected', False))
        item_name = str(row_data.get('商品名', ''))
        is_cont = '🔄' in item_name
        bulk_val = row_data.get('BULK在庫', 0)
        try:
            bulk_qty = int(bulk_val) if pd.notna(bulk_val) else 0
        except:
            bulk_qty = 0

        # 行の基本背景色
        if is_sel:
            row_fill = styles['fill_yellow_selected']
        elif is_cont:
            row_fill = styles['fill_cont_blue']
        else:
            row_fill = None

        for col_idx, (col_name, _, align) in enumerate(columns, start=1):
            if col_name == '選定対象':
                val = '○' if is_sel else ''
            else:
                val = row_data.get(col_name, '')

            cell = ws.cell(row=row_idx, column=col_idx, value=val)
            cell.font = styles['font_bold'] if is_sel and col_name in ('選定対象', '移動推奨数', '出荷店舗', '受入店舗') else styles['font_main']
            cell.border = styles['border_all_thin']

            # アライメント
            if align == 'center':
                cell.alignment = styles['align_center']
            elif align == 'right':
                cell.alignment = styles['align_right']
            else:
                cell.alignment = styles['align_left']

            # 背景色
            if row_fill:
                cell.fill = row_fill

            # BULK在庫が > 0 の場合の強調
            if col_name == 'BULK在庫' and bulk_qty > 0:
                cell.fill = styles['fill_bulk_orange']

            # 選定対象列の文字色
            if col_name == '選定対象' and is_sel:
                cell.font = Font(name='Meiryo UI', size=11, bold=True, color='C00000')

    ws.freeze_panes = 'A2'

def build_store_shipping_sheets(wb, df_selected, styles):
    """シート3〜11: 出荷_<店舗名> シートの作成"""
    store_cols = [
        ('商品コード', 16, 'center'),
        ('商品名', 28, 'left'),
        ('カラー', 20, 'left'),
        ('出荷指示数', 12, 'center'),
        ('発送先店舗', 15, 'center'),
        ('優先度', 10, 'center'),
        ('自店現在庫', 12, 'right'),
        ('自店出荷前WOS', 13, 'right'),
        ('自店出荷後WOS', 13, 'right'),
        ('受入先在庫', 12, 'right'),
        ('受入前WOS', 12, 'right'),
        ('受入後WOS', 12, 'right'),
        ('消化率(%)', 12, 'right'),
        ('BULK在庫', 12, 'right'),
        ('移動理由', 55, 'left')
    ]

    col_map_from_raw = {
        '商品コード': '商品コード',
        '商品名': '商品名',
        'カラー': 'カラー',
        '出荷指示数': '移動推奨数',
        '発送先店舗': '受入店舗',
        '優先度': '優先度',
        '自店現在庫': '出荷元在庫',
        '自店出荷前WOS': '出荷前WOS',
        '自店出荷後WOS': '出荷後WOS',
        '受入先在庫': '受入先在庫',
        '受入前WOS': '受入前WOS',
        '受入後WOS': '受入後WOS',
        '消化率(%)': '消化率(%)',
        'BULK在庫': 'BULK在庫',
        '移動理由': '理由'
    }

    for store_name in ALL_STORES:
        store_df = df_selected[df_selected['出荷店舗'] == store_name].copy()
        sheet_title = f"出荷_{store_name}"
        ws = wb.create_sheet(title=sheet_title)
        ws.views.sheetView[0].showGridLines = True

        # タイトル情報
        ws['A1'] = f"【出荷指示書】 {store_name} ➔ 各受入店舗"
        ws['A1'].font = styles['font_title']
        tot_qty_val = int(store_df['移動推奨数'].sum()) if not store_df.empty else 0
        ws['A2'] = f"対象出荷件数: {len(store_df)} 件 / 合計出荷点数: {tot_qty_val} 点"
        ws['A2'].font = styles['font_subtitle']

        header_row = 4
        for col_idx, (col_name, col_width, align) in enumerate(store_cols, start=1):
            cell = ws.cell(row=header_row, column=col_idx, value=col_name)
            cell.font = styles['font_header']
            cell.fill = styles['fill_header_blue']
            cell.alignment = styles['align_center']
            cell.border = styles['border_all_thin']
            col_letter = get_column_letter(col_idx)
            ws.column_dimensions[col_letter].width = col_width

        if store_df.empty:
            ws.cell(row=5, column=1, value="※この店舗からの出荷対象アイテムはありません。").font = styles['font_main']
            continue

        store_df['priority_rank'] = store_df['優先度'].apply(lambda x: 0 if x == '優先' else 1)
        store_df = store_df.sort_values(by=['priority_rank', '受入店舗', '商品コード']).drop(columns=['priority_rank'])

        current_row = header_row + 1
        total_ship_qty = 0

        for _, row_data in store_df.iterrows():
            item_name = str(row_data.get('商品名', ''))
            is_cont = '🔄' in item_name
            move_qty = int(row_data.get('移動推奨数', 0)) if pd.notna(row_data.get('移動推奨数')) else 0
            total_ship_qty += move_qty

            bulk_val = row_data.get('BULK在庫', 0)
            try:
                bulk_qty = int(bulk_val) if pd.notna(bulk_val) else 0
            except:
                bulk_qty = 0

            row_fill = styles['fill_cont_blue'] if is_cont else None

            for col_idx, (col_name, _, align) in enumerate(store_cols, start=1):
                raw_field = col_map_from_raw[col_name]
                val = row_data.get(raw_field, '')

                cell = ws.cell(row=current_row, column=col_idx, value=val)
                cell.font = styles['font_main']
                cell.border = styles['border_all_thin']

                if row_fill:
                    cell.fill = row_fill

                if align == 'center':
                    cell.alignment = styles['align_center']
                elif align == 'right':
                    cell.alignment = styles['align_right']
                else:
                    cell.alignment = styles['align_left']

                if col_name == '出荷指示数':
                    cell.font = Font(name='Meiryo UI', size=11, bold=True, color='002060')
                    cell.fill = styles['fill_yellow_selected']
                elif col_name == '発送先店舗':
                    cell.font = styles['font_bold']
                    cell.fill = styles['fill_yellow_selected']
                elif col_name == 'BULK在庫' and bulk_qty > 0:
                    cell.fill = styles['fill_bulk_orange']

            current_row += 1

        # 合計行
        tot_label = ws.cell(row=current_row, column=1, value="合計出荷点数")
        tot_label.font = styles['font_bold']
        tot_label.alignment = styles['align_center']
        tot_label.fill = styles['fill_total_row']
        tot_label.border = styles['border_total_row']

        for c in range(2, 4):
            blank_cell = ws.cell(row=current_row, column=c, value="")
            blank_cell.fill = styles['fill_total_row']
            blank_cell.border = styles['border_total_row']

        tot_val = ws.cell(row=current_row, column=4, value=total_ship_qty)
        tot_val.font = Font(name='Meiryo UI', size=11, bold=True, color='C00000')
        tot_val.alignment = styles['align_center']
        tot_val.fill = styles['fill_yellow_selected']
        tot_val.border = styles['border_total_row']

        for c in range(5, len(store_cols) + 1):
            blank_cell = ws.cell(row=current_row, column=c, value="")
            blank_cell.fill = styles['fill_total_row']
            blank_cell.border = styles['border_total_row']

        ws.freeze_panes = 'A5'

def main():
    if sys.platform == 'win32':
        sys.stdout.reconfigure(encoding='utf-8', errors='replace')

    output_file = "WOS_Report_店舗出荷.xlsx"

    # WOS_Report_元データ.xlsx と WOS_Report.xlsx のうち、最新のものをソースとして使用
    candidates = ["WOS_Report_元データ.xlsx", "WOS_Report.xlsx"]
    valid_candidates = [f for f in candidates if os.path.exists(f)]
    if not valid_candidates:
        print("エラー: 元データファイルが見つかりません。")
        return

    # 最新の更新日時を持つファイルを選択
    source_file = max(valid_candidates, key=os.path.getmtime)

    print("=== WOS 店舗別出荷指示レポート生成処理 開始 ===")
    print(f"データ読み込み元: {source_file}")

    df_all = load_data_from_excel(source_file)
    df_selected = df_all[df_all['is_selected'] == True].copy()

    print(f"全体アイテム数: {len(df_all)} 件")
    print(f"ユーザー選定（黄色ハイライト）アイテム数: {len(df_selected)} 件")
    print(f"  - 優先: {len(df_selected[df_selected['優先度'] == '優先'])} 件")
    print(f"  - 通常: {len(df_selected[df_selected['優先度'] != '優先'])} 件")
    print(f"総出荷推奨数: {df_selected['移動推奨数'].sum()} 点")

    wb = openpyxl.Workbook()
    default_sheet = wb.active

    styles = setup_styles()

    print("生成中: 店舗別出荷サマリー シート")
    build_shipping_summary_sheet(wb, df_selected, styles)

    print("生成中: 統合集約リスト シート")
    build_merged_summary_sheet(wb, df_all, styles)

    print("生成中: 各店舗別出荷シート（全9店舗）")
    build_store_shipping_sheets(wb, df_selected, styles)

    if default_sheet.title in wb.sheetnames:
        wb.remove(default_sheet)

    print(f"Excelファイルを保存しています: {output_file}")
    wb.save(output_file)
    print("=== レポート生成が正常に完了しました ===")

if __name__ == '__main__':
    main()
