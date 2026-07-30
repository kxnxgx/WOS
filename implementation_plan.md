# WOS アイテム集約ツール — 全面リビルド実装計画

## 背景と目的

既存ツール（run.py / src/）は以下の問題を抱えていた：
1. **アイテム集約リスト（出荷/受入れ店舗の可視化）が存在しない**
2. **UIが煩雑**（HTML+CSV操作が重い）

今回は「売上実績 → WOS算出 → 店舗間集約リスト生成 → HTML+Excel出力」の新パイプラインを一から再設計する。

---

## 確定要件サマリー

| 項目 | 決定内容 |
|---|---|
| WOS計算式 | `現在在庫 ÷ 最近4週平均販売数`（予測なし） |
| WOS表示 | 小数点1桁（例：1.3 WOS）。平均販売数=0 の場合は `—`（出荷不可） |
| 出荷判定 | WOS < 全店舗WOS平均 → 「出荷店舗」 |
| 受入れ判定 | WOS ≥ 全店舗WOS平均 → 「受入れ店舗」 |
| 移動推奨数量 | 「WOS平均に近付ける」算出（現在在庫のみ。入荷予定は考慮しない） |
| 地理的制約 | Phaseで後実装（まず店舗定義なしでコアパイプライン） |
| 小口制約 | 同上（store_config.csv追加後に対応） |
| SKU粒度 | 商品コード単位 |
| HTML | 自包含（外部CSS・JS不要）、ダブルクリックで開ける |
| Excel | 別途出力 |

---

## User Review Required

> [!IMPORTANT]
> **フェーズ分割で進めます。**
> - **Phase 1（今回）**: コアパイプライン + HTML/Excelレポート（店舗定義・エリア制約なし）
> - **Phase 2（後回し）**: `store_config.csv` 追加 → エリア内マッチング + 小口制約

> [!WARNING]
> 既存の `src/` 以下のファイルは**置換**します。`run.py` も新実装に差し替えます。
> 旧ファイルが必要な場合は事前に別ディレクトリへバックアップしてください。

---

## Open Questions（確認済み・クローズ）

すべてインタビュー済み。追加の未解決事項はなし。

---

## Proposed Changes

### Phase 1 アーキテクチャ図

```
CSVファイル群
  ├── 営業日付別売上分析.csv   → DataLoader
  └── 在庫一覧.csv            → DataLoader
         ↓
   WOSCalculator              # 最近4週平均、WOS計算
         ↓
   ItemAllocator              # 出荷/受入れ判定 + 移動数算出
         ↓
   Reporter                   # HTML（自包含） + Excel出力
```

---

### パイプラインコアモジュール

#### [DELETE] src/forecaster.py
予測ロジックは廃止（最近4週平均に置換）

#### [DELETE] src/simulator.py
週次シミュレーションは廃止（シンプルな現在在庫÷平均に置換）

#### [MODIFY] src/data_loader.py
- `load_sales_history()` : 既存ロジックを維持（必要列のみ変更可能性あり）
- `load_current_stock()` : 既存ロジックを維持
- `load_incoming_schedule()` : Phase 2 まで削除または無効化

#### [NEW] src/wos_calculator.py
- `WOSCalculator.calculate(sales_df, stock_df) → pd.DataFrame`
- 最近4週間（営業日付の最大日から28日前まで）の販売数を集計
- 店舗×SKU ごとの平均販売数を算出
- 在庫データとJOINしてWOS計算
- 平均=0のとき `wos=None` で管理

#### [NEW] src/item_allocator.py
- `ItemAllocator.allocate(wos_df) → pd.DataFrame`
- SKU単位で出荷候補（WOS < 平均）と受入れ候補（WOS ≥ 平均）を分類
- 移動推奨数量の算出ロジック：
  ```
  # 出荷側：WOS平均まで持ち上げるのに必要な在庫を計算
  # 受入れ側：WOS平均との余剰分を算出
  # 実際の移動数 = min(出荷側不足, 受入れ側余剰)
  
  target_stock_shipper = wos_avg * avg_sales_shipper   # 目標在庫
  surplus_shipper = current_stock - target_stock_shipper
  
  target_stock_receiver = wos_avg * avg_sales_receiver  
  deficit_receiver = target_stock_receiver - current_stock_receiver
  
  move_qty = min(surplus_shipper, deficit_receiver)     # 移動数（整数切り上げ）
  ```
- 理由テキスト自動生成：「WOS X.X週（平均Y.Y週）のためA店舗からB店舗へZ点移動を推奨」

#### [MODIFY] src/reporter.py
- 既存HTMLレポートを全面刷新
- **3タブ構成**：
  1. **アイテム集約リスト** : SKU × 出荷店舗 × 受入れ店舗 × 移動数 × 理由
  2. **店舗別WOSサマリー** : SKU-店舗のWOSヒートマップ（赤=不足、青=余剰）
  3. **全社サマリー** : 全SKUのWOS分布テーブル
- 各セルに「計算根拠ツールチップ」を追加
- 計算式説明ドキュメントをHTMLの `<details>` タグ内に埋め込む
- 外部リソース一切なし（Chart.jsはインライン埋め込みまたはCanvas手書き）
- Excelも同時出力

#### [MODIFY] run.py
- 引数: `--sales`, `--stock` （入荷は Phase 2）
- 処理フロー: DataLoader → WOSCalculator → ItemAllocator → Reporter
- 既存の対話型パス入力UIを維持

---

## ファイル変更一覧

| ファイル | アクション | 概要 |
|---|---|---|
| `src/data_loader.py` | MODIFY | 入荷関連を無効化、既存の読み込みロジック維持 |
| `src/forecaster.py` | DELETE | 廃止 |
| `src/simulator.py` | DELETE | 廃止 |
| `src/wos_calculator.py` | NEW | WOS計算コア |
| `src/item_allocator.py` | NEW | 出荷/受入れ判定・移動数算出 |
| `src/reporter.py` | MODIFY | HTML/Excelレポート全面刷新 |
| `src/main.py` | DELETE | run.pyに統合済みのため不要 |
| `run.py` | MODIFY | 新パイプライン接続 |
| `store_config.csv` | NEW（Phase 2） | 店舗定義・エリア・小口制約（後回し） |

---

## WOS計算ロジック詳細

$$WOS_{store, sku} = \frac{stock_{store, sku}}{avg\_sales\_4w_{store, sku}}$$

$$avg\_sales\_4w = \frac{\sum_{t=T-3}^{T} sales_t}{4}$$

ここで $T$ は売上データの最終週（week_start_date の最大値）。

- **ゼロ除算**: `avg_sales_4w = 0` の場合 `WOS = None`（HTML表示 `—`、Excel空白）
- **平均の定義**: 全店舗WOS平均（Noneを除く）でSKU単位の閾値を決める

---

## 移動推奨数量ロジック詳細

```python
# SKU単位でループ
for sku in all_skus:
    stores = wos_df[wos_df.sku == sku]
    valid = stores[stores.wos.notna()]       # 平均0除外
    wos_avg = valid['wos'].mean()

    shippers  = valid[valid.wos < wos_avg]   # 出荷候補
    receivers = valid[valid.wos >= wos_avg]  # 受入れ候補

    for shipper in shippers:
        surplus = shipper.stock - wos_avg * shipper.avg_sales  # 余剰在庫
        for receiver in receivers:
            deficit = wos_avg * receiver.avg_sales - receiver.stock  # 不足在庫
            move = min(surplus, deficit, surplus)  # 実移動数
            move = math.ceil(move)  # 切り上げ（端数は受入れ側へ）
            # 移動数 > 0 の場合のみリストに追加
```

---

## Verification Plan

### 自動テスト
```bash
python -m pytest tests/ -v
```
- `tests/test_wos_calculator.py` : WOS計算の単体テスト（ゼロ除算、正常系）
- `tests/test_item_allocator.py` : 移動数算出の単体テスト（境界値、マルチSKU）

### 手動検証
1. `python run.py` を実行し、HTMLとExcelが生成されることを確認
2. HTML をダブルクリックしてブラウザで開き、3タブが正常表示されることを確認
3. アイテム集約リストのWOS値をExcelで手計算と照合（3件以上）
4. 移動数のゼロ件・正常件・境界件のレビュー
