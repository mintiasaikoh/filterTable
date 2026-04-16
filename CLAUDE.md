# filterTable プロジェクトメモ

Codex will review your output once you are done

## アーキテクチャ: クロスフィルターとスライサー同期

**SelectionManager と applyJsonFilter は役割が違う。両方必要。** `applyJsonFilter` は BasicFilter / AdvancedFilter を使い分ける。

| 仕組み | 範囲 | 用途 |
|---|---|---|
| `ISelectionManager.select(ids)` | 同一ページ内 | 行単位の正確なクロスフィルター |
| `applyJsonFilter(BasicFilter[])` | ページ間・スライサー | 値ベース同期（手動行選択・外部スライサー由来） |
| `applyJsonFilter(AdvancedFilter[])` | ページ間・スライサー | 条件ベース同期（Contains/DoesNotContain + AND/OR） |

`applyDatasetFilter()` で SelectionManager と applyJsonFilter を同時に発火している。詳細は `~/.claude/skills/powerbi-custom-visuals/SKILL.md`。

## BasicFilter vs AdvancedFilter の使い分け

`applyDatasetFilter()` は `canUseAdvancedFilter()` で分岐する。条件入力時は AdvancedFilter、それ以外（手動行選択）は BasicFilter。

| 状況 | 経路 | 理由 |
|---|---|---|
| 条件（contains/notContains）あり、列横断 AND | AdvancedFilter | Power BI エンジンが全件に条件評価 → 全行読込を待たずに他ページへ即時反映できる |
| 条件あり、同一列 ≤2 条件（AND/OR） | AdvancedFilter | 列内の AND/OR は logicalOperator で表現可能 |
| 条件あり、列横断 OR | BasicFilter（fallback） | AdvancedFilter は列間を暗黙 AND で結合するため OR 表現不可 |
| 条件なし、手動行選択 | BasicFilter | 行単位の絞り込みは値ベース列挙が必要 |

### AdvancedFilter 経路のルール

1. **列内は最大 2 条件**（API 仕様）。UI 側で 3 条件目を追加不能にし、`restoreState` でも切り詰める
2. **列間は暗黙 AND**。複数列に filters を並べると Power BI は AND で結合する
3. **初回チャンクで即時発火可能**: 行列挙に依存しないので `host.fetchMoreData` 完了を待たずに `applyJsonFilter` できる。`advFilterEmitted` フラグで多重発火を防ぐ
4. **受信時の条件 UI 復元**: `filterType === FilterType.Advanced` を分岐し、`conditions[]` を `FilterCondition` に写像。UI 未対応のオペレーター（`GreaterThan` 等）は無視
5. **経路遷移（Advanced ↔ Basic）**: 前のフィルタを `FilterAction.remove` で確実に消してから新経路で `merge`。`lastFilterMode` で遷移検知

### エコー判定（共通）

`lastFilterJson` は **プレフィックス付きシグネチャ** で保存する: `ADV|...` / `BASIC|...`。混ざると遷移検知が壊れる。`lastFilterMode` も `"ADV" | "BASIC" | null` で保持。

## BasicFilter 実装の重要ルール

1. **全列で発火**: 単一列だと他ページテーブルで一意性が取れない（先頭列の値が重複すると複数行ヒット）
2. **raw 値（型付き）を渡す**: `tableData.rawRows[i][ci]` を使う。`String(v)` 化すると数値/日付列で型ミスマッチして他ビジュアルで何もヒットしない
3. **incremental モードでは自前蓄積データを参照**: `lastDataView.table.rows` は最新チャンクのみ。`tableData.rawRows` から取得
4. **数値列の自動集計（Sum）を対象に含める**: `isMeasure=true` だけで弾くと Sale Price 等の数値列が BasicFilter から抜け落ち、他ページで絞り込めない。`queryName` が `Sum(Table.Col)` 形式なら中身を剥がして target にする。**column 名は常に queryName の後半を使う**（displayName はユーザーリネームでズレる）
5. **日時列は Date オブジェクト**: `normalizeValue()` で ISO 文字列化して比較・signature 生成。`String(date)` はロケール依存で受信エコーが一致せず無限ループになる。`BasicFilter` には raw Date を渡す（型は `string|number|boolean[]` にキャストが必要）

## jsonFilters 受信

- `options.jsonFilters` から取得（`dv.metadata.jsonFilters` は**存在しない**。型アサーションで誤魔化すと無言で壊れる）
- エコー判定は `filterSignature()` による意味的比較（`JSON.stringify(filter.toJSON())` は不安定）
- `lastFilterJson` は emit / restore の双方で **sorted な signature 配列**で保存する（unsorted で保存→sorted で比較、の混在は余計な再発火を招く）
- 受信時に一致行が 0 件なら **selection を破壊しない**（early return）。ウィンドウ外 / 未ロードで起きうる
- 受信時は `selectionManager.select(ids)` も呼んで SelectionId 側も同期する

## capabilities.json 必須項目

```json
"objects": {
  "general": { "properties": { "filter": { "type": { "filter": true } } } }
},
"supportsSynchronizingFilterState": true
```

## データ保持構造

```typescript
type ColumnType = "date" | "text";
interface TableData {
    columns: string[];
    rows: string[][];              // 描画用（文字列化済み）
    rawRows: PrimitiveValue[][];   // BasicFilter 用（型保持）
    types: ColumnType[];           // 列型: date/text。UI 分岐・演算子マップ用
}
```

`extractTableData` / `appendIncrementalData` の両方で `rawRows` / `types` も蓄積すること。`types` は `dv.table.columns[i].type?.dateTime` で判定。

## テーブル表示ルール

- **1 列目（データ列 index 0）のみ** セル内改行 `\n` を尊重（`white-space: pre-line`）
- **2 列目以降** は常に単一行（`nowrap` + `text-overflow: ellipsis`）
- 理由: 複数列が改行を持つと、1 列目の行高さが引っ張られて表示が崩れる
- `wordWrap` Format 設定は「1 列目の折り返し許可」の意味（OFF で全列 nowrap）

## 条件フィルタ UI 制約

- 1 列につき最大 2 条件（AdvancedFilter API 仕様）
- 「+ 条件を追加」は空きのある先頭列を自動選択。全列 2 条件で埋まると disabled
- 列セレクタで既に 2 条件ある他列は `disabled` option（ラベル末尾に "（上限）"）
- `restoreState` で永続化された 3 条件以上は先頭 2 件に切り詰め

## 日時列のカレンダーフィルター

- `dv.table.columns[i].type?.dateTime === true` の列は `types[i] === "date"` として扱う
- フィルター UI では値入力を `<input type="date">`（カレンダー）にし、演算子は `eq/neq/gte/lte/gt/lt`（「と同じ」「以外」「以降」「以前」「より後」「より前」）
- **日付単位で比較（時刻は無視）**: 行側は `new Date(y, m, d).getTime()`、入力側は `new Date("YYYY-MM-DD" + "T00:00:00").getTime()`。両者 0:00 に揃えて epoch 比較
- **範囲指定**: 同一列に `gte YYYY-MM-DD` + `lte YYYY-MM-DD` の 2 条件で表現（1 列 2 条件制限の範囲内）
- **AdvancedFilter マップ**: `eq→Is` / `neq→IsNot` / `lt→LessThan` / `lte→LessThanOrEqual` / `gt→GreaterThan` / `gte→GreaterThanOrEqual`。値は `Date` オブジェクトで渡す（powerbi-models の型は `string|number|boolean` だが実行時 Date 受理。キャストが必要）
- **エコー signature**: 値は `YYYY-MM-DD` 表記で統一（ISO 時刻込みは不安定）
- **列型変更時のリセット**: 条件行で列を切り替えた際に列型が変わったら、演算子と値をデフォルトにリセット（date→eq, text→contains）
- **受信時の UI 未対応オペレーターはドロップ**（`In`/`NotIn` など）
- `restoreState` の sanitize で、列型と演算子の不整合・日付値フォーマット不正を検出して修復する
- **表示フォーマット**: 日付列セルは `formatLocalDate(Date) → "YYYY-MM-DD"`（ローカル TZ）で統一。`cellToString()` が `extractTableData` / `appendIncrementalData` の両方で適用する。ロケール依存の `String(Date)` は使わない
- **共通ヘルパー**: 条件のアクティブ判定は `isConditionActive(c)`、日付比較は `toDateEpochFromString` / `toLocalDateEpoch` / `formatLocalDate` に集約。重複ロジックを書かない
