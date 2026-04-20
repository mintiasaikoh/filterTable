# filterTable プロジェクトメモ

Codex will review your output once you are done

## アーキテクチャ: クロスフィルターとスライサー同期

**SelectionManager と applyJsonFilter は役割が違う。両方必要。** `applyJsonFilter` は BasicFilter / AdvancedFilter を使い分ける。

| 仕組み | 範囲 | 用途 |
|---|---|---|
| `ISelectionManager.select(ids)` | 同一ページ内 | 行単位の正確なクロスフィルター |
| `applyJsonFilter(BasicFilter[])` | ページ間・スライサー | 値ベース同期（手動行選択・外部スライサー由来） |
| `applyJsonFilter(AdvancedFilter[])` | ページ間・スライサー | 条件ベース同期（Contains/DoesNotContain + AND/OR） |

`applyDatasetFilter(mode)` で SelectionManager と applyJsonFilter を同時に発火している。詳細は `~/.claude/skills/powerbi-custom-visuals/SKILL.md`。

### mode="search" / "selection" の責務分離

`applyDatasetFilter(mode)` はクロスフィルターのソースを切り替える:

| mode | ソース | 呼び出し元 |
|---|---|---|
| `"search"` | `filteredOrigIdx`（検索ヒット行） | `commitFilter()` / `update()` 最終チャンク |
| `"selection"` | `selectedOrigIdx`（チェックされた行） | `commitSelection()` / `toggleSelectAll()` / `clearSelection()` |

**検索ヒットは `selectedOrigIdx` を汚さない**。検索は「ローカル絞り込み + 条件ベース AdvancedFilter」だけに留め、該当行は検索結果から手でチェックしてもらう。チェックされた行のみが BasicFilter として他ページに送られる（mode="selection" 経路）。`emitBasicFilterForSync(srcIdx)` は引数のインデックス配列から値集合を生成する（`selectedOrigIdx` を直接参照しない）。

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

## RLS とセキュリティ

本ビジュアルは Row-Level Security 有効環境で安全に動作することが前提。以下の原則を死守。

1. **`persistProperties({ selector: null })` はレポート定義保存 = 全ユーザー共有**。ユーザー単位の永続化 API は Power BI に存在しない（`ILocalVisualStorageV2Service` は certified visual 限定で対象外）
2. **永続化してよいもの**: ユーザーが書いたフィルター条件（`conditions` / `logic` / `applied` / `appliedLogic`）のみ
3. **永続化してはいけないもの**:
   - **行 index 配列** (`selectionIdx` 等): RLS / sort / incremental 読込で意味が崩壊
   - **行の値サンプル / SelectionId シリアライズ**: レポートメタデータ経由で機密値が低権限ユーザーに露出しうる
4. **真実源は `applyJsonFilter` / `options.jsonFilters` に一本化**（ChicletSlicer / Timeline パターン）。行選択のクロスセッション復元は `restoreFromJsonFilters()` が jsonFilters から自動で再構築する。`selectedOrigIdx` はセッション内の derived state
5. **`applyJsonFilter` は RLS-safe**: モデル層で RLS と AND 結合される。バイパス経路にはならない（データ漏洩リスクなし）
6. **PBIX 直接配布では RLS が効かない**。Service 公開を前提とする

出典: `learn.microsoft.com/en-us/power-bi/developer/visuals/objects-properties`, `/local-storage`, `/power-bi/guidance/rls-guidance`, `microsoft/PowerBI-visuals-ChicletSlicer/src/webBehavior.ts:216-236`, `microsoft/powerbi-visuals-timeline/src/timeLine.ts:958-977`

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
- フィルター UI では値入力をカスタムカレンダーポップアップにし、演算子は `eq/neq/gte/lte/gt/lt`（「と同じ」「以外」「以降」「以前」「より後」「より前」）
- **すべて UTC 基準**: PBI は日時値を UTC epoch の Date で渡す。`getUTCFullYear/Month/Date` で日付部分を取り出す。ローカル TZ（JST 等）の `getDate()` を使うと日付が 1 日ズレるため禁止
- **日付単位で比較（時刻は無視）**: 行側は `Date.UTC(getUTCFullYear, getUTCMonth, getUTCDate)`、入力側は `Date.UTC(y, m-1, d)`。両者 UTC 0:00 に揃えて epoch 比較
- **文字列値にも対応**: PBI が ISO 文字列（`"2014-01-01T15:00:00.000Z"`）で渡すケースがある。`toDateEpoch` / `cellToString` で先頭 10 文字を正規表現抽出して処理
- **範囲指定**: 同一列に `gte YYYY-MM-DD` + `lte YYYY-MM-DD` の 2 条件で表現（1 列 2 条件制限の範囲内）
- **AdvancedFilter は半開区間 `[start, next)` で発行する**（DateTime 列で時間成分のせいで境界漏れするのを回避）:
  - `eq  YYYY-MM-DD` → `GreaterThanOrEqual YYYY-MM-DD AND LessThan YYYY-MM-(DD+1)` （列内 logical=And 強制）
  - `neq YYYY-MM-DD` → `LessThan YYYY-MM-DD OR GreaterThanOrEqual YYYY-MM-(DD+1)` （列内 logical=Or 強制）
  - `gte` / `lt` は単純マップ（`GreaterThanOrEqual` / `LessThan`）
  - `gt  YYYY-MM-DD` → `GreaterThanOrEqual YYYY-MM-(DD+1)`
  - `lte YYYY-MM-DD` → `LessThan YYYY-MM-(DD+1)`
  - 同一列に 2 条件並ぶ場合（ユーザー指定範囲）は eq/neq を除外し範囲演算子のみ個別マップ（列内 2 条件制限の範囲で完結）
  - 値は UTC midnight の `Date` オブジェクトで渡す。powerbi-models 型は `string|number|boolean` だが実行時 Date 受理（キャスト必要）
  - 単純な `Is`/`IsNot` を使うと DateTime 列で時間込み完全一致を要求し 9 割外すため禁止
  - エコー signature は emit / receive 両側で `${AdvancedFilterConditionOperators}:${YYYY-MM-DD}` 形式に揃える（半開区間展開後の実条件で signature を作る）
- **エコー signature**: 値は `YYYY-MM-DD` 表記で統一（ISO 時刻込みは不安定）
- **列型変更時のリセット**: 条件行で列を切り替えた際に列型が変わったら、演算子と値をデフォルトにリセット（date→eq, text→contains）
- **受信時の UI 未対応オペレーターはドロップ**（`In`/`NotIn` など）
- `restoreState` の sanitize で、列型と演算子の不整合・日付値フォーマット不正を検出して修復する
- **表示フォーマット**: 日付列セルは `formatDateUTC(Date) → "YYYY-MM-DD"`（UTC）で統一。`cellToString()` が `extractTableData` / `appendIncrementalData` の両方で適用。ISO 文字列の場合は先頭 10 文字を抽出
- **カレンダー UI**: `position: fixed` でカスタム DOM カレンダーポップアップを表示（PBI iframe 内で `<input type="date">` のネイティブピッカーが出ないため）。"today" ハイライトはローカル TZ、日付選択値は `YYYY-MM-DD` 文字列
- **共通ヘルパー**: 条件のアクティブ判定は `isConditionActive(c)`、日付比較は `toDateEpochFromString` / `toDateEpoch` / `formatDateUTC` に集約。重複ロジックを書かない
