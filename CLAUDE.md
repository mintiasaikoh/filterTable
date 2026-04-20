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
5. **日付機能は本ビジュアルでは扱わない**: 日付フィルターは別ビジュアル `dateCalendar` に分離した。date 列がバインドされても本ビジュアルは **text 扱い（String 化して contains マッチ）** でフォールバックする

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
interface TableData {
    columns: string[];
    rows: string[][];              // 描画用（文字列化済み）
    rawRows: PrimitiveValue[][];   // BasicFilter 用（型保持）
}
```

`extractTableData` / `appendIncrementalData` の両方で `rawRows` も蓄積すること。日付列は String 化して text 扱い（日付 UI は `dateCalendar` 側）。

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

## 日付フィルターは別ビジュアル

本ビジュアルは **text 系の contains / notContains 条件のみ**を扱う。日付範囲・単一日の絞り込みは別ビジュアル **`dateCalendar`** に分離した（`/Users/mymac/Developer/dateCalendar/`）。

- date 列がバインドされても本ビジュアルはカレンダー UI を出さず、値を `String()` 化した上で text 扱いの contains マッチにフォールバックする
- AdvancedFilter の半開区間 `[start, next)` 展開などの日付特有ロジックは本プロジェクトに残していない
- 日付でクロスフィルターしたい場合は `dateCalendar` ビジュアルを同じページに配置し、同じデータモデル列にバインドする（AdvancedFilter 経由で他ビジュアルへ反映される）
