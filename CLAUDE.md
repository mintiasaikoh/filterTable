# filterTable プロジェクトメモ

Codex will review your output once you are done

## v2.0 破壊的変更

本ビジュアルは **行選択チェックボックステーブル + 検索ボックス** 専用に縮小された。以前の「contains / notContains 条件フィルタ UI」は別ビジュアル **`filterCondition`** (`/Users/mymac/Developer/filterCondition/`) に分離し、日付フィルタは **`dateCalendar`** に分離済み。

- 条件で他ページを絞り込みたい → filterCondition を同ページに配置し、同じ列にバインド
- 日付で他ページを絞り込みたい → dateCalendar を同ページに配置
- 本ビジュアルは BasicFilter 経路のみ。AdvancedFilter は一切発火しない

**v1 からの移行**: persistProperties 永続化スキーマから `conditions` / `logic` / `appliedConditions` / `appliedLogic` / `lastFilterMode` が消えた。既存レポートを開くと条件設定は静かに失われる。filterCondition を追加配置して再入力する必要がある。

## アーキテクチャ: クロスフィルターとスライサー同期

**SelectionManager と applyJsonFilter は役割が違う。両方必要。** 本ビジュアルでは `applyJsonFilter` は BasicFilter 専用。

| 仕組み | 範囲 | 用途 |
|---|---|---|
| `ISelectionManager.select(ids)` | 同一ページ内 | 行単位の正確なクロスフィルター |
| `applyJsonFilter(BasicFilter[])` | ページ間・スライサー | 値ベース同期（チェック行・外部スライサー由来） |

`applyDatasetFilter()` が SelectionManager と applyJsonFilter を同時に発火する。詳細は `~/.claude/skills/powerbi-custom-visuals/SKILL.md`。

### 検索 / 選択の責務分離

- **検索ボックス**: `filteredOrigIdx` に反映するローカル絞り込みのみ。**他ページには一切発火しない**。`applyJsonFilter` は呼ばない
- **チェックボックス**: `selectedOrigIdx` を更新し `applyDatasetFilter()` で BasicFilter を発火。他ページ・スライサーへ同期する

検索ヒットは `selectedOrigIdx` を汚さない。検索で絞った表示の中からユーザーが手でチェックする運用。

## BasicFilter 実装の重要ルール

1. **全列で発火**: 単一列だと他ページテーブルで一意性が取れない（先頭列の値が重複すると複数行ヒット）
2. **raw 値（型付き）を渡す**: `tableData.rawRows[i][ci]` を使う。`String(v)` 化すると数値/日付列で型ミスマッチして他ビジュアルで何もヒットしない
3. **incremental モードでは自前蓄積データを参照**: `lastDataView.table.rows` は最新チャンクのみ。`tableData.rawRows` から取得
4. **数値列の自動集計（Sum）を対象に含める**: `isMeasure=true` だけで弾くと Sale Price 等の数値列が BasicFilter から抜け落ちる。`queryName` が `Sum(Table.Col)` 形式なら中身を剥がして target にする。**column 名は常に queryName の後半を使う**（displayName はユーザーリネームでズレる）
5. **日付列は text 扱い**: 日付フィルターは `dateCalendar` に分離。date 列がバインドされても本ビジュアルは `String()` 化して contains マッチにフォールバックする

## RLS とセキュリティ

本ビジュアルは Row-Level Security 有効環境で安全に動作することが前提。

1. **`persistProperties({ selector: null })` はレポート定義保存 = 全ユーザー共有**。ユーザー単位の永続化 API は Power BI に存在しない
2. **永続化してよいもの**: ユーザーが書いた検索文字列 (`searchText`) のみ
3. **永続化してはいけないもの**:
   - **行 index 配列** (`selectionIdx` 等): RLS / sort / incremental 読込で意味が崩壊
   - **行の値サンプル / SelectionId シリアライズ**: レポートメタデータ経由で機密値が低権限ユーザーに露出しうる
4. **真実源は `applyJsonFilter` / `options.jsonFilters` に一本化**（ChicletSlicer / Timeline パターン）。行選択のクロスセッション復元は `restoreFromJsonFilters()` が jsonFilters から自動で再構築する。`selectedOrigIdx` はセッション内の derived state
5. **`applyJsonFilter` は RLS-safe**: モデル層で RLS と AND 結合される
6. **PBIX 直接配布では RLS が効かない**。Service 公開を前提とする

出典: `learn.microsoft.com/en-us/power-bi/developer/visuals/objects-properties`, `/local-storage`, `/power-bi/guidance/rls-guidance`, `microsoft/PowerBI-visuals-ChicletSlicer/src/webBehavior.ts:216-236`, `microsoft/powerbi-visuals-timeline/src/timeLine.ts:958-977`

## エコー判定

`lastFilterJson` は BasicFilter の意味的 signature (`filterSignature()`) を sorted 配列で join した文字列を保存。`JSON.stringify(filter.toJSON())` は不安定なので使わない。経路が BasicFilter 固定になったのでプレフィックスは不要。

## jsonFilters 受信

- `options.jsonFilters` から取得（`dv.metadata.jsonFilters` は**存在しない**。型アサーションで誤魔化すと無言で壊れる）
- BasicFilter のみを処理。AdvancedFilter（filterCondition / dateCalendar 由来）は値列挙できないので受信時は **selection を破壊しない**（値ベースで一致行を抽出できないため early return）
- 受信時に一致行が 0 件なら selection を破壊しない（ウィンドウ外 / 未ロードで起きうる）
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

## 条件・日付フィルターは別ビジュアル

- **条件 (contains / notContains)** → `filterCondition` (`/Users/mymac/Developer/filterCondition/`)
- **日付範囲・単一日** → `dateCalendar` (`/Users/mymac/Developer/dateCalendar/`)

3 ビジュアルを同じページに並べ、同じデータモデル列にバインドすることで v1 相当の UX を再現できる（Power BI が jsonFilters を AND 結合する）。
