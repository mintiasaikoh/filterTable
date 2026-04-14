# filterTable プロジェクトメモ

Codex will review your output once you are done

## アーキテクチャ: クロスフィルターとスライサー同期

**SelectionManager と BasicFilter は役割が違う。両方必要。**

| 仕組み | 範囲 | 用途 |
|---|---|---|
| `ISelectionManager.select(ids)` | 同一ページ内 | 行単位の正確なクロスフィルター |
| `host.applyJsonFilter(BasicFilter[])` | ページ間・スライサー | 値ベース同期 |

`applyDatasetFilter()` で両方を同時に発火している。詳細は `~/.claude/skills/powerbi-custom-visuals/SKILL.md`。

## BasicFilter 実装の重要ルール

1. **全列で発火**: 単一列だと他ページテーブルで一意性が取れない（先頭列の値が重複すると複数行ヒット）
2. **raw 値（型付き）を渡す**: `tableData.rawRows[i][ci]` を使う。`String(v)` 化すると数値/日付列で型ミスマッチして他ビジュアルで何もヒットしない
3. **incremental モードでは自前蓄積データを参照**: `lastDataView.table.rows` は最新チャンクのみ。`tableData.rawRows` から取得
4. **数値列の自動集計（Sum）を対象に含める**: `isMeasure=true` だけで弾くと Sale Price 等の数値列が BasicFilter から抜け落ち、他ページで絞り込めない。`queryName` が `Sum(Table.Col)` 形式なら中身を剥がして target にする。column 名は `displayName`（"Sum of X"）でなく queryName の後半を使う

## jsonFilters 受信

- `options.jsonFilters` から取得（`dv.metadata.jsonFilters` は**存在しない**。型アサーションで誤魔化すと無言で壊れる）
- エコー判定は `filterSignature()` による意味的比較（`JSON.stringify(filter.toJSON())` は不安定）
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

`extractTableData` / `appendIncrementalData` の両方で `rawRows` も蓄積すること。
