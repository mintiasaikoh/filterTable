"use strict";

import powerbi from "powerbi-visuals-api";
import { IFilterColumnTarget } from "powerbi-models";

// ==========================================================
// 共有型
// ==========================================================

export type PrimitiveValue = string | number | boolean | Date | null;
export type FilterValue = string | number | boolean | Date;

export interface TableData {
    columns: string[];
    rows: string[][];
    rawRows: PrimitiveValue[][]; // BasicFilter 用に型を保ったまま保持
}

// ==========================================================
// 値の正規化（BasicFilter の重複排除・エコー signature 用）
// ==========================================================

export function normalizeValue(v: unknown): string {
    if (v instanceof Date) return v.toISOString();
    if (typeof v === "string" && /^\d{4}-\d{2}-\d{2}T/.test(v)) {
        const d = new Date(v);
        if (!isNaN(d.getTime())) return d.toISOString();
    }
    return String(v);
}

/** target + 正規化済みキー配列の比較キー */
export function filterSignature(target: IFilterColumnTarget, normalizedKeys: string[]): string {
    const sorted = normalizedKeys.slice().sort();
    return `${target.table}\0${target.column}\0${sorted.join("\0")}`;
}

/**
 * DataViewMetadataColumn から BasicFilter の target を生成。
 * 集計ラッパー "Sum(Table.Column)" 等は中身を剥がす。DAX メジャーは対象外。
 */
export function buildFilterTarget(col: powerbi.DataViewMetadataColumn): IFilterColumnTarget | null {
    if (!col?.queryName) return null;
    let qn = col.queryName;
    const aggMatch = qn.match(/^\w+\((.+)\)$/);
    const hasAgg = !!aggMatch;
    if (hasAgg) qn = aggMatch[1];
    if (!hasAgg && col.isMeasure) return null;
    const dotIdx = qn.indexOf(".");
    if (dotIdx < 1) return null;
    return { table: qn.substring(0, dotIdx), column: qn.substring(dotIdx + 1) };
}
