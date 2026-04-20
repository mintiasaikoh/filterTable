"use strict";

import powerbi from "powerbi-visuals-api";
import { IFilterColumnTarget } from "powerbi-models";

// ==========================================================
// 共有型
// ==========================================================

export type TextOp = "contains" | "notContains";
export type DateOp = "eq" | "neq" | "lt" | "gt" | "lte" | "gte";
export type FilterOp = TextOp | DateOp;

export interface FilterCondition {
    columnIndex: number;
    operator: FilterOp;
    value: string; // 日付条件は "YYYY-MM-DD" 固定、テキストは任意
}

export type PrimitiveValue = string | number | boolean | Date | null;
export type FilterValue = string | number | boolean | Date;

export type ColumnType = "date" | "text";

export interface TableData {
    columns: string[];
    rows: string[][];
    rawRows: PrimitiveValue[][]; // BasicFilter 用に型を保ったまま保持（Date 含む）
    types: ColumnType[];         // 列型（date/text）: UI 分岐と演算子マップに使用
}

// ==========================================================
// 値の正規化（Date / ISO 文字列）
// エコー判定・キー化用。toJSON の差異や TZ による不一致を吸収する
// ==========================================================

export function normalizeValue(v: unknown): string {
    if (v instanceof Date) return v.toISOString();
    if (typeof v === "string" && /^\d{4}-\d{2}-\d{2}T/.test(v)) {
        // Power BI が jsonFilters で ISO 文字列化した Date を受け取るケース
        const d = new Date(v);
        if (!isNaN(d.getTime())) return d.toISOString();
    }
    return String(v);
}

// ==========================================================
// 日付ユーティリティ（すべて UTC 基準で動作）
// Power BI は日時値を UTC epoch の Date オブジェクト（or ISO 文字列）
// で渡す。getUTCFullYear/Month/Date で日付部分を取り出すことで、
// JST 等のローカル TZ による日付ズレを回避する。
// ==========================================================

/** Date → "YYYY-MM-DD"（UTC の年月日）。 */
export function formatDateUTC(d: Date): string {
    const y = d.getUTCFullYear();
    const m = String(d.getUTCMonth() + 1).padStart(2, "0");
    const day = String(d.getUTCDate()).padStart(2, "0");
    return `${y}-${m}-${day}`;
}

/** "YYYY-MM-DD" → UTC 0:00 の epoch。不正値は NaN */
export function toDateEpochFromString(s: string): number {
    if (!/^\d{4}-\d{2}-\d{2}$/.test(s)) return NaN;
    const [y, m, d] = s.split("-").map(Number);
    return Date.UTC(y, m - 1, d);
}

/**
 * 行の値 → UTC 0:00 epoch（時刻部分を切り捨て）。
 * PBI が渡す可能性のある全形態に対応:
 *   - Date オブジェクト
 *   - ISO 先頭一致文字列 "YYYY-MM-DD..."
 *   - 非 ISO 文字列（"2014/04/20" 等）→ Date コンストラクタにフォールバック
 *   - 数値（epoch ms）
 * いずれも UTC の年月日に揃えた epoch を返す。判定不能のみ NaN。
 */
export function toDateEpoch(v: unknown): number {
    if (v == null) return NaN;
    if (v instanceof Date) {
        const t = v.getTime();
        if (!Number.isFinite(t)) return NaN;
        return Date.UTC(v.getUTCFullYear(), v.getUTCMonth(), v.getUTCDate());
    }
    if (typeof v === "string") {
        const m = v.match(/^(\d{4})-(\d{2})-(\d{2})/);
        if (m) return Date.UTC(+m[1], +m[2] - 1, +m[3]);
        // ISO 以外の文字列（ロケール表記など）は Date パーサに委ねる
        const d = new Date(v);
        const t = d.getTime();
        if (Number.isFinite(t)) return Date.UTC(d.getUTCFullYear(), d.getUTCMonth(), d.getUTCDate());
        return NaN;
    }
    if (typeof v === "number" && Number.isFinite(v)) {
        const d = new Date(v);
        return Date.UTC(d.getUTCFullYear(), d.getUTCMonth(), d.getUTCDate());
    }
    return NaN;
}

/** 条件が「検索に寄与するアクティブな条件」か共通判定 */
export function isConditionActive(c: FilterCondition, types: ColumnType[]): boolean {
    const v = c.value.trim();
    if (v === "") return false;
    if (types[c.columnIndex] === "date") {
        return /^\d{4}-\d{2}-\d{2}$/.test(v);
    }
    return true;
}

/** target + 正規化済みキー配列の比較キー（toJSON の差異を回避） */
export function filterSignature(target: IFilterColumnTarget, normalizedKeys: string[]): string {
    const sorted = normalizedKeys.slice().sort();
    return `${target.table}\0${target.column}\0${sorted.join("\0")}`;
}

/**
 * DataViewMetadataColumn から BasicFilter / AdvancedFilter の target を生成。
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
