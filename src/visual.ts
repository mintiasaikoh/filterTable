"use strict";

import powerbi from "powerbi-visuals-api";
import { FormattingSettingsService } from "powerbi-visuals-utils-formattingmodel";
import {
    BasicFilter, IFilterColumnTarget, IBasicFilter,
    AdvancedFilter, IAdvancedFilter, IAdvancedFilterCondition,
    AdvancedFilterLogicalOperators, AdvancedFilterConditionOperators,
    FilterType,
} from "powerbi-models";
import "./../style/visual.less";

import VisualConstructorOptions = powerbi.extensibility.visual.VisualConstructorOptions;
import VisualUpdateOptions      = powerbi.extensibility.visual.VisualUpdateOptions;
import IVisual                  = powerbi.extensibility.visual.IVisual;
import IVisualHost              = powerbi.extensibility.visual.IVisualHost;
import ISelectionManager        = powerbi.extensibility.ISelectionManager;
import ISelectionId             = powerbi.visuals.ISelectionId;
import DataView                 = powerbi.DataView;
import FilterAction             = powerbi.FilterAction;
import VisualUpdateType         = powerbi.VisualUpdateType;
import VisualDataChangeOperationKind = powerbi.VisualDataChangeOperationKind;
import DataViewTable               = powerbi.DataViewTable;

import { VisualFormattingSettingsModel } from "./settings";
import {
    FilterOp, FilterCondition, PrimitiveValue, FilterValue, ColumnType, TableData,
    normalizeValue, formatDateUTC, toDateEpochFromString, toDateEpoch,
    isConditionActive, filterSignature, buildFilterTarget,
} from "./filterEngine";

const ROW_H  = 24;   // px（tbody 行の高さ）
const BUFFER = 8;    // ビューポート外に余分に描画しておく行数

export class Visual implements IVisual {
    private host: IVisualHost;
    private formattingSettings: VisualFormattingSettingsModel;
    private formattingSettingsService: FormattingSettingsService;

    // ---- データ状態 ----
    private conditions: FilterCondition[]        = [];
    private logic: "AND" | "OR"                  = "AND";
    private tableData: TableData                 = { columns: [], rows: [], rawRows: [], types: [] };
    private appliedConditions: FilterCondition[] = [];
    private appliedLogic: "AND" | "OR"           = "AND";
    private filteredRows: string[][]             = [];
    private filteredOrigIdx: number[]            = [];
    private selectedOrigIdx: Set<number>         = new Set();
    private selectionIds: ISelectionId[]         = [];        // 行ごとの一意な SelectionId
    private selectionManager: ISelectionManager;
    private lastDataView: DataView | null        = null;      // BasicFilter ターゲット生成用
    private lastFilterJson = "";                              // 自己発火 BasicFilter/AdvancedFilter の検出用（エコー除外、プレフィックス付きシグネチャ）
    private lastFilterMode: "ADV" | "BASIC" | null = null;    // 前回発火したフィルタ経路（遷移検知用）
    private advFilterEmitted = false;                         // 中間チャンクで AdvancedFilter を発火済みか
    private activeColTab  = -1;   // -1=全列表示, 0..n-1=指定列のみ表示
    private prevColKey    = "";   // 列構成変化検知用（列名結合文字列）

    // ---- DOM ----
    private filterPanel:  HTMLElement;
    private colToggleBar: HTMLElement;
    private statusBar:    HTMLElement;
    private tableWrapper: HTMLElement;
    private scrollEl:     HTMLElement;
    private table:        HTMLTableElement;
    private colGroup:     HTMLElement;
    private thead:        HTMLTableSectionElement;
    private tbody:        HTMLTableSectionElement;

    // ---- 制御フラグ ----
    private hasInteracted     = false;
    private hasAppliedFilter  = false; // selectionManager.clear() の無駄撃ちを防ぐ
    private isLoadingMore     = false; // fetchMoreData 読み込み中フラグ
    private dataLimitReached  = false; // 100MB メモリ制限到達フラグ
    private persistTimer: number | null = null;
    private scrollRaf:    number | null = null;
    private rootEl:       HTMLElement;
    private rowHeight     = ROW_H;
    private colWidths: Map<number, number> = new Map(); // 列インデックス → px幅
    private sortColIdx = -1;                           // ソート対象列（-1=なし）
    private sortDir: "asc" | "desc" | null = null;     // ソート方向
    private lastClickedRi = -1;                        // Shift+クリック用の起点行
    private calendarOverlay: HTMLElement;               // カレンダーポップアップ（position:fixed 1 個を使い回す）
    private activeCalendarIdx = -1;                     // 現在開いているカレンダーの条件 index（-1=閉）
    private activeCalendarAnchor: HTMLInputElement | null = null; // ARIA 同期用

    constructor(options: VisualConstructorOptions) {
        this.host = options.host;
        this.selectionManager = options.host.createSelectionManager();
        this.formattingSettingsService = new FormattingSettingsService();
        this.rootEl = options.element;
        this.rootEl.className = "filter-table-visual";
        this.buildDOM(this.rootEl);
    }

    private buildDOM(root: HTMLElement): void {
        this.filterPanel  = this.el("div", "filter-panel");
        this.colToggleBar = this.el("div", "col-toggle-bar");
        this.statusBar    = this.el("div", "status-bar");
        this.tableWrapper = this.el("div", "table-wrapper");
        this.scrollEl     = this.el("div", "table-scroll");
        this.table        = this.el("table", "data-table") as HTMLTableElement;
        this.colGroup     = this.el("colgroup", "");
        this.thead        = this.el("thead", "") as HTMLTableSectionElement;
        this.tbody        = this.el("tbody", "") as HTMLTableSectionElement;

        this.table.appendChild(this.colGroup);
        this.table.appendChild(this.thead);
        this.table.appendChild(this.tbody);
        this.scrollEl.appendChild(this.table);
        this.tableWrapper.appendChild(this.scrollEl);
        // カレンダーオーバーレイ（fixed 配置で overflow:hidden を回避）
        this.calendarOverlay = this.el("div", "date-calendar-overlay");
        this.calendarOverlay.style.display = "none";
        this.calendarOverlay.setAttribute("role", "dialog");
        this.calendarOverlay.setAttribute("aria-label", "日付選択カレンダー");

        [this.filterPanel, this.colToggleBar, this.statusBar, this.tableWrapper, this.calendarOverlay]
            .forEach(e => root.appendChild(e));

        // カレンダー外クリックで閉じる
        root.addEventListener("mousedown", (e) => {
            if (this.calendarOverlay.style.display !== "none"
                && !this.calendarOverlay.contains(e.target as Node)) {
                this.closeCalendar();
            }
        });
        // ESC でカレンダーを閉じる（WCAG 対応）。anchor にフォーカスを戻す
        root.addEventListener("keydown", (e) => {
            if (e.key === "Escape" && this.calendarOverlay.style.display !== "none") {
                const anchor = this.activeCalendarAnchor;
                this.closeCalendar();
                anchor?.focus();
                e.stopPropagation();
            }
        });

        this.scrollEl.addEventListener("scroll", () => {
            if (this.scrollRaf !== null) cancelAnimationFrame(this.scrollRaf);
            this.scrollRaf = requestAnimationFrame(() => {
                this.scrollRaf = null;
                this.renderVirtualRows();
            });
        });
    }

    private el<K extends keyof HTMLElementTagNameMap>(tag: K, cls: string): HTMLElementTagNameMap[K] {
        const e = document.createElement(tag);
        if (cls) e.className = cls;
        return e;
    }

    private clear(el: HTMLElement): void {
        while (el.firstChild) el.removeChild(el.firstChild);
    }

    // ==========================================================
    // update — VisualUpdateType で分岐し、不要な再描画を抑制
    // ==========================================================
    public update(options: VisualUpdateOptions): void {
        this.formattingSettings = this.formattingSettingsService
            .populateFormattingSettingsModel(VisualFormattingSettingsModel, options.dataViews?.[0]);

        // リサイズのみの場合はスクロール再描画だけで済む（Data/Style が無ければ早期 return）
        const hasResize = !!(options.type & VisualUpdateType.Resize);
        const hasDataOrStyle = !!(options.type & (VisualUpdateType.Data | VisualUpdateType.Style));
        if (hasResize && !hasDataOrStyle) {
            this.renderVirtualRows();
            return;
        }

        // Style のみ → 書式設定だけ反映して返る
        if (!(options.type & VisualUpdateType.Data)) {
            if (options.type & VisualUpdateType.Style) {
                this.applyTableStyles();
                this.renderTableHeader();
                this.renderVirtualRows();
            }
            return;
        }

        const dv: DataView = options.dataViews?.[0];
        this.lastDataView = dv;

        // --- fetchMoreData: incremental mode ---
        const isSegment = options.operationKind === VisualDataChangeOperationKind.Segment;
        const hasMoreSegments = !!(dv?.metadata?.segment);

        if (isSegment && dv?.table) {
            this.appendIncrementalData(dv.table);
        } else {
            this.tableData = this.extractTableData(dv);
        }

        if (hasMoreSegments) {
            const accepted = this.host.fetchMoreData(false);
            this.isLoadingMore = accepted;
            if (!accepted) this.dataLimitReached = true;
        } else {
            this.isLoadingMore = false;
            this.dataLimitReached = false;
        }

        // 読み込み中の中間チャンク
        if (isSegment && this.isLoadingMore) {
            this.runFilter();
            const hasActiveSearch = this.appliedConditions.some(c => isConditionActive(c, this.tableData.types));
            // AdvancedFilter 経路は全件条件評価なので、途中チャンクで即発火して他ページへ反映。
            // selectedOrigIdx は触らない（中間の不完全な行集合でページ内 SelectionManager を
            // 絞ると瞬間的な表示ブレが起きるため）。
            if (hasActiveSearch && !this.advFilterEmitted && this.canUseAdvancedFilter()) {
                this.emitAdvancedFilterForSync();
                this.advFilterEmitted = true;
            }
            // 中間チャンクは描画を抑制（ステータスのみ更新）。最終チャンクで一括レンダリング。
            this.renderStatus();
            return;
        }
        // 最終チャンク完了後: 蓄積した検索ヒットをフィルターとして適用
        if (isSegment && !this.isLoadingMore) {
            this.runFilter();
            const hasActiveSearch = this.appliedConditions.some(c => isConditionActive(c, this.tableData.types));
            if (hasActiveSearch) {
                // 最終チャンクで全件揃ったので selectedOrigIdx を確定させる
                this.filteredOrigIdx.forEach(i => this.selectedOrigIdx.add(i));
                this.applyDatasetFilter();
            }
            this.advFilterEmitted = false; // 次の検索に備えてリセット
        }

        // --- 列変化検知 ---
        const colKey = this.tableData.columns.join("\0");
        const colsChanged = colKey !== this.prevColKey;
        this.prevColKey = colKey;
        if (colsChanged) {
            this.activeColTab = -1;
            this.colWidths.clear();
            this.sortColIdx = -1;
            this.sortDir = null;
            this.lastClickedRi = -1;
            this.conditions = [];
            this.appliedConditions = [];
            // 列構成が変わったらフィルタ経路フラグをリセット。
            // lastFilterJson が残っていると次回同一 signature の remove が「エコー扱い」で skip される
            this.lastFilterJson = "";
            this.lastFilterMode = null;
            this.advFilterEmitted = false;
            // 既存フィルタを明示的に解除（次の emit が発火できるように）
            if (this.hasAppliedFilter) this.removeFilter();
            this.selectedOrigIdx.clear();
        }

        // --- 状態復元（初回 or 列構成変化時のみ）---
        // 行選択は永続化しないので restoreState は conditions のみ復元する。
        // 真実源は options.jsonFilters に一本化されているので、restoreFromJsonFilters で
        // 行選択（selectedOrigIdx）を毎回再構築する（ChicletSlicer パターン）。
        const isFirstLoad = !this.hasInteracted;
        if (isFirstLoad || colsChanged) {
            this.restoreState(dv);
        }

        // 列構成が変わっていなければ常に jsonFilters から行選択を再構築。
        // 自己発火分は restoreFromJsonFilters 内の signature チェックで除外される。
        if (!colsChanged) {
            this.restoreFromJsonFilters(options.jsonFilters, dv);
        }

        // 範囲外インデックスを除去（行数が減った場合）
        const maxIdx = this.tableData.rows.length;
        this.selectedOrigIdx.forEach(i => { if (i >= maxIdx) this.selectedOrigIdx.delete(i); });
        if (this.selectedOrigIdx.size === 0 && this.hasAppliedFilter) {
            this.removeFilter();
        }

        if (isFirstLoad) {
            this.scrollEl.scrollTop = 0;
        }

        // --- レンダリング ---
        if (!this.filterPanel.querySelector(".value-input:focus")) {
            this.renderFilterPanel();
        }
        this.renderColToggleBar();
        this.runFilter();
        this.applyTableStyles();
        this.renderTableHeader();
        this.renderVirtualRows();
        this.renderStatus();
    }

    private restoreState(dv: DataView): void {
        const m   = dv?.metadata?.objects?.["filterState"];
        const len = this.tableData.columns.length;
        const types = this.tableData.types;
        const textOps = new Set<FilterOp>(["contains", "notContains"]);
        const dateOps = new Set<FilterOp>(["eq", "neq", "lt", "gt", "lte", "gte"]);
        // 範囲内チェック + 列型と演算子の整合 + 列ごと 2 条件までで切り詰め
        const sanitize = (arr: FilterCondition[]): FilterCondition[] => {
            const filtered = arr.filter(c => c.columnIndex >= 0 && c.columnIndex < len);
            const counts = new Map<number, number>();
            const out: FilterCondition[] = [];
            for (const c of filtered) {
                const n = counts.get(c.columnIndex) ?? 0;
                if (n >= 2) continue;
                const colType: ColumnType = types[c.columnIndex] ?? "text";
                const repaired: FilterCondition = { ...c };
                if (colType === "date") {
                    // date 列なのにテキスト演算子 → eq にリセットして値クリア
                    if (!dateOps.has(repaired.operator)) {
                        repaired.operator = "eq";
                        repaired.value = "";
                    }
                    // 日付値フォーマット不正は空に
                    if (repaired.value && !/^\d{4}-\d{2}-\d{2}$/.test(repaired.value)) {
                        repaired.value = "";
                    }
                } else {
                    // text 列なのに date 演算子 → contains にリセットして値クリア
                    if (!textOps.has(repaired.operator)) {
                        repaired.operator = "contains";
                        repaired.value = "";
                    }
                }
                counts.set(c.columnIndex, n + 1);
                out.push(repaired);
            }
            return out;
        };
        try   { this.conditions = sanitize(m?.["conditions"] ? JSON.parse(m["conditions"] as string) : []); }
        catch { this.conditions = []; }
        this.logic = (m?.["logic"] as string) === "OR" ? "OR" : "AND";
        try   { this.appliedConditions = sanitize(m?.["applied"] ? JSON.parse(m["applied"] as string) : []); }
        catch { this.appliedConditions = []; }
        this.appliedLogic = (m?.["appliedLogic"] as string) === "OR" ? "OR" : "AND";

        // 行選択 (selectedOrigIdx) はここで復元しない。
        // 真実源は applyJsonFilter / options.jsonFilters 側に一本化しており、
        // restoreFromJsonFilters() が外部受信経路で自然に再構築する（ChicletSlicer パターン）。
        // 行 index を persistProperties({selector:null}) に保存するとレポート全体共有となり、
        // RLS で可視行集合が異なるユーザー間で意味が壊れるため廃止した。
    }

    /**
     * 表示用セル値の生成。日付列は Date → "YYYY-MM-DD"（ローカル TZ）に統一。
     * これにより JST 環境では JST 基準の年月日がそのまま表示され、
     * ロケール依存の `String(Date)`（"Thu Apr 14 2026 ..."）を回避する。
     */
    private cellToString(v: PrimitiveValue, colType: ColumnType): string {
        if (v == null) return "";
        if (colType === "date") {
            // toDateEpoch と同じ解釈規則で "YYYY-MM-DD" に揃える。
            // 表示と filter 比較を必ず同じ日付に寄せるため、分岐を増やさず
            // epoch 経由で統一（Date / ISO / 非 ISO 文字列 / 数値 epoch ms すべて対応）。
            const ep = toDateEpoch(v);
            if (Number.isFinite(ep)) return formatDateUTC(new Date(ep));
        }
        return String(v);
    }

    private extractTableData(dv: DataView): TableData {
        if (!dv?.table) { this.selectionIds = []; return { columns: [], rows: [], rawRows: [], types: [] }; }
        this.selectionIds = dv.table.rows.map((_, i) =>
            this.host.createSelectionIdBuilder().withTable(dv.table, i).createSelectionId()
        );
        const types: ColumnType[] = dv.table.columns.map(c => c?.type?.dateTime ? "date" : "text");
        return {
            columns: dv.table.columns.map(c => c.displayName || ""),
            rows:    dv.table.rows.map(r => r.map((c, ci) => this.cellToString(c as PrimitiveValue, types[ci] ?? "text"))),
            rawRows: dv.table.rows.map(r => r.map(c => (c == null ? null : c)) as PrimitiveValue[]),
            types,
        };
    }

    private appendIncrementalData(table: DataViewTable): void {
        const offset = (table as unknown as Record<string, unknown>)["lastMergeIndex"] as number | undefined;
        const startIdx = (offset === undefined) ? 0 : offset + 1;

        if (this.tableData.columns.length === 0) {
            this.tableData.columns = table.columns.map(c => c.displayName || "");
            this.tableData.types   = table.columns.map(c => c?.type?.dateTime ? "date" : "text");
        }
        const types = this.tableData.types;

        for (let i = startIdx; i < table.rows.length; i++) {
            this.tableData.rows.push(table.rows[i].map((c, ci) => this.cellToString(c as PrimitiveValue, types[ci] ?? "text")));
            this.tableData.rawRows.push(table.rows[i].map(c => (c == null ? null : c)) as PrimitiveValue[]);
            this.selectionIds.push(
                this.host.createSelectionIdBuilder().withTable(table, i).createSelectionId()
            );
        }
    }

    // ==========================================================
    // フィルターパネル
    // ==========================================================
    private renderFilterPanel(): void {
        // 再描画前にカレンダーとデバウンスタイマーをクリア
        this.closeCalendar();
        if (this.persistTimer !== null) { clearTimeout(this.persistTimer); this.persistTimer = null; }
        this.clear(this.filterPanel);

        const hdr = this.el("div", "filter-header");
        const ttl = this.el("span", "filter-title");
        ttl.textContent = "フィルター";
        hdr.appendChild(ttl);

        const help = this.el("span", "filter-help") as HTMLSpanElement;
        help.textContent = "ⓘ";
        help.title = "1 つの列につき最大 2 条件まで指定できます";
        hdr.appendChild(help);

        const tog = this.el("div", "logic-toggle");
        for (const v of ["AND", "OR"] as const) {
            const b = this.el("button", "logic-btn" + (this.logic === v ? " active" : ""));
            b.textContent = v;
            b.onclick = () => { this.logic = v; this.debounceSave(); this.renderFilterPanel(); };
            tog.appendChild(b);
        }
        hdr.appendChild(tog);
        this.filterPanel.appendChild(hdr);

        const list = this.el("div", "condition-list");
        this.conditions.forEach((c, i) => list.appendChild(this.makeConditionRow(c, i)));
        this.filterPanel.appendChild(list);

        const footer = this.el("div", "filter-footer");

        // 列ごとの条件数カウント → 空きのある最初の列を次の追加先にする
        const countByCol = new Map<number, number>();
        this.conditions.forEach(c =>
            countByCol.set(c.columnIndex, (countByCol.get(c.columnIndex) ?? 0) + 1));
        const availableCol = this.tableData.columns.findIndex(
            (_, i) => (countByCol.get(i) ?? 0) < 2);

        const addBtn = this.el("button", "add-condition-btn") as HTMLButtonElement;
        addBtn.textContent = "+ 条件を追加";
        addBtn.disabled = availableCol < 0;
        addBtn.title = addBtn.disabled
            ? "すべての列が上限（2 条件）に達しました"
            : "新しい条件行を追加";
        addBtn.onclick = () => {
            if (addBtn.disabled) return;
            const defOp: FilterOp = (this.tableData.types[availableCol] === "date") ? "eq" : "contains";
            this.conditions.push({ columnIndex: availableCol, operator: defOp, value: "" });
            this.debounceSave(); this.renderFilterPanel();
        };

        const clearBtn = this.el("button", "clear-btn") as HTMLButtonElement;
        clearBtn.textContent = "解除";
        clearBtn.title = "フィルターを解除して全件表示";
        clearBtn.disabled = this.appliedConditions.length === 0;
        clearBtn.onclick = () => this.clearFilter();

        const runBtn = this.el("button", "run-btn");
        runBtn.textContent = "実行";
        runBtn.onclick = () => this.executeSearch();

        footer.appendChild(addBtn); footer.appendChild(clearBtn); footer.appendChild(runBtn);
        this.filterPanel.appendChild(footer);

        if (availableCol < 0 && this.tableData.columns.length > 0) {
            const note = this.el("div", "filter-note");
            note.textContent = "列ごとの条件は最大 2 個までです";
            this.filterPanel.appendChild(note);
        }
    }

    private makeConditionRow(cond: FilterCondition, idx: number): HTMLElement {
        const row = this.el("div", "condition-row");
        const colType: ColumnType = this.tableData.types[cond.columnIndex] ?? "text";

        const colSel = this.el("select", "col-select");
        this.tableData.columns.forEach((col, i) => {
            const o = this.el("option", "") as HTMLOptionElement;
            o.value = String(i); o.textContent = col;
            if (i === cond.columnIndex) o.selected = true;
            const usedByOthers = this.conditions.filter((c, j) => j !== idx && c.columnIndex === i).length;
            if (usedByOthers >= 2) {
                o.disabled = true;
                o.textContent = col + "（上限）";
                o.title = "この列は既に 2 条件が設定されています";
            }
            colSel.appendChild(o);
        });
        colSel.onchange = () => {
            const newColIdx = +colSel.value;
            const newType: ColumnType = this.tableData.types[newColIdx] ?? "text";
            this.conditions[idx].columnIndex = newColIdx;
            if (colType !== newType) {
                // 列型が変わったら演算子と値をデフォルトにリセット
                this.conditions[idx].operator = (newType === "date") ? "eq" : "contains";
                this.conditions[idx].value = "";
            }
            this.debounceSave(); this.renderFilterPanel();
        };

        const opSel = this.el("select", "op-select");
        const opOptions: { v: FilterOp; l: string }[] = (colType === "date")
            ? [
                { v: "eq",  l: "と同じ" },
                { v: "neq", l: "以外" },
                { v: "gte", l: "以降" },
                { v: "lte", l: "以前" },
                { v: "gt",  l: "より後" },
                { v: "lt",  l: "より前" },
              ]
            : [
                { v: "contains",    l: "を含む" },
                { v: "notContains", l: "を含まない" },
              ];
        for (const { v, l } of opOptions) {
            const o = this.el("option", "") as HTMLOptionElement;
            o.value = v; o.textContent = l;
            if (v === cond.operator) o.selected = true;
            opSel.appendChild(o);
        }
        opSel.onchange = () => {
            this.conditions[idx].operator = opSel.value as FilterOp;
            this.debounceSave();
        };

        let valueEl: HTMLElement;
        if (colType === "date") {
            // カレンダーポップアップを開くクリッカブルな入力欄
            const inp = this.el("input", "value-input date-trigger") as HTMLInputElement;
            inp.type = "text"; inp.readOnly = true;
            inp.placeholder = "📅 日付を選択";
            inp.value = cond.value || "";
            // ARIA: スクリーンリーダーにカレンダー起動可能であることを伝える
            inp.setAttribute("role", "combobox");
            inp.setAttribute("aria-haspopup", "dialog");
            inp.setAttribute("aria-expanded", "false");
            inp.setAttribute("aria-label", "日付を選択");
            // mousedown も止めないと、root の outside-click handler が先に発火して
            // 直後の click での toggle close が「close → open」に化ける（Bug 1）
            inp.onmousedown = (e) => e.stopPropagation();
            inp.onclick = (e) => { e.stopPropagation(); this.showCalendar(inp, cond, idx); };
            inp.onkeydown = (e: KeyboardEvent) => {
                if (e.key === "Enter") { this.executeSearch(); return; }
                // Space / ArrowDown でもカレンダーを開く（キーボード操作）
                if (e.key === " " || e.key === "ArrowDown") {
                    e.preventDefault();
                    this.showCalendar(inp, cond, idx);
                }
            };
            valueEl = inp;
        } else {
            const inp = this.el("input", "value-input") as HTMLInputElement;
            inp.type = "text";
            inp.placeholder = "検索キーワード";
            inp.value = cond.value;
            inp.oninput   = () => { this.conditions[idx].value = inp.value; this.debounceSave(); };
            inp.onkeydown = (e: KeyboardEvent) => { if (e.key === "Enter") this.executeSearch(); };
            valueEl = inp;
        }

        const del = this.el("button", "remove-btn"); del.textContent = "×";
        del.onclick = () => { this.conditions.splice(idx, 1); this.debounceSave(); this.renderFilterPanel(); };

        row.appendChild(colSel); row.appendChild(opSel); row.appendChild(valueEl); row.appendChild(del);
        return row;
    }

    // ==========================================================
    // カレンダーポップアップ
    // position:fixed で filter-panel の overflow:hidden を回避。
    // Power BI iframe 内でも getBoundingClientRect はビューポート基準
    // で返るので fixed 配置と一致する。
    // ==========================================================

    private closeCalendar(): void {
        this.calendarOverlay.style.display = "none";
        this.activeCalendarIdx = -1;
        if (this.activeCalendarAnchor) {
            this.activeCalendarAnchor.setAttribute("aria-expanded", "false");
            this.activeCalendarAnchor = null;
        }
    }

    private showCalendar(anchor: HTMLInputElement, cond: FilterCondition, idx: number): void {
        if (this.activeCalendarIdx === idx) { this.closeCalendar(); return; }
        // 別の input から切替える場合は、前の anchor の aria-expanded を戻す
        if (this.activeCalendarAnchor && this.activeCalendarAnchor !== anchor) {
            this.activeCalendarAnchor.setAttribute("aria-expanded", "false");
        }
        this.activeCalendarIdx = idx;
        this.activeCalendarAnchor = anchor;
        anchor.setAttribute("aria-expanded", "true");

        const rect = anchor.getBoundingClientRect();
        const overlay = this.calendarOverlay;
        // 測定中は非表示にして、旧 top/left 位置での一瞬表示（フリッカー）を防ぐ
        overlay.style.visibility = "hidden";
        overlay.style.display = "block";

        // 描画後にサイズを取得して下か上か決める
        this.renderCalendarContent(cond, idx, anchor);
        const oh = overlay.offsetHeight;
        const spaceBelow = window.innerHeight - rect.bottom;
        overlay.style.left = rect.left + "px";
        overlay.style.top = (spaceBelow >= oh + 4)
            ? (rect.bottom + 2) + "px"   // 下に余裕がある → 下向き
            : (rect.top - oh - 2) + "px"; // 上向き
        overlay.style.visibility = "visible";
    }

    private renderCalendarContent(cond: FilterCondition, idx: number, anchor: HTMLInputElement): void {
        const overlay = this.calendarOverlay;
        this.clear(overlay);

        // 選択済みの値から初期表示月を決定（Date 変換せず直接パース）
        let viewYear: number, viewMonth: number;
        const valMatch = cond.value?.match(/^(\d{4})-(\d{2})-(\d{2})$/);
        if (valMatch) {
            viewYear = +valMatch[1];
            viewMonth = +valMatch[2] - 1;
        } else {
            const now = new Date();
            viewYear = now.getFullYear();
            viewMonth = now.getMonth();
        }

        const render = () => {
            this.clear(overlay);

            // ---- ヘッダー: ◀  2026年4月  ▶ ----
            const header = this.el("div", "cal-header");
            const prev = this.el("button", "cal-nav") as HTMLButtonElement;
            prev.type = "button";
            prev.textContent = "◀";
            prev.onclick = (e) => {
                e.stopPropagation(); viewMonth--;
                if (viewMonth < 0) { viewMonth = 11; viewYear--; }
                render();
            };
            const title = this.el("span", "cal-title");
            title.textContent = `${viewYear}年${viewMonth + 1}月`;
            const next = this.el("button", "cal-nav") as HTMLButtonElement;
            next.type = "button";
            next.textContent = "▶";
            next.onclick = (e) => {
                e.stopPropagation(); viewMonth++;
                if (viewMonth > 11) { viewMonth = 0; viewYear++; }
                render();
            };
            header.appendChild(prev);
            header.appendChild(title);
            header.appendChild(next);
            overlay.appendChild(header);

            // ---- 曜日行 ----
            const dayRow = this.el("div", "cal-weekdays");
            for (const name of ["日", "月", "火", "水", "木", "金", "土"]) {
                const cell = this.el("span", "cal-weekday");
                cell.textContent = name;
                dayRow.appendChild(cell);
            }
            overlay.appendChild(dayRow);

            // ---- 日付グリッド ----
            const grid = this.el("div", "cal-grid");
            const firstDow = new Date(viewYear, viewMonth, 1).getDay();
            const daysInMonth = new Date(viewYear, viewMonth + 1, 0).getDate();
            const pad2 = (n: number) => String(n).padStart(2, "0");

            for (let i = 0; i < firstDow; i++) {
                grid.appendChild(this.el("span", "cal-cell cal-empty"));
            }

            // today はユーザーのローカル TZ（直観と一致させるため）
            const now = new Date();
            const todayStr = `${now.getFullYear()}-${pad2(now.getMonth() + 1)}-${pad2(now.getDate())}`;
            const selectedStr = cond.value;

            for (let d = 1; d <= daysInMonth; d++) {
                // カレンダーは純粋な暦 → TZ 変換不要、直接文字列生成
                const dateStr = `${viewYear}-${pad2(viewMonth + 1)}-${pad2(d)}`;
                const btn = this.el("button", "cal-cell") as HTMLButtonElement;
                btn.type = "button";
                btn.textContent = String(d);
                if (dateStr === selectedStr) btn.classList.add("cal-selected");
                if (dateStr === todayStr) btn.classList.add("cal-today");

                btn.onclick = (e) => {
                    e.stopPropagation();
                    this.conditions[idx].value = dateStr;
                    anchor.value = dateStr;
                    this.closeCalendar();
                    this.debounceSave();
                };
                grid.appendChild(btn);
            }
            overlay.appendChild(grid);
        };
        render();
    }

    // ==========================================================
    // 列トグルバー（タブ動作：排他選択）
    // ==========================================================
    private renderColToggleBar(): void {
        this.clear(this.colToggleBar);
        const multi = this.tableData.columns.length > 1;
        this.colToggleBar.style.display = multi ? "flex" : "none";
        if (!multi) return;

        const allChip = this.el("button", "col-chip" + (this.activeColTab === -1 ? " active" : ""));
        allChip.textContent = "全列";
        allChip.onclick = () => this.switchColTab(-1);
        this.colToggleBar.appendChild(allChip);

        this.tableData.columns.forEach((col, i) => {
            const chip = this.el("button", "col-chip" + (this.activeColTab === i ? " active" : ""));
            chip.textContent = col;
            chip.onclick = () => this.switchColTab(i);
            this.colToggleBar.appendChild(chip);
        });
    }

    private switchColTab(idx: number): void {
        if (this.activeColTab === idx) return;
        this.activeColTab = idx;
        this.renderColToggleBar();
        this.renderTableHeader();
        this.renderVirtualRows();
        // BasicFilter は全列で発火しているので列タブ切替時の再発火は不要
    }

    private isColVisible(i: number): boolean {
        return this.activeColTab === -1 || this.activeColTab === i;
    }

    // ==========================================================
    // 検索
    // ==========================================================
    private executeSearch(): void {
        this.appliedConditions = this.conditions.map(c => ({ ...c }));
        this.appliedLogic = this.logic;
        this.advFilterEmitted = false; // 条件が変わったので中間チャンク発火フラグをリセット
        this.commitFilter();
    }

    private clearFilter(): void {
        this.appliedConditions = []; this.appliedLogic = "AND";
        this.advFilterEmitted = false;
        this.commitFilter();
    }

    private commitFilter(): void {
        this.hasInteracted = true;
        this.selectedOrigIdx.clear();
        this.lastClickedRi = -1;

        this.runFilter();

        const hasActiveSearch = this.appliedConditions.some(c => isConditionActive(c, this.tableData.types));

        // 検索結果がある場合、全結果行を自動選択してクロスフィルター適用
        if (hasActiveSearch && this.filteredRows.length > 0) {
            this.filteredOrigIdx.forEach(i => this.selectedOrigIdx.add(i));
            this.applyDatasetFilter();
        } else {
            // 検索解除時はフィルターも解除
            this.removeFilter();
        }

        this.renderTableHeader();
        this.scrollEl.scrollTop = 0;
        this.renderVirtualRows();
        this.renderFilterPanel();
        this.renderStatus();
        requestAnimationFrame(() => this.persist());
    }

    private runFilter(): void {
        const active = this.appliedConditions.filter(c => isConditionActive(c, this.tableData.types));
        this.filteredRows = []; this.filteredOrigIdx = [];

        if (active.length === 0) {
            this.filteredRows    = this.tableData.rows.slice();
            this.filteredOrigIdx = this.tableData.rows.map((_, i) => i);
            return;
        }

        // 条件ごとに事前計算: text は小文字化キーワード、date は "YYYY-MM-DD" 値の epoch
        interface Prepared {
            cond: FilterCondition;
            kind: "text" | "date";
            keyword: string;       // text 用
            valueEpoch: number;    // date 用
        }
        const prep: Prepared[] = active.map(c => {
            const isDate = this.tableData.types[c.columnIndex] === "date";
            return {
                cond: c,
                kind: isDate ? "date" : "text",
                keyword: isDate ? "" : c.value.toLowerCase(),
                valueEpoch: isDate ? toDateEpochFromString(c.value.trim()) : NaN,
            };
        });

        const isAnd = this.appliedLogic === "AND";

        const evalOne = (p: Prepared, oi: number, row: string[]): boolean => {
            if (p.kind === "text") {
                const hit = (row[p.cond.columnIndex] ?? "").toLowerCase().includes(p.keyword);
                return hit === (p.cond.operator === "contains");
            }
            // date: rawRows から Date を取り出し、時刻を 0:00 に丸めてローカル epoch 比較
            const raw = this.tableData.rawRows[oi]?.[p.cond.columnIndex];
            const rowEp = toDateEpoch(raw);
            if (isNaN(p.valueEpoch)) return false; // 入力が不正ならマッチしない
            if (isNaN(rowEp)) {
                // SQL 三値論理に合わせ、null/非 Date 行は全演算子で不一致扱い
                // （Power BI 純正の日付スライサーも ≠ で null 行を残さない）
                return false;
            }
            switch (p.cond.operator) {
                case "eq":  return rowEp === p.valueEpoch;
                case "neq": return rowEp !== p.valueEpoch;
                case "lt":  return rowEp <  p.valueEpoch;
                case "lte": return rowEp <= p.valueEpoch;
                case "gt":  return rowEp >  p.valueEpoch;
                case "gte": return rowEp >= p.valueEpoch;
                default:    return false;
            }
        };

        this.tableData.rows.forEach((row, oi) => {
            let pass = isAnd;
            for (let k = 0; k < prep.length; k++) {
                const match = evalOne(prep[k], oi, row);
                if (match !== isAnd) { pass = match; break; }
            }
            if (pass) { this.filteredRows.push(row); this.filteredOrigIdx.push(oi); }
        });

        this.applySort();
    }

    private applySort(): void {
        if (this.sortColIdx < 0 || !this.sortDir) return;
        const ci = this.sortColIdx;
        const dir = this.sortDir === "asc" ? 1 : -1;

        // 比較キー生成: Date → getTime, 数値 → Number, それ以外 → 文字列
        const keyOf = (oi: number): number | string => {
            const raw = this.tableData.rawRows[oi]?.[ci];
            if (raw == null) return "";
            if (raw instanceof Date) return raw.getTime();
            if (typeof raw === "number") return raw;
            if (typeof raw === "boolean") return raw ? 1 : 0;
            const s = String(raw);
            const n = Number(s);
            return (s !== "" && !isNaN(n)) ? n : s;
        };

        const indices = this.filteredOrigIdx.map((_, i) => i);
        indices.sort((a, b) => {
            const ka = keyOf(this.filteredOrigIdx[a]);
            const kb = keyOf(this.filteredOrigIdx[b]);
            if (typeof ka === "number" && typeof kb === "number") return (ka - kb) * dir;
            return String(ka).localeCompare(String(kb), undefined, { numeric: true }) * dir;
        });

        this.filteredRows    = indices.map(i => this.filteredRows[i]);
        this.filteredOrigIdx = indices.map(i => this.filteredOrigIdx[i]);
    }

    // ==========================================================
    // テーブル描画（DOM仮想スクロール）
    // ==========================================================
    private renderTableHeader(): void {
        this.clear(this.colGroup);
        this.clear(this.thead);
        if (!this.tableData.columns.length) return;

        const cbCol = this.el("col", ""); cbCol.style.width = "32px";
        this.colGroup.appendChild(cbCol);
        this.tableData.columns.forEach((_, i) => {
            if (!this.isColVisible(i)) return;
            const col = this.el("col", "");
            const w = this.colWidths.get(i);
            if (w) col.style.width = w + "px";
            this.colGroup.appendChild(col);
        });

        const tr = this.el("tr", "");

        const cbTh = this.el("th", "cb-col");
        const allSel = this.filteredOrigIdx.length > 0
            && this.filteredOrigIdx.every(i => this.selectedOrigIdx.has(i));
        const someSel = !allSel && this.filteredOrigIdx.some(i => this.selectedOrigIdx.has(i));
        const allCb = this.el("input", "") as HTMLInputElement;
        allCb.type = "checkbox"; allCb.checked = allSel; allCb.indeterminate = someSel;
        allCb.onchange = () => this.toggleSelectAll();
        cbTh.appendChild(allCb);
        tr.appendChild(cbTh);

        this.tableData.columns.forEach((col, i) => {
            if (!this.isColVisible(i)) return;
            const th = this.el("th", "");

            const label = this.el("span", "col-label");
            label.textContent = col;
            th.appendChild(label);

            // ソートインジケータ
            const arrow = this.el("span", "sort-indicator");
            if (this.sortColIdx === i && this.sortDir === "asc")  arrow.textContent = " ▲";
            else if (this.sortColIdx === i && this.sortDir === "desc") arrow.textContent = " ▼";
            else arrow.textContent = " △"; // 未ソート
            th.appendChild(arrow);

            // ヘッダークリックでソート切替
            th.addEventListener("click", (e) => {
                if ((e.target as HTMLElement).classList.contains("col-resize-handle")) return;
                if (this.sortColIdx === i) {
                    this.sortDir = this.sortDir === "asc" ? "desc" : this.sortDir === "desc" ? null : "asc";
                    if (!this.sortDir) this.sortColIdx = -1;
                } else {
                    this.sortColIdx = i;
                    this.sortDir = "asc";
                }
                this.lastClickedRi = -1; // ソート変更で行順が変わるのでリセット
                this.runFilter();
                this.renderTableHeader();
                this.renderVirtualRows();
            });

            // リサイズハンドル
            const handle = this.el("div", "col-resize-handle");
            handle.addEventListener("mousedown", (e) => this.onColResizeStart(e, i));
            th.appendChild(handle);

            tr.appendChild(th);
        });
        this.thead.appendChild(tr);

        // 列幅が指定されている場合、テーブル幅を合計に設定して横スクロールを有効化
        if (this.colWidths.size > 0) {
            let total = 32; // cb col
            this.tableData.columns.forEach((_, i) => {
                if (!this.isColVisible(i)) return;
                total += this.colWidths.get(i) || 120;
            });
            this.table.style.width = total + "px";
        } else {
            this.table.style.width = "";
        }
    }

    private renderVirtualRows(): void {
        const scrollTop = this.scrollEl.scrollTop;
        const viewH     = this.scrollEl.clientHeight;
        const total     = this.filteredRows.length;

        if (total === 0) {
            this.clear(this.tbody);
            const tr = this.el("tr", ""); const td = this.el("td", "no-data") as HTMLTableCellElement;
            const visCols = this.tableData.columns.filter((_, i) => this.isColVisible(i)).length;
            td.colSpan = visCols + 1;
            td.textContent = this.tableData.columns.length === 0
                ? "データをフィールドに追加してください"
                : "該当するデータがありません";
            tr.appendChild(td); this.tbody.appendChild(tr);
            return;
        }

        const rh = this.rowHeight;
        // wordWrap ON の場合、行の実高さが可変になるためバッファを大幅に拡大
        const isWordWrap = this.formattingSettings?.valuesCard?.wordWrap?.value ?? false;
        const buf = isWordWrap ? BUFFER * 4 : BUFFER;
        const startRow = Math.max(0, Math.floor(scrollTop / rh) - buf);
        const endRow   = Math.min(total, startRow + Math.ceil(viewH / rh) + buf * 2);
        const beforeH  = startRow * rh;
        const afterH   = Math.max(0, (total - endRow) * rh);
        const span     = this.tableData.columns.filter((_, i) => this.isColVisible(i)).length + 1;

        this.clear(this.tbody);
        const frag = document.createDocumentFragment();

        if (beforeH > 0) frag.appendChild(this.makeSpacerRow(beforeH, span));
        for (let ri = startRow; ri < endRow; ri++) frag.appendChild(this.makeDataRow(ri));
        if (afterH  > 0) frag.appendChild(this.makeSpacerRow(afterH,  span));

        this.tbody.appendChild(frag);
    }

    private makeSpacerRow(h: number, span: number): HTMLTableRowElement {
        const tr = this.el("tr", "spacer-row") as HTMLTableRowElement;
        const td = this.el("td", "") as HTMLTableCellElement;
        td.colSpan = span; td.style.height = h + "px";
        td.style.padding = "0"; td.style.border = "none";
        tr.appendChild(td);
        return tr;
    }

    private makeDataRow(ri: number): HTMLTableRowElement {
        const oi  = this.filteredOrigIdx[ri];
        const sel = this.selectedOrigIdx.has(oi);
        const tr  = this.el("tr", ri % 2 === 0 ? "row-even" : "row-odd") as HTMLTableRowElement;
        tr.dataset.ri = String(ri);
        if (sel) tr.classList.add("row-selected");

        const cbTd = this.el("td", "cb-col") as HTMLTableCellElement;
        const cb   = this.el("input", "") as HTMLInputElement;
        cb.type = "checkbox"; cb.checked = sel;
        // checkbox のネイティブ toggle を無効化（tr click ハンドラに任せる）
        cb.addEventListener("click", (ev) => { ev.preventDefault(); });
        cbTd.appendChild(cb); tr.appendChild(cbTd);

        // 行全体のクリックで選択（Ctrl/Shift 対応）
        tr.addEventListener("click", (e) => {
            this.handleRowClick(ri, e);
        });

        const row = this.filteredRows[ri];
        this.tableData.columns.forEach((_, i) => {
            if (!this.isColVisible(i)) return;
            const td = this.el("td", i === 0 ? "first-data-col" : "") as HTMLTableCellElement;
            td.textContent = row[i] ?? "";
            tr.appendChild(td);
        });
        return tr;
    }

    // ==========================================================
    // 選択
    // ==========================================================
    private handleRowClick(ri: number, e: MouseEvent): void {
        if (ri >= this.filteredOrigIdx.length) return;
        const oi = this.filteredOrigIdx[ri];
        const ctrlOrMeta = e.ctrlKey || e.metaKey;

        if (e.shiftKey && this.lastClickedRi >= 0) {
            // Shift+クリック: 範囲選択
            const from = Math.min(this.lastClickedRi, ri);
            const to   = Math.max(this.lastClickedRi, ri);
            if (!ctrlOrMeta) this.selectedOrigIdx.clear();
            for (let r = from; r <= to; r++) {
                this.selectedOrigIdx.add(this.filteredOrigIdx[r]);
            }
        } else if (ctrlOrMeta) {
            // Ctrl/Cmd+クリック: トグル追加/解除
            this.selectedOrigIdx.has(oi) ? this.selectedOrigIdx.delete(oi) : this.selectedOrigIdx.add(oi);
        } else {
            // 通常クリック: 1件だけ選択中で同じ行なら解除、それ以外は単一選択
            const onlyThis = this.selectedOrigIdx.size === 1 && this.selectedOrigIdx.has(oi);
            this.selectedOrigIdx.clear();
            if (!onlyThis) this.selectedOrigIdx.add(oi);
        }

        this.lastClickedRi = ri;
        this.commitSelection();
    }

    private toggleSelectAll(): void {
        if (this.filteredOrigIdx.length === 0) return;
        const allSel = this.filteredOrigIdx.every(i => this.selectedOrigIdx.has(i));
        this.filteredOrigIdx.forEach(i =>
            allSel ? this.selectedOrigIdx.delete(i) : this.selectedOrigIdx.add(i)
        );
        this.commitSelection();
    }

    private clearSelection(): void {
        this.selectedOrigIdx.clear();
        this.commitSelection();
    }

    private commitSelection(): void {
        this.hasInteracted = true;
        this.applyDatasetFilter();
        this.updateSelectionUI();
        this.renderStatus();
        // 行選択は persistProperties には保存しない（真実源は applyJsonFilter / jsonFilters）。
        // ここで persist() は呼ばない。条件変更は commitSelection を経由しない別経路で persist される。
    }

    private applyDatasetFilter(): void {
        if (this.selectedOrigIdx.size === 0) {
            this.removeFilter();
            return;
        }
        // SelectionId ベースで正確な行をクロスフィルター（同一ページ内）
        const ids = Array.from(this.selectedOrigIdx)
            .filter(i => i < this.selectionIds.length)
            .map(i => this.selectionIds[i]);
        if (ids.length === 0) {
            this.removeFilter();
            return;
        }
        this.hasAppliedFilter = true;
        this.selectionManager.select(ids);

        // ページ間同期: 条件ベースなら AdvancedFilter、それ以外は BasicFilter
        if (this.canUseAdvancedFilter()) {
            this.emitAdvancedFilterForSync();
        } else {
            this.emitBasicFilterForSync();
        }
    }

    /**
     * AdvancedFilter で条件を表現可能か判定。
     * - 条件 0 件: false（条件フィルタでない → BasicFilter で行値同期）
     * - 列横断 OR: false（Power BI AdvancedFilter は列間を暗黙 AND で結合するため表現不能）
     * - 同一列 3 条件以上: false（API 仕様で 1 列 2 条件まで。UI 側で制限しているが永続化復元の保険）
     */
    private canUseAdvancedFilter(): boolean {
        const active = this.appliedConditions.filter(c => isConditionActive(c, this.tableData.types));
        if (active.length === 0) return false;

        const byCol = new Map<number, number>();
        for (const c of active) byCol.set(c.columnIndex, (byCol.get(c.columnIndex) ?? 0) + 1);

        for (const count of byCol.values()) {
            if (count > 2) return false;
        }
        if (byCol.size >= 2 && this.appliedLogic === "OR") return false;
        return true;
    }

    private emitAdvancedFilterForSync(): void {
        const dv = this.lastDataView;
        if (!dv?.table?.columns?.length) return;

        const active = this.appliedConditions.filter(c => isConditionActive(c, this.tableData.types));
        if (active.length === 0) return;

        // 列ごとにグループ化
        const byCol = new Map<number, FilterCondition[]>();
        for (const c of active) {
            if (!byCol.has(c.columnIndex)) byCol.set(c.columnIndex, []);
            byCol.get(c.columnIndex).push(c);
        }

        const cols = dv.table.columns;
        const filters: AdvancedFilter[] = [];
        const sigParts: string[] = [];
        // 列間は Power BI が暗黙 AND で結合。列内の logicalOperator は基本 appliedLogic。
        // ただし date 列で eq/neq 単独の場合のみ、半開区間展開に合わせて And/Or を強制上書きする。
        const globalLogical: AdvancedFilterLogicalOperators =
            this.appliedLogic === "OR" ? "Or" : "And";

        const ONE_DAY = 86400000;

        // text 列の演算子マッピング（date は半開区間展開で個別処理するのでここに入れない）
        const opMapText = (op: FilterOp): AdvancedFilterConditionOperators | null => {
            if (op === "contains")    return "Contains";
            if (op === "notContains") return "DoesNotContain";
            return null;
        };

        type CondPair = { op: AdvancedFilterConditionOperators; value: string | number | boolean | Date; sig: string };

        // date 条件を「DateTime 列でも日単位で正しく効く」形に展開する。
        // Is / IsNot / LessThanOrEqual / GreaterThan は時刻成分のせいで境界漏れを起こすため、
        // すべて半開区間 [startOfDay, startOfNextDay) ベースで再マップする。
        // 戻り値が 2 要素のときは forceLogical で列内演算子を強制する必要がある。
        // sig の形式は restoreFromAdvancedFilters 側と一致させる:
        //   `${AdvancedFilterConditionOperators}:${YYYY-MM-DD}`
        // これにより自己発火エコー判定（lastFilterJson 比較）が emit / receive で揃う。
        const expandDateCond = (c: FilterCondition): { pair: CondPair[]; forceLogical: AdvancedFilterLogicalOperators | null } => {
            const epoch = toDateEpochFromString(c.value);
            if (!Number.isFinite(epoch)) return { pair: [], forceLogical: null };
            const start = new Date(epoch);
            const next  = new Date(epoch + ONE_DAY);
            const ymdStart = c.value;
            const ymdNext  = formatDateUTC(next);
            switch (c.operator) {
                case "eq":
                    // [start, next) = gte start AND lt next
                    return {
                        pair: [
                            { op: "GreaterThanOrEqual", value: start, sig: `GreaterThanOrEqual:${ymdStart}` },
                            { op: "LessThan",           value: next,  sig: `LessThan:${ymdNext}` },
                        ],
                        forceLogical: "And",
                    };
                case "neq":
                    // NOT [start, next) = lt start OR gte next
                    return {
                        pair: [
                            { op: "LessThan",           value: start, sig: `LessThan:${ymdStart}` },
                            { op: "GreaterThanOrEqual", value: next,  sig: `GreaterThanOrEqual:${ymdNext}` },
                        ],
                        forceLogical: "Or",
                    };
                case "gte":
                    return { pair: [{ op: "GreaterThanOrEqual", value: start, sig: `GreaterThanOrEqual:${ymdStart}` }], forceLogical: null };
                case "lt":
                    return { pair: [{ op: "LessThan",           value: start, sig: `LessThan:${ymdStart}` }],         forceLogical: null };
                case "gt":
                    // x > YYYY-MM-DD  ⇔  x >= YYYY-MM-(DD+1)
                    return { pair: [{ op: "GreaterThanOrEqual", value: next,  sig: `GreaterThanOrEqual:${ymdNext}` }], forceLogical: null };
                case "lte":
                    // x <= YYYY-MM-DD ⇔  x <  YYYY-MM-(DD+1)
                    return { pair: [{ op: "LessThan",           value: next,  sig: `LessThan:${ymdNext}` }],          forceLogical: null };
                default:
                    return { pair: [], forceLogical: null };
            }
        };

        for (const [ci, conds] of byCol) {
            const col = cols[ci];
            if (!col) continue;
            const target = buildFilterTarget(col);
            if (!target) continue;

            const colType: ColumnType = this.tableData.types[ci] ?? "text";
            const advConds: IAdvancedFilterCondition[] = [];
            const sigItems: string[] = [];
            let colLogical: AdvancedFilterLogicalOperators = globalLogical;

            if (colType === "date") {
                // date 列: 単独条件なら半開区間展開を適用（列内 2 条件制限の範囲で完結）。
                // 複数条件が並んでる場合は展開せずに個別マップ（gte + lte などのユーザー指定範囲は既に正しい）。
                if (conds.length === 1) {
                    const { pair, forceLogical } = expandDateCond(conds[0]);
                    for (const p of pair) {
                        advConds.push({ operator: p.op, value: p.value } as unknown as IAdvancedFilterCondition);
                        sigItems.push(p.sig);
                    }
                    if (forceLogical) colLogical = forceLogical;
                } else {
                    // 2 条件同居: eq/neq は単発でしか意味をなさないので除外、範囲演算子のみ採用
                    for (const c of conds) {
                        if (c.operator === "eq" || c.operator === "neq") continue;
                        const { pair } = expandDateCond(c);
                        for (const p of pair) {
                            advConds.push({ operator: p.op, value: p.value } as unknown as IAdvancedFilterCondition);
                            sigItems.push(p.sig);
                        }
                    }
                }
            } else {
                for (const c of conds) {
                    const op = opMapText(c.operator);
                    if (!op) continue;
                    advConds.push({ operator: op, value: c.value } as unknown as IAdvancedFilterCondition);
                    sigItems.push(`${op}:${c.value}`);
                }
            }
            if (advConds.length === 0) continue;

            filters.push(new AdvancedFilter(target, colLogical, ...advConds));
            const condSig = sigItems.slice().sort().join(",");
            sigParts.push(`${target.table}\0${target.column}\0${colLogical}\0${condSig}`);
        }

        if (filters.length === 0) return;

        const key = "ADV|" + sigParts.slice().sort().join("|");
        if (key === this.lastFilterJson && this.lastFilterMode === "ADV") return;

        // 経路遷移時は前のフィルタを確実に除去してから新経路で merge
        if (this.lastFilterMode === "BASIC") {
            this.host.applyJsonFilter(null, "general", "filter", FilterAction.remove);
        }
        this.lastFilterJson = key;
        this.lastFilterMode = "ADV";
        this.host.applyJsonFilter(filters, "general", "filter", FilterAction.merge);
    }

    private emitBasicFilterForSync(): void {
        const dv = this.lastDataView;
        if (!dv?.table?.columns?.length) return;

        const cols = dv.table.columns;
        const selArr = Array.from(this.selectedOrigIdx);
        if (selArr.length === 0) return;

        // 全列に対して BasicFilter を生成（他ページのスライサー/テーブルの絞り込み精度を最大化）
        const filters: BasicFilter[] = [];
        const sigParts: string[] = [];

        for (let ci = 0; ci < cols.length; ci++) {
            const target = buildFilterTarget(cols[ci]);
            if (!target) continue;

            // 正規化キーで重複排除しつつ、BasicFilter には raw（Date 含む）を渡す
            const valueMap = new Map<string, FilterValue>();
            for (const i of selArr) {
                const raw = this.tableData.rawRows[i]?.[ci];
                if (raw == null || raw === "") continue;
                const key = normalizeValue(raw);
                if (!valueMap.has(key)) valueMap.set(key, raw as FilterValue);
            }
            if (valueMap.size === 0) continue;

            // powerbi-models の型は Date を受け付けないが、実行時は受理される
            const rawValues = Array.from(valueMap.values()) as (string | number | boolean)[];
            filters.push(new BasicFilter(target, "In", ...rawValues));
            sigParts.push(filterSignature(target, Array.from(valueMap.keys())));
        }

        if (filters.length === 0) return;

        // エコー比較のため常にソート済み＋プレフィックスで保存
        const key = "BASIC|" + sigParts.slice().sort().join("|");
        if (key === this.lastFilterJson && this.lastFilterMode === "BASIC") return;

        // 経路遷移時は前のフィルタを確実に除去してから新経路で merge
        if (this.lastFilterMode === "ADV") {
            this.host.applyJsonFilter(null, "general", "filter", FilterAction.remove);
        }
        this.lastFilterJson = key;
        this.lastFilterMode = "BASIC";
        this.host.applyJsonFilter(filters, "general", "filter", FilterAction.merge);
    }

    private removeFilter(): void {
        if (!this.hasAppliedFilter) return;
        this.selectionManager.clear();
        this.lastFilterJson = "";
        this.lastFilterMode = null;
        this.advFilterEmitted = false;
        this.host.applyJsonFilter(null, "general", "filter", FilterAction.remove);
        this.hasAppliedFilter = false;
    }

    /** 外部からの BasicFilter / AdvancedFilter を受信した場合に選択・条件を復元 */
    private restoreFromJsonFilters(jsonFilters: powerbi.IFilter[] | undefined, dv: DataView): boolean {
        if (!jsonFilters || jsonFilters.length === 0) return false;

        // FilterType で分類（powerbi.IFilter は filterType を持たないので powerbi-models の IFilter に経由）
        const advanced: IAdvancedFilter[] = [];
        const basic: IBasicFilter[] = [];
        for (const f of jsonFilters) {
            const ft = (f as unknown as { filterType?: FilterType })?.filterType;
            if (ft === FilterType.Advanced) advanced.push(f as unknown as IAdvancedFilter);
            else if (ft === FilterType.Basic) basic.push(f as unknown as IBasicFilter);
        }

        // AdvancedFilter が含まれる場合はそちらを優先して復元
        if (advanced.length > 0) {
            return this.restoreFromAdvancedFilters(advanced, dv);
        }
        if (basic.length > 0) {
            return this.restoreFromBasicFilters(basic, dv);
        }
        return false;
    }

    private restoreFromBasicFilters(basicFilters: IBasicFilter[], dv: DataView): boolean {
        const cols = dv?.table?.columns || [];

        interface Parsed { colIdx: number; valueSet: Set<string>; sig: string; }
        const parsed: Parsed[] = [];
        for (const bf of basicFilters) {
            const tgt = bf.target as IFilterColumnTarget;
            const values = bf.values;
            if (!tgt || !values || values.length === 0) continue;

            let colIdx = -1;
            for (let i = 0; i < cols.length; i++) {
                const t = buildFilterTarget(cols[i]);
                if (t && t.table === tgt.table && t.column === tgt.column) { colIdx = i; break; }
            }
            if (colIdx < 0) continue;

            const normalized = values.map(v => normalizeValue(v));
            parsed.push({
                colIdx,
                valueSet: new Set(normalized),
                sig: filterSignature(tgt, normalized),
            });
        }
        if (parsed.length === 0) return false;

        const incomingKey = "BASIC|" + parsed.map(p => p.sig).sort().join("|");
        if (incomingKey === this.lastFilterJson) return false;

        // 全ての filter に一致する行を集計（AND）- rawRows + normalizeValue で型差異を吸収
        const matched = new Set<number>();
        this.tableData.rawRows.forEach((row, i) => {
            for (const p of parsed) {
                const raw = row[p.colIdx];
                if (raw == null) return;
                if (!p.valueSet.has(normalizeValue(raw))) return;
            }
            matched.add(i);
        });

        // 一致ゼロ = 現在のウィンドウに該当行が無いだけの可能性が高い。
        // 現セッションの行選択 (selectedOrigIdx) を破壊しないよう early return
        // （lastFilterJson も更新しない）。
        if (matched.size === 0) return false;

        this.selectedOrigIdx = matched;
        this.lastFilterJson = incomingKey;
        this.lastFilterMode = "BASIC";

        // SelectionManager 側も同期（他ビジュアルへのクロスフィルター）
        const ids = Array.from(this.selectedOrigIdx)
            .filter(i => i < this.selectionIds.length)
            .map(i => this.selectionIds[i]);
        if (ids.length > 0) {
            this.hasAppliedFilter = true;
            this.selectionManager.select(ids);
        }
        return true;
    }

    /** AdvancedFilter 受信 → 条件 UI を復元し、ローカル行にもフィルタを適用 */
    private restoreFromAdvancedFilters(advFilters: IAdvancedFilter[], dv: DataView): boolean {
        const cols = dv?.table?.columns || [];

        const opMapIn = (op: string, colType: ColumnType): FilterOp | null => {
            if (colType === "date") {
                switch (op) {
                    case "Is":                  return "eq";
                    case "IsNot":               return "neq";
                    case "LessThan":            return "lt";
                    case "LessThanOrEqual":     return "lte";
                    case "GreaterThan":         return "gt";
                    case "GreaterThanOrEqual":  return "gte";
                    default:                    return null;
                }
            }
            if (op === "Contains")       return "contains";
            if (op === "DoesNotContain") return "notContains";
            return null;
        };
        const toDateStr = (v: unknown): string => {
            if (v instanceof Date) return v.toISOString().slice(0, 10);
            if (typeof v === "string") {
                const m = v.match(/^(\d{4}-\d{2}-\d{2})/);
                if (m) return m[1];
                const d = new Date(v);
                if (!isNaN(d.getTime())) return d.toISOString().slice(0, 10);
            }
            return "";
        };

        interface RestoredCond { op: FilterOp; value: string; sigItem: string; }
        interface Restored { colIdx: number; logic: AdvancedFilterLogicalOperators; conds: RestoredCond[]; sig: string; }
        const restored: Restored[] = [];
        let globalLogic: "AND" | "OR" = "AND";

        for (const af of advFilters) {
            const tgt = af.target as IFilterColumnTarget;
            if (!tgt || !af.conditions || af.conditions.length === 0) continue;

            let colIdx = -1;
            for (let i = 0; i < cols.length; i++) {
                const t = buildFilterTarget(cols[i]);
                if (t && t.table === tgt.table && t.column === tgt.column) { colIdx = i; break; }
            }
            if (colIdx < 0) continue;

            const colType: ColumnType = this.tableData.types[colIdx] ?? "text";
            const condsRaw: RestoredCond[] = [];
            for (const c of af.conditions) {
                const mapped = opMapIn(String(c.operator), colType);
                if (!mapped) continue; // UI 未対応のオペレーターはドロップ
                const valStr = (colType === "date")
                    ? toDateStr(c.value)
                    : String(c.value ?? "");
                if (valStr === "") continue;
                condsRaw.push({
                    op: mapped,
                    value: valStr,
                    sigItem: `${c.operator}:${valStr}`,
                });
            }
            if (condsRaw.length === 0) continue;

            // 1 列 2 条件まで（UI / API 制約）
            const kept = condsRaw.slice(0, 2);
            const logic = (af.logicalOperator || "And") as AdvancedFilterLogicalOperators;
            if (kept.length >= 2) globalLogic = logic === "Or" ? "OR" : "AND";

            const condSig = kept.map(k => k.sigItem).sort().join(",");
            restored.push({
                colIdx, logic, conds: kept,
                sig: `${tgt.table}\0${tgt.column}\0${logic}\0${condSig}`,
            });
        }
        if (restored.length === 0) return false;

        const incomingKey = "ADV|" + restored.map(r => r.sig).sort().join("|");
        if (incomingKey === this.lastFilterJson) return false;

        // FilterCondition 配列を再構築
        const newConds: FilterCondition[] = [];
        for (const r of restored) {
            for (const c of r.conds) {
                newConds.push({
                    columnIndex: r.colIdx,
                    operator: c.op,
                    value: c.value,
                });
            }
        }

        // 編集中の UI 状態（this.conditions / this.logic）は上書きしない。
        // ユーザーが条件入力中に他ビジュアルからの同期が飛んできても入力が消えないように、
        // 「適用済み」側 (appliedConditions / appliedLogic) のみ同期する。
        // ただし UI が空（未タッチ）の場合は UI にも同期してユーザーに見せる。
        const uiIsEmpty = this.conditions.length === 0
            || this.conditions.every(c => !isConditionActive(c, this.tableData.types));
        if (uiIsEmpty) {
            this.conditions = newConds.map(c => ({ ...c }));
            this.logic      = globalLogic;
        }
        this.appliedConditions = newConds;
        this.appliedLogic      = globalLogic;

        // ローカル表示用にもフィルタを適用し、一致行を selection に反映
        this.runFilter();
        const matched = new Set<number>(this.filteredOrigIdx);

        // 一致 0 件は未ロード等の可能性があるので selection は破壊せず、
        // lastFilterJson も更新しない（次回データが揃った時の再マッチを許す）。
        if (matched.size === 0) return true;

        this.lastFilterJson = incomingKey;
        this.lastFilterMode = "ADV";

        this.selectedOrigIdx = matched;
        const ids = Array.from(this.selectedOrigIdx)
            .filter(i => i < this.selectionIds.length)
            .map(i => this.selectionIds[i]);
        if (ids.length > 0) {
            this.hasAppliedFilter = true;
            this.selectionManager.select(ids);
        }
        return true;
    }



    private updateSelectionUI(): void {
        this.tbody.querySelectorAll("tr[data-ri]").forEach((el: Element) => {
            const ri  = parseInt((el as HTMLElement).dataset.ri, 10);
            const oi  = this.filteredOrigIdx[ri];
            const sel = this.selectedOrigIdx.has(oi);
            const cb  = el.querySelector("input") as HTMLInputElement;
            if (cb) cb.checked = sel;
            (el as HTMLElement).classList.toggle("row-selected", sel);
        });
        const allCb = this.thead.querySelector("input") as HTMLInputElement;
        if (allCb) {
            const allSel = this.filteredOrigIdx.length > 0
                && this.filteredOrigIdx.every(i => this.selectedOrigIdx.has(i));
            const someSel = !allSel && this.filteredOrigIdx.some(i => this.selectedOrigIdx.has(i));
            allCb.checked = allSel; allCb.indeterminate = someSel;
        }
    }

    // ==========================================================
    // ステータスバー
    // ==========================================================
    private renderStatus(): void {
        this.clear(this.statusBar);
        const f = this.filteredRows.length, t = this.tableData.rows.length;
        const countText = f === t ? `${t} 件` : `${f} / ${t} 件`;

        if (this.isLoadingMore) {
            this.statusBar.appendChild(document.createTextNode(`${countText}（読み込み中…）`));
        } else if (this.dataLimitReached) {
            this.statusBar.appendChild(document.createTextNode(`${countText}（データ制限到達）`));
        } else {
            this.statusBar.appendChild(document.createTextNode(countText));
        }

        if (this.selectedOrigIdx.size > 0) {
            const info = this.el("span", "sel-info");
            info.textContent = `　${this.selectedOrigIdx.size} 件選択中`;
            this.statusBar.appendChild(info);
            const clr = this.el("button", "clear-sel-btn");
            clr.textContent = "選択解除";
            clr.onclick = () => this.clearSelection();
            this.statusBar.appendChild(clr);
        }
    }


    // ==========================================================
    // 列幅リサイズ
    // ==========================================================
    private onColResizeStart(e: MouseEvent, colIdx: number): void {
        e.preventDefault();
        e.stopPropagation();
        const startX = e.clientX;

        // th 要素から実際の描画幅を取得（<col> は offsetWidth が常に 0）
        const ths = this.thead.querySelectorAll("th");
        let visIdx = 0;
        for (let i = 0; i < this.tableData.columns.length; i++) {
            if (!this.isColVisible(i)) continue;
            if (i === colIdx) break;
            visIdx++;
        }
        const thEl = ths[visIdx + 1] as HTMLElement; // +1 for cb col
        const startW = thEl ? thEl.getBoundingClientRect().width : 80;

        // 対応する col 要素も同時に更新
        const colEls = this.colGroup.querySelectorAll("col");
        const colEl = colEls[visIdx + 1] as HTMLElement;

        const onMove = (ev: MouseEvent) => {
            const newW = Math.max(40, startW + ev.clientX - startX);
            if (colEl) colEl.style.width = newW + "px";
            this.colWidths.set(colIdx, newW);
        };

        const onUp = () => {
            document.removeEventListener("mousemove", onMove);
            document.removeEventListener("mouseup", onUp);
            this.rootEl.classList.remove("col-resizing");
        };

        this.rootEl.classList.add("col-resizing");
        document.addEventListener("mousemove", onMove);
        document.addEventListener("mouseup", onUp);
    }

    // ==========================================================
    // 書式設定の適用
    // ==========================================================
    private applyTableStyles(): void {
        if (!this.formattingSettings) return;
        const v = this.formattingSettings.valuesCard;
        const h = this.formattingSettings.columnHeaderCard;
        const s = this.rootEl.style;

        // 値（セル）のスタイル
        const vSize = v.font.fontSize.value;
        s.setProperty("--val-font-family", v.font.fontFamily.value);
        s.setProperty("--val-font-size", vSize + "pt");
        s.setProperty("--val-font-weight", v.font.bold.value ? "bold" : "normal");
        s.setProperty("--val-font-style", v.font.italic.value ? "italic" : "normal");
        s.setProperty("--val-text-decoration", v.font.underline.value ? "underline" : "none");
        s.setProperty("--val-color", v.fontColor.value.value);
        s.setProperty("--val-bg", v.backgroundColor.value.value);
        s.setProperty("--val-alt-color", v.altFontColor.value.value);
        s.setProperty("--val-alt-bg", v.altBackgroundColor.value.value);
        s.setProperty("--val-white-space", v.wordWrap.value ? "pre-line" : "nowrap");

        // 行高さをフォントサイズに連動（pt→px換算 * 1.6 + padding）
        this.rowHeight = Math.max(ROW_H, Math.round(vSize * 1.333 * 1.6 + 4));
        s.setProperty("--val-row-height", this.rowHeight + "px");

        // 列見出しのスタイル
        s.setProperty("--hdr-font-family", h.font.fontFamily.value);
        s.setProperty("--hdr-font-size", h.font.fontSize.value + "pt");
        s.setProperty("--hdr-font-weight", h.font.bold.value ? "bold" : "normal");
        s.setProperty("--hdr-font-style", h.font.italic.value ? "italic" : "normal");
        s.setProperty("--hdr-text-decoration", h.font.underline.value ? "underline" : "none");
        s.setProperty("--hdr-color", h.fontColor.value.value);
        s.setProperty("--hdr-bg", h.backgroundColor.value.value);
    }

    // ==========================================================
    // Persist
    // ==========================================================
    private debounceSave(): void {
        if (this.persistTimer !== null) clearTimeout(this.persistTimer);
        this.persistTimer = window.setTimeout(() => { this.persistTimer = null; this.persist(); }, 800);
    }

    private persist(): void {
        this.hasInteracted = true;
        // selectionIdx はレポート全体共有で RLS で壊れるため保存しない。
        // 行選択は applyJsonFilter / options.jsonFilters 経由で自然に復元される（真実源一本化）。
        this.host.persistProperties({ merge: [{ objectName: "filterState", selector: null, properties: {
            conditions: JSON.stringify(this.conditions), logic: this.logic,
            applied: JSON.stringify(this.appliedConditions), appliedLogic: this.appliedLogic,
        }}]});
    }

    public getFormattingModel(): powerbi.visuals.FormattingModel {
        if (!this.formattingSettings) {
            this.formattingSettings = new VisualFormattingSettingsModel();
        }
        return this.formattingSettingsService.buildFormattingModel(this.formattingSettings);
    }

    public destroy(): void {
        if (this.persistTimer !== null) { clearTimeout(this.persistTimer); this.persistTimer = null; }
        if (this.scrollRaf !== null) { cancelAnimationFrame(this.scrollRaf); this.scrollRaf = null; }
    }
}
