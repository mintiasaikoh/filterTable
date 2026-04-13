"use strict";

import powerbi from "powerbi-visuals-api";
import { FormattingSettingsService } from "powerbi-visuals-utils-formattingmodel";
import {
    BasicFilter, AdvancedFilter,
    IFilterColumnTarget, IAdvancedFilterCondition,
    IBasicFilter, IAdvancedFilter, FilterType,
    AdvancedFilterLogicalOperators, AdvancedFilterConditionOperators,
} from "powerbi-models";
import "./../style/visual.less";

import VisualConstructorOptions = powerbi.extensibility.visual.VisualConstructorOptions;
import VisualUpdateOptions      = powerbi.extensibility.visual.VisualUpdateOptions;
import IVisual                  = powerbi.extensibility.visual.IVisual;
import IVisualHost              = powerbi.extensibility.visual.IVisualHost;
import DataView                 = powerbi.DataView;
import FilterAction             = powerbi.FilterAction;
import VisualUpdateType         = powerbi.VisualUpdateType;
import VisualDataChangeOperationKind = powerbi.VisualDataChangeOperationKind;

import { VisualFormattingSettingsModel } from "./settings";

const ROW_H  = 24;   // px（tbody 行の高さ）
const BUFFER = 8;    // ビューポート外に余分に描画しておく行数

interface FilterCondition {
    columnIndex: number;
    operator: "contains" | "notContains";
    value: string;
}

type PrimitiveValue = string | number | boolean | null;

interface TableData {
    columns: string[];
    rows: string[][];
    rawRows: PrimitiveValue[][];
}

export class Visual implements IVisual {
    private host: IVisualHost;
    private formattingSettings: VisualFormattingSettingsModel;
    private formattingSettingsService: FormattingSettingsService;

    // ---- データ状態 ----
    private conditions: FilterCondition[]        = [];
    private logic: "AND" | "OR"                  = "AND";
    private tableData: TableData                 = { columns: [], rows: [], rawRows: [] };
    private appliedConditions: FilterCondition[] = [];
    private appliedLogic: "AND" | "OR"           = "AND";
    private filteredRows: string[][]             = [];
    private filteredOrigIdx: number[]            = [];
    private selectedOrigIdx: Set<number>         = new Set();
    private selectedValues: Set<string>          = new Set(); // フィルター同期用の選択値
    private activeColTab  = -1;   // -1=全列表示, 0..n-1=指定列のみ表示
    private colCount      = 0;    // 列数変化検知用
    private lastDataView: DataView | null        = null;      // フィルター生成用

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
    private hasAppliedFilter  = false; // applyJsonFilter(remove) の無駄撃ちを防ぐ
    private isLoadingMore     = false; // fetchMoreData 読み込み中フラグ
    private loadAllRequested  = false; // ユーザーが全件読み込みを要求したか
    private needsFullData     = false; // 現在の検索が in-memory フォールバックを必要としているか
    private lastFilterJson    = "";    // 自分が適用したフィルターの JSON（自己 update 判定用）
    private persistTimer: number | null = null;
    private scrollRaf:    number | null = null;
    private rootEl:       HTMLElement;
    private rowHeight     = ROW_H;

    constructor(options: VisualConstructorOptions) {
        this.host = options.host;
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
        [this.filterPanel, this.colToggleBar, this.statusBar, this.tableWrapper]
            .forEach(e => root.appendChild(e));

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

        // リサイズのみの場合はスクロール再描画だけで済む
        if (options.type === VisualUpdateType.Resize
            || options.type === (VisualUpdateType.Resize | VisualUpdateType.ResizeEnd)) {
            this.renderVirtualRows();
            return;
        }

        // データ変更を含まない場合は何もしない（Style, ViewMode 等）
        if (!(options.type & VisualUpdateType.Data)) return;

        const dv: DataView = options.dataViews?.[0];
        this.lastDataView = dv;
        this.tableData = this.extractTableData(dv);

        // --- 自分が適用したフィルターの応答か判定 ---
        const currentFilterJson = JSON.stringify(options.jsonFilters ?? []);
        const isSelfFilterUpdate = this.lastFilterJson !== ""
            && currentFilterJson === this.lastFilterJson;
        if (isSelfFilterUpdate) this.lastFilterJson = "";

        // --- fetchMoreData（Append/Segment）---
        const isAppend = options.operationKind === VisualDataChangeOperationKind.Append;
        const hasMoreSegments = !!(dv?.metadata?.segment);
        if (hasMoreSegments && this.loadAllRequested && this.host.fetchMoreData(true)) {
            this.isLoadingMore = true;
        } else {
            this.isLoadingMore = false;
            if (!hasMoreSegments) this.loadAllRequested = false;
        }

        // fetchMoreData 中の中間 update: テーブルだけ更新して返る
        if (isAppend && this.isLoadingMore) {
            this.runFilter();
            this.renderVirtualRows();
            this.renderStatus();
            return;
        }

        // --- 列変化検知 ---
        const colsChanged = this.tableData.columns.length !== this.colCount;
        this.colCount = this.tableData.columns.length;
        if (colsChanged) this.activeColTab = -1;

        // --- 状態復元（初回 or 列構成変化時のみ）---
        const isFirstLoad = !this.hasInteracted;
        if (isFirstLoad || colsChanged) {
            this.restoreState(dv);
            this.restoreFromJsonFilters(options.jsonFilters);
        }

        // --- 選択インデックスの再構築 ---
        if (isSelfFilterUpdate || this.selectedValues.size > 0) {
            this.rebuildSelectionFromValues();
        } else if (!isSelfFilterUpdate) {
            this.selectedOrigIdx.clear();
        }
        // スクロールリセット: 初回ロードまたは外部フィルターでデータが変わった場合のみ
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
        const sanitize = (arr: FilterCondition[]) =>
            arr.filter(c => c.columnIndex >= 0 && c.columnIndex < len);
        try   { this.conditions = sanitize(m?.["conditions"] ? JSON.parse(m["conditions"] as string) : []); }
        catch { this.conditions = []; }
        this.logic = (m?.["logic"] as string) === "OR" ? "OR" : "AND";
        try   { this.appliedConditions = sanitize(m?.["applied"] ? JSON.parse(m["applied"] as string) : []); }
        catch { this.appliedConditions = []; }
        this.appliedLogic = (m?.["appliedLogic"] as string) === "OR" ? "OR" : "AND";

        // 選択値の復元
        try {
            const selStr = m?.["selection"] as string;
            const selArr: string[] = selStr ? JSON.parse(selStr) : [];
            if (selArr.length > 0) {
                this.selectedValues = new Set(selArr);
                this.hasAppliedFilter = true;
            }
        } catch { /* 復元失敗時は空のまま */ }
    }

    // スライサー同期: 他ページから同期されたフィルターを読み取って UI に反映
    private restoreFromJsonFilters(jsonFilters: powerbi.IFilter[] | undefined): void {
        if (!jsonFilters?.length || !this.tableData.columns.length) return;

        for (const f of jsonFilters) {
            const raw = f as unknown as Record<string, unknown>;
            const ft = raw.filterType as number | undefined;

            if (ft === FilterType.Basic) {
                // BasicFilter（選択）の復元
                const bf = raw as unknown as IBasicFilter;
                if (bf.operator === "In" && Array.isArray(bf.values)) {
                    this.selectedValues = new Set(bf.values.map(String));
                    this.hasAppliedFilter = true;
                }
            } else if (ft === FilterType.Advanced) {
                // AdvancedFilter（検索）の復元（操作中でなければ）
                const af = raw as unknown as IAdvancedFilter;
                if (!af.conditions?.length) continue;
                const target = af.target as IFilterColumnTarget | undefined;
                if (!target?.column) continue;

                let colIdx = this.tableData.columns.findIndex(c => c === target.column);
                if (colIdx < 0) {
                    colIdx = this.tableData.columns.findIndex((_, i) =>
                        this.lastDataView?.table?.columns?.[i]?.queryName?.endsWith("." + target.column));
                }
                if (colIdx < 0) continue;

                const mapOp = (op: string): "contains" | "notContains" =>
                    op === "DoesNotContain" ? "notContains" : "contains";

                const restored: FilterCondition[] = af.conditions.map(c => ({
                    columnIndex: colIdx,
                    operator: mapOp(c.operator as string),
                    value: String(c.value ?? ""),
                }));

                this.appliedConditions = restored;
                this.conditions = restored.map(c => ({ ...c }));
                this.appliedLogic = af.logicalOperator === "Or" ? "OR" : "AND";
                this.logic = this.appliedLogic;
                this.hasAppliedFilter = true;
            }
        }
    }

    private extractTableData(dv: DataView): TableData {
        if (!dv?.table) return { columns: [], rows: [], rawRows: [] };
        return {
            columns: dv.table.columns.map(c => c.displayName || ""),
            rows:    dv.table.rows.map(r => r.map(c => (c == null) ? "" : String(c))),
            rawRows: dv.table.rows.map(r => r.map(c => (c == null) ? null : c as PrimitiveValue)),
        };
    }

    // ==========================================================
    // フィルターパネル
    // ==========================================================
    private renderFilterPanel(): void {
        // 再描画前にデバウンスタイマーをキャンセル（削除後の古いクロージャが誤った条件を書き換えるのを防ぐ）
        if (this.persistTimer !== null) { clearTimeout(this.persistTimer); this.persistTimer = null; }
        this.clear(this.filterPanel);

        const hdr = this.el("div", "filter-header");
        const ttl = this.el("span", "filter-title");
        ttl.textContent = "フィルター";
        hdr.appendChild(ttl);

        const tog = this.el("div", "logic-toggle");
        for (const v of ["AND", "OR"] as const) {
            const b = this.el("button", "logic-btn" + (this.logic === v ? " active" : ""));
            b.textContent = v;
            b.onclick = () => { this.logic = v; this.persist(); this.renderFilterPanel(); };
            tog.appendChild(b);
        }
        hdr.appendChild(tog);
        this.filterPanel.appendChild(hdr);

        const list = this.el("div", "condition-list");
        this.conditions.forEach((c, i) => list.appendChild(this.makeConditionRow(c, i)));
        this.filterPanel.appendChild(list);

        const footer = this.el("div", "filter-footer");

        const addBtn = this.el("button", "add-condition-btn");
        addBtn.textContent = "+ 条件を追加";
        addBtn.onclick = () => {
            this.conditions.push({ columnIndex: 0, operator: "contains", value: "" });
            this.persist(); this.renderFilterPanel();
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
    }

    private makeConditionRow(cond: FilterCondition, idx: number): HTMLElement {
        const row = this.el("div", "condition-row");

        const colSel = this.el("select", "col-select");
        this.tableData.columns.forEach((col, i) => {
            const o = this.el("option", ""); o.value = String(i); o.textContent = col;
            if (i === cond.columnIndex) o.selected = true;
            colSel.appendChild(o);
        });
        colSel.onchange = () => { this.conditions[idx].columnIndex = +colSel.value; this.persist(); };

        const opSel = this.el("select", "op-select");
        for (const { v, l } of [{ v: "contains", l: "を含む" }, { v: "notContains", l: "を含まない" }]) {
            const o = this.el("option", ""); o.value = v; o.textContent = l;
            if (v === cond.operator) o.selected = true;
            opSel.appendChild(o);
        }
        opSel.onchange = () => {
            this.conditions[idx].operator = opSel.value as "contains" | "notContains";
            this.persist();
        };

        const inp = this.el("input", "value-input") as HTMLInputElement;
        inp.type = "text"; inp.placeholder = "検索キーワード"; inp.value = cond.value;
        inp.oninput   = () => { this.conditions[idx].value = inp.value; this.debounceSave(); };
        inp.onkeydown = (e: KeyboardEvent) => { if (e.key === "Enter") this.executeSearch(); };

        const del = this.el("button", "remove-btn"); del.textContent = "×";
        del.onclick = () => { this.conditions.splice(idx, 1); this.persist(); this.renderFilterPanel(); };

        row.appendChild(colSel); row.appendChild(opSel); row.appendChild(inp); row.appendChild(del);
        return row;
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
        allChip.onclick = () => {
            this.activeColTab = -1;
            this.renderColToggleBar();
            this.renderTableHeader();
            this.renderVirtualRows();
        };
        this.colToggleBar.appendChild(allChip);

        this.tableData.columns.forEach((col, i) => {
            const chip = this.el("button", "col-chip" + (this.activeColTab === i ? " active" : ""));
            chip.textContent = col;
            chip.onclick = () => {
                this.activeColTab = i;
                this.renderColToggleBar();
                this.renderTableHeader();
                this.renderVirtualRows();
            };
            this.colToggleBar.appendChild(chip);
        });
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
        this.commitFilter();
    }

    private clearFilter(): void {
        this.appliedConditions = []; this.appliedLogic = "AND";
        this.commitFilter();
    }

    private commitFilter(): void {
        this.selectedOrigIdx.clear();
        this.selectedValues.clear();
        this.needsFullData = this.applySearchFilter();
        this.runFilter();

        const hasMoreSegments = !!(this.lastDataView?.metadata?.segment);
        if (this.needsFullData && hasMoreSegments) {
            this.loadAllRequested = true;
            this.host.fetchMoreData(true);
            this.isLoadingMore = true;
        }

        this.renderTableHeader();
        this.scrollEl.scrollTop = 0;
        this.renderVirtualRows();
        this.renderFilterPanel();
        this.renderStatus();
        // persist は applyJsonFilter と同じ同期ブロックで呼ぶとフィルターが消える既知問題があるため遅延実行
        requestAnimationFrame(() => this.persist());
    }

    // PBI クエリエンジンに検索条件を渡す（100k 件超のデータでも全件検索可能にする）
    // 戻り値: in-memory フォールバックが必要（= PBI フィルターで表現しきれない）場合 true
    private applySearchFilter(): boolean {
        const active = this.appliedConditions.filter(c => c.value.trim() !== "");

        if (active.length === 0) {
            this.removeFilter();
            return false;
        }

        const dv = this.lastDataView;
        if (!dv?.table?.columns?.length) return false;

        const byCol = new Map<number, FilterCondition[]>();
        active.forEach(c => {
            if (!byCol.has(c.columnIndex)) byCol.set(c.columnIndex, []);
            byCol.get(c.columnIndex)!.push(c);
        });

        // OR かつ複数列は PBI フィルターで表現不可 → in-memory のみ
        if (byCol.size > 1 && this.appliedLogic === "OR") {
            this.removeFilter();
            return true;
        }

        const mapOp = (op: "contains" | "notContains"): AdvancedFilterConditionOperators =>
            op === "contains" ? "Contains" : "DoesNotContain";

        const filters: (BasicFilter | AdvancedFilter)[] = [];

        byCol.forEach((conds, colIdx) => {
            const target = this.buildFilterTarget(dv.table.columns[colIdx]);
            if (!target) return;

            // AdvancedFilter は 1 回の呼び出しで最大 2 条件しか受け付けない。
            // 3 件以上は PBI 側で先頭 2 件を絞り、in-memory runFilter() が残りを絞る。
            const logic: AdvancedFilterLogicalOperators = this.appliedLogic === "OR" ? "Or" : "And";
            const c0: IAdvancedFilterCondition = { operator: mapOp(conds[0].operator), value: conds[0].value };
            filters.push(
                conds.length === 1
                    ? new AdvancedFilter(target, "And", c0)
                    : new AdvancedFilter(
                        target, logic, c0,
                        { operator: mapOp(conds[1].operator), value: conds[1].value },
                    ),
            );
        });

        if (filters.length === 0) return false;

        this.hasAppliedFilter = true;
        const filterPayload = filters.length === 1 ? filters[0] : filters;
        this.lastFilterJson = JSON.stringify(filters.map(f => f.toJSON()));
        this.host.applyJsonFilter(filterPayload, "general", "filter", FilterAction.merge);

        // 同一列に3条件以上 → PBI は先頭2件のみ、残りは in-memory
        const needsInMemory = Array.from(byCol.values()).some(conds => conds.length > 2);
        return needsInMemory;
    }

    private runFilter(): void {
        const active = this.appliedConditions.filter(c => c.value.trim() !== "");
        this.filteredRows = []; this.filteredOrigIdx = [];

        if (active.length === 0) {
            this.filteredRows    = this.tableData.rows.slice();
            this.filteredOrigIdx = this.tableData.rows.map((_, i) => i);
            return;
        }

        // キーワードを事前に小文字化して行ごとの toLowerCase を省く
        const keywords = active.map(c => c.value.toLowerCase());
        const isAnd    = this.appliedLogic === "AND";

        this.tableData.rows.forEach((row, oi) => {
            // AND: 最初の false で即 fail、OR: 最初の true で即 pass（短絡評価）
            let pass = isAnd;
            for (let k = 0; k < active.length; k++) {
                const c     = active[k];
                const match = (row[c.columnIndex] ?? "").toLowerCase().includes(keywords[k])
                    === (c.operator === "contains");
                if (match !== isAnd) { pass = match; break; }
            }
            if (pass) { this.filteredRows.push(row); this.filteredOrigIdx.push(oi); }
        });
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
            if (this.isColVisible(i)) this.colGroup.appendChild(this.el("col", ""));
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
            const th = this.el("th", ""); th.textContent = col;
            tr.appendChild(th);
        });
        this.thead.appendChild(tr);
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
        const startRow = Math.max(0, Math.floor(scrollTop / rh) - BUFFER);
        const endRow   = Math.min(total, startRow + Math.ceil(viewH / rh) + BUFFER * 2);
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
        cb.onchange = () => this.toggleRowSelection(ri);
        cbTd.appendChild(cb); tr.appendChild(cbTd);

        const row = this.filteredRows[ri];
        this.tableData.columns.forEach((_, i) => {
            if (!this.isColVisible(i)) return;
            const td = this.el("td", "") as HTMLTableCellElement;
            td.textContent = row[i] ?? "";
            tr.appendChild(td);
        });
        return tr;
    }

    // ==========================================================
    // 選択
    // ==========================================================
    private toggleRowSelection(ri: number): void {
        if (ri >= this.filteredOrigIdx.length) return;
        const oi = this.filteredOrigIdx[ri];
        this.selectedOrigIdx.has(oi) ? this.selectedOrigIdx.delete(oi) : this.selectedOrigIdx.add(oi);
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
        this.selectedValues.clear();
        this.commitSelection();
    }

    private commitSelection(): void {
        this.applyDatasetFilter();
        this.updateSelectionUI();
        this.renderStatus();
        // persist は applyJsonFilter と同じ同期ブロックで呼ぶとフィルターが消える既知問題があるため遅延実行
        requestAnimationFrame(() => this.persist());
    }

    private applyDatasetFilter(): void {
        const dv = this.lastDataView;
        if (!dv?.table?.columns?.length) return;

        // フィルター可能な列を決定（指定列 → フォールバック先頭の非メジャー列）
        let colIdx = this.activeColTab >= 0 ? this.activeColTab : 0;
        let col = dv.table.columns[colIdx];
        let target = this.buildFilterTarget(col);
        if (!target) {
            // 指定列がメジャー等でフィルター不可 → 他の列を探す
            for (let i = 0; i < dv.table.columns.length; i++) {
                const t = this.buildFilterTarget(dv.table.columns[i]);
                if (t) { colIdx = i; col = dv.table.columns[i]; target = t; break; }
            }
        }

        this.selectedValues.clear();
        this.selectedOrigIdx.forEach(i => {
            const v = this.tableData.rows[i]?.[colIdx];
            if (v != null && v !== "") this.selectedValues.add(v);
        });

        if (!target) return;

        if (this.selectedValues.size === 0) {
            this.removeFilter();
            return;
        }

        // 元の型付きデータから一意な値を収集（PBI フィルターは型一致が必須）
        const rawValuesSet = new Set<string | number | boolean>();
        this.selectedOrigIdx.forEach(i => {
            const raw = this.tableData.rawRows[i]?.[colIdx];
            if (raw != null) rawValuesSet.add(raw as string | number | boolean);
        });
        const filterValues = Array.from(rawValuesSet);

        const filter = new BasicFilter(target, "In", filterValues);
        this.hasAppliedFilter = true;
        this.lastFilterJson = JSON.stringify([filter.toJSON()]);
        this.host.applyJsonFilter(filter, "general", "filter", FilterAction.merge);
    }

    private buildFilterTarget(col: powerbi.DataViewMetadataColumn): IFilterColumnTarget | null {
        if (!col?.queryName) return null;

        // メジャー列（集計値）はフィルターターゲットに使えない
        if (col.isMeasure) return null;

        let qn = col.queryName;

        // 集計関数でラップされている場合 (e.g. "Sum(Table.Column)") → 中身を取り出す
        const aggMatch = qn.match(/^\w+\((.+)\)$/);
        if (aggMatch) qn = aggMatch[1];

        const dotIdx = qn.indexOf(".");
        if (dotIdx < 1) return null;

        // table: queryName のドット前、column: displayName（公式推奨パターン）
        const target: IFilterColumnTarget = {
            table:  qn.substring(0, dotIdx),
            column: col.displayName,
        };
        return target;
    }

    private removeFilter(): void {
        if (!this.hasAppliedFilter) return;
        this.host.applyJsonFilter(null, "general", "filter", FilterAction.remove);
        this.hasAppliedFilter = false;
    }

    private rebuildSelectionFromValues(): void {
        if (this.selectedValues.size === 0) { this.selectedOrigIdx.clear(); return; }
        const colIdx = this.activeColTab >= 0 ? this.activeColTab : 0;
        this.selectedOrigIdx.clear();
        this.tableData.rows.forEach((row, i) => {
            if (this.selectedValues.has(row[colIdx] ?? "")) this.selectedOrigIdx.add(i);
        });
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
            this.statusBar.appendChild(document.createTextNode(`${countText}（全件読み込み中…）`));
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
    // 書式設定の適用
    // ==========================================================
    private applyTableStyles(): void {
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
        s.setProperty("--val-white-space", v.wordWrap.value ? "normal" : "nowrap");

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
        const selArr = this.selectedValues.size > 0 ? Array.from(this.selectedValues) : [];
        const selCol = this.activeColTab >= 0 ? this.activeColTab : 0;
        this.host.persistProperties({ merge: [{ objectName: "filterState", selector: null, properties: {
            conditions: JSON.stringify(this.conditions), logic: this.logic,
            applied: JSON.stringify(this.appliedConditions), appliedLogic: this.appliedLogic,
            selection: JSON.stringify(selArr), selectionCol: selCol,
        }}]});
    }

    public getFormattingModel(): powerbi.visuals.FormattingModel {
        return this.formattingSettingsService.buildFormattingModel(this.formattingSettings);
    }
}
