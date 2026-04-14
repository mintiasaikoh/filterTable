"use strict";

import powerbi from "powerbi-visuals-api";
import { FormattingSettingsService } from "powerbi-visuals-utils-formattingmodel";
import { BasicFilter, IFilterColumnTarget, IBasicFilter } from "powerbi-models";
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
    rawRows: PrimitiveValue[][]; // BasicFilter 用に型を保ったまま保持
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
    private selectionIds: ISelectionId[]         = [];        // 行ごとの一意な SelectionId
    private selectionManager: ISelectionManager;
    private lastDataView: DataView | null        = null;      // BasicFilter ターゲット生成用
    private lastFilterJson = "";                              // 自己発火 BasicFilter の検出用（エコー除外）
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

        // 読み込み中の中間チャンク: テーブルだけ更新して返る
        if (isSegment && this.isLoadingMore) {
            this.runFilter();
            const hasActiveSearch = this.appliedConditions.some(c => c.value.trim() !== "");
            if (hasActiveSearch) {
                this.filteredOrigIdx.forEach(i => this.selectedOrigIdx.add(i));
            }
            this.renderVirtualRows();
            this.renderStatus();
            return;
        }
        // 最終チャンク完了後: 蓄積した検索ヒットをフィルターとして適用
        if (isSegment && !this.isLoadingMore) {
            this.runFilter();
            const hasActiveSearch = this.appliedConditions.some(c => c.value.trim() !== "");
            if (hasActiveSearch) {
                this.filteredOrigIdx.forEach(i => this.selectedOrigIdx.add(i));
            }
            if (this.selectedOrigIdx.size > 0) this.applyDatasetFilter();
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
        }

        // --- 状態復元（初回 or 列構成変化時のみ）---
        const isFirstLoad = !this.hasInteracted;
        if (isFirstLoad || colsChanged) {
            this.restoreState(dv);
        }

        // 外部スライサーからの BasicFilter を検出 → 選択に反映（自己発火のエコーは除外）
        // 初回ロードでも、永続化された選択がなければ slicer filter を優先
        const hasPersistedSelection = this.selectedOrigIdx.size > 0;
        if (!colsChanged && (!isFirstLoad || !hasPersistedSelection)) {
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
        const sanitize = (arr: FilterCondition[]) =>
            arr.filter(c => c.columnIndex >= 0 && c.columnIndex < len);
        try   { this.conditions = sanitize(m?.["conditions"] ? JSON.parse(m["conditions"] as string) : []); }
        catch { this.conditions = []; }
        this.logic = (m?.["logic"] as string) === "OR" ? "OR" : "AND";
        try   { this.appliedConditions = sanitize(m?.["applied"] ? JSON.parse(m["applied"] as string) : []); }
        catch { this.appliedConditions = []; }
        this.appliedLogic = (m?.["appliedLogic"] as string) === "OR" ? "OR" : "AND";

        // 選択インデックスの復元
        try {
            const idxStr = m?.["selectionIdx"] as string;
            const idxArr: number[] = idxStr ? JSON.parse(idxStr) : [];
            if (idxArr.length > 0) {
                this.selectedOrigIdx = new Set(idxArr.filter(i => i >= 0 && i < this.tableData.rows.length));
                this.hasAppliedFilter = true;
            }
        } catch { /* 復元失敗時は空のまま */ }
    }

    private extractTableData(dv: DataView): TableData {
        if (!dv?.table) { this.selectionIds = []; return { columns: [], rows: [], rawRows: [] }; }
        this.selectionIds = dv.table.rows.map((_, i) =>
            this.host.createSelectionIdBuilder().withTable(dv.table, i).createSelectionId()
        );
        return {
            columns: dv.table.columns.map(c => c.displayName || ""),
            rows:    dv.table.rows.map(r => r.map(c => (c == null) ? "" : String(c))),
            rawRows: dv.table.rows.map(r => r.map(c => (c == null ? null : c)) as PrimitiveValue[]),
        };
    }

    private appendIncrementalData(table: DataViewTable): void {
        const offset = (table as unknown as Record<string, unknown>)["lastMergeIndex"] as number | undefined;
        const startIdx = (offset === undefined) ? 0 : offset + 1;

        if (this.tableData.columns.length === 0) {
            this.tableData.columns = table.columns.map(c => c.displayName || "");
        }

        for (let i = startIdx; i < table.rows.length; i++) {
            this.tableData.rows.push(table.rows[i].map(c => (c == null) ? "" : String(c)));
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
        this.commitFilter();
    }

    private clearFilter(): void {
        this.appliedConditions = []; this.appliedLogic = "AND";
        this.commitFilter();
    }

    private commitFilter(): void {
        this.hasInteracted = true;
        this.selectedOrigIdx.clear();
        this.lastClickedRi = -1;

        this.runFilter();

        const hasActiveSearch = this.appliedConditions.some(c => c.value.trim() !== "");

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
            let pass = isAnd;
            for (let k = 0; k < active.length; k++) {
                const c     = active[k];
                const match = (row[c.columnIndex] ?? "").toLowerCase().includes(keywords[k])
                    === (c.operator === "contains");
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

        // インデックス配列を並べ替え、filteredRows も連動
        const indices = this.filteredOrigIdx.map((oi, i) => i);
        indices.sort((a, b) => {
            const va = this.filteredRows[a][ci] ?? "";
            const vb = this.filteredRows[b][ci] ?? "";
            // 数値として比較可能ならば数値比較
            const na = Number(va), nb = Number(vb);
            if (va !== "" && vb !== "" && !isNaN(na) && !isNaN(nb)) return (na - nb) * dir;
            return va.localeCompare(vb, undefined, { numeric: true }) * dir;
        });

        const newRows = indices.map(i => this.filteredRows[i]);
        const newIdx  = indices.map(i => this.filteredOrigIdx[i]);
        this.filteredRows    = newRows;
        this.filteredOrigIdx = newIdx;
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
            const td = this.el("td", "") as HTMLTableCellElement;
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
        requestAnimationFrame(() => this.persist());
    }

    private applyDatasetFilter(): void {
        if (this.selectedOrigIdx.size === 0) {
            this.removeFilter();
            return;
        }
        // SelectionId ベースで正確な行をクロスフィルター
        const ids = Array.from(this.selectedOrigIdx)
            .filter(i => i < this.selectionIds.length)
            .map(i => this.selectionIds[i]);
        if (ids.length === 0) {
            this.removeFilter();
            return;
        }
        this.hasAppliedFilter = true;
        this.selectionManager.select(ids);

        // スライサー同期用に BasicFilter も発火（値ベース）
        this.emitBasicFilterForSync();
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
            const target = this.buildFilterTarget(cols[ci]);
            if (!target) continue;

            const valueSet = new Set<string | number | boolean>();
            for (const i of selArr) {
                const raw = this.tableData.rawRows[i]?.[ci];
                if (raw != null && raw !== "") valueSet.add(raw as string | number | boolean);
            }
            if (valueSet.size === 0) continue;

            const values = Array.from(valueSet);
            filters.push(new BasicFilter(target, "In", values));
            sigParts.push(this.filterSignature(target, values));
        }

        if (filters.length === 0) return;

        const key = sigParts.join("|");
        if (key === this.lastFilterJson) return;
        this.lastFilterJson = key;
        this.host.applyJsonFilter(filters, "general", "filter", FilterAction.merge);
    }

    /** target + values を正規化した比較キー（toJSON の差異を回避） */
    private filterSignature(target: IFilterColumnTarget, values: (string | number | boolean)[]): string {
        const sorted = values.map(v => String(v)).sort();
        return `${target.table}\0${target.column}\0${sorted.join("\0")}`;
    }

    private buildFilterTarget(col: powerbi.DataViewMetadataColumn): IFilterColumnTarget | null {
        if (!col?.queryName) return null;
        let qn = col.queryName;
        // 集計ラッパー "Sum(Table.Column)" 等を剥がして元列を特定
        const aggMatch = qn.match(/^\w+\((.+)\)$/);
        const hasAgg = !!aggMatch;
        if (hasAgg) qn = aggMatch[1];
        // ラッパー無しかつ isMeasure は DAX メジャーなど元列不明 → 対象外
        if (!hasAgg && col.isMeasure) return null;
        const dotIdx = qn.indexOf(".");
        if (dotIdx < 1) return null;
        // 集計列は displayName が "Sum of X" のようになるので queryName 後半を使う
        const columnName = hasAgg ? qn.substring(dotIdx + 1) : col.displayName;
        return { table: qn.substring(0, dotIdx), column: columnName };
    }

    private removeFilter(): void {
        if (!this.hasAppliedFilter) return;
        this.selectionManager.clear();
        this.lastFilterJson = "";
        this.host.applyJsonFilter(null, "general", "filter", FilterAction.remove);
        this.hasAppliedFilter = false;
    }

    /** 外部スライサーからの BasicFilter を受信した場合に選択を復元 */
    private restoreFromJsonFilters(jsonFilters: powerbi.IFilter[] | undefined, dv: DataView): boolean {
        if (!jsonFilters || jsonFilters.length === 0) return false;

        const cols = dv?.table?.columns || [];

        // 受信 filter を {colIdx, valueSet, sig} に正規化
        interface Parsed { colIdx: number; valueSet: Set<string>; sig: string; }
        const parsed: Parsed[] = [];
        for (const f of jsonFilters) {
            const bf = f as IBasicFilter;
            const tgt = bf.target as IFilterColumnTarget;
            const values = bf.values;
            if (!tgt || !values || values.length === 0) continue;

            let colIdx = -1;
            for (let i = 0; i < cols.length; i++) {
                const t = this.buildFilterTarget(cols[i]);
                if (t && t.table === tgt.table && t.column === tgt.column) { colIdx = i; break; }
            }
            if (colIdx < 0) continue;

            parsed.push({
                colIdx,
                valueSet: new Set(values.map(v => String(v))),
                sig: this.filterSignature(tgt, values as (string | number | boolean)[]),
            });
        }
        if (parsed.length === 0) return false;

        // 自己発火エコー判定（同一 signature 集合）
        const incomingKey = parsed.map(p => p.sig).sort().join("|");
        const selfKey = this.lastFilterJson.split("|").sort().join("|");
        if (incomingKey === selfKey) return false;

        // 全ての filter に一致する行のみ選択（AND）
        this.selectedOrigIdx.clear();
        this.tableData.rows.forEach((row, i) => {
            for (const p of parsed) {
                const v = row[p.colIdx] ?? "";
                if (v === "" || !p.valueSet.has(v)) return;
            }
            this.selectedOrigIdx.add(i);
        });

        this.lastFilterJson = incomingKey;

        // SelectionManager 側も同期（他ビジュアルへのクロスフィルター）
        if (this.selectedOrigIdx.size > 0) {
            const ids = Array.from(this.selectedOrigIdx)
                .filter(i => i < this.selectionIds.length)
                .map(i => this.selectionIds[i]);
            if (ids.length > 0) {
                this.hasAppliedFilter = true;
                this.selectionManager.select(ids);
            }
        } else if (this.hasAppliedFilter) {
            this.selectionManager.clear();
            this.hasAppliedFilter = false;
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
        const selIdx = this.selectedOrigIdx.size > 0 ? Array.from(this.selectedOrigIdx) : [];
        this.host.persistProperties({ merge: [{ objectName: "filterState", selector: null, properties: {
            conditions: JSON.stringify(this.conditions), logic: this.logic,
            applied: JSON.stringify(this.appliedConditions), appliedLogic: this.appliedLogic,
            selectionIdx: JSON.stringify(selIdx),
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
