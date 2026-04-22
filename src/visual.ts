"use strict";

import powerbi from "powerbi-visuals-api";
import { FormattingSettingsService } from "powerbi-visuals-utils-formattingmodel";
import {
    BasicFilter, IFilterColumnTarget, IBasicFilter,
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
    PrimitiveValue, FilterValue, TableData,
    normalizeValue, filterSignature, buildFilterTarget,
} from "./filterEngine";

const ROW_H  = 24;   // px（tbody 行の高さ）
const BUFFER = 8;    // ビューポート外に余分に描画しておく行数

export class Visual implements IVisual {
    private host: IVisualHost;
    private formattingSettings: VisualFormattingSettingsModel;
    private formattingSettingsService: FormattingSettingsService;

    // ---- データ状態 ----
    private searchText: string                   = "";
    private appliedSearchText: string            = "";
    private tableData: TableData                 = { columns: [], rows: [], rawRows: [] };
    private filteredRows: string[][]             = [];
    private filteredOrigIdx: number[]            = [];
    private selectedOrigIdx: Set<number>         = new Set();
    private selectionIds: ISelectionId[]         = [];
    private selectionManager: ISelectionManager;
    private lastDataView: DataView | null        = null;
    private lastFilterJson = ""; // BasicFilter 自己発火検出
    private activeColTab  = -1;
    private prevColKey    = "";

    // ---- DOM ----
    private filterPanel:  HTMLElement;
    private searchInput:  HTMLInputElement;
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
    private hasAppliedFilter  = false;
    private isLoadingMore     = false;
    private dataLimitReached  = false;
    private persistTimer: number | null = null;
    private scrollRaf:    number | null = null;
    private emitTimer: number | null = null;
    private searchDebounceTimer: number | null = null;
    private rootEl:       HTMLElement;
    private rowHeight     = ROW_H;
    private colWidths: Map<number, number> = new Map();
    private sortColIdx = -1;
    private sortDir: "asc" | "desc" | null = null;
    private lastClickedRi = -1;

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
    // update
    // ==========================================================
    public update(options: VisualUpdateOptions): void {
        this.formattingSettings = this.formattingSettingsService
            .populateFormattingSettingsModel(VisualFormattingSettingsModel, options.dataViews?.[0]);

        const hasResize = !!(options.type & VisualUpdateType.Resize);
        const hasDataOrStyle = !!(options.type & (VisualUpdateType.Data | VisualUpdateType.Style));
        if (hasResize && !hasDataOrStyle) {
            this.renderVirtualRows();
            return;
        }

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

        if (isSegment && this.isLoadingMore) {
            // 中間チャンクは runFilter だけ（DOM 再構築は抑制）
            this.runFilter();
            this.renderStatus();
            return;
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
            this.searchText = "";
            this.appliedSearchText = "";
            this.lastFilterJson = "";
            if (this.hasAppliedFilter) this.removeFilter();
            this.selectedOrigIdx.clear();
        }

        const isFirstLoad = !this.hasInteracted;
        if (isFirstLoad || colsChanged) {
            this.restoreState(dv);
        }

        // jsonFilters から行選択を毎回再構築
        this.restoreFromJsonFilters(options.jsonFilters, dv);

        // 範囲外インデックスを除去
        const maxIdx = this.tableData.rows.length;
        this.selectedOrigIdx.forEach(i => { if (i >= maxIdx) this.selectedOrigIdx.delete(i); });
        if (this.selectedOrigIdx.size === 0 && this.hasAppliedFilter) {
            this.removeFilter();
        }

        if (isFirstLoad) {
            this.scrollEl.scrollTop = 0;
        }

        // --- レンダリング ---
        if (!this.filterPanel.querySelector(".search-input:focus")) {
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
        const m = dv?.metadata?.objects?.["filterState"];
        this.searchText = String(m?.["searchText"] ?? "");
        this.appliedSearchText = this.searchText;
        // 行選択 (selectedOrigIdx) は永続化しない（RLS 原則）。
        // 真実源は applyJsonFilter / options.jsonFilters。restoreFromJsonFilters で復元される。
    }

    private cellToString(v: PrimitiveValue): string {
        return v == null ? "" : String(v);
    }

    private extractTableData(dv: DataView): TableData {
        if (!dv?.table) { this.selectionIds = []; return { columns: [], rows: [], rawRows: [] }; }
        this.selectionIds = dv.table.rows.map((_, i) =>
            this.host.createSelectionIdBuilder().withTable(dv.table, i).createSelectionId()
        );
        return {
            columns: dv.table.columns.map(c => c.displayName || ""),
            rows:    dv.table.rows.map(r => r.map(c => this.cellToString(c as PrimitiveValue))),
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
            this.tableData.rows.push(table.rows[i].map(c => this.cellToString(c as PrimitiveValue)));
            this.tableData.rawRows.push(table.rows[i].map(c => (c == null ? null : c)) as PrimitiveValue[]);
            this.selectionIds.push(
                this.host.createSelectionIdBuilder().withTable(table, i).createSelectionId()
            );
        }
    }

    // ==========================================================
    // 検索パネル（単一の検索ボックス / 全列 substring マッチ）
    // ==========================================================
    private renderFilterPanel(): void {
        if (this.persistTimer !== null) { clearTimeout(this.persistTimer); this.persistTimer = null; }
        this.clear(this.filterPanel);

        const hdr = this.el("div", "filter-header");
        const ttl = this.el("span", "filter-title");
        ttl.textContent = "検索";
        hdr.appendChild(ttl);

        const help = this.el("span", "filter-help") as HTMLSpanElement;
        help.textContent = "ⓘ";
        help.title = "全列対象の部分一致（ローカル絞り込み）。条件フィルターは filterCondition ビジュアルを利用してください。";
        hdr.appendChild(help);

        this.filterPanel.appendChild(hdr);

        const row = this.el("div", "search-row");
        this.searchInput = this.el("input", "search-input") as HTMLInputElement;
        this.searchInput.type = "text";
        this.searchInput.placeholder = "全列を部分一致検索";
        this.searchInput.value = this.searchText;
        this.searchInput.oninput = () => {
            this.searchText = this.searchInput.value;
            if (this.searchDebounceTimer !== null) clearTimeout(this.searchDebounceTimer);
            this.searchDebounceTimer = window.setTimeout(() => {
                this.searchDebounceTimer = null;
                this.applySearchLocally();
            }, 150);
        };
        this.searchInput.onkeydown = (e: KeyboardEvent) => {
            if (e.key === "Enter") {
                if (this.searchDebounceTimer !== null) { clearTimeout(this.searchDebounceTimer); this.searchDebounceTimer = null; }
                this.applySearchLocally();
            }
        };
        row.appendChild(this.searchInput);

        const clearBtn = this.el("button", "clear-btn") as HTMLButtonElement;
        clearBtn.textContent = "×";
        clearBtn.title = "検索解除";
        clearBtn.onclick = () => {
            this.searchText = "";
            this.searchInput.value = "";
            this.applySearchLocally();
        };
        row.appendChild(clearBtn);

        this.filterPanel.appendChild(row);
    }

    private applySearchLocally(): void {
        this.hasInteracted = true;
        this.appliedSearchText = this.searchText;
        this.lastClickedRi = -1;
        this.runFilter();
        this.scrollEl.scrollTop = 0;
        this.renderTableHeader();
        this.renderVirtualRows();
        this.renderStatus();
        this.debounceSave();
    }

    // ==========================================================
    // 列トグルバー（タブ動作）
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
    }

    private isColVisible(i: number): boolean {
        return this.activeColTab === -1 || this.activeColTab === i;
    }

    // ==========================================================
    // ローカル絞り込み（全列 substring）
    // ==========================================================
    private runFilter(): void {
        this.filteredRows = []; this.filteredOrigIdx = [];
        const q = this.appliedSearchText.trim().toLowerCase();

        if (!q) {
            this.filteredRows    = this.tableData.rows.slice();
            this.filteredOrigIdx = this.tableData.rows.map((_, i) => i);
        } else {
            this.tableData.rows.forEach((row, oi) => {
                let hit = false;
                for (let ci = 0; ci < row.length; ci++) {
                    if ((row[ci] ?? "").toLowerCase().includes(q)) { hit = true; break; }
                }
                if (hit) { this.filteredRows.push(row); this.filteredOrigIdx.push(oi); }
            });
        }
        this.applySort();
        this.liftSelectedToTop();
    }

    /** 選択行を stable に最上位へ（選択内は元順維持） */
    private liftSelectedToTop(): void {
        if (this.selectedOrigIdx.size === 0) return;
        const selRows: string[][] = [];
        const selIdx: number[] = [];
        const restRows: string[][] = [];
        const restIdx: number[] = [];
        for (let i = 0; i < this.filteredOrigIdx.length; i++) {
            const oi = this.filteredOrigIdx[i];
            if (this.selectedOrigIdx.has(oi)) {
                selRows.push(this.filteredRows[i]);
                selIdx.push(oi);
            } else {
                restRows.push(this.filteredRows[i]);
                restIdx.push(oi);
            }
        }
        this.filteredRows    = selRows.concat(restRows);
        this.filteredOrigIdx = selIdx.concat(restIdx);
    }

    private applySort(): void {
        if (this.sortColIdx < 0 || !this.sortDir) return;
        const ci = this.sortColIdx;
        const dir = this.sortDir === "asc" ? 1 : -1;

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

            const arrow = this.el("span", "sort-indicator");
            if (this.sortColIdx === i && this.sortDir === "asc")  arrow.textContent = " ▲";
            else if (this.sortColIdx === i && this.sortDir === "desc") arrow.textContent = " ▼";
            else arrow.textContent = " △";
            th.appendChild(arrow);

            th.addEventListener("click", (e) => {
                if ((e.target as HTMLElement).classList.contains("col-resize-handle")) return;
                if (this.sortColIdx === i) {
                    this.sortDir = this.sortDir === "asc" ? "desc" : this.sortDir === "desc" ? null : "asc";
                    if (!this.sortDir) this.sortColIdx = -1;
                } else {
                    this.sortColIdx = i;
                    this.sortDir = "asc";
                }
                this.lastClickedRi = -1;
                this.runFilter();
                this.renderTableHeader();
                this.renderVirtualRows();
            });

            const handle = this.el("div", "col-resize-handle");
            handle.addEventListener("mousedown", (e) => this.onColResizeStart(e, i));
            th.appendChild(handle);

            tr.appendChild(th);
        });
        this.thead.appendChild(tr);

        if (this.colWidths.size > 0) {
            let total = 32;
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
        cb.addEventListener("click", (ev) => { ev.preventDefault(); });
        cbTd.appendChild(cb); tr.appendChild(cbTd);

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
            const from = Math.min(this.lastClickedRi, ri);
            const to   = Math.max(this.lastClickedRi, ri);
            if (!ctrlOrMeta) this.selectedOrigIdx.clear();
            for (let r = from; r <= to; r++) {
                this.selectedOrigIdx.add(this.filteredOrigIdx[r]);
            }
        } else if (ctrlOrMeta) {
            if (this.selectedOrigIdx.has(oi)) this.selectedOrigIdx.delete(oi);
            else this.selectedOrigIdx.add(oi);
        } else {
            const onlyThis = this.selectedOrigIdx.size === 1 && this.selectedOrigIdx.has(oi);
            this.selectedOrigIdx.clear();
            if (!onlyThis) this.selectedOrigIdx.add(oi);
        }

        this.commitSelection();
        // 並び替え後に同じ行を指す ri を再取得
        this.lastClickedRi = this.filteredOrigIdx.indexOf(oi);
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
        // 選択行を最上位へ並べ替え直す
        this.runFilter();
        // 並べ替えた結果を実際に見える位置（最上段）へスクロールリセット
        if (this.scrollEl) this.scrollEl.scrollTop = 0;
        this.renderTableHeader();
        this.renderVirtualRows();
        this.renderStatus();
    }

    /**
     * 選択行の値で BasicFilter を発火（全列）+ SelectionManager（同ページ）
     * 連続クリックに備えて applyJsonFilter は 150ms デバウンス。
     */
    private applyDatasetFilter(): void {
        const srcIdx = Array.from(this.selectedOrigIdx);
        if (srcIdx.length === 0) {
            this.removeFilter();
            return;
        }
        const ids = srcIdx
            .filter(i => i < this.selectionIds.length)
            .map(i => this.selectionIds[i]);
        if (ids.length === 0) {
            this.removeFilter();
            return;
        }
        this.hasAppliedFilter = true;
        this.selectionManager.select(ids);

        if (this.emitTimer !== null) clearTimeout(this.emitTimer);
        this.emitTimer = window.setTimeout(() => this.flushJsonFilterEmit(srcIdx), 150);
    }

    private flushJsonFilterEmit(_snapshotSrc: number[]): void {
        this.emitTimer = null;
        const srcIdx = Array.from(this.selectedOrigIdx);
        if (srcIdx.length === 0) return;
        this.emitBasicFilterForSync(srcIdx);
    }

    private emitBasicFilterForSync(srcIdx: number[]): void {
        const dv = this.lastDataView;
        if (!dv?.table?.columns?.length) return;

        const cols = dv.table.columns;
        const filters: BasicFilter[] = [];
        const sigParts: string[] = [];

        for (let ci = 0; ci < cols.length; ci++) {
            const target = buildFilterTarget(cols[ci]);
            if (!target) continue;

            const valueMap = new Map<string, FilterValue>();
            for (const i of srcIdx) {
                const raw = this.tableData.rawRows[i]?.[ci];
                if (raw == null || raw === "") continue;
                const key = normalizeValue(raw);
                if (!valueMap.has(key)) valueMap.set(key, raw as FilterValue);
            }
            if (valueMap.size === 0) continue;

            const rawValues = Array.from(valueMap.values()) as (string | number | boolean)[];
            filters.push(new BasicFilter(target, "In", ...rawValues));
            sigParts.push(filterSignature(target, Array.from(valueMap.keys())));
        }

        if (filters.length === 0) return;

        const key = "BASIC|" + sigParts.slice().sort().join("|");
        if (key === this.lastFilterJson) return;

        this.lastFilterJson = key;
        this.host.applyJsonFilter(filters, "general", "filter", FilterAction.merge);
    }

    private removeFilter(): void {
        if (this.emitTimer !== null) { clearTimeout(this.emitTimer); this.emitTimer = null; }
        if (!this.hasAppliedFilter) return;
        this.selectionManager.clear();
        this.lastFilterJson = "";
        this.host.applyJsonFilter(null, "general", "filter", FilterAction.remove);
        this.hasAppliedFilter = false;
    }

    /** 外部 BasicFilter を受信して行選択を復元（AdvancedFilter は filterCondition 側の責務なので無視） */
    private restoreFromJsonFilters(jsonFilters: powerbi.IFilter[] | undefined, dv: DataView): boolean {
        if (!jsonFilters || jsonFilters.length === 0) return false;

        const basic: IBasicFilter[] = [];
        for (const f of jsonFilters) {
            const ft = (f as unknown as { filterType?: FilterType })?.filterType;
            if (ft === FilterType.Basic) basic.push(f as unknown as IBasicFilter);
        }
        if (basic.length === 0) return false;
        return this.restoreFromBasicFilters(basic, dv);
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

        const matched = new Set<number>();
        this.tableData.rawRows.forEach((row, i) => {
            for (const p of parsed) {
                const raw = row[p.colIdx];
                if (raw == null) return;
                if (!p.valueSet.has(normalizeValue(raw))) return;
            }
            matched.add(i);
        });

        if (matched.size === 0) return false;

        this.selectedOrigIdx = matched;
        this.lastFilterJson = incomingKey;

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

        const selSize = this.selectedOrigIdx.size;
        if (selSize > 0) {
            const info = this.el("span", "sel-info");
            info.textContent = `　${selSize} 件選択中`;
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

        const ths = this.thead.querySelectorAll("th");
        let visIdx = 0;
        for (let i = 0; i < this.tableData.columns.length; i++) {
            if (!this.isColVisible(i)) continue;
            if (i === colIdx) break;
            visIdx++;
        }
        const thEl = ths[visIdx + 1] as HTMLElement;
        const startW = thEl ? thEl.getBoundingClientRect().width : 80;

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

        this.rowHeight = Math.max(ROW_H, Math.round(vSize * 1.333 * 1.6 + 4));
        s.setProperty("--val-row-height", this.rowHeight + "px");

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
        // 検索語のみ永続化（行選択は真実源 = jsonFilters で復元）
        this.host.persistProperties({ merge: [{ objectName: "filterState", selector: null, properties: {
            searchText: this.searchText,
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
        if (this.searchDebounceTimer !== null) { clearTimeout(this.searchDebounceTimer); this.searchDebounceTimer = null; }
    }
}
