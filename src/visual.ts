"use strict";

import powerbi from "powerbi-visuals-api";
import { FormattingSettingsService } from "powerbi-visuals-utils-formattingmodel";
import {
    BasicFilter, IFilterColumnTarget, IBasicFilter, FilterType,
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
import DataViewTable            = powerbi.DataViewTable;

import { VisualFormattingSettingsModel } from "./settings";
import {
    PrimitiveValue, FilterValue, TableData,
    normalizeValue, filterSignature, buildFilterTarget,
} from "./filterEngine";

interface PoolRow {
    tr: HTMLTableRowElement;
    cb: HTMLInputElement;
    cells: HTMLTableCellElement[];
    oi: number; // -1 if unbound
}

const BUFFER_ROWS = 6;

export class Visual implements IVisual {
    private host: IVisualHost;
    private formattingSettings: VisualFormattingSettingsModel;
    private formattingSettingsService: FormattingSettingsService;

    private tableData: TableData = { columns: [], rows: [], rawRows: [] };
    private selectedOrigIdx = new Set<number>();
    private selectionIdCache = new Map<number, ISelectionId>();
    private selectionManager: ISelectionManager;
    private lastDataView: DataView | null = null;
    private lastFilterJson = "";
    private prevColKey = "";

    // DOM
    private rootEl: HTMLElement;
    private statusBar: HTMLElement;
    private tableScroll: HTMLElement;
    private table: HTMLTableElement;
    private thead: HTMLTableSectionElement;
    private tbody: HTMLTableSectionElement;
    private topSpacer: HTMLTableRowElement;
    private bottomSpacer: HTMLTableRowElement;
    private headerCb: HTMLInputElement | null = null;

    // 仮想スクロール
    private pool: PoolRow[] = [];
    private poolColCount = -1; // 現在プール行が持っているセル数
    private rowHeight = 24;
    private scrollRaf: number | null = null;

    // 制御
    private hasAppliedFilter = false;
    private isLoadingMore = false;
    private dataLimitReached = false;
    private emitTimer: number | null = null;
    private lastClickedOi = -1;

    constructor(options: VisualConstructorOptions) {
        this.host = options.host;
        this.selectionManager = options.host.createSelectionManager();
        this.formattingSettingsService = new FormattingSettingsService();
        this.rootEl = options.element;
        this.rootEl.className = "filter-table-visual";
        this.buildDOM();
    }

    private buildDOM(): void {
        this.statusBar = document.createElement("div");
        this.statusBar.className = "status-bar";

        this.tableScroll = document.createElement("div");
        this.tableScroll.className = "table-scroll";

        this.table = document.createElement("table");
        this.table.className = "data-table";
        this.thead = document.createElement("thead");
        this.tbody = document.createElement("tbody");
        this.table.appendChild(this.thead);
        this.table.appendChild(this.tbody);
        this.tableScroll.appendChild(this.table);

        this.rootEl.appendChild(this.statusBar);
        this.rootEl.appendChild(this.tableScroll);

        this.tableScroll.addEventListener("scroll", () => this.onScroll(), { passive: true });
    }

    // ==========================================================
    // update
    // ==========================================================
    public update(options: VisualUpdateOptions): void {
        this.formattingSettings = this.formattingSettingsService
            .populateFormattingSettingsModel(VisualFormattingSettingsModel, options.dataViews?.[0]);

        if (options.type & VisualUpdateType.Resize) {
            this.renderVirtual();
        }

        if (!(options.type & VisualUpdateType.Data)) {
            if (options.type & VisualUpdateType.Style) {
                this.applyStyles();
                this.renderVirtual();
            }
            return;
        }

        const dv = options.dataViews?.[0];
        this.lastDataView = dv ?? null;

        const isSegment = options.operationKind === VisualDataChangeOperationKind.Segment;
        const hasMoreSegments = !!(dv?.metadata?.segment);

        if (isSegment && dv?.table) {
            this.appendIncremental(dv.table);
        } else {
            this.tableData = this.extractTable(dv);
            this.selectionIdCache.clear();
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
            // 中間チャンクは下部 spacer だけ伸ばす（見えない領域なので）
            this.updateSpacerHeights();
            this.renderStatus();
            return;
        }

        const colKey = this.tableData.columns.join("\0");
        const colsChanged = colKey !== this.prevColKey;
        if (colsChanged) {
            this.prevColKey = colKey;
            this.selectedOrigIdx.clear();
            this.lastClickedOi = -1;
            this.lastFilterJson = "";
            if (this.hasAppliedFilter) this.removeFilter();
        }

        this.restoreFromJsonFilters(options.jsonFilters, dv);

        const max = this.tableData.rows.length;
        this.selectedOrigIdx.forEach(i => { if (i >= max) this.selectedOrigIdx.delete(i); });

        this.applyStyles();
        this.renderHeader();
        this.ensurePool();
        this.renderVirtual();
        this.renderStatus();
    }

    // ==========================================================
    // データ
    // ==========================================================
    private cellStr(v: PrimitiveValue): string {
        return v == null ? "" : String(v);
    }

    private extractTable(dv: DataView | undefined): TableData {
        if (!dv?.table) return { columns: [], rows: [], rawRows: [] };
        return {
            columns: dv.table.columns.map(c => c.displayName || ""),
            rows:    dv.table.rows.map(r => r.map(c => this.cellStr(c as PrimitiveValue))),
            rawRows: dv.table.rows.map(r => r.map(c => (c == null ? null : c)) as PrimitiveValue[]),
        };
    }

    private appendIncremental(table: DataViewTable): void {
        const offset = (table as unknown as Record<string, unknown>)["lastMergeIndex"] as number | undefined;
        const startIdx = (offset === undefined) ? 0 : offset + 1;
        if (this.tableData.columns.length === 0) {
            this.tableData.columns = table.columns.map(c => c.displayName || "");
        }
        for (let i = startIdx; i < table.rows.length; i++) {
            this.tableData.rows.push(table.rows[i].map(c => this.cellStr(c as PrimitiveValue)));
            this.tableData.rawRows.push(table.rows[i].map(c => (c == null ? null : c)) as PrimitiveValue[]);
        }
    }

    private getSelectionId(origIdx: number): ISelectionId | null {
        const dv = this.lastDataView;
        if (!dv?.table) return null;
        let id = this.selectionIdCache.get(origIdx);
        if (!id) {
            id = this.host.createSelectionIdBuilder()
                .withTable(dv.table, origIdx)
                .createSelectionId();
            this.selectionIdCache.set(origIdx, id);
        }
        return id;
    }

    // ==========================================================
    // ヘッダ描画
    // ==========================================================
    private renderHeader(): void {
        while (this.thead.firstChild) this.thead.removeChild(this.thead.firstChild);
        if (this.tableData.columns.length === 0) return;

        const tr = document.createElement("tr");

        const cbTh = document.createElement("th");
        cbTh.className = "cb-col";
        this.headerCb = document.createElement("input");
        this.headerCb.type = "checkbox";
        this.headerCb.onchange = () => this.toggleSelectAll();
        cbTh.appendChild(this.headerCb);
        tr.appendChild(cbTh);

        this.tableData.columns.forEach(col => {
            const th = document.createElement("th");
            th.textContent = col;
            tr.appendChild(th);
        });
        this.thead.appendChild(tr);
        this.refreshHeaderCb();
    }

    // ==========================================================
    // 行プール構築（仮想スクロールの肝）
    // ==========================================================
    private ensurePool(): void {
        const colCount = this.tableData.columns.length;

        // 列数が変わったらプール再構築
        if (this.poolColCount !== colCount) {
            while (this.tbody.firstChild) this.tbody.removeChild(this.tbody.firstChild);
            this.pool = [];
            this.poolColCount = colCount;

            this.topSpacer = document.createElement("tr");
            this.topSpacer.className = "spacer-row";
            const topTd = document.createElement("td");
            topTd.colSpan = colCount + 1;
            this.topSpacer.appendChild(topTd);

            this.bottomSpacer = document.createElement("tr");
            this.bottomSpacer.className = "spacer-row";
            const botTd = document.createElement("td");
            botTd.colSpan = colCount + 1;
            this.bottomSpacer.appendChild(botTd);

            this.tbody.appendChild(this.topSpacer);
            this.tbody.appendChild(this.bottomSpacer);
        }

        // no-data / empty の場合はプール不要
        if (this.tableData.rows.length === 0 || colCount === 0) {
            return;
        }

        // プールサイズ = ビューポートに必要な行数 + バッファ。最小 20
        const viewH = Math.max(this.tableScroll.clientHeight, 200);
        const need = Math.ceil(viewH / this.rowHeight) + BUFFER_ROWS * 2;
        const target = Math.max(20, need);

        while (this.pool.length < target) {
            this.pool.push(this.createPoolRow(colCount));
        }
        // プール行が多すぎたら末尾を削除
        while (this.pool.length > target) {
            const row = this.pool.pop();
            if (row?.tr.parentElement) row.tr.remove();
        }
    }

    private createPoolRow(colCount: number): PoolRow {
        const tr = document.createElement("tr");
        tr.style.display = "none";
        tr.dataset.oi = "-1";

        const cbTd = document.createElement("td");
        cbTd.className = "cb-col";
        const cb = document.createElement("input");
        cb.type = "checkbox";
        cb.addEventListener("click", e => e.stopPropagation());
        cb.addEventListener("change", () => {
            const oi = Number(tr.dataset.oi);
            if (oi >= 0) this.onRowToggle(oi, cb.checked);
        });
        cbTd.appendChild(cb);
        tr.appendChild(cbTd);

        const cells: HTMLTableCellElement[] = [];
        for (let i = 0; i < colCount; i++) {
            const td = document.createElement("td");
            tr.appendChild(td);
            cells.push(td);
        }

        tr.addEventListener("click", e => {
            const oi = Number(tr.dataset.oi);
            if (oi >= 0) this.onRowClick(oi, e);
        });

        // bottomSpacer の前に差し込む
        this.tbody.insertBefore(tr, this.bottomSpacer);

        return { tr, cb, cells, oi: -1 };
    }

    // ==========================================================
    // 仮想スクロール描画
    // ==========================================================
    private onScroll(): void {
        if (this.scrollRaf !== null) return;
        this.scrollRaf = requestAnimationFrame(() => {
            this.scrollRaf = null;
            this.renderVirtual();
        });
    }

    private updateSpacerHeights(): void {
        const total = this.tableData.rows.length;
        if (!this.topSpacer || !this.bottomSpacer) return;
        // no-data 時は spacer 不要
        if (total === 0) {
            this.topSpacer.style.height = "0px";
            this.bottomSpacer.style.height = "0px";
        }
    }

    private renderVirtual(): void {
        const total = this.tableData.rows.length;

        // 空 or 列なし → no-data 行を表示
        if (total === 0 || this.tableData.columns.length === 0) {
            this.showNoDataRow();
            return;
        }
        this.hideNoDataRow();

        const rh = this.rowHeight;
        const viewH = this.tableScroll.clientHeight;
        const scrollTop = this.tableScroll.scrollTop;

        const startRow = Math.max(0, Math.floor(scrollTop / rh) - BUFFER_ROWS);
        const visCount = Math.min(total - startRow, this.pool.length);
        const endRow = startRow + visCount;

        this.topSpacer.style.height = (startRow * rh) + "px";
        this.bottomSpacer.style.height = Math.max(0, (total - endRow) * rh) + "px";

        for (let p = 0; p < this.pool.length; p++) {
            const row = this.pool[p];
            const oi = startRow + p;
            if (p < visCount) {
                this.bindRow(row, oi);
            } else {
                if (row.oi !== -1) {
                    row.tr.style.display = "none";
                    row.oi = -1;
                }
            }
        }
    }

    private bindRow(row: PoolRow, oi: number): void {
        const sel = this.selectedOrigIdx.has(oi);
        row.tr.style.display = "";
        row.tr.dataset.oi = String(oi);
        row.oi = oi;
        row.tr.className = (oi % 2 === 0 ? "row-even" : "row-odd") + (sel ? " row-selected" : "");
        row.cb.checked = sel;

        const data = this.tableData.rows[oi];
        for (let c = 0; c < row.cells.length; c++) {
            const txt = data[c] ?? "";
            if (row.cells[c].textContent !== txt) row.cells[c].textContent = txt;
        }
    }

    private showNoDataRow(): void {
        // プールを全て非表示、spacer も 0 にし no-data 行だけ出す
        this.pool.forEach(r => { r.tr.style.display = "none"; r.oi = -1; });
        if (this.topSpacer) this.topSpacer.style.height = "0px";
        if (this.bottomSpacer) this.bottomSpacer.style.height = "0px";

        // 既存 no-data 行があれば使い回す
        let nd = this.tbody.querySelector(".no-data-row") as HTMLTableRowElement | null;
        if (!nd) {
            nd = document.createElement("tr");
            nd.className = "no-data-row";
            const td = document.createElement("td");
            td.className = "no-data";
            td.colSpan = Math.max(1, this.tableData.columns.length + 1);
            nd.appendChild(td);
            if (this.bottomSpacer) this.tbody.insertBefore(nd, this.bottomSpacer);
            else this.tbody.appendChild(nd);
        }
        const td = nd.firstChild as HTMLTableCellElement;
        td.colSpan = Math.max(1, this.tableData.columns.length + 1);
        td.textContent = this.tableData.columns.length === 0
            ? "データをフィールドに追加してください"
            : "該当するデータがありません";
    }

    private hideNoDataRow(): void {
        const nd = this.tbody.querySelector(".no-data-row");
        if (nd) nd.remove();
    }

    // ==========================================================
    // 選択
    // ==========================================================
    private onRowClick(oi: number, e: MouseEvent): void {
        const ctrlOrMeta = e.ctrlKey || e.metaKey;
        const changed = new Set<number>();

        if (e.shiftKey && this.lastClickedOi >= 0) {
            const from = Math.min(this.lastClickedOi, oi);
            const to   = Math.max(this.lastClickedOi, oi);
            if (!ctrlOrMeta) {
                this.selectedOrigIdx.forEach(i => changed.add(i));
                this.selectedOrigIdx.clear();
            }
            for (let i = from; i <= to; i++) {
                if (!this.selectedOrigIdx.has(i)) {
                    this.selectedOrigIdx.add(i);
                    changed.add(i);
                }
            }
        } else if (ctrlOrMeta) {
            if (this.selectedOrigIdx.has(oi)) this.selectedOrigIdx.delete(oi);
            else this.selectedOrigIdx.add(oi);
            changed.add(oi);
        } else {
            const onlyThis = this.selectedOrigIdx.size === 1 && this.selectedOrigIdx.has(oi);
            this.selectedOrigIdx.forEach(i => changed.add(i));
            this.selectedOrigIdx.clear();
            if (!onlyThis) {
                this.selectedOrigIdx.add(oi);
                changed.add(oi);
            }
        }

        this.lastClickedOi = oi;
        this.refreshVisibleSelection(changed);
        this.commitSelection();
    }

    private onRowToggle(oi: number, checked: boolean): void {
        if (checked) this.selectedOrigIdx.add(oi);
        else this.selectedOrigIdx.delete(oi);
        this.lastClickedOi = oi;
        this.refreshVisibleSelection(new Set([oi]));
        this.commitSelection();
    }

    private toggleSelectAll(): void {
        const total = this.tableData.rows.length;
        if (total === 0) return;
        const allSel = this.selectedOrigIdx.size === total;
        if (allSel) this.selectedOrigIdx.clear();
        else for (let i = 0; i < total; i++) this.selectedOrigIdx.add(i);
        this.refreshAllVisibleSelection();
        this.commitSelection();
    }

    private clearSelection(): void {
        this.selectedOrigIdx.clear();
        this.refreshAllVisibleSelection();
        this.commitSelection();
    }

    /** changed に含まれる oi のうち、プール上に表示されているものだけ DOM 同期 */
    private refreshVisibleSelection(changed: Set<number>): void {
        this.pool.forEach(row => {
            if (row.oi < 0 || !changed.has(row.oi)) return;
            const sel = this.selectedOrigIdx.has(row.oi);
            row.tr.classList.toggle("row-selected", sel);
            row.cb.checked = sel;
        });
        this.refreshHeaderCb();
    }

    /** 全選択/全解除など変更行が多い時はプール全体を走査 */
    private refreshAllVisibleSelection(): void {
        this.pool.forEach(row => {
            if (row.oi < 0) return;
            const sel = this.selectedOrigIdx.has(row.oi);
            row.tr.classList.toggle("row-selected", sel);
            row.cb.checked = sel;
        });
        this.refreshHeaderCb();
    }

    private refreshHeaderCb(): void {
        if (!this.headerCb) return;
        const total = this.tableData.rows.length;
        const size = this.selectedOrigIdx.size;
        this.headerCb.checked = total > 0 && size === total;
        this.headerCb.indeterminate = size > 0 && size < total;
    }

    private commitSelection(): void {
        this.applyDatasetFilter();
        this.renderStatus();
    }

    // ==========================================================
    // フィルター発火
    // ==========================================================
    private applyDatasetFilter(): void {
        const srcIdx = Array.from(this.selectedOrigIdx);
        if (srcIdx.length === 0) {
            this.removeFilter();
            return;
        }
        const ids: ISelectionId[] = [];
        for (const i of srcIdx) {
            const id = this.getSelectionId(i);
            if (id) ids.push(id);
        }
        if (ids.length === 0) { this.removeFilter(); return; }
        this.hasAppliedFilter = true;
        this.selectionManager.select(ids);

        if (this.emitTimer !== null) clearTimeout(this.emitTimer);
        this.emitTimer = window.setTimeout(() => this.flushEmit(), 120);
    }

    private flushEmit(): void {
        this.emitTimer = null;
        const srcIdx = Array.from(this.selectedOrigIdx);
        if (srcIdx.length === 0) return;
        this.emitBasicFilter(srcIdx);
    }

    private emitBasicFilter(srcIdx: number[]): void {
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

    // ==========================================================
    // 外部 jsonFilters 受信
    // ==========================================================
    private restoreFromJsonFilters(jsonFilters: powerbi.IFilter[] | undefined, dv: DataView | undefined): boolean {
        if (!jsonFilters || jsonFilters.length === 0) return false;
        const basic: IBasicFilter[] = [];
        for (const f of jsonFilters) {
            const ft = (f as unknown as { filterType?: FilterType })?.filterType;
            if (ft === FilterType.Basic) basic.push(f as unknown as IBasicFilter);
        }
        if (basic.length === 0) return false;
        return this.restoreFromBasic(basic, dv);
    }

    private restoreFromBasic(basicFilters: IBasicFilter[], dv: DataView | undefined): boolean {
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
            parsed.push({ colIdx, valueSet: new Set(normalized), sig: filterSignature(tgt, normalized) });
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

        const ids: ISelectionId[] = [];
        for (const i of matched) {
            const id = this.getSelectionId(i);
            if (id) ids.push(id);
        }
        if (ids.length > 0) {
            this.hasAppliedFilter = true;
            this.selectionManager.select(ids);
        }
        return true;
    }

    // ==========================================================
    // ステータスバー
    // ==========================================================
    private renderStatus(): void {
        while (this.statusBar.firstChild) this.statusBar.removeChild(this.statusBar.firstChild);
        const t = this.tableData.rows.length;
        let text = `${t} 件`;
        if (this.isLoadingMore) text += "（読み込み中…）";
        else if (this.dataLimitReached) text += "（データ制限到達）";
        this.statusBar.appendChild(document.createTextNode(text));

        const selSize = this.selectedOrigIdx.size;
        if (selSize > 0) {
            const info = document.createElement("span");
            info.className = "sel-info";
            info.textContent = `　${selSize} 件選択中`;
            this.statusBar.appendChild(info);
            const clr = document.createElement("button");
            clr.className = "clear-sel-btn";
            clr.textContent = "選択解除";
            clr.onclick = () => this.clearSelection();
            this.statusBar.appendChild(clr);
        }
    }

    // ==========================================================
    // 書式適用
    // ==========================================================
    private applyStyles(): void {
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

        const rh = Math.max(24, Math.round(vSize * 1.333 * 1.6 + 4));
        this.rowHeight = rh;
        s.setProperty("--val-row-height", rh + "px");

        s.setProperty("--hdr-font-family", h.font.fontFamily.value);
        s.setProperty("--hdr-font-size", h.font.fontSize.value + "pt");
        s.setProperty("--hdr-font-weight", h.font.bold.value ? "bold" : "normal");
        s.setProperty("--hdr-font-style", h.font.italic.value ? "italic" : "normal");
        s.setProperty("--hdr-text-decoration", h.font.underline.value ? "underline" : "none");
        s.setProperty("--hdr-color", h.fontColor.value.value);
        s.setProperty("--hdr-bg", h.backgroundColor.value.value);
    }

    public getFormattingModel(): powerbi.visuals.FormattingModel {
        if (!this.formattingSettings) this.formattingSettings = new VisualFormattingSettingsModel();
        return this.formattingSettingsService.buildFormattingModel(this.formattingSettings);
    }

    public destroy(): void {
        if (this.emitTimer !== null) { clearTimeout(this.emitTimer); this.emitTimer = null; }
        if (this.scrollRaf !== null) { cancelAnimationFrame(this.scrollRaf); this.scrollRaf = null; }
    }
}
