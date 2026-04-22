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

export class Visual implements IVisual {
    private host: IVisualHost;
    private formattingSettings: VisualFormattingSettingsModel;
    private formattingSettingsService: FormattingSettingsService;

    // ---- データ ----
    private tableData: TableData = { columns: [], rows: [], rawRows: [] };
    private selectedOrigIdx = new Set<number>();
    // SelectionId は必要になった時だけ生成してキャッシュ
    private selectionIdCache = new Map<number, ISelectionId>();
    private selectionManager: ISelectionManager;
    private lastDataView: DataView | null = null;
    private lastFilterJson = "";
    private prevColKey = "";

    // ---- DOM ----
    private rootEl: HTMLElement;
    private statusBar: HTMLElement;
    private tableScroll: HTMLElement;
    private table: HTMLTableElement;
    private thead: HTMLTableSectionElement;
    private tbody: HTMLTableSectionElement;
    private headerCb: HTMLInputElement | null = null;
    private rowNodes = new Map<number, { tr: HTMLTableRowElement; cb: HTMLInputElement }>();

    // ---- 制御 ----
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
    }

    // ==========================================================
    // update
    // ==========================================================
    public update(options: VisualUpdateOptions): void {
        this.formattingSettings = this.formattingSettingsService
            .populateFormattingSettingsModel(VisualFormattingSettingsModel, options.dataViews?.[0]);

        if (!(options.type & VisualUpdateType.Data)) {
            if (options.type & VisualUpdateType.Style) this.applyStyles();
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

        // 中間チャンクは status だけ更新
        if (isSegment && this.isLoadingMore) {
            this.renderStatus();
            return;
        }

        // 列構成変化で選択状態リセット
        const colKey = this.tableData.columns.join("\0");
        if (colKey !== this.prevColKey) {
            this.prevColKey = colKey;
            this.selectedOrigIdx.clear();
            this.lastClickedOi = -1;
            this.lastFilterJson = "";
            if (this.hasAppliedFilter) this.removeFilter();
        }

        // 外部 jsonFilters から行選択を復元
        this.restoreFromJsonFilters(options.jsonFilters, dv);

        // 範囲外 index を削除
        const max = this.tableData.rows.length;
        this.selectedOrigIdx.forEach(i => { if (i >= max) this.selectedOrigIdx.delete(i); });

        this.applyStyles();
        this.renderAll();
    }

    // ==========================================================
    // データ抽出
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
    // 描画
    // ==========================================================
    private renderAll(): void {
        this.renderHeader();
        this.renderRows();
        this.renderStatus();
    }

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

    private renderRows(): void {
        while (this.tbody.firstChild) this.tbody.removeChild(this.tbody.firstChild);
        this.rowNodes.clear();

        const total = this.tableData.rows.length;
        if (total === 0) {
            const tr = document.createElement("tr");
            const td = document.createElement("td");
            td.className = "no-data";
            td.colSpan = this.tableData.columns.length + 1;
            td.textContent = this.tableData.columns.length === 0
                ? "データをフィールドに追加してください"
                : "該当するデータがありません";
            tr.appendChild(td);
            this.tbody.appendChild(tr);
            return;
        }

        const frag = document.createDocumentFragment();
        for (let i = 0; i < total; i++) {
            frag.appendChild(this.makeRow(i));
        }
        this.tbody.appendChild(frag);
    }

    private makeRow(oi: number): HTMLTableRowElement {
        const sel = this.selectedOrigIdx.has(oi);
        const tr = document.createElement("tr");
        tr.className = oi % 2 === 0 ? "row-even" : "row-odd";
        if (sel) tr.classList.add("row-selected");

        const cbTd = document.createElement("td");
        cbTd.className = "cb-col";
        const cb = document.createElement("input");
        cb.type = "checkbox";
        cb.checked = sel;
        cb.addEventListener("click", e => e.stopPropagation());
        cb.addEventListener("change", () => this.onRowToggle(oi, cb.checked));
        cbTd.appendChild(cb);
        tr.appendChild(cbTd);

        tr.addEventListener("click", e => this.onRowClick(oi, e));

        const row = this.tableData.rows[oi];
        for (let ci = 0; ci < this.tableData.columns.length; ci++) {
            const td = document.createElement("td");
            td.textContent = row[ci] ?? "";
            tr.appendChild(td);
        }

        this.rowNodes.set(oi, { tr, cb });
        return tr;
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
        this.refreshRowsUI(changed);
        this.commitSelection();
    }

    private onRowToggle(oi: number, checked: boolean): void {
        if (checked) this.selectedOrigIdx.add(oi);
        else this.selectedOrigIdx.delete(oi);
        this.lastClickedOi = oi;
        this.refreshRowsUI(new Set([oi]));
        this.commitSelection();
    }

    private toggleSelectAll(): void {
        const total = this.tableData.rows.length;
        if (total === 0) return;
        const allSel = this.selectedOrigIdx.size === total;
        const changed = new Set<number>();
        if (allSel) {
            this.selectedOrigIdx.forEach(i => changed.add(i));
            this.selectedOrigIdx.clear();
        } else {
            for (let i = 0; i < total; i++) {
                if (!this.selectedOrigIdx.has(i)) {
                    this.selectedOrigIdx.add(i);
                    changed.add(i);
                }
            }
        }
        this.refreshRowsUI(changed);
        this.commitSelection();
    }

    private clearSelection(): void {
        const changed = new Set<number>(this.selectedOrigIdx);
        this.selectedOrigIdx.clear();
        this.refreshRowsUI(changed);
        this.commitSelection();
    }

    private refreshRowsUI(changedOi: Set<number>): void {
        changedOi.forEach(oi => {
            const node = this.rowNodes.get(oi);
            if (!node) return;
            const sel = this.selectedOrigIdx.has(oi);
            node.tr.classList.toggle("row-selected", sel);
            node.cb.checked = sel;
        });
        this.refreshHeaderCb();
    }

    private refreshHeaderCb(): void {
        if (!this.headerCb) return;
        const total = this.tableData.rows.length;
        const allSel = total > 0 && this.selectedOrigIdx.size === total;
        const someSel = !allSel && this.selectedOrigIdx.size > 0;
        this.headerCb.checked = allSel;
        this.headerCb.indeterminate = someSel;
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
        if (ids.length === 0) {
            this.removeFilter();
            return;
        }
        this.hasAppliedFilter = true;
        this.selectionManager.select(ids);

        if (this.emitTimer !== null) clearTimeout(this.emitTimer);
        this.emitTimer = window.setTimeout(() => this.flushEmit(srcIdx), 120);
    }

    private flushEmit(_snapshot: number[]): void {
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
        s.setProperty("--val-white-space", v.wordWrap.value ? "pre-line" : "nowrap");
        s.setProperty("--val-row-height", Math.max(24, Math.round(vSize * 1.333 * 1.6 + 4)) + "px");

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
    }
}
