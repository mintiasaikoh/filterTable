"use strict";

import powerbi from "powerbi-visuals-api";
import { FormattingSettingsService } from "powerbi-visuals-utils-formattingmodel";
import { BasicFilter, IFilterColumnTarget } from "powerbi-models";
import "./../style/visual.less";

import VisualConstructorOptions = powerbi.extensibility.visual.VisualConstructorOptions;
import VisualUpdateOptions      = powerbi.extensibility.visual.VisualUpdateOptions;
import IVisual                  = powerbi.extensibility.visual.IVisual;
import IVisualHost              = powerbi.extensibility.visual.IVisualHost;
import ISelectionManager        = powerbi.extensibility.ISelectionManager;
import DataView                 = powerbi.DataView;
import FilterAction             = powerbi.FilterAction;

import { VisualFormattingSettingsModel } from "./settings";

const ROW_H  = 24;   // px（tbody 行の高さ）
const BUFFER = 8;    // ビューポート外に余分に描画しておく行数

interface FilterCondition {
    columnIndex: number;
    operator: "contains" | "notContains";
    value: string;
}

interface TableData {
    columns: string[];
    rows: string[][];
}

export class Visual implements IVisual {
    private host: IVisualHost;
    private formattingSettings: VisualFormattingSettingsModel;
    private formattingSettingsService: FormattingSettingsService;
    private selectionManager: ISelectionManager;

    // ---- データ状態 ----
    private conditions: FilterCondition[]        = [];
    private logic: "AND" | "OR"                  = "AND";
    private tableData: TableData                 = { columns: [], rows: [] };
    private appliedConditions: FilterCondition[] = [];
    private appliedLogic: "AND" | "OR"           = "AND";
    private filteredRows: string[][]             = [];
    private filteredOrigIdx: number[]            = [];
    private selectionIds: powerbi.visuals.ISelectionId[] = [];
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
    private skipRender        = false;
    private hasInteracted     = false;
    private selfFilterApplied = false; // applyJsonFilter 後の update で選択クリアを防ぐ
    private prevRowCount      = -1;    // データ変化検知用
    private persistTimer: number | null = null;
    private scrollRaf:    number | null = null;

    constructor(options: VisualConstructorOptions) {
        this.host = options.host;
        this.formattingSettingsService = new FormattingSettingsService();
        this.selectionManager = options.host.createSelectionManager();
        options.element.className = "filter-table-visual";
        this.buildDOM(options.element);
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

        const dv: DataView = options.dataViews?.[0];

        // 自分の persist による update はパネルのみ更新してスキップ
        if (this.skipRender) {
            this.skipRender = false;
            if (!this.filterPanel.querySelector(".value-input:focus")) this.renderFilterPanel();
            return;
        }

        this.lastDataView = dv;
        this.tableData = this.extractTableData(dv);
        this.buildSelectionIds(dv);

        // 列数が変わったらタブをリセット
        const colsChanged = this.tableData.columns.length !== this.colCount;
        this.colCount = this.tableData.columns.length;
        if (colsChanged) this.activeColTab = -1;

        // 行数が変わったとき（外部スライサー等）：スクロールリセット＋選択クリア
        // ただし自分で applyJsonFilter した直後は選択状態を維持してインデックスを再マップ
        const rowCount = this.tableData.rows.length;
        const rowsChanged = rowCount !== this.prevRowCount;
        this.prevRowCount = rowCount;
        if (rowsChanged) {
            this.scrollEl.scrollTop = 0;
            if (this.selfFilterApplied) {
                // 自分で applyJsonFilter した直後：selectedValues を元にインデックスを再マップ
                this.rebuildSelectionFromValues();
            } else {
                // 外部スライサー等によるデータ変化：選択状態をクリア
                this.selectedOrigIdx.clear();
                this.selectedValues.clear();
            }
            this.selfFilterApplied = false; // rowsChanged 時にだけリセット（早期リセット防止）
        }

        if (!this.filterPanel.querySelector(".value-input:focus")) {
            // ユーザーが操作済みかつ列構成が変わっていない場合は状態を上書きしない
            if (!this.hasInteracted || colsChanged) this.restoreState(dv);
            this.renderFilterPanel();
        }

        this.renderColToggleBar();
        this.runFilter();
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
    }

    private extractTableData(dv: DataView): TableData {
        if (!dv?.table) return { columns: [], rows: [] };
        return {
            columns: dv.table.columns.map(c => c.displayName || ""),
            rows:    dv.table.rows.map(r => r.map(c => (c == null) ? "" : String(c))),
        };
    }

    private buildSelectionIds(dv: DataView): void {
        this.selectionIds = [];
        if (!dv?.table) return;
        for (let i = 0; i < dv.table.rows.length; i++) {
            this.selectionIds.push(
                this.host.createSelectionIdBuilder().withTable(dv.table, i).createSelectionId()
            );
        }
    }

    // ==========================================================
    // フィルターパネル
    // ==========================================================
    private renderFilterPanel(): void {
        this.clear(this.filterPanel);

        const hdr = this.el("div", "filter-header");
        const ttl = this.el("span", "filter-title");
        ttl.textContent = "フィルター";
        hdr.appendChild(ttl);

        const tog = this.el("div", "logic-toggle");
        for (const v of ["AND", "OR"] as const) {
            const b = this.el("button", "logic-btn" + (this.logic === v ? " active" : ""));
            b.textContent = v;
            b.onclick = () => { this.logic = v; this.saveState(); this.renderFilterPanel(); };
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
            this.saveState(); this.renderFilterPanel();
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
        colSel.onchange = () => { this.conditions[idx].columnIndex = +colSel.value; this.saveState(); };

        const opSel = this.el("select", "op-select");
        for (const { v, l } of [{ v: "contains", l: "を含む" }, { v: "notContains", l: "を含まない" }]) {
            const o = this.el("option", ""); o.value = v; o.textContent = l;
            if (v === cond.operator) o.selected = true;
            opSel.appendChild(o);
        }
        opSel.onchange = () => {
            this.conditions[idx].operator = opSel.value as "contains" | "notContains";
            this.saveState();
        };

        const inp = this.el("input", "value-input") as HTMLInputElement;
        inp.type = "text"; inp.placeholder = "検索キーワード"; inp.value = cond.value;
        inp.oninput   = () => { this.conditions[idx].value = inp.value; this.debounceSave(); };
        inp.onkeydown = (e: KeyboardEvent) => { if (e.key === "Enter") this.executeSearch(); };

        const del = this.el("button", "remove-btn"); del.textContent = "×";
        del.onclick = () => { this.conditions.splice(idx, 1); this.saveState(); this.renderFilterPanel(); };

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

        // 「全列」チップ
        const allChip = this.el("button", "col-chip" + (this.activeColTab === -1 ? " active" : ""));
        allChip.textContent = "全列";
        allChip.onclick = () => {
            this.activeColTab = -1;
            this.renderColToggleBar();
            this.renderTableHeader();
            this.renderVirtualRows();
        };
        this.colToggleBar.appendChild(allChip);

        // 各列チップ（クリックでその列だけ表示）
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
        // テキストフィルター変更時は選択状態をリセット（データセットフィルターとの矛盾防止）
        this.selectedOrigIdx.clear();
        this.selectedValues.clear();
        this.selectionManager.clear();
        this.host.applyJsonFilter(null, "general", "filter", FilterAction.remove);
        this.runFilter();
        this.renderTableHeader();
        this.scrollEl.scrollTop = 0;
        this.renderVirtualRows();
        this.persist(); this.renderFilterPanel(); this.renderStatus();
    }

    private clearFilter(): void {
        this.appliedConditions = []; this.appliedLogic = "AND";
        // フィルター解除時も選択をリセット
        this.selectedOrigIdx.clear();
        this.selectedValues.clear();
        this.selectionManager.clear();
        this.host.applyJsonFilter(null, "general", "filter", FilterAction.remove);
        this.runFilter();
        this.renderTableHeader();
        this.scrollEl.scrollTop = 0;
        this.renderVirtualRows();
        this.persist(); this.renderFilterPanel(); this.renderStatus();
    }

    private runFilter(): void {
        const active = this.appliedConditions.filter(c => c.value.trim() !== "");
        this.filteredRows = []; this.filteredOrigIdx = [];
        this.tableData.rows.forEach((row, oi) => {
            if (!active.length) { this.filteredRows.push(row); this.filteredOrigIdx.push(oi); return; }
            const res = active.map(c => {
                const cell = (row[c.columnIndex] ?? "").toLowerCase();
                const kw   = c.value.toLowerCase();
                return c.operator === "contains" ? cell.includes(kw) : !cell.includes(kw);
            });
            if (this.appliedLogic === "AND" ? res.every(Boolean) : res.some(Boolean)) {
                this.filteredRows.push(row); this.filteredOrigIdx.push(oi);
            }
        });
    }

    // ==========================================================
    // テーブル描画（DOM仮想スクロール）
    // ==========================================================
    private renderTableHeader(): void {
        this.clear(this.colGroup);
        this.clear(this.thead);
        if (!this.tableData.columns.length) return;

        // colgroup（チェックボックス列 + データ列）
        const cbCol = this.el("col", ""); cbCol.style.width = "32px";
        this.colGroup.appendChild(cbCol);
        this.tableData.columns.forEach((_, i) => {
            if (this.isColVisible(i)) this.colGroup.appendChild(this.el("col", ""));
        });

        // thead 行
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

        const startRow = Math.max(0, Math.floor(scrollTop / ROW_H) - BUFFER);
        const endRow   = Math.min(total, startRow + Math.ceil(viewH / ROW_H) + BUFFER * 2);
        const beforeH  = startRow * ROW_H;
        const afterH   = Math.max(0, (total - endRow) * ROW_H);
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
        // 同ページクロスフィルター（ハイライト）
        const ids = Array.from(this.selectedOrigIdx).map(i => this.selectionIds[i]).filter(Boolean);
        ids.length ? this.selectionManager.select(ids) : this.selectionManager.clear();

        // データセットレベルフィルター（スライサー同期に必要）
        this.applyDatasetFilter();

        this.updateSelectionUI();
        this.renderStatus();
    }

    // 選択値でデータセットフィルターを適用（スライサーの同期パネル対応）
    private applyDatasetFilter(): void {
        const dv = this.lastDataView;
        if (!dv?.table?.columns?.length) return;

        // フィルターキー列：アクティブなタブ列、なければ先頭列
        const colIdx = this.activeColTab >= 0 ? this.activeColTab : 0;
        const col    = dv.table.columns[colIdx];
        if (!col?.queryName) return;

        // 選択値を更新
        this.selectedValues.clear();
        this.selectedOrigIdx.forEach(i => {
            const v = this.tableData.rows[i]?.[colIdx];
            if (v != null && v !== "") this.selectedValues.add(v);
        });

        if (this.selectedValues.size === 0) {
            this.host.applyJsonFilter(null, "general", "filter", FilterAction.remove);
            return;
        }

        // queryName は "テーブル名.列名" 形式
        const parts  = col.queryName.split(".");
        const target: IFilterColumnTarget = {
            table:  parts[0],
            column: parts.slice(1).join(".") || col.displayName,
        };
        const values = Array.from(this.selectedValues);
        const filter = new BasicFilter(target, "In", values);

        this.selfFilterApplied = true;
        this.host.applyJsonFilter(filter.toJSON(), "general", "filter", FilterAction.merge);
    }

    // フィルター後に新しいデータで選択インデックスを再マップ
    private rebuildSelectionFromValues(): void {
        if (this.selectedValues.size === 0) { this.selectedOrigIdx.clear(); return; }
        const colIdx = this.activeColTab >= 0 ? this.activeColTab : 0;
        this.selectedOrigIdx.clear();
        this.tableData.rows.forEach((row, i) => {
            if (this.selectedValues.has(row[colIdx] ?? "")) this.selectedOrigIdx.add(i);
        });
    }

    // 選択変更時：全行再生成せず、可視行のチェックボックスだけ更新
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
        this.statusBar.appendChild(document.createTextNode(f === t ? `${t} 件` : `${f} / ${t} 件`));
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
    // Persist
    // ==========================================================
    private saveState(): void { this.persist(); }

    private debounceSave(): void {
        if (this.persistTimer !== null) clearTimeout(this.persistTimer);
        this.persistTimer = window.setTimeout(() => { this.persistTimer = null; this.persist(); }, 800);
    }

    private persist(): void {
        this.skipRender = true;
        this.hasInteracted = true;
        this.host.persistProperties({ merge: [{ objectName: "filterState", selector: null, properties: {
            conditions: JSON.stringify(this.conditions), logic: this.logic,
            applied: JSON.stringify(this.appliedConditions), appliedLogic: this.appliedLogic,
        }}]});
    }

    public getFormattingModel(): powerbi.visuals.FormattingModel {
        return this.formattingSettingsService.buildFormattingModel(this.formattingSettings);
    }
}
