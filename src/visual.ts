"use strict";

import powerbi from "powerbi-visuals-api";
import { FormattingSettingsService } from "powerbi-visuals-utils-formattingmodel";
import "./../style/visual.less";

import VisualConstructorOptions = powerbi.extensibility.visual.VisualConstructorOptions;
import VisualUpdateOptions = powerbi.extensibility.visual.VisualUpdateOptions;
import IVisual = powerbi.extensibility.visual.IVisual;
import IVisualHost = powerbi.extensibility.visual.IVisualHost;
import DataView = powerbi.DataView;

import { VisualFormattingSettingsModel } from "./settings";

const PAGE_SIZE = 100;

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
    private target: HTMLElement;
    private formattingSettings: VisualFormattingSettingsModel;
    private formattingSettingsService: FormattingSettingsService;

    private conditions: FilterCondition[] = [];
    private logic: "AND" | "OR" = "AND";
    private tableData: TableData = { columns: [], rows: [] };
    private appliedConditions: FilterCondition[] = [];
    private appliedLogic: "AND" | "OR" = "AND";
    private filteredRows: string[][] = [];
    private currentPage = 0;

    private filterPanel: HTMLElement;
    private tableContainer: HTMLElement;
    private statusBar: HTMLElement;
    private pager: HTMLElement;

    // 自分の persist 呼び出しによる update サイクルでテーブルを再描画しないフラグ
    private skipTableRender = false;
    private persistTimer: number | null = null;

    constructor(options: VisualConstructorOptions) {
        this.host = options.host;
        this.formattingSettingsService = new FormattingSettingsService();
        this.target = options.element;
        this.target.className = "filter-table-visual";
        this.buildLayout();
    }

    private buildLayout(): void {
        this.filterPanel = document.createElement("div");
        this.filterPanel.className = "filter-panel";

        this.statusBar = document.createElement("div");
        this.statusBar.className = "status-bar";

        this.tableContainer = document.createElement("div");
        this.tableContainer.className = "table-container";

        this.pager = document.createElement("div");
        this.pager.className = "pager";

        this.target.appendChild(this.filterPanel);
        this.target.appendChild(this.statusBar);
        this.target.appendChild(this.tableContainer);
        this.target.appendChild(this.pager);
    }

    public update(options: VisualUpdateOptions): void {
        this.formattingSettings = this.formattingSettingsService
            .populateFormattingSettingsModel(VisualFormattingSettingsModel, options.dataViews?.[0]);

        const dataView: DataView = options.dataViews?.[0];

        // 自分の persist による update は無視（テーブル再描画なし）
        if (this.skipTableRender) {
            this.skipTableRender = false;
            // パネルだけ再描画（データは変わっていないので tableData は更新不要）
            const isTyping = this.filterPanel.querySelector("input:focus") !== null;
            if (!isTyping) this.renderFilterPanel();
            return;
        }

        this.tableData = this.extractTableData(dataView);

        const isTyping = this.filterPanel.querySelector("input:focus") !== null;
        if (!isTyping) {
            this.restoreState(dataView);
            this.renderFilterPanel();
        }

        this.filteredRows = this.applyFilters();
        this.currentPage = 0;
        this.renderTable();
        this.renderPager();
    }

    private restoreState(dataView: DataView): void {
        const meta = dataView?.metadata?.objects?.["filterState"];
        try { this.conditions = meta?.["conditions"] ? JSON.parse(meta["conditions"] as string) : []; }
        catch { this.conditions = []; }
        this.logic = (meta?.["logic"] as string) === "OR" ? "OR" : "AND";
        try { this.appliedConditions = meta?.["applied"] ? JSON.parse(meta["applied"] as string) : []; }
        catch { this.appliedConditions = []; }
        this.appliedLogic = (meta?.["appliedLogic"] as string) === "OR" ? "OR" : "AND";
    }

    private extractTableData(dataView: DataView): TableData {
        if (!dataView?.table) return { columns: [], rows: [] };
        const columns = dataView.table.columns.map(c => c.displayName || "");
        const rows = dataView.table.rows.map(row =>
            row.map(cell => (cell == null) ? "" : String(cell))
        );
        return { columns, rows };
    }

    private clearElement(el: HTMLElement): void {
        while (el.firstChild) el.removeChild(el.firstChild);
    }

    // ---- フィルターパネル ----

    private renderFilterPanel(): void {
        this.clearElement(this.filterPanel);

        const header = document.createElement("div");
        header.className = "filter-header";

        const title = document.createElement("span");
        title.className = "filter-title";
        title.textContent = "フィルター";
        header.appendChild(title);

        const logicToggle = document.createElement("div");
        logicToggle.className = "logic-toggle";
        for (const val of ["AND", "OR"] as const) {
            const btn = document.createElement("button");
            btn.textContent = val;
            btn.className = "logic-btn" + (this.logic === val ? " active" : "");
            btn.onclick = () => { this.logic = val; this.saveState(); this.renderFilterPanel(); };
            logicToggle.appendChild(btn);
        }
        header.appendChild(logicToggle);
        this.filterPanel.appendChild(header);

        const conditionList = document.createElement("div");
        conditionList.className = "condition-list";
        this.conditions.forEach((cond, idx) => {
            conditionList.appendChild(this.createConditionRow(cond, idx));
        });
        this.filterPanel.appendChild(conditionList);

        const footer = document.createElement("div");
        footer.className = "filter-footer";

        const addBtn = document.createElement("button");
        addBtn.className = "add-condition-btn";
        addBtn.textContent = "+ 条件を追加";
        addBtn.onclick = () => {
            this.conditions.push({ columnIndex: 0, operator: "contains", value: "" });
            this.saveState();
            this.renderFilterPanel();
        };

        const runBtn = document.createElement("button");
        runBtn.className = "run-btn";
        runBtn.textContent = "実行";
        runBtn.onclick = () => this.executeSearch();

        footer.appendChild(addBtn);
        footer.appendChild(runBtn);
        this.filterPanel.appendChild(footer);
    }

    private createConditionRow(cond: FilterCondition, idx: number): HTMLElement {
        const row = document.createElement("div");
        row.className = "condition-row";

        const colSelect = document.createElement("select");
        colSelect.className = "col-select";
        this.tableData.columns.forEach((col, i) => {
            const opt = document.createElement("option");
            opt.value = String(i);
            opt.textContent = col;
            if (i === cond.columnIndex) opt.selected = true;
            colSelect.appendChild(opt);
        });
        colSelect.onchange = () => {
            this.conditions[idx].columnIndex = parseInt(colSelect.value, 10);
            this.saveState();
        };

        const opSelect = document.createElement("select");
        opSelect.className = "op-select";
        for (const op of [{ value: "contains", label: "を含む" }, { value: "notContains", label: "を含まない" }]) {
            const opt = document.createElement("option");
            opt.value = op.value;
            opt.textContent = op.label;
            if (op.value === cond.operator) opt.selected = true;
            opSelect.appendChild(opt);
        }
        opSelect.onchange = () => {
            this.conditions[idx].operator = opSelect.value as "contains" | "notContains";
            this.saveState();
        };

        const valueInput = document.createElement("input");
        valueInput.type = "text";
        valueInput.className = "value-input";
        valueInput.placeholder = "検索キーワード";
        valueInput.value = cond.value;
        valueInput.oninput = () => {
            this.conditions[idx].value = valueInput.value;
            this.debounceSave();
        };
        valueInput.onkeydown = (e: KeyboardEvent) => {
            if (e.key === "Enter") this.executeSearch();
        };

        const removeBtn = document.createElement("button");
        removeBtn.className = "remove-btn";
        removeBtn.textContent = "×";
        removeBtn.onclick = () => {
            this.conditions.splice(idx, 1);
            this.saveState();
            this.renderFilterPanel();
        };

        row.appendChild(colSelect);
        row.appendChild(opSelect);
        row.appendChild(valueInput);
        row.appendChild(removeBtn);
        return row;
    }

    // ---- 実行 ----

    private executeSearch(): void {
        this.appliedConditions = this.conditions.map(c => ({ ...c }));
        this.appliedLogic = this.logic;
        this.filteredRows = this.applyFilters();
        this.currentPage = 0;
        this.persist();          // skipTableRender = true にした上で persist
        this.renderTable();      // 実行後だけテーブルを更新
        this.renderPager();
    }

    // ---- 保存 ----

    private saveState(): void {
        this.persist();
    }

    private debounceSave(): void {
        if (this.persistTimer !== null) clearTimeout(this.persistTimer);
        this.persistTimer = window.setTimeout(() => {
            this.persistTimer = null;
            this.persist();
        }, 800);
    }

    private persist(): void {
        this.skipTableRender = true;  // 次の update() でテーブル再描画をスキップ
        this.host.persistProperties({
            merge: [{
                objectName: "filterState",
                selector: null,
                properties: {
                    conditions: JSON.stringify(this.conditions),
                    logic: this.logic,
                    applied: JSON.stringify(this.appliedConditions),
                    appliedLogic: this.appliedLogic,
                },
            }],
        });
    }

    // ---- フィルター処理 ----

    private applyFilters(): string[][] {
        const active = this.appliedConditions.filter(c => c.value.trim() !== "");
        if (active.length === 0) return this.tableData.rows;

        return this.tableData.rows.filter(row => {
            const results = active.map(cond => {
                const cell = (row[cond.columnIndex] ?? "").toLowerCase();
                const kw = cond.value.toLowerCase();
                return cond.operator === "contains" ? cell.includes(kw) : !cell.includes(kw);
            });
            return this.appliedLogic === "AND" ? results.every(Boolean) : results.some(Boolean);
        });
    }

    // ---- テーブル描画（現在ページのみ） ----

    private renderTable(): void {
        this.clearElement(this.tableContainer);

        if (this.tableData.columns.length === 0) {
            const msg = document.createElement("div");
            msg.className = "no-data";
            msg.textContent = "データをフィールドに追加してください";
            this.tableContainer.appendChild(msg);
            this.statusBar.textContent = "";
            return;
        }

        const total = this.tableData.rows.length;
        const filtered = this.filteredRows.length;
        this.statusBar.textContent = filtered === total
            ? `${total} 件`
            : `${filtered} / ${total} 件`;

        const table = document.createElement("table");
        table.className = "data-table";

        // ヘッダー
        const thead = document.createElement("thead");
        const headerRow = document.createElement("tr");
        const hFrag = document.createDocumentFragment();
        this.tableData.columns.forEach(col => {
            const th = document.createElement("th");
            th.textContent = col;
            hFrag.appendChild(th);
        });
        headerRow.appendChild(hFrag);
        thead.appendChild(headerRow);
        table.appendChild(thead);

        // ボディ：現在ページ分だけ
        const start = this.currentPage * PAGE_SIZE;
        const pageRows = this.filteredRows.slice(start, start + PAGE_SIZE);

        const tbody = document.createElement("tbody");
        const bFrag = document.createDocumentFragment();
        pageRows.forEach((row, i) => {
            const tr = document.createElement("tr");
            tr.className = (start + i) % 2 === 0 ? "row-even" : "row-odd";
            const rFrag = document.createDocumentFragment();
            row.forEach(cell => {
                const td = document.createElement("td");
                td.textContent = cell;
                rFrag.appendChild(td);
            });
            tr.appendChild(rFrag);
            bFrag.appendChild(tr);
        });
        tbody.appendChild(bFrag);
        table.appendChild(tbody);
        this.tableContainer.appendChild(table);
    }

    // ---- ページネーション ----

    private renderPager(): void {
        this.clearElement(this.pager);
        const totalPages = Math.ceil(this.filteredRows.length / PAGE_SIZE);
        if (totalPages <= 1) return;

        const info = document.createElement("span");
        info.className = "pager-info";
        info.textContent = `${this.currentPage + 1} / ${totalPages} ページ`;

        const prev = document.createElement("button");
        prev.className = "pager-btn";
        prev.textContent = "‹";
        prev.disabled = this.currentPage === 0;
        prev.onclick = () => { this.currentPage--; this.renderTable(); this.renderPager(); };

        const next = document.createElement("button");
        next.className = "pager-btn";
        next.textContent = "›";
        next.disabled = this.currentPage >= totalPages - 1;
        next.onclick = () => { this.currentPage++; this.renderTable(); this.renderPager(); };

        const frag = document.createDocumentFragment();
        frag.appendChild(prev);
        frag.appendChild(info);
        frag.appendChild(next);
        this.pager.appendChild(frag);
    }

    public getFormattingModel(): powerbi.visuals.FormattingModel {
        return this.formattingSettingsService.buildFormattingModel(this.formattingSettings);
    }
}
