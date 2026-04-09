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
    // 実行ボタンを押した時点の条件スナップショット
    private appliedConditions: FilterCondition[] = [];
    private appliedLogic: "AND" | "OR" = "AND";

    private filterPanel: HTMLElement;
    private tableContainer: HTMLElement;
    private statusBar: HTMLElement;

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

        this.tableContainer = document.createElement("div");
        this.tableContainer.className = "table-container";

        this.statusBar = document.createElement("div");
        this.statusBar.className = "status-bar";

        this.target.appendChild(this.filterPanel);
        this.target.appendChild(this.statusBar);
        this.target.appendChild(this.tableContainer);
    }

    public update(options: VisualUpdateOptions): void {
        this.formattingSettings = this.formattingSettingsService
            .populateFormattingSettingsModel(VisualFormattingSettingsModel, options.dataViews?.[0]);

        const dataView: DataView = options.dataViews?.[0];
        this.tableData = this.extractTableData(dataView);

        // 入力中はパネルを再構築しない（フォーカス維持）
        const isTyping = this.filterPanel.querySelector("input:focus") !== null;
        if (!isTyping) {
            const savedConditions = dataView?.metadata?.objects
                ?.["filterState"]?.["conditions"] as string ?? "";
            const savedLogic = dataView?.metadata?.objects
                ?.["filterState"]?.["logic"] as string ?? "AND";
            const savedApplied = dataView?.metadata?.objects
                ?.["filterState"]?.["applied"] as string ?? "";
            const savedAppliedLogic = dataView?.metadata?.objects
                ?.["filterState"]?.["appliedLogic"] as string ?? "AND";

            try {
                this.conditions = savedConditions ? JSON.parse(savedConditions) : [];
            } catch {
                this.conditions = [];
            }
            this.logic = savedLogic === "OR" ? "OR" : "AND";

            try {
                this.appliedConditions = savedApplied ? JSON.parse(savedApplied) : [];
            } catch {
                this.appliedConditions = [];
            }
            this.appliedLogic = savedAppliedLogic === "OR" ? "OR" : "AND";

            this.renderFilterPanel();
        }

        this.renderTable();
    }

    private extractTableData(dataView: DataView): TableData {
        if (!dataView?.table) return { columns: [], rows: [] };

        const columns = dataView.table.columns.map(c => c.displayName || "");
        const rows = dataView.table.rows.map(row =>
            row.map(cell => (cell === null || cell === undefined) ? "" : String(cell))
        );
        return { columns, rows };
    }

    private clearElement(el: HTMLElement): void {
        while (el.firstChild) el.removeChild(el.firstChild);
    }

    private renderFilterPanel(): void {
        this.clearElement(this.filterPanel);

        // ヘッダー行（タイトル + AND/OR）
        const header = document.createElement("div");
        header.className = "filter-header";

        const title = document.createElement("span");
        title.className = "filter-title";
        title.textContent = "フィルター";
        header.appendChild(title);

        const logicToggle = document.createElement("div");
        logicToggle.className = "logic-toggle";

        const andBtn = document.createElement("button");
        andBtn.textContent = "AND";
        andBtn.className = "logic-btn" + (this.logic === "AND" ? " active" : "");
        andBtn.onclick = () => { this.logic = "AND"; this.saveState(); this.renderFilterPanel(); };

        const orBtn = document.createElement("button");
        orBtn.textContent = "OR";
        orBtn.className = "logic-btn" + (this.logic === "OR" ? " active" : "");
        orBtn.onclick = () => { this.logic = "OR"; this.saveState(); this.renderFilterPanel(); };

        logicToggle.appendChild(andBtn);
        logicToggle.appendChild(orBtn);
        header.appendChild(logicToggle);
        this.filterPanel.appendChild(header);

        // 条件リスト
        const conditionList = document.createElement("div");
        conditionList.className = "condition-list";
        this.conditions.forEach((cond, idx) => {
            conditionList.appendChild(this.createConditionRow(cond, idx));
        });
        this.filterPanel.appendChild(conditionList);

        // フッター行（条件追加 + 実行ボタン）
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
        [
            { value: "contains", label: "を含む" },
            { value: "notContains", label: "を含まない" },
        ].forEach(op => {
            const opt = document.createElement("option");
            opt.value = op.value;
            opt.textContent = op.label;
            if (op.value === cond.operator) opt.selected = true;
            opSelect.appendChild(opt);
        });
        opSelect.onchange = () => {
            this.conditions[idx].operator = opSelect.value as "contains" | "notContains";
            this.saveState();
        };

        const valueInput = document.createElement("input");
        valueInput.type = "text";
        valueInput.className = "value-input";
        valueInput.placeholder = "検索キーワード";
        valueInput.value = cond.value;
        // 入力中は条件をメモリ更新＋デバウンス保存のみ。テーブルは触らない
        valueInput.oninput = () => {
            this.conditions[idx].value = valueInput.value;
            this.debounceSave();
        };
        // Enter キーで実行
        valueInput.onkeydown = (e: KeyboardEvent) => {
            if (e.key === "Enter") this.executeSearch();
        };

        const removeBtn = document.createElement("button");
        removeBtn.className = "remove-btn";
        removeBtn.textContent = "×";
        removeBtn.title = "条件を削除";
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

    // 実行ボタン or Enter → 現在の conditions を applied にコピーしてテーブル更新
    private executeSearch(): void {
        this.appliedConditions = this.conditions.map(c => ({ ...c }));
        this.appliedLogic = this.logic;
        this.persist();
        this.renderTable();
    }

    // 条件の構造変更（追加・削除・列・演算子・AND/OR）を即座に保存
    private saveState(): void {
        this.persist();
    }

    // テキスト入力は 800ms デバウンスで保存
    private debounceSave(): void {
        if (this.persistTimer !== null) clearTimeout(this.persistTimer);
        this.persistTimer = window.setTimeout(() => {
            this.persistTimer = null;
            this.persist();
        }, 800);
    }

    private persist(): void {
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

    private applyFilters(rows: string[][]): string[][] {
        const active = this.appliedConditions.filter(c => c.value.trim() !== "");
        if (active.length === 0) return rows;

        return rows.filter(row => {
            const results = active.map(cond => {
                const cell = (row[cond.columnIndex] ?? "").toLowerCase();
                const kw = cond.value.toLowerCase();
                return cond.operator === "contains" ? cell.includes(kw) : !cell.includes(kw);
            });
            return this.appliedLogic === "AND" ? results.every(Boolean) : results.some(Boolean);
        });
    }

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

        const filteredRows = this.applyFilters(this.tableData.rows);
        this.statusBar.textContent = `${filteredRows.length} / ${this.tableData.rows.length} 件`;

        const table = document.createElement("table");
        table.className = "data-table";

        const thead = document.createElement("thead");
        const headerRow = document.createElement("tr");
        this.tableData.columns.forEach(col => {
            const th = document.createElement("th");
            th.textContent = col;
            headerRow.appendChild(th);
        });
        thead.appendChild(headerRow);
        table.appendChild(thead);

        const tbody = document.createElement("tbody");
        filteredRows.forEach((row, rowIdx) => {
            const tr = document.createElement("tr");
            tr.className = rowIdx % 2 === 0 ? "row-even" : "row-odd";
            row.forEach(cell => {
                const td = document.createElement("td");
                td.textContent = cell;
                tr.appendChild(td);
            });
            tbody.appendChild(tr);
        });
        table.appendChild(tbody);
        this.tableContainer.appendChild(table);
    }

    public getFormattingModel(): powerbi.visuals.FormattingModel {
        return this.formattingSettingsService.buildFormattingModel(this.formattingSettings);
    }
}
