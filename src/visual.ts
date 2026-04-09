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

        // 入力中（フォーカスあり）は条件を上書きしない → フォーカス・スクロール位置を維持
        const isTyping = this.filterPanel.querySelector("input:focus") !== null;
        if (!isTyping) {
            const savedConditions = dataView?.metadata?.objects
                ?.["filterState"]?.["conditions"] as string ?? "";
            const savedLogic = dataView?.metadata?.objects
                ?.["filterState"]?.["logic"] as string ?? "AND";

            if (savedConditions) {
                try {
                    this.conditions = JSON.parse(savedConditions);
                } catch {
                    this.conditions = [];
                }
            } else {
                this.conditions = [];
            }
            this.logic = (savedLogic === "OR") ? "OR" : "AND";
            this.renderFilterPanel();
        }

        this.renderTable();
    }

    private extractTableData(dataView: DataView): TableData {
        if (!dataView?.table) return { columns: [], rows: [] };

        const columns = dataView.table.columns.map(c => c.displayName || "");
        const rows = dataView.table.rows.map(row =>
            row.map(cell => {
                if (cell === null || cell === undefined) return "";
                return String(cell);
            })
        );
        return { columns, rows };
    }

    private clearElement(el: HTMLElement): void {
        while (el.firstChild) el.removeChild(el.firstChild);
    }

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

        const andBtn = document.createElement("button");
        andBtn.textContent = "AND";
        andBtn.className = "logic-btn" + (this.logic === "AND" ? " active" : "");
        andBtn.onclick = () => this.setLogic("AND");

        const orBtn = document.createElement("button");
        orBtn.textContent = "OR";
        orBtn.className = "logic-btn" + (this.logic === "OR" ? " active" : "");
        orBtn.onclick = () => this.setLogic("OR");

        logicToggle.appendChild(andBtn);
        logicToggle.appendChild(orBtn);
        header.appendChild(logicToggle);
        this.filterPanel.appendChild(header);

        const conditionList = document.createElement("div");
        conditionList.className = "condition-list";

        this.conditions.forEach((cond, idx) => {
            const row = this.createConditionRow(cond, idx);
            conditionList.appendChild(row);
        });

        this.filterPanel.appendChild(conditionList);

        const addBtn = document.createElement("button");
        addBtn.className = "add-condition-btn";
        addBtn.textContent = "+ 条件を追加";
        addBtn.onclick = () => this.addCondition();
        this.filterPanel.appendChild(addBtn);
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
            this.saveAndRender();
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
            this.saveAndRender();
        };

        const valueInput = document.createElement("input");
        valueInput.type = "text";
        valueInput.className = "value-input";
        valueInput.placeholder = "検索キーワード";
        valueInput.value = cond.value;
        valueInput.oninput = () => this.onTextInput(idx, valueInput.value);

        const removeBtn = document.createElement("button");
        removeBtn.className = "remove-btn";
        removeBtn.textContent = "×";
        removeBtn.title = "条件を削除";
        removeBtn.onclick = () => {
            this.conditions.splice(idx, 1);
            this.saveAndRender();
        };

        row.appendChild(colSelect);
        row.appendChild(opSelect);
        row.appendChild(valueInput);
        row.appendChild(removeBtn);
        return row;
    }

    private addCondition(): void {
        this.conditions.push({
            columnIndex: 0,
            operator: "contains",
            value: "",
        });
        this.saveAndRender();
    }

    private setLogic(logic: "AND" | "OR"): void {
        this.logic = logic;
        this.saveAndRender();
    }

    private persist(): void {
        this.host.persistProperties({
            merge: [{
                objectName: "filterState",
                selector: null,
                properties: {
                    conditions: JSON.stringify(this.conditions),
                    logic: this.logic,
                },
            }],
        });
    }

    // 列変更・演算子変更・条件追加削除・AND/OR切り替えはすぐ保存＋全体再描画
    private saveAndRender(): void {
        this.persist();
        this.renderFilterPanel();
        this.renderTable();
    }

    // テキスト入力はテーブルだけ即時更新、保存は 800ms デバウンス
    private onTextInput(idx: number, value: string): void {
        this.conditions[idx].value = value;
        this.renderTable();
        if (this.persistTimer !== null) clearTimeout(this.persistTimer);
        this.persistTimer = window.setTimeout(() => {
            this.persistTimer = null;
            this.persist();
        }, 800);
    }

    private applyFilters(rows: string[][]): string[][] {
        const activeConditions = this.conditions.filter(c => c.value.trim() !== "");
        if (activeConditions.length === 0) return rows;

        return rows.filter(row => {
            const results = activeConditions.map(cond => {
                const cellValue = (row[cond.columnIndex] ?? "").toLowerCase();
                const keyword = cond.value.toLowerCase();
                if (cond.operator === "contains") {
                    return cellValue.includes(keyword);
                } else {
                    return !cellValue.includes(keyword);
                }
            });

            if (this.logic === "AND") {
                return results.every(Boolean);
            } else {
                return results.some(Boolean);
            }
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

        this.statusBar.textContent =
            `${filteredRows.length} / ${this.tableData.rows.length} 件`;

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
