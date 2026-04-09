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

// --- Canvas 定数 ---
const ROW_H      = 22;
const HEADER_H   = 28;
const SB_W       = 10;   // スクロールバー幅
const CELL_PAD   = 8;
const FONT       = "12px 'Segoe UI',sans-serif";
const FONT_HDR   = "600 12px 'Segoe UI',sans-serif";
const C_HDR_BG   = "#f2f2f2";
const C_ROW_EVEN = "#ffffff";
const C_ROW_ODD  = "#fafafa";
const C_ROW_HOV  = "#e8f4fd";
const C_TEXT     = "#252423";
const C_BORDER   = "#e0e0e0";
const C_HDR_LINE = "#d1d1d1";
const C_SB_TRACK = "#f0f0f0";
const C_SB_THUMB = "#c8c8c8";

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

    private conditions: FilterCondition[]        = [];
    private logic: "AND" | "OR"                  = "AND";
    private tableData: TableData                 = { columns: [], rows: [] };
    private appliedConditions: FilterCondition[] = [];
    private appliedLogic: "AND" | "OR"           = "AND";
    private filteredRows: string[][]             = [];

    // --- DOM ---
    private filterPanel:   HTMLElement;
    private statusBar:     HTMLElement;
    private canvasArea:    HTMLElement;
    private canvas:        HTMLCanvasElement;
    private ctx:           CanvasRenderingContext2D;

    // --- Canvas 状態 ---
    private scrollTop    = 0;
    private hoveredRow   = -1;
    private colWidths:     number[] = [];
    private logicalW     = 0;   // CSS px
    private logicalH     = 0;
    private drawPending  = false;

    // --- スクロールバードラッグ ---
    private isDragging   = false;
    private dragStartY   = 0;
    private dragStartScroll = 0;

    // --- 自分の persist による update を無視するフラグ ---
    private skipRender   = false;
    private persistTimer: number | null = null;

    constructor(options: VisualConstructorOptions) {
        this.host = options.host;
        this.formattingSettingsService = new FormattingSettingsService();
        this.target = options.element;
        this.target.className = "filter-table-visual";
        this.buildLayout();
        this.setupCanvasEvents();
    }

    private buildLayout(): void {
        this.filterPanel = document.createElement("div");
        this.filterPanel.className = "filter-panel";

        this.statusBar = document.createElement("div");
        this.statusBar.className = "status-bar";

        this.canvasArea = document.createElement("div");
        this.canvasArea.className = "canvas-area";

        this.canvas = document.createElement("canvas");
        this.canvasArea.appendChild(this.canvas);
        this.ctx = this.canvas.getContext("2d");

        this.target.appendChild(this.filterPanel);
        this.target.appendChild(this.statusBar);
        this.target.appendChild(this.canvasArea);
    }

    // =========================================================
    // update
    // =========================================================
    public update(options: VisualUpdateOptions): void {
        this.formattingSettings = this.formattingSettingsService
            .populateFormattingSettingsModel(VisualFormattingSettingsModel, options.dataViews?.[0]);

        const dataView: DataView = options.dataViews?.[0];

        if (this.skipRender) {
            this.skipRender = false;
            const isTyping = !!this.filterPanel.querySelector("input:focus");
            if (!isTyping) this.renderFilterPanel();
            return;
        }

        this.tableData = this.extractTableData(dataView);

        const isTyping = !!this.filterPanel.querySelector("input:focus");
        if (!isTyping) {
            this.restoreState(dataView);
            this.renderFilterPanel();
        }

        this.filteredRows = this.applyFilters();
        this.scrollTop = 0;
        this.updateStatus();
        this.resizeCanvas();   // canvas サイズを合わせてから描画
        this.scheduleRedraw();
    }

    private restoreState(dataView: DataView): void {
        const m = dataView?.metadata?.objects?.["filterState"];
        try   { this.conditions = m?.["conditions"] ? JSON.parse(m["conditions"] as string) : []; }
        catch { this.conditions = []; }
        this.logic = (m?.["logic"] as string) === "OR" ? "OR" : "AND";
        try   { this.appliedConditions = m?.["applied"] ? JSON.parse(m["applied"] as string) : []; }
        catch { this.appliedConditions = []; }
        this.appliedLogic = (m?.["appliedLogic"] as string) === "OR" ? "OR" : "AND";
    }

    private extractTableData(dataView: DataView): TableData {
        if (!dataView?.table) return { columns: [], rows: [] };
        const columns = dataView.table.columns.map(c => c.displayName || "");
        const rows = dataView.table.rows.map(row =>
            row.map(cell => (cell == null) ? "" : String(cell))
        );
        return { columns, rows };
    }

    // =========================================================
    // フィルターパネル（HTML DOM）
    // =========================================================
    private clearEl(el: HTMLElement): void {
        while (el.firstChild) el.removeChild(el.firstChild);
    }

    private renderFilterPanel(): void {
        this.clearEl(this.filterPanel);

        const header = document.createElement("div");
        header.className = "filter-header";

        const title = document.createElement("span");
        title.className = "filter-title";
        title.textContent = "フィルター";
        header.appendChild(title);

        const toggle = document.createElement("div");
        toggle.className = "logic-toggle";
        for (const v of ["AND", "OR"] as const) {
            const btn = document.createElement("button");
            btn.textContent = v;
            btn.className = "logic-btn" + (this.logic === v ? " active" : "");
            btn.onclick = () => { this.logic = v; this.saveState(); this.renderFilterPanel(); };
            toggle.appendChild(btn);
        }
        header.appendChild(toggle);
        this.filterPanel.appendChild(header);

        const list = document.createElement("div");
        list.className = "condition-list";
        this.conditions.forEach((c, i) => list.appendChild(this.makeConditionRow(c, i)));
        this.filterPanel.appendChild(list);

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

    private makeConditionRow(cond: FilterCondition, idx: number): HTMLElement {
        const row = document.createElement("div");
        row.className = "condition-row";

        const colSel = document.createElement("select");
        colSel.className = "col-select";
        this.tableData.columns.forEach((col, i) => {
            const o = document.createElement("option");
            o.value = String(i); o.textContent = col;
            if (i === cond.columnIndex) o.selected = true;
            colSel.appendChild(o);
        });
        colSel.onchange = () => { this.conditions[idx].columnIndex = +colSel.value; this.saveState(); };

        const opSel = document.createElement("select");
        opSel.className = "op-select";
        for (const op of [{ v: "contains", l: "を含む" }, { v: "notContains", l: "を含まない" }]) {
            const o = document.createElement("option");
            o.value = op.v; o.textContent = op.l;
            if (op.v === cond.operator) o.selected = true;
            opSel.appendChild(o);
        }
        opSel.onchange = () => { this.conditions[idx].operator = opSel.value as "contains" | "notContains"; this.saveState(); };

        const inp = document.createElement("input");
        inp.type = "text"; inp.className = "value-input";
        inp.placeholder = "検索キーワード"; inp.value = cond.value;
        inp.oninput    = () => { this.conditions[idx].value = inp.value; this.debounceSave(); };
        inp.onkeydown  = (e: KeyboardEvent) => { if (e.key === "Enter") this.executeSearch(); };

        const del = document.createElement("button");
        del.className = "remove-btn"; del.textContent = "×";
        del.onclick = () => { this.conditions.splice(idx, 1); this.saveState(); this.renderFilterPanel(); };

        row.appendChild(colSel); row.appendChild(opSel); row.appendChild(inp); row.appendChild(del);
        return row;
    }

    private executeSearch(): void {
        this.appliedConditions = this.conditions.map(c => ({ ...c }));
        this.appliedLogic = this.logic;
        this.filteredRows = this.applyFilters();
        this.scrollTop = 0;
        this.persist();
        this.updateStatus();
        this.scheduleRedraw();
    }

    private saveState(): void  { this.persist(); }

    private debounceSave(): void {
        if (this.persistTimer !== null) clearTimeout(this.persistTimer);
        this.persistTimer = window.setTimeout(() => { this.persistTimer = null; this.persist(); }, 800);
    }

    private persist(): void {
        this.skipRender = true;
        this.host.persistProperties({
            merge: [{ objectName: "filterState", selector: null, properties: {
                conditions:    JSON.stringify(this.conditions),
                logic:         this.logic,
                applied:       JSON.stringify(this.appliedConditions),
                appliedLogic:  this.appliedLogic,
            }}],
        });
    }

    private applyFilters(): string[][] {
        const active = this.appliedConditions.filter(c => c.value.trim() !== "");
        if (!active.length) return this.tableData.rows;
        return this.tableData.rows.filter(row => {
            const res = active.map(c => {
                const cell = (row[c.columnIndex] ?? "").toLowerCase();
                const kw = c.value.toLowerCase();
                return c.operator === "contains" ? cell.includes(kw) : !cell.includes(kw);
            });
            return this.appliedLogic === "AND" ? res.every(Boolean) : res.some(Boolean);
        });
    }

    private updateStatus(): void {
        const f = this.filteredRows.length, t = this.tableData.rows.length;
        this.statusBar.textContent = f === t ? `${t} 件` : `${f} / ${t} 件`;
    }

    // =========================================================
    // Canvas 描画
    // =========================================================
    private resizeCanvas(): void {
        const w = this.canvasArea.clientWidth;
        const h = this.canvasArea.clientHeight;
        if (w <= 0 || h <= 0) return;
        const dpr = window.devicePixelRatio || 1;
        this.canvas.width  = Math.round(w * dpr);
        this.canvas.height = Math.round(h * dpr);
        this.canvas.style.width  = w + "px";
        this.canvas.style.height = h + "px";
        this.logicalW = w;
        this.logicalH = h;
        this.ctx.setTransform(dpr, 0, 0, dpr, 0, 0);
        this.calcColWidths(w - SB_W);
    }

    private calcColWidths(tableW: number): void {
        const cols = this.tableData.columns;
        if (!cols.length) { this.colWidths = []; return; }

        const ctx = this.ctx;
        const min = 50, max = 280;

        ctx.font = FONT_HDR;
        const ws = cols.map(c => Math.min(max, Math.max(min, ctx.measureText(c).width + CELL_PAD * 2 + 12)));

        ctx.font = FONT;
        const sample = this.filteredRows.length > 300 ? this.filteredRows.slice(0, 300) : this.filteredRows;
        for (const row of sample) {
            for (let i = 0; i < row.length; i++) {
                const w = Math.min(max, ctx.measureText(row[i]).width + CELL_PAD * 2);
                if (w > ws[i]) ws[i] = w;
            }
        }

        // 余白があれば均等に広げる
        const total = ws.reduce((s, w) => s + w, 0);
        if (total < tableW && cols.length > 0) {
            const extra = (tableW - total) / cols.length;
            for (let i = 0; i < ws.length; i++) ws[i] += extra;
        }
        this.colWidths = ws;
    }

    private scheduleRedraw(): void {
        if (this.drawPending) return;
        this.drawPending = true;
        requestAnimationFrame(() => { this.drawPending = false; this.draw(); });
    }

    private maxScroll(): number {
        return Math.max(0, this.filteredRows.length * ROW_H - (this.logicalH - HEADER_H));
    }

    private draw(): void {
        const ctx = this.ctx;
        const W = this.logicalW, H = this.logicalH;
        if (W <= 0 || H <= 0) return;

        ctx.clearRect(0, 0, W, H);

        if (!this.tableData.columns.length) {
            ctx.fillStyle = "#a19f9d"; ctx.font = FONT;
            ctx.textAlign = "center"; ctx.textBaseline = "middle";
            ctx.fillText("データをフィールドに追加してください", W / 2, H / 2);
            return;
        }

        const tW = W - SB_W;   // テーブル描画幅

        this.drawHeader(ctx, tW);
        this.drawBody(ctx, tW);
        this.drawScrollbar(ctx, W, H);

        // ヘッダー下のボーダー
        ctx.strokeStyle = C_HDR_LINE; ctx.lineWidth = 1;
        ctx.beginPath(); ctx.moveTo(0, HEADER_H); ctx.lineTo(tW, HEADER_H); ctx.stroke();
    }

    private drawHeader(ctx: CanvasRenderingContext2D, tW: number): void {
        ctx.fillStyle = C_HDR_BG;
        ctx.fillRect(0, 0, tW, HEADER_H);

        ctx.font = FONT_HDR; ctx.fillStyle = C_TEXT;
        ctx.textBaseline = "middle"; ctx.textAlign = "left";

        let x = 0;
        this.tableData.columns.forEach((col, i) => {
            const cw = this.colWidths[i];
            ctx.save();
            ctx.beginPath(); ctx.rect(x + CELL_PAD, 0, cw - CELL_PAD * 2, HEADER_H); ctx.clip();
            ctx.fillText(col, x + CELL_PAD, HEADER_H / 2);
            ctx.restore();
            // 列区切り線
            ctx.strokeStyle = C_HDR_LINE; ctx.lineWidth = 1;
            ctx.beginPath(); ctx.moveTo(x + cw - 0.5, 2); ctx.lineTo(x + cw - 0.5, HEADER_H - 2); ctx.stroke();
            x += cw;
        });
    }

    private drawBody(ctx: CanvasRenderingContext2D, tW: number): void {
        const bodyH = this.logicalH - HEADER_H;
        const first = Math.floor(this.scrollTop / ROW_H);
        const last  = Math.min(this.filteredRows.length, first + Math.ceil(bodyH / ROW_H) + 1);

        ctx.save();
        ctx.beginPath(); ctx.rect(0, HEADER_H, tW, bodyH); ctx.clip();
        ctx.font = FONT; ctx.textBaseline = "middle"; ctx.textAlign = "left";

        for (let ri = first; ri < last; ri++) {
            const y = HEADER_H + ri * ROW_H - this.scrollTop;
            ctx.fillStyle = ri === this.hoveredRow ? C_ROW_HOV : (ri % 2 === 0 ? C_ROW_EVEN : C_ROW_ODD);
            ctx.fillRect(0, y, tW, ROW_H);

            // 行下線
            ctx.strokeStyle = C_BORDER; ctx.lineWidth = 1;
            ctx.beginPath(); ctx.moveTo(0, y + ROW_H); ctx.lineTo(tW, y + ROW_H); ctx.stroke();

            ctx.fillStyle = C_TEXT;
            let x = 0;
            const row = this.filteredRows[ri];
            for (let ci = 0; ci < row.length; ci++) {
                const cw = this.colWidths[ci];
                ctx.save();
                ctx.beginPath(); ctx.rect(x + CELL_PAD, y, cw - CELL_PAD * 2, ROW_H); ctx.clip();
                ctx.fillText(row[ci], x + CELL_PAD, y + ROW_H / 2);
                ctx.restore();
                x += cw;
            }
        }
        ctx.restore();
    }

    private drawScrollbar(ctx: CanvasRenderingContext2D, W: number, H: number): void {
        const bodyH   = H - HEADER_H;
        const totalH  = this.filteredRows.length * ROW_H;
        const trackX  = W - SB_W;
        const trackY  = HEADER_H;

        // トラック
        ctx.fillStyle = C_SB_TRACK;
        ctx.fillRect(trackX, trackY, SB_W, bodyH);

        if (totalH <= bodyH) return;

        const thumbH = Math.max(24, (bodyH / totalH) * bodyH);
        const thumbY = trackY + (this.scrollTop / this.maxScroll()) * (bodyH - thumbH);

        ctx.fillStyle = C_SB_THUMB;
        ctx.beginPath();
        ctx.roundRect(trackX + 2, thumbY + 1, SB_W - 4, thumbH - 2, 3);
        ctx.fill();
    }

    // =========================================================
    // Canvas イベント
    // =========================================================
    private setupCanvasEvents(): void {
        this.canvas.addEventListener("wheel", (e: WheelEvent) => {
            e.preventDefault();
            this.scrollTop = Math.max(0, Math.min(this.scrollTop + e.deltaY, this.maxScroll()));
            this.scheduleRedraw();
        }, { passive: false });

        this.canvas.addEventListener("mousemove", (e: MouseEvent) => {
            const rect = this.canvas.getBoundingClientRect();
            const y    = e.clientY - rect.top;
            const prev = this.hoveredRow;
            this.hoveredRow = y < HEADER_H ? -1 : Math.floor((y - HEADER_H + this.scrollTop) / ROW_H);
            if (this.hoveredRow !== prev) this.scheduleRedraw();

            // スクロールバードラッグ中
            if (this.isDragging) {
                const dy     = y - this.dragStartY;
                const bodyH  = this.logicalH - HEADER_H;
                const totalH = this.filteredRows.length * ROW_H;
                const ratio  = totalH / bodyH;
                this.scrollTop = Math.max(0, Math.min(this.dragStartScroll + dy * ratio, this.maxScroll()));
                this.scheduleRedraw();
            }
        });

        this.canvas.addEventListener("mousedown", (e: MouseEvent) => {
            const rect = this.canvas.getBoundingClientRect();
            const x    = e.clientX - rect.left;
            if (x >= this.logicalW - SB_W) {
                this.isDragging      = true;
                this.dragStartY      = e.clientY - rect.top;
                this.dragStartScroll = this.scrollTop;
            }
        });

        const endDrag = () => { this.isDragging = false; };
        this.canvas.addEventListener("mouseup",    endDrag);
        this.canvas.addEventListener("mouseleave", () => { endDrag(); this.hoveredRow = -1; this.scheduleRedraw(); });
    }

    public getFormattingModel(): powerbi.visuals.FormattingModel {
        return this.formattingSettingsService.buildFormattingModel(this.formattingSettings);
    }
}
