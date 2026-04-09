"use strict";

import powerbi from "powerbi-visuals-api";
import { FormattingSettingsService } from "powerbi-visuals-utils-formattingmodel";
import "./../style/visual.less";

import VisualConstructorOptions = powerbi.extensibility.visual.VisualConstructorOptions;
import VisualUpdateOptions      = powerbi.extensibility.visual.VisualUpdateOptions;
import IVisual                  = powerbi.extensibility.visual.IVisual;
import IVisualHost              = powerbi.extensibility.visual.IVisualHost;
import ISelectionManager        = powerbi.extensibility.ISelectionManager;
import DataView                 = powerbi.DataView;

import { VisualFormattingSettingsModel } from "./settings";

// ---- Canvas 定数 ----
const ROW_H    = 22;
const HEADER_H = 28;
const CB_W     = 26;   // チェックボックス列幅
const SB_W     = 10;   // スクロールバー幅
const PAD      = 8;
const FONT     = "12px 'Segoe UI',sans-serif";
const FONT_HDR = "600 12px 'Segoe UI',sans-serif";

const C = {
    hdrBg:   "#f2f2f2", rowEven: "#ffffff", rowOdd: "#fafafa",
    rowHov:  "#e8f4fd", rowSel:  "#deecf9",
    text:    "#252423", border:  "#e0e0e0", hdrLine: "#d1d1d1",
    sbTrack: "#f0f0f0", sbThumb: "#c8c8c8",
    cbBorder:"#8a8886", cbFill:  "#0078d4", cbCheck: "#ffffff",
};

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
    private filteredOrigIdx: number[]            = [];   // filteredRows[i] → tableData.rows の元インデックス
    private selectionIds: powerbi.visuals.ISelectionId[] = [];
    private selectedOrigIdx: Set<number>         = new Set();
    private visibleCols: boolean[]               = [];

    // ---- DOM ----
    private filterPanel:  HTMLElement;
    private colToggleBar: HTMLElement;
    private statusBar:    HTMLElement;
    private canvasArea:   HTMLElement;
    private canvas:       HTMLCanvasElement;
    private ctx:          CanvasRenderingContext2D;

    // ---- Canvas 状態 ----
    private scrollTop   = 0;
    private hoveredRow  = -1;
    private colWidths:    number[] = [];
    private logicalW    = 0;
    private logicalH    = 0;
    private drawPending = false;

    // ---- スクロールバードラッグ ----
    private isDragging      = false;
    private dragStartY      = 0;
    private dragStartScroll = 0;

    // ---- 自分の persist サイクルをスキップ ----
    private skipRender   = false;
    private persistTimer: number | null = null;

    constructor(options: VisualConstructorOptions) {
        this.host = options.host;
        this.formattingSettingsService = new FormattingSettingsService();
        this.selectionManager = options.host.createSelectionManager();
        options.element.className = "filter-table-visual";
        this.buildDOM(options.element);
        this.setupCanvasEvents();
    }

    private buildDOM(root: HTMLElement): void {
        this.filterPanel  = this.el("div", "filter-panel");
        this.colToggleBar = this.el("div", "col-toggle-bar");
        this.statusBar    = this.el("div", "status-bar");
        this.canvasArea   = this.el("div", "canvas-area");
        this.canvas       = this.el("canvas", "") as HTMLCanvasElement;
        this.ctx          = this.canvas.getContext("2d");
        this.canvasArea.appendChild(this.canvas);
        [this.filterPanel, this.colToggleBar, this.statusBar, this.canvasArea]
            .forEach(e => root.appendChild(e));
    }

    private el<K extends keyof HTMLElementTagNameMap>(
        tag: K, cls: string
    ): HTMLElementTagNameMap[K] {
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
            if (!this.filterPanel.querySelector("input:focus")) this.renderFilterPanel();
            return;
        }

        this.tableData = this.extractTableData(dv);
        this.buildSelectionIds(dv);

        // 列が増えた分は true、減った分は切り詰め（既存の表示設定を保持）
        if (this.visibleCols.length !== this.tableData.columns.length) {
            this.visibleCols = this.tableData.columns.map((_, i) =>
                this.visibleCols[i] !== undefined ? this.visibleCols[i] : true
            );
        }

        if (!this.filterPanel.querySelector("input:focus")) {
            this.restoreState(dv);
            this.renderFilterPanel();
        }

        this.renderColToggleBar();
        this.runFilter();
        this.scrollTop = 0;
        this.renderStatus();
        this.resizeCanvas();
        this.scheduleRedraw();
    }

    private restoreState(dv: DataView): void {
        const m   = dv?.metadata?.objects?.["filterState"];
        const len = this.tableData.columns.length;

        // 列数を超える columnIndex の条件は除外（列を削減したときの不整合対策）
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
    // フィルターパネル（HTML DOM）
    // ==========================================================
    private renderFilterPanel(): void {
        this.clear(this.filterPanel);

        // ヘッダー（タイトル + AND/OR）
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

        // 条件リスト
        const list = this.el("div", "condition-list");
        this.conditions.forEach((c, i) => list.appendChild(this.makeConditionRow(c, i)));
        this.filterPanel.appendChild(list);

        // フッター（追加 / 解除 / 実行）
        const footer = this.el("div", "filter-footer");

        const addBtn = this.el("button", "add-condition-btn");
        addBtn.textContent = "+ 条件を追加";
        addBtn.onclick = () => {
            this.conditions.push({ columnIndex: 0, operator: "contains", value: "" });
            this.saveState(); this.renderFilterPanel();
        };

        const clearBtn = this.el("button", "clear-btn");
        clearBtn.textContent = "解除";
        clearBtn.title = "フィルターを解除して全件表示";
        (clearBtn as HTMLButtonElement).disabled = this.appliedConditions.length === 0;
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
    // 列トグルバー
    // ==========================================================
    private renderColToggleBar(): void {
        this.clear(this.colToggleBar);
        const multi = this.tableData.columns.length > 1;
        this.colToggleBar.style.display = multi ? "flex" : "none";
        if (!multi) return;

        this.tableData.columns.forEach((col, i) => {
            const chip = this.el("button", "col-chip" + (this.visibleCols[i] ? " active" : ""));
            chip.textContent = col;
            chip.title = this.visibleCols[i] ? "非表示にする" : "表示する";
            chip.onclick = () => {
                const visCount = this.visibleCols.filter(Boolean).length;
                if (visCount === 1 && this.visibleCols[i]) return; // 最後の1列は守る
                this.visibleCols[i] = !this.visibleCols[i];
                chip.className = "col-chip" + (this.visibleCols[i] ? " active" : "");
                chip.title = this.visibleCols[i] ? "非表示にする" : "表示する";
                this.calcColWidths(this.logicalW - SB_W - CB_W);
                this.scheduleRedraw();
            };
            this.colToggleBar.appendChild(chip);
        });
    }

    // ==========================================================
    // 検索ロジック
    // ==========================================================
    private executeSearch(): void {
        this.appliedConditions = this.conditions.map(c => ({ ...c }));
        this.appliedLogic = this.logic;
        this.runFilter(); this.scrollTop = 0;
        this.persist(); this.renderFilterPanel(); this.renderStatus(); this.scheduleRedraw();
    }

    private clearFilter(): void {
        this.appliedConditions = []; this.appliedLogic = "AND";
        this.runFilter(); this.scrollTop = 0;
        this.persist(); this.renderFilterPanel(); this.renderStatus(); this.scheduleRedraw();
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
    // 選択（クロスフィルター）
    // ==========================================================
    private toggleRowSelection(fi: number): void {
        const oi = this.filteredOrigIdx[fi];
        this.selectedOrigIdx.has(oi) ? this.selectedOrigIdx.delete(oi) : this.selectedOrigIdx.add(oi);
        this.commitSelection();
    }

    private toggleSelectAll(): void {
        const allSel = this.filteredOrigIdx.every(i => this.selectedOrigIdx.has(i));
        this.filteredOrigIdx.forEach(i => allSel ? this.selectedOrigIdx.delete(i) : this.selectedOrigIdx.add(i));
        this.commitSelection();
    }

    private clearSelection(): void {
        this.selectedOrigIdx.clear(); this.commitSelection();
    }

    private commitSelection(): void {
        const ids = Array.from(this.selectedOrigIdx).map(i => this.selectionIds[i]).filter(Boolean);
        ids.length ? this.selectionManager.select(ids) : this.selectionManager.clear();
        this.renderStatus(); this.scheduleRedraw();
    }

    // ==========================================================
    // ステータスバー
    // ==========================================================
    private renderStatus(): void {
        this.clear(this.statusBar);
        const f = this.filteredRows.length, t = this.tableData.rows.length;
        this.statusBar.appendChild(
            document.createTextNode(f === t ? `${t} 件` : `${f} / ${t} 件`)
        );
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
        this.host.persistProperties({ merge: [{ objectName: "filterState", selector: null, properties: {
            conditions: JSON.stringify(this.conditions), logic: this.logic,
            applied: JSON.stringify(this.appliedConditions), appliedLogic: this.appliedLogic,
        }}]});
    }

    // ==========================================================
    // Canvas
    // ==========================================================
    private resizeCanvas(): void {
        const w = this.canvasArea.clientWidth, h = this.canvasArea.clientHeight;
        if (w <= 0 || h <= 0) return;
        const dpr = window.devicePixelRatio || 1;
        this.canvas.width  = Math.round(w * dpr); this.canvas.height = Math.round(h * dpr);
        this.canvas.style.width  = w + "px";       this.canvas.style.height = h + "px";
        this.logicalW = w; this.logicalH = h;
        this.ctx.setTransform(dpr, 0, 0, dpr, 0, 0);
        this.calcColWidths(w - SB_W - CB_W);
    }

    private calcColWidths(availW: number): void {
        const vis = this.tableData.columns.map((_, i) => i).filter(i => this.visibleCols[i]);
        if (!vis.length || availW <= 0) { this.colWidths = this.tableData.columns.map(() => 0); return; }

        const ctx = this.ctx;
        const ws: Record<number, number> = {};
        ctx.font = FONT_HDR;
        vis.forEach(i => {
            ws[i] = Math.min(280, Math.max(50, ctx.measureText(this.tableData.columns[i]).width + PAD * 2 + 12));
        });
        ctx.font = FONT;
        for (const row of this.filteredRows.slice(0, 300)) {
            vis.forEach(i => {
                const w = Math.min(280, ctx.measureText(row[i] ?? "").width + PAD * 2);
                if (w > ws[i]) ws[i] = w;
            });
        }

        const total = vis.reduce((s, i) => s + ws[i], 0);
        if (total <= availW) {
            // 余白を均等配分
            const extra = (availW - total) / vis.length;
            vis.forEach(i => ws[i] += extra);
        } else {
            // 列が多すぎる場合は比率で縮小（最小 40px を保証）
            const scale = availW / total;
            vis.forEach(i => { ws[i] = Math.max(40, Math.floor(ws[i] * scale)); });
        }

        this.colWidths = this.tableData.columns.map((_, i) => ws[i] ?? 0);
    }

    private scheduleRedraw(): void {
        if (this.drawPending) return;
        this.drawPending = true;
        requestAnimationFrame(() => { this.drawPending = false; this.draw(); });
    }

    private maxScroll(): number {
        return Math.max(0, this.filteredRows.length * ROW_H - (this.logicalH - HEADER_H));
    }

    // ---- メイン描画 ----
    private draw(): void {
        const ctx = this.ctx, W = this.logicalW, H = this.logicalH;
        if (W <= 0 || H <= 0) return;
        ctx.clearRect(0, 0, W, H);
        if (!this.tableData.columns.length) {
            ctx.fillStyle = "#a19f9d"; ctx.font = FONT; ctx.textAlign = "center"; ctx.textBaseline = "middle";
            ctx.fillText("データをフィールドに追加してください", W / 2, H / 2);
            return;
        }
        const tW = W - SB_W;
        this.drawHeader(ctx, tW); this.drawBody(ctx, tW); this.drawScrollbar(ctx, W, H);
        ctx.strokeStyle = C.hdrLine; ctx.lineWidth = 1;
        ctx.beginPath(); ctx.moveTo(0, HEADER_H); ctx.lineTo(tW, HEADER_H); ctx.stroke();
    }

    private drawHeader(ctx: CanvasRenderingContext2D, tW: number): void {
        ctx.fillStyle = C.hdrBg; ctx.fillRect(0, 0, tW, HEADER_H);

        // 全選択チェックボックス
        const allSel = this.filteredOrigIdx.length > 0
            && this.filteredOrigIdx.every(i => this.selectedOrigIdx.has(i));
        const someSel = !allSel && this.filteredOrigIdx.some(i => this.selectedOrigIdx.has(i));
        this.drawCB(ctx, (CB_W - 14) / 2, (HEADER_H - 14) / 2, allSel, someSel);

        ctx.font = FONT_HDR; ctx.fillStyle = C.text; ctx.textBaseline = "middle"; ctx.textAlign = "left";
        const vis = this.tableData.columns.map((_, i) => i).filter(i => this.visibleCols[i]);
        let x = CB_W;
        vis.forEach(i => {
            const cw = this.colWidths[i] ?? 0;
            if (cw <= 0) return;
            ctx.save(); ctx.beginPath(); ctx.rect(x + PAD, 0, cw - PAD * 2, HEADER_H); ctx.clip();
            ctx.fillText(this.tableData.columns[i], x + PAD, HEADER_H / 2); ctx.restore();
            ctx.strokeStyle = C.hdrLine; ctx.lineWidth = 1;
            ctx.beginPath(); ctx.moveTo(x + cw - 0.5, 4); ctx.lineTo(x + cw - 0.5, HEADER_H - 4); ctx.stroke();
            x += cw;
        });
    }

    private drawBody(ctx: CanvasRenderingContext2D, tW: number): void {
        const bodyH = this.logicalH - HEADER_H;
        const first = Math.floor(this.scrollTop / ROW_H);
        const last  = Math.min(this.filteredRows.length, first + Math.ceil(bodyH / ROW_H) + 1);
        const vis   = this.tableData.columns.map((_, i) => i).filter(i => this.visibleCols[i]);

        ctx.save(); ctx.beginPath(); ctx.rect(0, HEADER_H, tW, bodyH); ctx.clip();
        ctx.font = FONT; ctx.textBaseline = "middle"; ctx.textAlign = "left";

        for (let ri = first; ri < last; ri++) {
            if (ri >= this.filteredOrigIdx.length) break;   // 境界保護
            const y   = HEADER_H + ri * ROW_H - this.scrollTop;
            const oi  = this.filteredOrigIdx[ri];
            const sel = this.selectedOrigIdx.has(oi);

            ctx.fillStyle = sel ? C.rowSel : (ri === this.hoveredRow ? C.rowHov : (ri % 2 === 0 ? C.rowEven : C.rowOdd));
            ctx.fillRect(0, y, tW, ROW_H);
            ctx.strokeStyle = C.border; ctx.lineWidth = 1;
            ctx.beginPath(); ctx.moveTo(0, y + ROW_H); ctx.lineTo(tW, y + ROW_H); ctx.stroke();

            this.drawCB(ctx, (CB_W - 14) / 2, y + (ROW_H - 14) / 2, sel, false);

            ctx.fillStyle = C.text;
            let x = CB_W;
            const row = this.filteredRows[ri];
            vis.forEach(i => {
                const cw = this.colWidths[i] ?? 0;
                if (cw <= 0) return;
                ctx.save(); ctx.beginPath(); ctx.rect(x + PAD, y, cw - PAD * 2, ROW_H); ctx.clip();
                ctx.fillText(row[i] ?? "", x + PAD, y + ROW_H / 2); ctx.restore();
                x += cw;
            });
        }
        ctx.restore();
    }

    private drawCB(ctx: CanvasRenderingContext2D, x: number, y: number, checked: boolean, indeterminate: boolean): void {
        const s = 14;
        ctx.fillStyle   = checked ? C.cbFill : "#fff";
        ctx.strokeStyle = checked ? C.cbFill : C.cbBorder;
        ctx.lineWidth   = 1;
        ctx.beginPath(); ctx.roundRect(x, y, s, s, 2); ctx.fill(); ctx.stroke();
        if (checked) {
            ctx.strokeStyle = C.cbCheck; ctx.lineWidth = 2; ctx.lineCap = "round"; ctx.lineJoin = "round";
            ctx.beginPath(); ctx.moveTo(x+3, y+7); ctx.lineTo(x+6, y+10); ctx.lineTo(x+11, y+4); ctx.stroke();
        } else if (indeterminate) {
            ctx.strokeStyle = C.cbFill; ctx.lineWidth = 2;
            ctx.beginPath(); ctx.moveTo(x+3, y+7); ctx.lineTo(x+11, y+7); ctx.stroke();
        }
    }

    private drawScrollbar(ctx: CanvasRenderingContext2D, W: number, H: number): void {
        const bodyH = H - HEADER_H, totalH = this.filteredRows.length * ROW_H, tx = W - SB_W;
        ctx.fillStyle = C.sbTrack; ctx.fillRect(tx, HEADER_H, SB_W, bodyH);
        if (totalH <= bodyH) return;
        const thumbH = Math.max(24, (bodyH / totalH) * bodyH);
        const thumbY = HEADER_H + (this.scrollTop / this.maxScroll()) * (bodyH - thumbH);
        ctx.fillStyle = C.sbThumb;
        ctx.beginPath(); ctx.roundRect(tx + 2, thumbY + 1, SB_W - 4, thumbH - 2, 3); ctx.fill();
    }

    // ==========================================================
    // Canvas イベント
    // ==========================================================
    private setupCanvasEvents(): void {
        this.canvas.addEventListener("wheel", (e: WheelEvent) => {
            e.preventDefault();
            this.scrollTop = Math.max(0, Math.min(this.scrollTop + e.deltaY, this.maxScroll()));
            this.scheduleRedraw();
        }, { passive: false });

        this.canvas.addEventListener("click", (e: MouseEvent) => {
            const r = this.canvas.getBoundingClientRect();
            const x = e.clientX - r.left, y = e.clientY - r.top;
            if (x >= this.logicalW - SB_W) return;
            if (y < HEADER_H) {
                if (x < CB_W) this.toggleSelectAll();
            } else if (x < CB_W) {
                const ri = Math.floor((y - HEADER_H + this.scrollTop) / ROW_H);
                if (ri >= 0 && ri < this.filteredRows.length) this.toggleRowSelection(ri);
            }
        });

        this.canvas.addEventListener("mousemove", (e: MouseEvent) => {
            const r = this.canvas.getBoundingClientRect();
            const y = e.clientY - r.top;
            const prev = this.hoveredRow;
            this.hoveredRow = y < HEADER_H ? -1 : Math.floor((y - HEADER_H + this.scrollTop) / ROW_H);
            if (this.hoveredRow !== prev) this.scheduleRedraw();
            if (this.isDragging) {
                const dy = y - this.dragStartY;
                const ratio = (this.filteredRows.length * ROW_H) / (this.logicalH - HEADER_H);
                this.scrollTop = Math.max(0, Math.min(this.dragStartScroll + dy * ratio, this.maxScroll()));
                this.scheduleRedraw();
            }
        });

        this.canvas.addEventListener("mousedown", (e: MouseEvent) => {
            const r = this.canvas.getBoundingClientRect();
            if (e.clientX - r.left >= this.logicalW - SB_W) {
                this.isDragging = true; this.dragStartY = e.clientY - r.top; this.dragStartScroll = this.scrollTop;
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
