/*
 *  Gantt Deneb — Power BI Custom Visual
 *  Renders a Vega v5 Gantt chart (David Bacci template).
 *
 *  Scroll behaviour (modification vs. original spec):
 *    • Mouse-wheel left/right  → pan the timeline left or right
 *    • Mouse-wheel up/down     → pan the row list up or down
 *    • No zoom on scroll (zoom removed)
 *    • Double-click still resets to the initial view
 *    • Drag still pans both axes
 *    • Toolbar buttons (All / Years / Months / Days) still work
 */
"use strict";

import powerbi from "powerbi-visuals-api";
import "./../style/visual.less";
import { BASE_SPEC } from "./ganttSpec";
import { VisualFormattingSettingsModel } from "./settings";
import * as vega from "vega";
import { FormattingSettingsService } from "powerbi-visuals-utils-formattingmodel";

import VisualConstructorOptions = powerbi.extensibility.visual.VisualConstructorOptions;
import VisualUpdateOptions   = powerbi.extensibility.visual.VisualUpdateOptions;
import IVisual               = powerbi.extensibility.visual.IVisual;
import IVisualHost           = powerbi.extensibility.visual.IVisualHost;
import DataView              = powerbi.DataView;

/* ─── Padding constants ────────────────────────────────────────────────────
 * The buttons in the spec are at y = -60 (above the chart origin).
 * We need enough top padding so the SVG viewBox includes them.
 * padTop = 70  →  viewBox starts at y = -70, so y = -60 is 10 px inside.
 * ─────────────────────────────────────────────────────────────────────── */
const PAD_LEFT   = 5;
const PAD_RIGHT  = 5;
const PAD_TOP    = 55;  // must be > button height (18) + button y-offset (30) + axis room
const PAD_BOTTOM = 18;  // room for horizontal scrollbar

/* ─── Build the patched Vega spec ──────────────────────────────────────── */
function buildSpec(width: number, height: number, data?: any[]): any {
    // Deep-clone so we never mutate BASE_SPEC between calls
    const spec: any = JSON.parse(JSON.stringify(BASE_SPEC));

    // ── Dimensions ──────────────────────────────────────────────────────
    spec.width   = Math.max(200, width  - PAD_LEFT  - PAD_RIGHT);
    spec.height  = Math.max(100, height - PAD_TOP   - PAD_BOTTOM);
    spec.padding = { left: PAD_LEFT, right: PAD_RIGHT, top: PAD_TOP, bottom: PAD_BOTTOM };

    const signals: any[] = spec.signals;

    // ── 1. Remove zoom-on-wheel ─────────────────────────────────────────
    // Original: zoom signal reacts to "wheel!" and applies pow() zoom.
    // New behaviour: constant 1 (no zoom), xDomPre will never fire from wheel.
    const zoomSig = signals.find((s: any) => s.name === "zoom");
    if (zoomSig) {
        delete zoomSig.on;       // remove the wheel! handler entirely
    }

    // ── 2. Remove anchor wheel handler ──────────────────────────────────
    // anchor was only needed to compute the zoom pivot point.
    const anchorSig = signals.find((s: any) => s.name === "anchor");
    if (anchorSig) {
        delete anchorSig.on;
    }

    // ── 3 & 4. Wheel scroll is handled entirely in TypeScript (RAF-throttled).
    // No wheel handlers are added to Vega signals; the native listener in
    // Visual.attachWheelListener() pushes signal values directly and batches
    // all events within one animation frame into a single runAsync() call.

    // ── 5. Inject Power BI data ─────────────────────────────────────────
    // Always replace the hardcoded sample values: either with real PBI data
    // or with an empty array so the chart never silently shows sample data.
    spec.data[0].values = (data && data.length > 0) ? data : [];

    // ── 6. Disable all animations ────────────────────────────────────────
    if (!spec.config) spec.config = {};
    spec.config.animation = { duration: 0 };

    return spec;
}

/* ─── Format a JS Date as "dd/MM/yyyy" ────────────────────────────────── */
function fmtDate(d: Date): string {
    const dd = String(d.getDate()).padStart(2, "0");
    const mm = String(d.getMonth() + 1).padStart(2, "0");
    const yyyy = String(d.getFullYear());
    return `${dd}/${mm}/${yyyy}`;
}

/* ─── Visual class ─────────────────────────────────────────────────────── */
export class Visual implements IVisual {
    private host: HTMLElement;
    private pbiHost: IVisualHost;
    private vegaView: any = null;
    private lastWidth  = 0;
    private lastHeight = 0;
    private lastDataHash = "";

    // RAF scroll state
    private wheelListener: ((e: WheelEvent) => void) | null = null;
    private rafPending   = false;
    private pendingXDom:   [number, number] | null = null;
    private pendingYRange: [number, number] | null = null;
    private vScrollEl:     HTMLInputElement | null = null;
    private hScrollEl:     HTMLInputElement | null = null;

    // Formatting
    private formattingSettings: VisualFormattingSettingsModel;
    private formattingSettingsService: FormattingSettingsService;

    // Title element
    private titleEl: HTMLDivElement | null = null;

    constructor(options: VisualConstructorOptions) {
        this.host = options.element;
        this.pbiHost = options.host;
        this.host.classList.add("gantt-deneb-host");
        this.formattingSettings = new VisualFormattingSettingsModel();
        this.formattingSettingsService = new FormattingSettingsService();
    }

    public update(options: VisualUpdateOptions): void {
        // Parse formatting settings from dataViews
        if (options.dataViews && options.dataViews[0]) {
            this.formattingSettings = this.formattingSettingsService.populateFormattingSettingsModel(
                VisualFormattingSettingsModel, options.dataViews[0]
            );
        }

        const vp = options.viewport;
        const w  = Math.max(200, Math.floor(vp.width));
        const h  = Math.max(100, Math.floor(vp.height));

        // Apply title
        this.applyTitle(w);

        // Calculate effective height (subtract title if visible)
        const titleH = this.titleEl ? this.titleEl.offsetHeight : 0;
        const effectiveH = Math.max(100, h - titleH);

        // Extract Power BI data (returns null when no data is bound)
        const pbData = this.extractData(options.dataViews);
        const dataHash = pbData ? JSON.stringify(pbData) : "";

        const sizeChanged = (w !== this.lastWidth || effectiveH !== this.lastHeight);
        const dataChanged = (dataHash !== this.lastDataHash);

        if (!this.vegaView || dataChanged) {
            this.createView(w, effectiveH, pbData ?? undefined);
        } else if (sizeChanged) {
            this.resizeView(w, effectiveH);
        }

        this.lastWidth    = w;
        this.lastHeight   = effectiveH;
        this.lastDataHash = dataHash;
    }

    private applyTitle(_w: number): void {
        const show = this.formattingSettings.titleCard.showTitle.value;
        if (!show) {
            if (this.titleEl) { this.titleEl.remove(); this.titleEl = null; }
            return;
        }
        if (!this.titleEl) {
            this.titleEl = document.createElement("div");
            this.titleEl.className = "gantt-title";
            this.host.insertBefore(this.titleEl, this.host.firstChild);
        }
        const tc = this.formattingSettings.titleCard;
        this.titleEl.textContent = tc.titleText.value || "Gantt Chart";
        this.titleEl.style.fontSize = (tc.titleFontSize.value || 14) + "px";
        this.titleEl.style.color = tc.titleColor.value.value || "#333333";
        this.titleEl.style.fontWeight = "600";
        this.titleEl.style.padding = "4px 8px";
        this.titleEl.style.fontFamily = "Segoe UI, sans-serif";
    }

    public getFormattingModel(): powerbi.visuals.FormattingModel {
        return this.formattingSettingsService.buildFormattingModel(this.formattingSettings);
    }

    private createView(w: number, h: number, data?: any[]): void {
        // Tear down existing listener + view
        this.detachWheelListener();
        if (this.vegaView) {
            try { this.vegaView.finalize(); } catch (_) { /* ignore */ }
            this.vegaView = null;
        }
        this.host.innerHTML = "";

        const spec    = buildSpec(w, h, data);
        const runtime = vega.parse(spec);

        this.vegaView = new vega.View(runtime, {
            renderer : "canvas",  // Canvas is 10-50x faster than SVG for many marks
            container: this.host,
            hover    : true,
            logLevel : vega.Warn
        });

        this.vegaView.runAsync().then(() => {
            this.attachWheelListener();
            this.buildScrollbar();
            this.buildHScrollbar();
            this.attachClickToNavigate();
        });
    }

    /* ─── RAF-throttled native wheel handler ─────────────────────────────
     * Batches all wheel events within one animation frame into a single
     * Vega signal update + runAsync(), keeping the render rate ≤ 60 fps.
     * Both axes are driven entirely by Vega signals (no data re-slicing),
     * which avoids the NaN-on-empty-dataset race condition.
     * ─────────────────────────────────────────────────────────────────── */
    private attachWheelListener(): void {
        this.wheelListener = (e: WheelEvent) => {
            e.preventDefault();

            const v = this.vegaView;
            if (!v) return;

            // Normalise deltaMode: 0 = px, 1 = line (≈16 px), 2 = page (≈256 px)
            const norm = e.deltaMode === 0 ? 1 : e.deltaMode === 1 ? 16 : 256;
            const dx = e.deltaX * norm;
            const dy = -(e.deltaY * norm);  // negate: natural scroll (two-finger-up → content down)

            // ── Horizontal pan (deltaX) ──────────────────────────────────
            if (dx !== 0) {
                const curXDom    = this.pendingXDom ?? (v.signal("xDom") as [number, number]);
                const ganttWidth: number = v.signal("ganttWidth") ?? 1;
                const span  = curXDom[1] - curXDom[0];
                const shift = dx * span / ganttWidth;
                this.pendingXDom = [curXDom[0] + shift, curXDom[1] + shift];
            }

            // ── Vertical pan (deltaY) ─────────────────────────────────────
            // Vega's yRange controls which portion of the row list is visible.
            // deltaY > 0 = scroll down = show lower rows = increase yRange.
            if (dy !== 0) {
                const curYRange     = this.pendingYRange ?? (v.signal("yRange") as [number, number]);
                const height:       number = v.signal("height")       ?? 500;
                const scaledHeight: number = v.signal("scaledHeight") ?? height;
                const span  = curYRange[1] - curYRange[0];
                const minY  = height >= scaledHeight ? 0             : height - scaledHeight;
                const maxY  = height >= scaledHeight ? height        : scaledHeight;
                const newY0 = Math.min(maxY - span, Math.max(minY, curYRange[0] + dy));
                this.pendingYRange = [newY0, newY0 + span];
            }

            // ── Flush once per animation frame ───────────────────────────
            if (!this.rafPending) {
                this.rafPending = true;
                requestAnimationFrame(() => {
                    if (!this.vegaView) { this.rafPending = false; return; }
                    if (this.pendingXDom)   { this.vegaView.signal("xDom",   this.pendingXDom);   this.pendingXDom   = null; }
                    if (this.pendingYRange) { this.vegaView.signal("yRange", this.pendingYRange); this.pendingYRange = null; }
                    this.vegaView.runAsync();
                    this.rafPending = false;
                });
            }
        };

        this.host.addEventListener("wheel", this.wheelListener, { passive: false });
    }

    /* ─── Native vertical scrollbar ──────────────────────────────────────
     * A range <input> overlaid on the right edge of the visual.
     *   • Scrollbar → Vega: dragging updates yRange directly.
     *   • Vega → Scrollbar: addSignalListener keeps the thumb in sync
     *     when the user pans by mouse-drag or double-click reset.
     * ─────────────────────────────────────────────────────────────────── */
    private buildScrollbar(): void {
        // Remove stale element from a previous createView()
        if (this.vScrollEl) { try { this.vScrollEl.remove(); } catch (_) { /* ok */ } }

        const v = this.vegaView;
        if (!v) return;

        const scaledH: number = v.signal("scaledHeight") ?? 0;
        const height:  number = v.signal("height")       ?? 0;
        // Only show scrollbar when content is taller than the viewport
        if (scaledH <= height) { this.vScrollEl = null; return; }

        const sb = document.createElement("input");
        sb.type  = "range";
        sb.min   = "0";
        sb.max   = "10000";
        sb.value = "0";
        sb.style.cssText = [
            "position:absolute", "top:0", "right:0",
            "width:14px", "height:100%",
            // Standards-based vertical orientation
            "writing-mode:vertical-lr",
            "direction:rtl",
            // Chrome / Edge
            "-webkit-appearance:slider-vertical",
            "cursor:pointer", "opacity:0.55", "z-index:20"
        ].join(";");

        // Scrollbar → Vega: drag moves the yRange
        sb.addEventListener("input", () => {
            const vw = this.vegaView;
            if (!vw || this.rafPending) return;
            const scaledHNow: number = vw.signal("scaledHeight") ?? 1;
            const heightNow:  number = vw.signal("height")       ?? 1;
            const span   = Math.min(heightNow, scaledHNow);
            const maxY0  = Math.max(0, scaledHNow - span);
            const newY0  = (Number(sb.value) / 10000) * maxY0;
            this.rafPending = true;
            requestAnimationFrame(() => {
                this.vegaView?.signal("yRange", [newY0, newY0 + span]).runAsync();
                this.rafPending = false;
            });
        });

        // Vega → Scrollbar: keep thumb in sync when yRange changes (pan / dblclick)
        v.addSignalListener("yRange", (_name: string, value: [number, number]) => {
            if (!this.vScrollEl) return;
            const vw = this.vegaView;
            if (!vw) return;
            const scaledHNow: number = vw.signal("scaledHeight") ?? 1;
            const heightNow:  number = vw.signal("height")       ?? 1;
            const span  = Math.min(heightNow, scaledHNow);
            const maxY0 = Math.max(1, scaledHNow - span);
            const y0    = value?.[0] ?? 0;
            this.vScrollEl.value = String(Math.round((y0 / maxY0) * 10000));
        });

        this.host.style.position = "relative";
        this.host.appendChild(sb);
        this.vScrollEl = sb;
    }

    /* ─── Native horizontal scrollbar ────────────────────────────────────
     * A range <input> overlaid at the bottom edge of the visual.
     * Synced with the xDom signal (timeline pan).
     * ─────────────────────────────────────────────────────────────────── */
    private buildHScrollbar(): void {
        if (this.hScrollEl) { try { this.hScrollEl.remove(); } catch (_) { /* ok */ } }

        const v = this.vegaView;
        if (!v) return;

        const xExt = v.data("xExt");
        if (!xExt || xExt.length === 0) { this.hScrollEl = null; return; }

        const sb = document.createElement("input");
        sb.type  = "range";
        sb.min   = "0";
        sb.max   = "10000";
        sb.value = "0";
        sb.style.cssText = [
            "position:absolute", "bottom:0", "left:0",
            "height:14px", "width:100%",
            "cursor:pointer", "opacity:0.55", "z-index:20"
        ].join(";");

        // Scrollbar → Vega: drag pans xDom
        sb.addEventListener("input", () => {
            const vw = this.vegaView;
            if (!vw) return;
            const xExtNow = vw.data("xExt");
            if (!xExtNow || xExtNow.length === 0) return;
            const fullMin = +xExtNow[0].s;  // xExt uses 's' (min start) and 'e' (max end)
            const fullMax = +xExtNow[0].e;
            const curXDom = vw.signal("xDom") as [number, number];
            const span = curXDom[1] - curXDom[0];
            const maxStart = Math.max(0, fullMax - span);
            const newStart = fullMin + (Number(sb.value) / 10000) * (maxStart - fullMin);
            vw.signal("xDom", [newStart, newStart + span]).runAsync();
        });

        // Vega → Scrollbar: keep thumb in sync when xDom changes
        v.addSignalListener("xDom", (_name: string, value: [number, number]) => {
            if (!this.hScrollEl) return;
            const vw = this.vegaView;
            if (!vw) return;
            const xExtNow = vw.data("xExt");
            if (!xExtNow || xExtNow.length === 0) return;
            const fullMin = +xExtNow[0].s;
            const fullMax = +xExtNow[0].e;
            const span = value[1] - value[0];
            const maxStart = Math.max(1, fullMax - span);
            const x0 = value?.[0] ?? fullMin;
            this.hScrollEl.value = String(Math.round(((x0 - fullMin) / (maxStart - fullMin)) * 10000));
        });

        this.host.appendChild(sb);
        this.hScrollEl = sb;
    }

    /* ─── Click on activity → scroll to its start ────────────────────────
     * Global click handler — skips phase header rows (phase === task)
     * so collapse/expand still works normally.
     * ─────────────────────────────────────────────────────────────────── */
    private attachClickToNavigate(): void {
        const v = this.vegaView;
        if (!v) return;

        const navigate = (_event: any, item: any) => {
            if (!item?.datum) return;
            // Phase headers have phase === task; skip them (handled by phaseClicked signal)
            if (String(item.datum.phase) === String(item.datum.task)) return;
            const startVal: any = item.datum.start;
            if (startVal == null) return;
            // datum.start is a numeric ms timestamp after Vega's input transform
            const startTs = typeof startVal === "number" ? startVal : +new Date(startVal);
            if (isNaN(startTs)) return;

            const curXDom = v.signal("xDom") as [number, number];
            const span = curXDom[1] - curXDom[0];
            const offset = span * 0.08;  // place start ~8% from left edge
            v.signal("xDom", [startTs - offset, startTs - offset + span]).runAsync();
        };

        v.addEventListener("click", navigate);
    }

    private detachWheelListener(): void {
        if (this.wheelListener) {
            this.host.removeEventListener("wheel", this.wheelListener);
            this.wheelListener = null;
        }
        this.rafPending    = false;
        this.pendingXDom   = null;
        this.pendingYRange = null;
    }

    private resizeView(w: number, h: number): void {
        const chartW = Math.max(200, w - PAD_LEFT  - PAD_RIGHT);
        const chartH = Math.max(100, h - PAD_TOP   - PAD_BOTTOM);
        this.vegaView
            .width(chartW)
            .height(chartH)
            .runAsync();
    }

    /* ── Map Power BI table rows → Vega dataset rows ── */
    private extractData(dataViews: DataView[]): any[] | null {
        if (!dataViews || !dataViews[0] || !dataViews[0].table) return null;

        const table = dataViews[0].table;
        if (!table.rows || table.rows.length === 0) return null;

        const cols = table.columns;
        const idx: Record<string, number> = {};
        const hierarchyColIndices: number[] = [];

        cols.forEach((col, i) => {
            const roles = Object.keys(col.roles || {});
            if (roles.includes("hierarchy")) {
                hierarchyColIndices.push(i);   // preserve binding order
            } else if (roles.length > 0) {
                idx[roles[0]] = i;
            }
        });

        const hasAny = Object.keys(idx).length > 0 || hierarchyColIndices.length > 0;
        if (!hasAny) return null;

        const todayStr = fmtDate(new Date());

        // Helper: numeric start/end timestamps from a raw row
        const tsOf = (raw: any): number => {
            if (raw == null) return Date.now();
            const d = raw instanceof Date ? raw : new Date(raw as string);
            return isNaN(d.getTime()) ? Date.now() : +d;
        };
        const getStartTs = (row: any) => tsOf(idx["startDate"] !== undefined ? row[idx["startDate"]] : null);
        const getEndTs   = (row: any) => tsOf(idx["endDate"]   !== undefined ? row[idx["endDate"]]   : null);

        // Helper: build one Vega dataset row
        const makeRow = (row: any, rowIdx: number, phaseName: string, groupOrder: number): any => {
            const getStr = (role: string) =>
                idx[role] !== undefined ? String(row[idx[role]] ?? "") : "";
            const getNum = (role: string) =>
                idx[role] !== undefined ? Number(row[idx[role]] ?? 0) : 0;
            const rawStart = idx["startDate"] !== undefined ? row[idx["startDate"]] : null;
            const rawEnd   = idx["endDate"]   !== undefined ? row[idx["endDate"]]   : null;
            const toDateStr = (raw: any, fallback: string): string => {
                if (!raw) return fallback;
                const d = raw instanceof Date ? raw : new Date(raw as string);
                return isNaN(d.getTime()) ? fallback : fmtDate(d);
            };
            const milestoneRaw = getStr("milestone").toLowerCase();
            const milestone    =
                milestoneRaw === "true" || milestoneRaw === "1" || milestoneRaw === "yes" ? true : null;
            const startStr = toDateStr(rawStart, todayStr);
            const endStr   = toDateStr(rawEnd, startStr);
            return {
                id          : idx["id"] !== undefined ? row[idx["id"]] : rowIdx + 1,
                phase       : phaseName,
                task        : getStr("task") || `Task ${rowIdx + 1}`,
                milestone   : milestone,
                start       : startStr,
                end         : endStr,
                completion  : Math.min(100, Math.max(0, Math.round(getNum("completion")))),
                dependencies: getStr("dependencies"),
                assignee    : getStr("assignee"),
                status      : getStr("status"),
                hyperlink   : "",
                _groupOrder : groupOrder
            };
        };

        // ── Multi-level hierarchy: Level1 = phase, Level2+ = sub-group headers ──
        if (hierarchyColIndices.length >= 2) {
            const level1Idx  = hierarchyColIndices[0];
            const subIndices = hierarchyColIndices.slice(1);

            // Pre-pass: group rows by (phase, subKey)
            type Group = { minStart: number; maxEnd: number; items: { row: any; rowIdx: number; startTs: number }[] };
            const groupMap = new Map<string, Map<string, Group>>();

            table.rows.forEach((row, rowIdx) => {
                const phase  = String(row[level1Idx] ?? "").trim() || "Tasks";
                const subKey = subIndices.map(i => String(row[i] ?? "").trim()).filter(v => v).join(" | ");
                if (!groupMap.has(phase))    groupMap.set(phase, new Map());
                const pg = groupMap.get(phase)!;
                if (!pg.has(subKey)) pg.set(subKey, { minStart: Infinity, maxEnd: -Infinity, items: [] });
                const g = pg.get(subKey)!;
                const s = getStartTs(row);
                const e = getEndTs(row);
                g.minStart = Math.min(g.minStart, s);
                g.maxEnd   = Math.max(g.maxEnd,   e);
                g.items.push({ row, rowIdx, startTs: s });
            });

            const result: any[] = [];

            for (const [phase, phaseGroups] of groupMap) {
                // Sort sub-groups chronologically
                const subEntries = Array.from(phaseGroups.entries())
                    .sort((a, b) => a[1].minStart - b[1].minStart);

                subEntries.forEach(([subKey, group], sgIdx) => {
                    const base = (sgIdx + 1) * 100000;

                    // Synthetic sub-header row (visible as a task bar spanning the sub-group)
                    result.push({
                        id          : `sub|${phase}|${subKey}`,
                        phase       : phase,
                        task        : `  \u25B8 ${subKey}`,  // "  ▸ Discipline"
                        milestone   : null,
                        start       : fmtDate(new Date(isFinite(group.minStart) ? group.minStart : Date.now())),
                        end         : fmtDate(new Date(isFinite(group.maxEnd)   ? group.maxEnd   : Date.now())),
                        completion  : 0,
                        dependencies: "",
                        assignee    : "",
                        status      : "",
                        hyperlink   : "",
                        _groupOrder : base
                    });

                    // Task rows sorted by start date
                    group.items
                        .sort((a, b) => a.startTs - b.startTs)
                        .forEach((item, ti) => {
                            result.push(makeRow(item.row, item.rowIdx, phase, base + ti + 1));
                        });
                });
            }

            return result.length > 0 ? result : null;
        }

        // ── Single hierarchy level: phase = level1, sort by date ──
        if (hierarchyColIndices.length === 1) {
            const level1Idx = hierarchyColIndices[0];
            return table.rows.map((row, rowIdx) => {
                const phase = String(row[level1Idx] ?? "").trim() || "Tasks";
                return makeRow(row, rowIdx, phase, getStartTs(row));
            });
        }

        // ── No hierarchy columns bound: default single group ──
        return table.rows.map((row, rowIdx) =>
            makeRow(row, rowIdx, "Tasks", getStartTs(row))
        );
    }

    public destroy(): void {
        this.detachWheelListener();
        if (this.vegaView) {
            try { this.vegaView.finalize(); } catch (_) { /* ignore */ }
            this.vegaView = null;
        }
    }
}
