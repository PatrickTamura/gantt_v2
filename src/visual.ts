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
import * as vega from "vega";

import VisualConstructorOptions = powerbi.extensibility.visual.VisualConstructorOptions;
import VisualUpdateOptions   = powerbi.extensibility.visual.VisualUpdateOptions;
import IVisual               = powerbi.extensibility.visual.IVisual;
import DataView              = powerbi.DataView;

/* ─── Padding constants ────────────────────────────────────────────────────
 * The buttons in the spec are at y = -60 (above the chart origin).
 * We need enough top padding so the SVG viewBox includes them.
 * padTop = 70  →  viewBox starts at y = -70, so y = -60 is 10 px inside.
 * ─────────────────────────────────────────────────────────────────────── */
const PAD_LEFT   = 5;
const PAD_RIGHT  = 5;
const PAD_TOP    = 70;
const PAD_BOTTOM = 5;

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
    private vegaView: any = null;
    private lastWidth  = 0;
    private lastHeight = 0;
    private lastDataHash = "";

    // RAF scroll state
    private wheelListener: ((e: WheelEvent) => void) | null = null;
    private rafPending = false;
    private pendingXDom:   [number, number] | null = null;
    private pendingYRange: [number, number] | null = null;

    constructor(options: VisualConstructorOptions) {
        this.host = options.element;
        this.host.classList.add("gantt-deneb-host");
    }

    public update(options: VisualUpdateOptions): void {
        const vp = options.viewport;
        const w  = Math.max(200, Math.floor(vp.width));
        const h  = Math.max(100, Math.floor(vp.height));

        // Extract Power BI data (returns null when no data is bound)
        const pbData = this.extractData(options.dataViews);
        const dataHash = pbData ? JSON.stringify(pbData) : "";

        const sizeChanged = (w !== this.lastWidth || h !== this.lastHeight);
        const dataChanged = (dataHash !== this.lastDataHash);

        if (!this.vegaView || dataChanged) {
            // (Re)build the view from scratch whenever data changes
            this.createView(w, h, pbData ?? undefined);
        } else if (sizeChanged) {
            // Only resize — preserves scroll/zoom state
            this.resizeView(w, h);
        }

        this.lastWidth    = w;
        this.lastHeight   = h;
        this.lastDataHash = dataHash;
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
            renderer : "svg",
            container: this.host,
            hover    : true,
            logLevel : vega.Warn
        });

        this.vegaView.runAsync().then(() => {
            // Attach the native RAF-throttled wheel listener only after first render
            this.attachWheelListener();
        });
    }

    /* ─── RAF-throttled native wheel handler ─────────────────────────────
     * Accumulates all wheel events within one animation frame (~16 ms) and
     * pushes a single signal update to Vega, instead of calling runAsync()
     * on every individual wheel tick (which is what makes it slow).
     * ─────────────────────────────────────────────────────────────────── */
    private attachWheelListener(): void {
        this.wheelListener = (e: WheelEvent) => {
            e.preventDefault();

            const v = this.vegaView;
            if (!v) return;

            // Normalise deltaMode: 0=px, 1=line(≈16px), 2=page(≈256px)
            const norm = e.deltaMode === 0 ? 1 : e.deltaMode === 1 ? 16 : 256;
            const dx = e.deltaX * norm;
            const dy = e.deltaY * norm;

            // Use pending values if a RAF is already queued (accumulate)
            const curXDom   = this.pendingXDom   ?? (v.signal("xDom")   as [number, number]);
            const curYRange = this.pendingYRange  ?? (v.signal("yRange") as [number, number]);

            // ── Horizontal pan (deltaX) ──────────────────────────────────
            if (dx !== 0) {
                const ganttWidth: number = v.signal("ganttWidth") ?? 1;
                const span = curXDom[1] - curXDom[0];
                const shift = dx * span / ganttWidth;
                this.pendingXDom = [curXDom[0] + shift, curXDom[1] + shift];
            }

            // ── Vertical pan (deltaY) ────────────────────────────────────
            if (dy !== 0) {
                const height:       number = v.signal("height")       ?? 500;
                const scaledHeight: number = v.signal("scaledHeight") ?? height;
                const span   = curYRange[1] - curYRange[0];
                const minY   = height >= scaledHeight ? 0              : height - scaledHeight;
                const maxY   = height >= scaledHeight ? height         : scaledHeight;
                const newY0  = Math.min(maxY - span, Math.max(minY, curYRange[0] + dy));
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

        // Build a column-name → column-index map from the roles metadata
        const cols = table.columns;
        const idx: Record<string, number> = {};
        cols.forEach((col, i) => {
            const roles = Object.keys(col.roles || {});
            if (roles.length > 0) idx[roles[0]] = i;
        });

        // Require at least one real field to be bound before replacing the built-in data.
        // This avoids the chart silently falling back to the hardcoded spec values when
        // the user has bound some — but not all — fields.
        const hasAny = Object.keys(idx).length > 0;
        if (!hasAny) return null;

        // Today's date string used as fallback when start/end are not bound
        const todayStr = fmtDate(new Date());

        return table.rows.map((row, rowIdx) => {
            const getStr = (role: string) =>
                idx[role] !== undefined ? String(row[idx[role]] ?? "") : "";
            const getNum = (role: string) =>
                idx[role] !== undefined ? Number(row[idx[role]] ?? 0) : 0;

            // Dates come from Power BI as Date objects or ISO strings
            const rawStart = idx["startDate"] !== undefined ? row[idx["startDate"]] : null;
            const rawEnd   = idx["endDate"]   !== undefined ? row[idx["endDate"]]   : null;
            const toDateStr = (raw: any, fallback: string): string => {
                if (!raw) return fallback;
                const d = raw instanceof Date ? raw : new Date(raw as string);
                return isNaN(d.getTime()) ? fallback : fmtDate(d);
            };

            const milestoneRaw = getStr("milestone").toLowerCase();
            const milestone    =
                milestoneRaw === "true" || milestoneRaw === "1" || milestoneRaw === "yes"
                    ? true : null;

            const taskName  = getStr("task") || `Task ${rowIdx + 1}`;
            // If no phase field is bound, group everything under a single "Tasks" phase
            const phaseName = idx["phase"] !== undefined ? getStr("phase") : "Tasks";
            const startStr  = toDateStr(rawStart, todayStr);
            // Default end = start (same-day bar) when not provided
            const endStr    = toDateStr(rawEnd, startStr);

            return {
                id          : idx["id"] !== undefined ? row[idx["id"]] : rowIdx + 1,
                phase       : phaseName,
                task        : taskName,
                milestone   : milestone,
                start       : startStr,
                end         : endStr,
                completion  : Math.min(100, Math.max(0, Math.round(getNum("completion")))),
                dependencies: getStr("dependencies"),
                assignee    : getStr("assignee"),
                status      : getStr("status"),
                hyperlink   : ""
            };
        });
    }

    public destroy(): void {
        this.detachWheelListener();
        if (this.vegaView) {
            try { this.vegaView.finalize(); } catch (_) { /* ignore */ }
            this.vegaView = null;
        }
    }
}
