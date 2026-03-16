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

/* ─── Virtualisation / performance constants ───────────────────────────── */
const ROW_HEIGHT     = 33;   // px – must match yRowHeight signal in spec
const VIRT_BUFFER    = 15;   // extra rows to render above/below viewport
const DEPS_THRESHOLD = 150;  // hide dependency lines when visible rows exceed this

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
    private vegaView: any = null;
    private lastWidth  = 0;
    private lastHeight = 0;
    private lastDataHash = "";

    // Full dataset kept in TypeScript; only a windowed slice is pushed to Vega
    private allRows: any[] = [];
    private virtualScrollY = 0;       // current vertical scroll offset in px
    private viewportHeight = 400;     // updated on every update() call

    // RAF scroll state
    private wheelListener: ((e: WheelEvent) => void) | null = null;
    private rafPending           = false;
    private pendingXDom: [number, number] | null = null;
    private pendingVirtualRefresh = false;  // true = push a new data slice in next RAF

    constructor(options: VisualConstructorOptions) {
        this.host = options.element;
        this.host.classList.add("gantt-deneb-host");
    }

    public update(options: VisualUpdateOptions): void {
        const vp = options.viewport;
        const w  = Math.max(200, Math.floor(vp.width));
        const h  = Math.max(100, Math.floor(vp.height));
        this.viewportHeight = h;

        // Extract Power BI data (returns null when no data is bound)
        const pbData = this.extractData(options.dataViews);
        const dataHash = pbData ? JSON.stringify(pbData) : "";

        const dataChanged = (dataHash !== this.lastDataHash);
        if (dataChanged) {
            this.allRows = pbData ?? [];
            this.virtualScrollY = 0;   // reset to top whenever data changes
        }

        const sizeChanged = (w !== this.lastWidth || h !== this.lastHeight);

        if (!this.vegaView || dataChanged) {
            // (Re)build the view from scratch whenever data changes
            this.createView(w, h, this.virtualSlice());
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
            renderer : "canvas",  // Canvas is 10-50x faster than SVG for many marks
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

            // ── Horizontal pan (deltaX) — push xDom signal directly ──────
            if (dx !== 0) {
                const curXDom    = this.pendingXDom ?? (v.signal("xDom") as [number, number]);
                const ganttWidth: number = v.signal("ganttWidth") ?? 1;
                const span  = curXDom[1] - curXDom[0];
                const shift = dx * span / ganttWidth;
                this.pendingXDom = [curXDom[0] + shift, curXDom[1] + shift];
            }

            // ── Vertical pan (deltaY) — managed entirely via data slicing ─
            // Instead of scrolling inside Vega (yRange), we shift the window
            // of rows we send to Vega. This keeps the rendered DOM tiny.
            if (dy !== 0) {
                const maxScrollY = Math.max(
                    0,
                    (this.allRows.length - Math.ceil(this.viewportHeight / ROW_HEIGHT)) * ROW_HEIGHT
                );
                this.virtualScrollY = Math.min(
                    maxScrollY,
                    Math.max(0, this.virtualScrollY + dy)
                );
                this.pendingVirtualRefresh = true;
            }

            // ── Flush once per animation frame ───────────────────────────
            if (!this.rafPending) {
                this.rafPending = true;
                requestAnimationFrame(() => {
                    if (!this.vegaView) { this.rafPending = false; return; }

                    if (this.pendingXDom) {
                        this.vegaView.signal("xDom", this.pendingXDom);
                        this.pendingXDom = null;
                    }

                    if (this.pendingVirtualRefresh) {
                        const slice = this.virtualSlice();
                        const cs = vega.changeset().remove(() => true).insert(slice);
                        this.vegaView.change("dataset", cs);
                        // Reset yRange origin to 0; the spec's update expression
                        // will correct the span to match the new scaledHeight.
                        this.vegaView.signal("yRange", [0, 9999]);
                        this.pendingVirtualRefresh = false;
                    }

                    this.vegaView.runAsync();
                    this.rafPending = false;
                });
            }
        };

        this.host.addEventListener("wheel", this.wheelListener, { passive: false });
    }

    /* ── Returns only the rows visible in the current virtual window ── */
    private virtualSlice(): any[] {
        if (this.allRows.length === 0) return [];
        const viewRows = Math.ceil(this.viewportHeight / ROW_HEIGHT);
        const startIdx = Math.max(0, Math.floor(this.virtualScrollY / ROW_HEIGHT) - VIRT_BUFFER);
        const endIdx   = Math.min(this.allRows.length, startIdx + viewRows + VIRT_BUFFER * 2);
        const slice    = this.allRows.slice(startIdx, endIdx);
        // Strip dependency arrows when too many rows — rendering them is expensive
        if (slice.length > DEPS_THRESHOLD) {
            return slice.map((r: any) => ({ ...r, dependencies: "" }));
        }
        return slice;
    }

    private detachWheelListener(): void {
        if (this.wheelListener) {
            this.host.removeEventListener("wheel", this.wheelListener);
            this.wheelListener = null;
        }
        this.rafPending           = false;
        this.pendingXDom          = null;
        this.pendingVirtualRefresh = false;
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

        // Build a column-name → column-index map from the roles metadata.
        // hierarchyLevels accepts multiple columns; collect them in binding order.
        const cols = table.columns;
        const idx: Record<string, number> = {};
        const hierarchyColIndices: number[] = [];

        cols.forEach((col, i) => {
            const roles = Object.keys(col.roles || {});
            if (roles.includes("hierarchyLevels")) {
                hierarchyColIndices.push(i);   // preserve order assigned by user
            } else if (roles.length > 0) {
                idx[roles[0]] = i;
            }
        });

        const hasAny = Object.keys(idx).length > 0 || hierarchyColIndices.length > 0;
        if (!hasAny) return null;

        const todayStr = fmtDate(new Date());

        // ── WBS mode: pre-build a lookup map for parent-chain walks ─────────────
        // Active when both "id" (= wbs_id) and "wbsParentId" (= parent_wbs_id)
        // are bound. The "task" column carries the wbs_name.
        const hasWbs = idx["wbsParentId"] !== undefined && idx["id"] !== undefined;
        const wbsNodeMap = new Map<string, { name: string; parentId: string }>();
        if (hasWbs) {
            for (const r of table.rows) {
                const nodeId   = String(r[idx["id"]]          ?? "").trim();
                const nodeName = String(r[idx["task"]]        ?? "").trim();
                const parentId = String(r[idx["wbsParentId"]] ?? "").trim();
                if (nodeId) wbsNodeMap.set(nodeId, { name: nodeName, parentId });
            }
        }

        return table.rows.map((row, rowIdx) => {
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
                milestoneRaw === "true" || milestoneRaw === "1" || milestoneRaw === "yes"
                    ? true : null;

            const taskName = getStr("task") || `Task ${rowIdx + 1}`;
            const startStr = toDateStr(rawStart, todayStr);
            const endStr   = toDateStr(rawEnd, startStr);

            // ── Phase resolution — priority: WBS > hierarchyLevels > phase > default
            let phaseName: string;

            if (hasWbs) {
                // Walk the parent chain upward, build path of ancestor names.
                // Guard against circular references with a visited set.
                const selfId = String(row[idx["id"]] ?? "").trim();
                const path: string[] = [];
                const visited = new Set<string>();
                let parentId = wbsNodeMap.get(selfId)?.parentId ?? "";
                while (parentId && !visited.has(parentId)) {
                    visited.add(parentId);
                    const node = wbsNodeMap.get(parentId);
                    if (!node) break;
                    path.unshift(node.name);   // prepend so root → leaf order
                    parentId = node.parentId;
                }
                phaseName = path.length > 0
                    ? path.join(" > ")
                    : (wbsNodeMap.get(selfId)?.name ?? "Tasks");

            } else if (hierarchyColIndices.length > 0) {
                // Concatenate non-empty values from each hierarchy column, in order.
                const parts = hierarchyColIndices
                    .map(i => String(row[i] ?? "").trim())
                    .filter(v => v !== "");
                phaseName = parts.join(" > ") || "Tasks";

            } else if (idx["phase"] !== undefined) {
                phaseName = getStr("phase") || "Tasks";

            } else {
                phaseName = "Tasks";
            }

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
