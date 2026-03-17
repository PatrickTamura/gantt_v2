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

    // RAF scroll state
    private wheelListener: ((e: WheelEvent) => void) | null = null;
    private rafPending   = false;
    private pendingXDom:   [number, number] | null = null;
    private pendingYRange: [number, number] | null = null;
    private vScrollEl:     HTMLInputElement | null = null;

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
            this.createView(w, h, pbData ?? undefined);
        } else if (sizeChanged) {
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
            this.attachWheelListener();
            this.buildScrollbar();
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
            const dy = e.deltaY * norm;

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
        const wbsNodeMap = new Map<string, { name: string; parentId: string; parentName: string }>();
        if (hasWbs) {
            for (const r of table.rows) {
                const nodeId     = String(r[idx["id"]]              ?? "").trim();
                const nodeName   = String(r[idx["task"]]            ?? "").trim();
                const parentId   = String(r[idx["wbsParentId"]]     ?? "").trim();
                // wbsParentName (optional): display name of the parent node supplied
                // directly on each row — avoids needing a full tree walk.
                const parentName = idx["wbsParentName"] !== undefined
                    ? String(r[idx["wbsParentName"]] ?? "").trim()
                    : "";
                if (nodeId) wbsNodeMap.set(nodeId, { name: nodeName, parentId, parentName });
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
                const selfId   = String(row[idx["id"]] ?? "").trim();
                const selfNode = wbsNodeMap.get(selfId);
                const parentId = selfNode?.parentId ?? "";

                if (idx["wbsParentName"] !== undefined) {
                    // Fast path: parent name is provided directly on each row.
                    // Prefer the stored parentName, fall back to the parent node's
                    // own name in the map, then the raw parentId as last resort.
                    const parent = wbsNodeMap.get(parentId);
                    phaseName = selfNode?.parentName || parent?.name || parentId || "Tasks";
                } else {
                    // Slow path: walk the full ancestor chain upward.
                    // Builds a path string like "Phase > Sub-Phase > Package".
                    // Visited set protects against circular references.
                    const path:    string[]  = [];
                    const visited: Set<string> = new Set();
                    let   pid                  = parentId;
                    while (pid && !visited.has(pid)) {
                        visited.add(pid);
                        const anc = wbsNodeMap.get(pid);
                        if (!anc) break;
                        path.unshift(anc.name);
                        pid = anc.parentId;
                    }
                    phaseName = path.length > 0
                        ? path.join(" > ")
                        : (selfNode?.name ?? "Tasks");
                }

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
