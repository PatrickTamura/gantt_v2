import powerbi from "powerbi-visuals-api";
import "./../style/visual.less";
import VisualConstructorOptions = powerbi.extensibility.visual.VisualConstructorOptions;
import VisualUpdateOptions = powerbi.extensibility.visual.VisualUpdateOptions;
import IVisual = powerbi.extensibility.visual.IVisual;
export declare class Visual implements IVisual {
    private host;
    private vegaView;
    private lastWidth;
    private lastHeight;
    private lastDataHash;
    private wheelListener;
    private rafPending;
    private pendingXDom;
    private pendingYRange;
    constructor(options: VisualConstructorOptions);
    update(options: VisualUpdateOptions): void;
    private createView;
    private attachWheelListener;
    private detachWheelListener;
    private resizeView;
    private extractData;
    destroy(): void;
}
