"use strict";

import { formattingSettings } from "powerbi-visuals-utils-formattingmodel";
import FormattingSettingsCard = formattingSettings.SimpleCard;
import FormattingSettingsSlice = formattingSettings.Slice;
import FormattingSettingsModel = formattingSettings.Model;

class TitleCard extends FormattingSettingsCard {
    name = "title";
    displayName = "Title";

    showTitle = new formattingSettings.ToggleSwitch({
        name: "showTitle",
        displayName: "Show Title",
        value: false
    });
    titleText = new formattingSettings.TextInput({
        name: "titleText",
        displayName: "Title Text",
        value: "Gantt Chart",
        placeholder: "Enter title"
    });
    titleFontSize = new formattingSettings.NumUpDown({
        name: "titleFontSize",
        displayName: "Font Size",
        value: 14
    });
    titleColor = new formattingSettings.ColorPicker({
        name: "titleColor",
        displayName: "Font Color",
        value: { value: "#333333" }
    });

    slices: FormattingSettingsSlice[] = [
        this.showTitle,
        this.titleText,
        this.titleFontSize,
        this.titleColor
    ];
}

class GanttOptionsCard extends FormattingSettingsCard {
    name = "ganttOptions";
    displayName = "Gantt Options";

    showTooltips = new formattingSettings.ToggleSwitch({
        name: "showTooltips",
        displayName: "Show Tooltips",
        value: true
    });
    rowHeight = new formattingSettings.NumUpDown({
        name: "rowHeight",
        displayName: "Row Height",
        value: 24
    });
    showGrid = new formattingSettings.ToggleSwitch({
        name: "showGrid",
        displayName: "Show Grid Lines",
        value: true
    });

    slices: FormattingSettingsSlice[] = [
        this.showTooltips,
        this.rowHeight,
        this.showGrid
    ];
}

class GeneralCard extends FormattingSettingsCard {
    name = "general";
    displayName = "General";
    slices: FormattingSettingsSlice[] = [];
}

export class VisualFormattingSettingsModel extends FormattingSettingsModel {
    titleCard = new TitleCard();
    ganttOptionsCard = new GanttOptionsCard();
    generalCard = new GeneralCard();
    cards = [this.titleCard, this.ganttOptionsCard, this.generalCard];
}
