import { formattingSettings } from "powerbi-visuals-utils-formattingmodel";
import FormattingSettingsCard = formattingSettings.SimpleCard;
import FormattingSettingsSlice = formattingSettings.Slice;
import FormattingSettingsModel = formattingSettings.Model;
declare class TitleCard extends FormattingSettingsCard {
    name: string;
    displayName: string;
    showTitle: formattingSettings.ToggleSwitch;
    titleText: formattingSettings.TextInput;
    titleFontSize: formattingSettings.NumUpDown;
    titleColor: formattingSettings.ColorPicker;
    slices: FormattingSettingsSlice[];
}
declare class GanttOptionsCard extends FormattingSettingsCard {
    name: string;
    displayName: string;
    showTooltips: formattingSettings.ToggleSwitch;
    rowHeight: formattingSettings.NumUpDown;
    showGrid: formattingSettings.ToggleSwitch;
    slices: FormattingSettingsSlice[];
}
declare class GeneralCard extends FormattingSettingsCard {
    name: string;
    displayName: string;
    slices: FormattingSettingsSlice[];
}
export declare class VisualFormattingSettingsModel extends FormattingSettingsModel {
    titleCard: TitleCard;
    ganttOptionsCard: GanttOptionsCard;
    generalCard: GeneralCard;
    cards: (TitleCard | GanttOptionsCard | GeneralCard)[];
}
export {};
