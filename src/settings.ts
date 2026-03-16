"use strict";

import { formattingSettings } from "powerbi-visuals-utils-formattingmodel";
import FormattingSettingsCard = formattingSettings.SimpleCard;
import FormattingSettingsModel = formattingSettings.Model;

class GeneralCard extends FormattingSettingsCard {
    name = "general";
    displayName = "General";
    slices = [];
}

export class VisualFormattingSettingsModel extends FormattingSettingsModel {
    generalCard = new GeneralCard();
    cards = [this.generalCard];
}
