"use strict";

import { formattingSettings } from "powerbi-visuals-utils-formattingmodel";

import FormattingSettingsCard  = formattingSettings.SimpleCard;
import FormattingSettingsSlice = formattingSettings.Slice;
import FormattingSettingsModel = formattingSettings.Model;

// 値（セル）の書式設定
class ValuesCardSettings extends FormattingSettingsCard {
    font = new formattingSettings.FontControl({
        name: "font",
        fontFamily: new formattingSettings.FontPicker({
            name: "fontFamily",
            displayName: "フォント",
            value: "Segoe UI",
        }),
        fontSize: new formattingSettings.NumUpDown({
            name: "fontSize",
            displayName: "テキスト サイズ",
            value: 12,
            options: { minValue: { value: 8, type: 0 as never }, maxValue: { value: 36, type: 1 as never } },
        }),
        bold: new formattingSettings.ToggleSwitch({
            name: "bold",
            displayName: "太字",
            value: false,
        }),
        italic: new formattingSettings.ToggleSwitch({
            name: "italic",
            displayName: "斜体",
            value: false,
        }),
        underline: new formattingSettings.ToggleSwitch({
            name: "underline",
            displayName: "下線",
            value: false,
        }),
    });

    fontColor = new formattingSettings.ColorPicker({
        name: "fontColor",
        displayName: "テキストの色",
        value: { value: "#252423" },
    });

    backgroundColor = new formattingSettings.ColorPicker({
        name: "backgroundColor",
        displayName: "背景色",
        value: { value: "#ffffff" },
    });

    altFontColor = new formattingSettings.ColorPicker({
        name: "altFontColor",
        displayName: "代替テキストの色",
        value: { value: "#252423" },
    });

    altBackgroundColor = new formattingSettings.ColorPicker({
        name: "altBackgroundColor",
        displayName: "代替の背景色",
        value: { value: "#fafafa" },
    });

    wordWrap = new formattingSettings.ToggleSwitch({
        name: "wordWrap",
        displayName: "テキストの折り返し",
        value: false,
    });

    name = "values";
    displayName = "値";
    slices: FormattingSettingsSlice[] = [
        this.font,
        this.fontColor, this.backgroundColor,
        this.altFontColor, this.altBackgroundColor,
        this.wordWrap,
    ];
}

// 列見出しの書式設定
class ColumnHeaderCardSettings extends FormattingSettingsCard {
    font = new formattingSettings.FontControl({
        name: "font",
        fontFamily: new formattingSettings.FontPicker({
            name: "fontFamily",
            displayName: "フォント",
            value: "Segoe UI",
        }),
        fontSize: new formattingSettings.NumUpDown({
            name: "fontSize",
            displayName: "テキスト サイズ",
            value: 13,
            options: { minValue: { value: 8, type: 0 as never }, maxValue: { value: 36, type: 1 as never } },
        }),
        bold: new formattingSettings.ToggleSwitch({
            name: "bold",
            displayName: "太字",
            value: true,
        }),
        italic: new formattingSettings.ToggleSwitch({
            name: "italic",
            displayName: "斜体",
            value: false,
        }),
        underline: new formattingSettings.ToggleSwitch({
            name: "underline",
            displayName: "下線",
            value: false,
        }),
    });

    fontColor = new formattingSettings.ColorPicker({
        name: "fontColor",
        displayName: "テキストの色",
        value: { value: "#252423" },
    });

    backgroundColor = new formattingSettings.ColorPicker({
        name: "backgroundColor",
        displayName: "背景色",
        value: { value: "#f2f2f2" },
    });

    name = "columnHeader";
    displayName = "列見出し";
    slices: FormattingSettingsSlice[] = [
        this.font,
        this.fontColor, this.backgroundColor,
    ];
}

export class VisualFormattingSettingsModel extends FormattingSettingsModel {
    valuesCard = new ValuesCardSettings();
    columnHeaderCard = new ColumnHeaderCardSettings();
    cards = [this.valuesCard, this.columnHeaderCard];
}
