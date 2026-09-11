import powerbiVisualsApi from "powerbi-visuals-api";
import { formattingSettings } from "powerbi-visuals-utils-formattingmodel";
import { legendInterfaces } from "powerbi-visuals-utils-chartutils";
import { ColorHelper } from "powerbi-visuals-utils-colorutils";
import LegendPosition = legendInterfaces.LegendPosition;

import { TornadoChartSeries } from "./interfaces"

import Card = formattingSettings.SimpleCard;
import CompositeCard = formattingSettings.CompositeCard;
import Model = formattingSettings.Model;

// Power BI represents an automatic numeric value as null, while NumUpDown.value is typed as number.
const AutoNumericValue = null as unknown as number;

import IEnumMember = powerbi.IEnumMember;
import ILocalizationManager = powerbi.extensibility.ILocalizationManager;
import { LegendData } from "powerbi-visuals-utils-chartutils/lib/legend/legendInterfaces";

export const enum TornadoObjectNames {
    Legend = "legend",
    LegendTitle = "legendTitle",
    Categories = "categories",
    DataPoint = "dataPoint",
    Labels = "labels",
    NegativeBars = "negativeBars",
    BarAppearance = "barAppearance",
    CenterLine = "centerLine",
    ChartArea = "chartArea",
    CategoryAxis = "categoryAxis",
}

class DataColorCardSettings extends Card {
    fill = new formattingSettings.ColorPicker({
        name: "fill",
        displayName: "Fill",
        displayNameKey: "Visual_Fill",
        value: { value: "" }
    });

    name: string = TornadoObjectNames.DataPoint;
    displayName: string = "Bars";
    displayNameKey: string = "Visual_Bars";
    description: string = "Display bar color options";
    descriptionKey: string = "Visual_Description_Bars";
    slices = [this.fill];
}

class CategoryAxisOptionsGroup extends Card {
    constructor(normalize: formattingSettings.ToggleSwitch) {
        super();
        this.slices = [normalize];
    }

    name: string = "categoryAxisOptions";
    displayName: string = "Options";
    displayNameKey: string = "Visual_Options";
    slices: formattingSettings.Slice[];
}

class CategoryAxisRangeGroup extends Card {
    constructor(
        index: number,
        displayName: string,
        slices: formattingSettings.Slice[]) {
        super();
        this.name = `categoryAxisRange${index}`;
        this.displayName = displayName;
        this.slices = slices;
    }

    name: string;
    displayName: string;
    slices: formattingSettings.Slice[];
}

class CategoryAxisCardSettings extends CompositeCard {
    normalize = new formattingSettings.ToggleSwitch({
        name: "normalize",
        displayName: "Normalize to 100%",
        displayNameKey: "Visual_Axis_Normalize",
        value: false
    });

    name: string = "categoryAxis";
    displayName: string = "X-axis";
    displayNameKey: string = "Visual_XAxis";
    optionsGroup = new CategoryAxisOptionsGroup(this.normalize);
    groups: formattingSettings.Group[] = [this.optionsGroup];
}

class NegativeBarsColorGroup extends Card {
    constructor(slices: formattingSettings.Slice[]) {
        super();
        this.slices = slices;
    }

    name: string = "negativeBarsColor";
    displayName: string = "Color";
    displayNameKey: string = "Visual_Color";
    slices: formattingSettings.Slice[];
}

class NegativeBarsBorderGroup extends Card {
    constructor(slices: formattingSettings.Slice[], topLevelSlice: formattingSettings.ToggleSwitch) {
        super();
        this.slices = slices;
        this.topLevelSlice = topLevelSlice;
    }

    name: string = "negativeBarsBorder";
    displayName: string = "Border";
    displayNameKey: string = "Visual_Border";
    slices: formattingSettings.Slice[];
    topLevelSlice: formattingSettings.ToggleSwitch;
}

class NegativeBarsCardSettings extends CompositeCard {
    show = new formattingSettings.ToggleSwitch({
        name: "show",
        displayName: "Show",
        displayNameKey: "Visual_Show",
        value: false
    });

    topLevelSlice? = this.show;

    fill = new formattingSettings.ColorPicker({
        name: "fill",
        displayName: "Fill",
        displayNameKey: "Visual_Fill",
        value: { value: "" }
    });

    transparency = new formattingSettings.Slider({
        name: "transparency",
        displayName: "Transparency",
        displayNameKey: "Visual_Transparency",
        value: 100,
        options: {
            unitSymbol: "%",
            unitSymbolAfterInput: true,
            minValue: {
                type: powerbiVisualsApi.visuals.ValidatorType.Min,
                value: 0,
            },
            maxValue: {
                type: powerbiVisualsApi.visuals.ValidatorType.Max,
                value: 100,
            }
        }
    });

    showBorder = new formattingSettings.ToggleSwitch({
        name: "showBorder",
        displayName: "Border",
        displayNameKey: "Visual_Border",
        value: true
    });

    borderColor = new formattingSettings.ColorPicker({
        name: "borderColor",
        displayName: "Color",
        displayNameKey: "Visual_Color",
        value: { value: "" }
    });

    borderWidth = new formattingSettings.Slider({
        name: "borderWidth",
        displayName: "Width",
        displayNameKey: "Visual_Width",
        value: 2,
        options: {
            unitSymbol: "px",
            unitSymbolAfterInput: true,
            minValue: {
                type: powerbiVisualsApi.visuals.ValidatorType.Min,
                value: 0,
            },
            maxValue: {
                type: powerbiVisualsApi.visuals.ValidatorType.Max,
                value: 10,
            }
        }
    });

    cornerRadius = new formattingSettings.Slider({
        name: "cornerRadius",
        displayName: "Rounded corners",
        displayNameKey: "Visual_CornerRadius",
        value: 0,
        options: {
            unitSymbol: "px",
            unitSymbolAfterInput: true,
            minValue: {
                type: powerbiVisualsApi.visuals.ValidatorType.Min,
                value: 0,
            },
            maxValue: {
                type: powerbiVisualsApi.visuals.ValidatorType.Max,
                value: 50,
            }
        }
    });

    name: string = "negativeBars";
    displayName: string = "Negative bars";
    displayNameKey: string = "Visual_NegativeBars";
    colorGroup = new NegativeBarsColorGroup([this.fill, this.transparency]);
    borderGroup = new NegativeBarsBorderGroup([this.borderColor, this.borderWidth, this.cornerRadius], this.showBorder);
    groups: formattingSettings.Group[] = [this.colorGroup, this.borderGroup];
}

class BarAppearanceBorderGroup extends Card {
    constructor(slices: formattingSettings.Slice[], topLevelSlice: formattingSettings.ToggleSwitch) {
        super();
        this.slices = slices;
        this.topLevelSlice = topLevelSlice;
    }

    name: string = "barAppearanceBorder";
    displayName: string = "Border";
    displayNameKey: string = "Visual_Border";
    slices: formattingSettings.Slice[];
    topLevelSlice: formattingSettings.ToggleSwitch;
}

class BarAppearanceLayoutGroup extends Card {
    constructor(slices: formattingSettings.Slice[]) {
        super();
        this.slices = slices;
    }

    name: string = "barAppearanceLayout";
    displayName: string = "Layout";
    displayNameKey: string = "Visual_Layout";
    slices: formattingSettings.Slice[];
}

class BarAppearanceCardSettings extends CompositeCard {
    showBorder = new formattingSettings.ToggleSwitch({
        name: "showBorder",
        displayName: "Border",
        displayNameKey: "Visual_Border",
        value: false
    });

    borderColor = new formattingSettings.ColorPicker({
        name: "borderColor",
        displayName: "Color",
        displayNameKey: "Visual_Color",
        value: { value: "" }
    });

    borderWidth = new formattingSettings.Slider({
        name: "borderWidth",
        displayName: "Width",
        displayNameKey: "Visual_Width",
        value: 2,
        options: {
            unitSymbol: "px",
            unitSymbolAfterInput: true,
            minValue: {
                type: powerbiVisualsApi.visuals.ValidatorType.Min,
                value: 0,
            },
            maxValue: {
                type: powerbiVisualsApi.visuals.ValidatorType.Max,
                value: 10,
            }
        }
    });

    cornerRadius = new formattingSettings.Slider({
        name: "cornerRadius",
        displayName: "Rounded corners",
        displayNameKey: "Visual_CornerRadius",
        value: 0,
        options: {
            unitSymbol: "px",
            unitSymbolAfterInput: true,
            minValue: {
                type: powerbiVisualsApi.visuals.ValidatorType.Min,
                value: 0,
            },
            maxValue: {
                type: powerbiVisualsApi.visuals.ValidatorType.Max,
                value: 50,
            }
        }
    });

    barSpacing = new formattingSettings.Slider({
        name: "barSpacing",
        displayName: "Space between bars",
        displayNameKey: "Visual_BarSpacing",
        value: 16,
        options: {
            unitSymbol: "%",
            unitSymbolAfterInput: true,
            minValue: {
                type: powerbiVisualsApi.visuals.ValidatorType.Min,
                value: 0,
            },
            maxValue: {
                type: powerbiVisualsApi.visuals.ValidatorType.Max,
                value: 75,
            }
        }
    });

    name: string = "barAppearance";
    displayName: string = "Bar appearance";
    displayNameKey: string = "Visual_BarAppearance";
    borderGroup = new BarAppearanceBorderGroup([this.borderColor, this.borderWidth, this.cornerRadius], this.showBorder);
    layoutGroup = new BarAppearanceLayoutGroup([this.barSpacing]);
    groups: formattingSettings.Group[] = [this.borderGroup, this.layoutGroup];
}

class CenterLineCardSettings extends Card {
    show = new formattingSettings.ToggleSwitch({
        name: "show",
        displayName: "Show",
        displayNameKey: "Visual_Show",
        value: true
    });

    topLevelSlice? = this.show;

    color = new formattingSettings.ColorPicker({
        name: "color",
        displayName: "Color",
        displayNameKey: "Visual_Color",
        value: { value: "" }
    });

    width = new formattingSettings.Slider({
        name: "width",
        displayName: "Width",
        displayNameKey: "Visual_Width",
        value: 1,
        options: {
            unitSymbol: "px",
            unitSymbolAfterInput: true,
            minValue: {
                type: powerbiVisualsApi.visuals.ValidatorType.Min,
                value: 1,
            },
            maxValue: {
                type: powerbiVisualsApi.visuals.ValidatorType.Max,
                value: 10,
            }
        }
    });

    name: string = "centerLine";
    displayName: string = "Center line";
    displayNameKey: string = "Visual_CenterLine";
    slices = [this.color, this.width];
}

class ChartAreaCardSettings extends Card {
    show = new formattingSettings.ToggleSwitch({
        name: "show",
        displayName: "Show",
        displayNameKey: "Visual_Show",
        value: false
    });

    topLevelSlice? = this.show;

    backgroundColor = new formattingSettings.ColorPicker({
        name: "backgroundColor",
        displayName: "Color",
        displayNameKey: "Visual_Color",
        value: { value: "" }
    });

    name: string = "chartArea";
    displayName: string = "Plot area background";
    displayNameKey: string = "Visual_PlotAreaBackground";
    slices = [this.backgroundColor];
}

export enum LabelDisplayMode {
    Value = "value",
    Percentage = "percentage",
    ValueAndPercentage = "valueAndPercentage",
}

export enum LabelPosition {
    Auto = "auto",
    OutsideEnd = "outsideEnd",
    InsideEnd = "insideEnd",
    InsideCenter = "insideCenter",
    InsideBase = "insideBase",
}

export const contentOptions: IEnumMemberWithDisplayNameKey[] = [
    { value: LabelDisplayMode.Value, displayName: "Value", key: "Visual_Value" },
    { value: LabelDisplayMode.Percentage, displayName: "%", key: "Visual_Percentage" },
    { value: LabelDisplayMode.ValueAndPercentage, displayName: "Value (%)", key: "Visual_ValueAndPercentage" },
];

export const dataLabelPositionOptions: IEnumMemberWithDisplayNameKey[] = [
    { value: LabelPosition.Auto, displayName: "Auto", key: "Visual_Position_Auto" },
    { value: LabelPosition.InsideEnd, displayName: "Inside end", key: "Visual_Position_InsideEnd" },
    { value: LabelPosition.OutsideEnd, displayName: "Outside end", key: "Visual_Position_OutsideEnd" },
    { value: LabelPosition.InsideCenter, displayName: "Inside center", key: "Visual_Position_InsideCenter" },
    { value: LabelPosition.InsideBase, displayName: "Inside base", key: "Visual_Position_InsideBase" },
];

class LabelsOptionsGroup extends Card {
    displayFormat = new formattingSettings.ItemDropdown({
        name: "displayFormat",
        displayName: "Content",
        displayNameKey: "Visual_Content",
        items: contentOptions,
        value: contentOptions[0]
    });

    position = new formattingSettings.ItemDropdown({
        name: "position",
        displayName: "Position",
        displayNameKey: "Visual_Position",
        items: dataLabelPositionOptions,
        value: dataLabelPositionOptions[0]
    });

    name: string = "options";
    displayName: string = "Options";
    displayNameKey: string = "Visual_Options";
    slices: formattingSettings.Slice[] = [this.displayFormat, this.position];
}

class LabelsValuesGroup extends Card {
    font: formattingSettings.FontControl = new BaseFontControlSettings(9);

    labelPrecision = new formattingSettings.NumUpDown({
        name: "labelPrecision",
        displayName: "Decimal places",
        displayNameKey: "Visual_DataLabels_DecimalPlaces",
        value: AutoNumericValue,
        options: {
            minValue: {
                type: powerbiVisualsApi.visuals.ValidatorType.Min,
                value: 0,
            },
            maxValue: {
                type: powerbiVisualsApi.visuals.ValidatorType.Max,
                value: 25,
            }
        }
    });

    percentagePrecision = new formattingSettings.NumUpDown({
        name: "percentagePrecision",
        displayName: "Percentage decimal places",
        displayNameKey: "Visual_DataLabels_PercentageDecimalPlaces",
        value: AutoNumericValue,
        options: {
            minValue: {
                type: powerbiVisualsApi.visuals.ValidatorType.Min,
                value: 0,
            },
            maxValue: {
                type: powerbiVisualsApi.visuals.ValidatorType.Max,
                value: 10,
            }
        }
    });

    labelDisplayUnits = new formattingSettings.AutoDropdown({
        name: "labelDisplayUnits",
        displayName: "Display units",
        displayNameKey: "Visual_DisplayUnits",
        value: 1
    });

    insideFill = new formattingSettings.ColorPicker({
        name: "insideFill",
        displayName: "Inside fill",
        displayNameKey: "Visual_DataLabels_InsideFill",
        value: { value: "" }
    });

    outsideFill = new formattingSettings.ColorPicker({
        name: "outsideFill",
        displayName: "Outside fill",
        displayNameKey: "Visual_DataLabels_OutsideFill",
        value: { value: "" }
    });

    negativeFill = new formattingSettings.ColorPicker({
        name: "negativeFill",
        displayName: "Negative fill",
        displayNameKey: "Visual_DataLabels_NegativeFill",
        value: { value: "" }
    });

    name: string = "values";
    displayName: string = "Values";
    displayNameKey: string = "Visual_Values";
    slices = [this.font, this.labelPrecision, this.percentagePrecision, this.labelDisplayUnits];
}

class LabelsColorGroup extends Card {
    constructor(valuesGroup: LabelsValuesGroup) {
        super();
        this.slices = [valuesGroup.insideFill, valuesGroup.outsideFill, valuesGroup.negativeFill];
    }

    name: string = "color";
    displayName: string = "Color";
    displayNameKey: string = "Visual_Color";
    slices: formattingSettings.Slice[];
}

export class DataLabelSettings extends CompositeCard {
    show = new formattingSettings.ToggleSwitch({
        name: "show",
        displayName: "Show",
        displayNameKey: "Visual_Show",
        value: true,
    });
    
    topLevelSlice = this.show;

    public labelsOptionsGroup: LabelsOptionsGroup = new LabelsOptionsGroup();
    public labelsValuesGroup: LabelsValuesGroup = new LabelsValuesGroup();
    public labelsColorGroup: LabelsColorGroup = new LabelsColorGroup(this.labelsValuesGroup);

    name: string = "labels";
    displayName: string = "Data labels";
    displayNameKey: string = "Visual_DataLabels";
    groups: formattingSettings.Group[] = [this.labelsOptionsGroup, this.labelsValuesGroup, this.labelsColorGroup];
}

interface IEnumMemberWithDisplayNameKey extends IEnumMember{
    key: string;
}

const positionOptions : IEnumMemberWithDisplayNameKey[] = [
    {value : LegendPosition[LegendPosition.Top], displayName : "Top", key: "Visual_Legend_Position_Top"}, 
    {value : LegendPosition[LegendPosition.Bottom], displayName : "Bottom", key: "Visual_Legend_Position_Bottom"},
    {value : LegendPosition[LegendPosition.Left], displayName : "Left", key: "Visual_Legend_Position_Left"}, 
    {value : LegendPosition[LegendPosition.Right], displayName : "Right", key: "Visual_Legend_Position_Right"}, 
    {value : LegendPosition[LegendPosition.TopCenter], displayName : "Top center", key: "Visual_Legend_Position_Top_Center"},
    {value : LegendPosition[LegendPosition.BottomCenter], displayName : "Bottom center", key: "Visual_Legend_Position_Bottom_Center"},
    {value : LegendPosition[LegendPosition.LeftCenter], displayName : "Left center", key: "Visual_Legend_Position_Left_Center"},
    {value : LegendPosition[LegendPosition.RightCenter], displayName : "Right center", key: "Visual_Legend_Position_Right_Center"},
];

class BaseFontCardSettings extends formattingSettings.FontControl {
    private static fontFamilyName: string = "fontFamily";
    private static fontSizeName: string = "fontSize";
    private static boldName: string = "fontBold";
    private static italicName: string = "fontItalic";
    private static underlineName: string = "fontUnderline";
    private static fontName: string = "font";
    public static defaultFontFamily: string = "wf_standard-font, helvetica, arial, sans-serif";
    public static minFontSize: number = 8;
    public static maxFontSize: number = 60;
    constructor(defaultFontSize: number, settingName: string = "", fontFamily: string = BaseFontCardSettings.defaultFontFamily){
        super(
            new formattingSettings.FontControl({
                name: BaseFontCardSettings.fontName + settingName,
                displayName: "Font",
                displayNameKey: "Visual_FontControl",
                fontFamily: new formattingSettings.FontPicker({
                    name: BaseFontCardSettings.fontFamilyName + settingName,
                    value: fontFamily
                }),
                fontSize: new formattingSettings.NumUpDown({
                    name: BaseFontCardSettings.fontSizeName + settingName,
                    displayNameKey: "Visual_TextSize",
                    value: defaultFontSize,
                    options: {
                        minValue: {
                            type: powerbi.visuals.ValidatorType.Min,
                            value: BaseFontCardSettings.minFontSize
                        },
                        maxValue: {
                            type: powerbi.visuals.ValidatorType.Max,
                            value: BaseFontCardSettings.maxFontSize
                        }
                    }
                }),
                bold: new formattingSettings.ToggleSwitch({
                    name: BaseFontCardSettings.boldName + settingName,
                    value: false
                }),
                italic: new formattingSettings.ToggleSwitch({
                    name: BaseFontCardSettings.italicName + settingName,
                    value: false
                }),
                underline: new formattingSettings.ToggleSwitch({
                    name: BaseFontCardSettings.underlineName + settingName,
                    value: false
                })
            })
        );
    }
}

class LegendOptionsGroup extends Card {
    public defaultPosition: IEnumMember = positionOptions[0];

    public position = new formattingSettings.ItemDropdown({
        name: "position",
        displayNameKey: "Visual_Position",
        items: positionOptions,
        value: this.defaultPosition,
    });

    name: string = "legendOptions";
    displayName: string = "Options";
    displayNameKey: string = "Visual_Options";
    slices = [this.position];
}

class LegendTextGroup extends Card {
    public defaultLabelColor: string = "";
    public defaultFontSize: number = 9;

    public labelColor = new formattingSettings.ColorPicker({
        name: "labelColor",
        displayName: "Color",
        displayNameKey: "Visual_Color",
        value: { value: this.defaultLabelColor },
    });

    public font = new BaseFontCardSettings(this.defaultFontSize, "", "Segoe UI");

    name: string = "legendText";
    displayName: string = "Text";
    displayNameKey: string = "Visual_Text";
    slices = [this.font, this.labelColor];
}

class LegendTitleGroup extends Card {
    public defaultShowTitle: boolean = false;
    public defaultTitleText: string = "Legend";

    public showTitle = new formattingSettings.ToggleSwitch({
        name: "showTitle",
        displayNameKey: "Visual_ShowTitle",
        value: this.defaultShowTitle,
    });

    topLevelSlice = this.showTitle;

    public titleText = new formattingSettings.TextInput({
        name: "titleText",
        displayName: "Text",
        displayNameKey: "Visual_Text",
        value: this.defaultTitleText,
        placeholder: "Text",
    });

    name: string = TornadoObjectNames.LegendTitle;
    displayName: string = "Title";
    displayNameKey: string = "Visual_Title";
    slices = [this.titleText];
}

export class LegendCardSettings extends CompositeCard {
    public defaultShow: boolean = true;

    public name: string = "legend";
    public displayNameKey: string = "Visual_Legend";
    public analyticsPane: boolean = false;

    public show = new formattingSettings.ToggleSwitch({
        name: "show",
        displayNameKey: "Visual_Legend_Show",
        value: this.defaultShow,
    });

    public topLevelSlice: formattingSettings.ToggleSwitch = this.show;

    public options: LegendOptionsGroup = new LegendOptionsGroup();
    public text: LegendTextGroup = new LegendTextGroup();
    public title: LegendTitleGroup = new LegendTitleGroup();

    public groups = [this.options, this.text, this.title];
}

const categoryPositionOptions : IEnumMemberWithDisplayNameKey[] = [
    {value : LegendPosition[LegendPosition.Left], displayName : "Left", key: "Visual_Group_Left"}, 
    {value : LegendPosition[LegendPosition.Right], displayName : "Right", key: "Visual_Group_Right"},
     
];

export class FontDefaultOptions {
    public static DefaultFontSizePt: number = 8;
    public static DefaultFontFamily: string = "Segoe UI, wf_segoe-ui_normal, helvetica, arial, sans-serif";
}

export class BaseFontControlSettings extends formattingSettings.FontControl {
    constructor(defaultFontSize: number){
        super(
            new formattingSettings.FontControl({
                name: "font",
                displayName: "Font",
                displayNameKey: "Visual_FontControl",
                fontFamily: new formattingSettings.FontPicker({
                    name: "fontFamily",
                    value: FontDefaultOptions.DefaultFontFamily
                }),
                fontSize: new formattingSettings.NumUpDown({
                    name: "fontSize",
                    displayName: "Text size",
                    displayNameKey: "Visual_TextSize",
                    value: defaultFontSize,
                    options: {
                        minValue: {
                            type: powerbiVisualsApi.visuals.ValidatorType.Min,
                            value: 8,
                        },
                        maxValue: {
                            type: powerbiVisualsApi.visuals.ValidatorType.Max,
                            value: 60,
                        }
                    }
                }),
                bold: new formattingSettings.ToggleSwitch({
                    name: "fontBold",
                    value: false
                }),
                italic: new formattingSettings.ToggleSwitch({
                    name: "fontItalic",
                    value: false
                }),
                underline: new formattingSettings.ToggleSwitch({
                    name: "fontUnderline",
                    value: false
                })
            })
        );
    }
}

class CategoryOptionsGroup extends Card {
    constructor(position: formattingSettings.ItemDropdown) {
        super();
        this.slices = [position];
    }

    name: string = "categoryOptions";
    displayName: string = "Options";
    displayNameKey: string = "Visual_Options";
    slices: formattingSettings.Slice[];
}

class CategoryTextGroup extends Card {
    constructor(font: formattingSettings.FontControl, fill: formattingSettings.ColorPicker) {
        super();
        this.slices = [font, fill];
    }

    name: string = "categoryText";
    displayName: string = "Text";
    displayNameKey: string = "Visual_Text";
    slices: formattingSettings.Slice[];
}

export class CategoryCardSettings extends CompositeCard {
    show = new formattingSettings.ToggleSwitch({
        name: "show",
        displayName: "Show",
        displayNameKey: "Visual_Show",
        value: true,
    });
    
    topLevelSlice? = this.show;

    font: formattingSettings.FontControl = new BaseFontControlSettings(9);

    fill = new formattingSettings.ColorPicker({
        name: "fill",
        displayName: "Color",
        displayNameKey: "Visual_Color",
        value: { value: "" }
    });

    positionDropdown = new formattingSettings.ItemDropdown({
        items: categoryPositionOptions,
        value: categoryPositionOptions[0],
        name: "position",
        displayName: "Position",
        displayNameKey: "Visual_Position"
    });

    name: string = TornadoObjectNames.Categories;
    displayName: string = "Group";
    displayNameKey: string = "Visual_Group";
    optionsGroup = new CategoryOptionsGroup(this.positionDropdown);
    textGroup = new CategoryTextGroup(this.font, this.fill);
    groups: formattingSettings.Group[] = [this.optionsGroup, this.textGroup];
}


export class TornadoChartSettingsModel extends Model {
    dataColors = new DataColorCardSettings();
    categoryAxis = new CategoryAxisCardSettings();
    barAppearance = new BarAppearanceCardSettings();
    negativeBars = new NegativeBarsCardSettings();
    centerLine = new CenterLineCardSettings();
    chartArea = new ChartAreaCardSettings();
    dataLabels = new DataLabelSettings();
    legend = new LegendCardSettings();
    category = new CategoryCardSettings();

    cards = [
        this.dataColors,
        this.barAppearance,
        this.negativeBars,
        this.centerLine,
        this.categoryAxis,
        this.dataLabels,
        this.legend,
        this.category,
        this.chartArea
    ];

    setLocalizedOptions(localizationManager: ILocalizationManager) {
        this.setLocalizedDisplayName(positionOptions, localizationManager);
        this.setLocalizedDisplayName(categoryPositionOptions, localizationManager);
        this.setLocalizedDisplayName(contentOptions, localizationManager);
        this.setLocalizedDisplayName(dataLabelPositionOptions, localizationManager);
    }   

    public setLocalizedDisplayName(options: IEnumMemberWithDisplayNameKey[], localizationManager: ILocalizationManager) {
        options.forEach(option => {
            option.displayName = localizationManager.getDisplayName(option.key)
        });
    }

    public setVisibilityOfLegendCardSettings(legend: LegendData){
        this.legend.visible = legend.dataPoints.length > 0;
    }

    public updateDataLabelControlsState(): void {
        const displayMode = this.dataLabels.labelsOptionsGroup.displayFormat.value.value;
        const includesValue = displayMode !== LabelDisplayMode.Percentage;
        const includesPercentage = displayMode !== LabelDisplayMode.Value;

        this.dataLabels.labelsValuesGroup.labelPrecision.disabled = !includesValue;
        this.dataLabels.labelsValuesGroup.labelDisplayUnits.disabled = !includesValue;
        this.dataLabels.labelsValuesGroup.percentagePrecision.disabled = !includesPercentage;
    }
    
    public populateDataColorSlice(dataPoints: TornadoChartSeries[]){
        this.dataColors.slices = [];
        for (const dataPoint of dataPoints) {
            this.dataColors.slices.push(
                new formattingSettings.ColorPicker(
                {
                    name: "fill",
                    displayName: dataPoint.name,
                    selector: dataPoint.selectionId.getSelector(),
                    value: { value: dataPoint.fill }
                })
            );
        }
    }

    public populateCategoryAxisSlice(dataPoints: TornadoChartSeries[]){
        const isNormalized = this.categoryAxis.normalize.value;
        this.categoryAxis.groups = [this.categoryAxis.optionsGroup];
        dataPoints.forEach((dataPoint, index) => {
            const selector = ColorHelper.normalizeSelector(
                dataPoint.selectionId.getSelector(),
                false);
            const startOptions: powerbi.visuals.NumUpDownFormat = {
                minValue: {
                    type: powerbiVisualsApi.visuals.ValidatorType.Min,
                    value: 0
                },
                ...(this.isValueSet(dataPoint.categoryAxisEnd) && dataPoint.categoryAxisEnd >= 0
                    ? { maxValue: {
                        type: powerbiVisualsApi.visuals.ValidatorType.Max,
                        value: dataPoint.categoryAxisEnd
                    } }
                    : {})
            };
            const endOptions: powerbi.visuals.NumUpDownFormat = {
                minValue: {
                    type: powerbiVisualsApi.visuals.ValidatorType.Min,
                    value: this.isValueSet(dataPoint.categoryAxisStart)
                        ? Math.max(0, dataPoint.categoryAxisStart)
                        : 0
                }
            };

            this.categoryAxis.groups.push(new CategoryAxisRangeGroup(
                index,
                dataPoint.name,
                [
                    new formattingSettings.NumUpDown({
                        name: "start",
                        displayName: "Start",
                        displayNameKey: "Visual_XAxisStart",
                        value: <number><unknown>dataPoint.categoryAxisStart,
                        selector,
                        disabled: isNormalized,
                        options: startOptions
                    }),
                    new formattingSettings.NumUpDown({
                        name: "end",
                        displayName: "End",
                        displayNameKey: "Visual_XAxisEnd",
                        value: <number><unknown>(dataPoint.categoryAxisEnd ?? null),
                        selector,
                        disabled: isNormalized,
                        options: endOptions
                    })
                ]));
        });
    }

    private isValueSet(value: number | null): value is number {
        return Number.isFinite(value);
    }
}