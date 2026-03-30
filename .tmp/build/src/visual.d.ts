import powerbi from "powerbi-visuals-api";
import "./../style/visual.less";
import VisualConstructorOptions = powerbi.extensibility.visual.VisualConstructorOptions;
import VisualUpdateOptions = powerbi.extensibility.visual.VisualUpdateOptions;
import IVisual = powerbi.extensibility.visual.IVisual;
export declare class Visual implements IVisual {
    private target;
    private formattingSettings;
    private formattingSettingsService;
    private host;
    private data;
    private oldData;
    private uniqueMonthYearSet;
    private storageV2Service;
    private startDateYearKey;
    private startDateMonthKey;
    private endDateYearKey;
    private endDateMonthKey;
    private validDates;
    private years;
    private months;
    private storedStartDateYear;
    private storedStartDateMonth;
    private storedEndDateYear;
    private storedEndDateMonth;
    private startDateMonthSelect;
    private startDateYearSelect;
    private endDateMonthSelect;
    private endDateYearSelect;
    private startDateContainer;
    private endDateContainer;
    private dateInputsContainer;
    constructor(options: VisualConstructorOptions);
    update(options: VisualUpdateOptions): void;
    private Message;
    dateConvert(dateVariable: string): string;
    private populateYear;
    private dateIsValid;
    private populateMonth;
    private lastValidMonthOfYear;
    private populateMonthDates;
    getDateReportProperties(persistedObjects: any): void;
    private setDateReportProperties;
    setDateLocalStorage(): Promise<void>;
    getDateLocalStorage(): Promise<void>;
    getDateForDay(date: string): Date;
    /**
     * Returns properties pane formatting model content hierarchies, properties and latest formatting values, Then populate properties pane.
     * This method is called once every time we open properties pane or when the user edit any format property.
     */
    getFormattingModel(): powerbi.visuals.FormattingModel;
}
