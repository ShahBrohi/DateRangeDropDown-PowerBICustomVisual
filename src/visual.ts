/*
 *  Power BI Visual CLI
 *
 *  Copyright (c) Microsoft Corporation
 *  All rights reserved.
 *  MIT License
 *
 *  Permission is hereby granted, free of charge, to any person obtaining a copy
 *  of this software and associated documentation files (the ""Software""), to deal
 *  in the Software without restriction, including without limitation the rights
 *  to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
 *  copies of the Software, and to permit persons to whom the Software is
 *  furnished to do so, subject to the following conditions:
 *
 *  The above copyright notice and this permission notice shall be included in
 *  all copies or substantial portions of the Software.
 *
 *  THE SOFTWARE IS PROVIDED *AS IS*, WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
 *  IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
 *  FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
 *  AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
 *  LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
 *  OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
 *  THE SOFTWARE.
 */
"use strict";

import powerbi from "powerbi-visuals-api";
import { FormattingSettingsService } from "powerbi-visuals-utils-formattingmodel";
import "./../style/visual.less";

import VisualConstructorOptions = powerbi.extensibility.visual.VisualConstructorOptions;
import VisualUpdateOptions = powerbi.extensibility.visual.VisualUpdateOptions;
import IVisual = powerbi.extensibility.visual.IVisual;
import IVisualHost = powerbi.extensibility.visual.IVisualHost;
import IVisualLocalStorageV2Service = powerbi.extensibility.IVisualLocalStorageV2Service;
import StorageV2ResultInfo = powerbi.extensibility.StorageV2ResultInfo;
import { dataViewObjects } from "powerbi-visuals-utils-dataviewutils";
import PrivilegeStatus = powerbi.PrivilegeStatus;

const PrivilegeStatusString = {
  [PrivilegeStatus.Allowed]: "Allowed",
  [PrivilegeStatus.NotDeclared]: "Not declared",
  [PrivilegeStatus.NotSupported]: "Not supported",
  [PrivilegeStatus.DisabledByAdmin]: "Disabled by admin",
};

import { VisualFormattingSettingsModel } from "./settings";

export class Visual implements IVisual {
  private target: HTMLElement;
  private formattingSettings: VisualFormattingSettingsModel;
  private formattingSettingsService: FormattingSettingsService;
  private host: IVisualHost;
  private data;
  private oldData;

  private uniqueMonthYearSet;
  private storageV2Service: IVisualLocalStorageV2Service;

  private startDateYearKey: string = "startdateyear";
  private startDateMonthKey: string = "startdatemonth";
  private endDateYearKey: string = "enddateyear";
  private endDateMonthKey: string = "enddatemonth";

  private validDates: { [year: string]: Set<string> } = {};

  private years = new Set<string>();
  private months = new Set<string>();

  private storedStartDateYear: string;
  private storedStartDateMonth: string;
  private storedEndDateYear: string;
  private storedEndDateMonth: string;

  private startDateMonthSelect!: HTMLSelectElement;
  private startDateYearSelect!: HTMLSelectElement;
  private endDateMonthSelect!: HTMLSelectElement;
  private endDateYearSelect!: HTMLSelectElement;

  private startDateContainer: HTMLElement;
  private endDateContainer: HTMLElement;

  private dateInputsContainer: HTMLElement;

  constructor(options: VisualConstructorOptions) {
    this.formattingSettingsService = new FormattingSettingsService();
    this.target = options.element;
    // this.storage = options.host.storageService;
    this.target = options.element;
    this.host = options.host;
    this.storageV2Service = options.host.storageV2Service;
  }

  public update(options: VisualUpdateOptions) {
    this.formattingSettings = this.formattingSettingsService.populateFormattingSettingsModel(VisualFormattingSettingsModel, options.dataViews[0]);

    const dataView = options.dataViews[0];

    if (dataView.categorical.categories.length !== 1) {
      this.Message("⚠️ Please add exactly 1 Date Column to the visual.");
      return;
    }

    const CompleteData = dataView.categorical.categories[0];

    let completeDataSet = new Set(CompleteData["values"]);

    let uniqueMonthYearSet = [];

    this.oldData = false;

    try {
      this.data.symmetricDifference(completeDataSet);
      this.oldData = true;
      uniqueMonthYearSet = this.uniqueMonthYearSet;
    } catch (exception) {
      const sortedDates = [];

      let dateValue: any;
      for (dateValue of completeDataSet) {
        try {
          const date = new Date(dateValue);
          dateValue.getTime();
          sortedDates.push(date);
        } catch {
          this.Message("⚠️ Non Date value exist in Column");
          return;
        }
      }

      sortedDates.sort((a, b) => a - b);

      let FinalDataSet = new Set();

      const years = new Set<string>();
      const months = new Set<string>();

      sortedDates.forEach((dateValue: any) => {
        const date = new Date(dateValue);
        const year = date.getFullYear();
        const month = (date.getMonth() + 1).toString().padStart(2, "0");
        const monthYear = `${year}-${month}`;
        FinalDataSet.add(monthYear);

        // will be used only for setting value which could be used for populating year
        years.add(year.toString());
        months.add(month.toString());

        // 1. Check if the year key already exists in the dictionary
        if (!this.validDates[year]) {
          // 2. If it does not exist, create a new Set and add the first month
          this.validDates[year] = new Set([month]);
        } else {
          // 3. If the year already exists, just add the new month to the existing Set
          this.validDates[year].add(month);
        }
      });
      uniqueMonthYearSet = Array.from(FinalDataSet);

      this.years = years;
      this.months = months;

      this.uniqueMonthYearSet = uniqueMonthYearSet;
      this.data = completeDataSet;
    }

    const persistedObjects = dataView.metadata.objects;
    
    //get data from report properties, so filters are saved in report whoever open the report will have same filter
    this.getDateReportProperties(persistedObjects);

    // // //sync all the visual same value properties
    this.getDateLocalStorage();

    // const [startDateYear,startDateMonth] = this.dateConvert(uniqueMonthYearSet[0]).toString().split("-");
    // const [endDateYear,endDateMonth] = this.dateConvert(uniqueMonthYearSet[uniqueMonthYearSet.length - 1]).toString().split("-");

    while (this.target.firstChild) {
      this.target.removeChild(this.target.firstChild);
    }

    this.startDateMonthSelect = document.createElement("select");
    this.startDateYearSelect = document.createElement("select");
    this.populateYear(this.startDateYearSelect);
    
    this.endDateMonthSelect = document.createElement("select");
    this.endDateYearSelect = document.createElement("select");
    this.populateYear(this.endDateYearSelect);

    if(this.storedStartDateYear != null && this.storedStartDateYear != undefined) this.startDateYearSelect.value = this.storedStartDateYear;
    if(this.storedEndDateYear != null && this.storedEndDateYear != undefined) this.endDateYearSelect.value = this.storedEndDateYear;

    if(this.storedStartDateMonth != null && this.storedStartDateMonth != undefined){
      this.populateMonthDates(this.startDateMonthSelect, this.startDateYearSelect.value);
      this.startDateMonthSelect.value = this.storedStartDateMonth;
    }
    else{
      this.populateMonth(this.startDateMonthSelect);
    }

    if(this.storedEndDateMonth != null && this.storedEndDateMonth != undefined){
      this.populateMonthDates(this.endDateMonthSelect, this.endDateYearSelect.value);
      this.endDateMonthSelect.value = this.storedEndDateMonth;
    }
    else{
      this.populateMonth(this.endDateMonthSelect);
    }
      
    this.startDateContainer = document.createElement("div");
    this.startDateContainer.style.display = "flex";
    this.startDateContainer.style.border = "1px solid #ccc";
    this.startDateContainer.style.borderRadius = "4px";
    this.startDateContainer.appendChild(this.startDateYearSelect);
    this.startDateContainer.appendChild(this.startDateMonthSelect);

    this.endDateContainer = document.createElement("div");
    this.endDateContainer.style.display = "flex";
    this.endDateContainer.style.border = "1px solid #ccc";
    this.endDateContainer.style.borderRadius = "4px";
    this.endDateContainer.appendChild(this.endDateYearSelect);
    this.endDateContainer.appendChild(this.endDateMonthSelect);

    this.dateInputsContainer = document.createElement("div");
    this.dateInputsContainer.style.display = "flex";
    this.dateInputsContainer.style.flexWrap = "wrap";
    this.dateInputsContainer.style.gap = "10px";
    this.dateInputsContainer.style.height = "90%";
    this.dateInputsContainer.appendChild(this.startDateContainer);
    this.dateInputsContainer.appendChild(this.endDateContainer);

    this.target.appendChild(this.dateInputsContainer);

    this.startDateYearSelect.addEventListener("change", () => {
      //if after changing of year date is valid then we need to pass month value so it should be selected after new dropdown creation
      //reason of creating dropdown again even after date is valid, for example month is selected to January and year is set to 2026, 
      // now for year 2026 there was only month january as valid and still it would show all the month which didnt even exist of 2026,
      if (this.dateIsValid(this.startDateYearSelect.value, this.startDateMonthSelect.value)) {
        this.populateMonthDates(this.startDateMonthSelect, this.startDateYearSelect.value.toString(),this.startDateMonthSelect.value);
      }
      else{
        this.populateMonthDates(this.startDateMonthSelect, this.startDateYearSelect.value.toString());
      }
    });

    this.endDateYearSelect.addEventListener("change", () => {
      if (this.dateIsValid(this.endDateYearSelect.value, this.endDateMonthSelect.value)) {
        this.populateMonthDates(this.endDateMonthSelect, this.endDateYearSelect.value.toString(),this.endDateMonthSelect.value);
      }
      else{
        this.populateMonthDates(this.endDateMonthSelect, this.endDateYearSelect.value.toString());        
      }
    });

    this.dateInputsContainer.addEventListener("change", async (event) => {
      const startYear = this.startDateYearSelect.value;
      let startMonth = this.startDateMonthSelect.value;
      const endYear = this.endDateYearSelect.value;
      let endMonth = this.endDateMonthSelect.value;

      let startForFilter = "";
      let endForFilter = "";

      //only apply filter if year exists
      if(startYear != ""){
        //if month value is not selected then just put month filter as 01 means it will filter from 1 month of selected year
        if (startMonth == "") {
          startMonth = "01";
          this.storedStartDateYear = startYear;
          // this.storedStartDateMonth = startMonth
          startForFilter = startYear + "-" + startMonth;
        }
        //when there is month available used same filter value to store in properties and local storage      
        else{
          this.storedStartDateYear = startYear;
          this.storedStartDateMonth = startMonth
          startForFilter = startYear + "-" + startMonth;
        }
      }
      
      if(endYear != ""){
        if (endMonth == "") {
          endMonth = this.lastValidMonthOfYear(endYear);
          this.storedEndDateYear = endYear
          // this.storedEndDateMonth = endMonth
          endForFilter =  endYear + "-" + endMonth;
        }
        else{
          this.storedEndDateYear = endYear;
          this.storedEndDateMonth = endMonth
          endForFilter = endYear + "-" + endMonth;
        }
      }

      this.setDateReportProperties();

      await this.setDateLocalStorage();

      const startDate = {
        operator: "GreaterThanOrEqual",
        value: startForFilter + "-01T00:00:00.000", //day included with timestamp as 01 means first day of month,
      };
      const endDate = {
        operator: "LessThanOrEqual",
        value: endForFilter + "-" + this.getDateForDay(endForFilter).getDate() + "T23:59:59.999", //day added seperately to get last day of month,
      };

      var conditionsList = [];

      if (startForFilter != "") {
        conditionsList.push(startDate);
      }

      if (endForFilter != "") {
        conditionsList.push(endDate);
      }

      const tableColumnNameList = CompleteData.source.queryName.split(".");

      const advancedFilter = {
        $schema: "http://powerbi.com/product/schema#advanced", // Required schema
        target: {
          table: tableColumnNameList[0], // Table name
          column: tableColumnNameList[1], // Column name
        },
        logicalOperator: "And", // Logical operator (And or Or)
        conditions: conditionsList,
      };

      try {
        await this.host.applyJsonFilter(advancedFilter, "general", "filter", powerbi.FilterAction.merge);
      } catch (error) {
        this.Message("Filter not applied properly");
      }
    });
  }

  private Message(message: string) {
    while (this.target.firstChild) {
      this.target.removeChild(this.target.firstChild);
    }
    // Create and append message
    const messageElement = document.createElement("div");
    messageElement.style.color = "red";
    messageElement.style.fontSize = "14px";
    messageElement.style.padding = "10px";
    messageElement.textContent = message;
    this.target.appendChild(messageElement);
  }

  public dateConvert(dateVariable: string) {
    const localISO = new Date(dateVariable).toISOString().slice(0, -1);
    const datelist = localISO.toString().split("-");
    return datelist.slice(0, 2).join("-");
  }

  private populateYear(selectElement: HTMLSelectElement) {
    selectElement.style.border = "none";

    const defaultOption = document.createElement("option");
    defaultOption.text = "Year";
    defaultOption.value = "";
    // defaultOption.disabled = true;
    defaultOption.selected = true;
    selectElement.appendChild(defaultOption);

    for (const year of this.years) {
      const option = document.createElement("option");
      option.value = year;
      option.text = year;
      selectElement.appendChild(option);
    }
  }

  private dateIsValid(year: string, month: string) {
    if (this.validDates[year] && this.validDates[year].has(month)) {
      return true;
    }
    return false;
  }

  private populateMonth(selectElement: HTMLSelectElement) {
    selectElement.style.border = "none";

    // Add a default "placeholder" option
    const defaultOption = document.createElement("option");
    defaultOption.text = "Month";
    defaultOption.value = "";
    // defaultOption.disabled = true;
    defaultOption.selected = true;
    selectElement.appendChild(defaultOption);

    const monthDictionary = {
        "01": "Jan",
        "02": "Feb",
        "03": "Mar",
        "04": "Apr",
        "05": "May",
        "06": "Jun",
        "07": "Jul",
        "08": "Aug",
        "09": "Sep",
        "10": "Oct",
        "11": "Nov",
        "12": "Dec",
      };

    for (const month of this.months) {
      const option = document.createElement("option");
      option.value = month;
      option.text = monthDictionary[month.toString()];
      selectElement.appendChild(option);
    }
  }

  private lastValidMonthOfYear(year: string) {
    let month = "12";
    for (const mon in this.validDates[year]) {
      month = mon;
    }
    return month;
  }

  private populateMonthDates(selectElement: HTMLSelectElement, year: string, month: string = "") {
    if (this.validDates[year]) {
      while (selectElement.firstChild) {
        selectElement.removeChild(selectElement.firstChild);
      }

      // this.populateMonth(selectElement);

      selectElement.style.border = "none";

      // Add a default "placeholder" option
      const defaultOption = document.createElement("option");
      defaultOption.text = "Month";
      defaultOption.value = "";
      // defaultOption.disabled = true;
      defaultOption.selected = true;
      selectElement.appendChild(defaultOption);

      const monthDictionary = {
        "01": "Jan",
        "02": "Feb",
        "03": "Mar",
        "04": "Apr",
        "05": "May",
        "06": "Jun",
        "07": "Jul",
        "08": "Aug",
        "09": "Sep",
        "10": "Oct",
        "11": "Nov",
        "12": "Dec",
      };

      this.validDates[year].forEach((month: string) => {
        const option = document.createElement("option");
        option.value = month;
        option.text = monthDictionary[month];

        selectElement.appendChild(option);
      });

      if(month != ""){
        selectElement.value = month;
      }
    }
  }

  public getDateReportProperties(persistedObjects : any){
        // // Retrieve the note, providing a default value if it doesn't exist yet
    const StartDateYear: string = dataViewObjects.getValue(
      persistedObjects,
      {
        objectName: "internalState",
        propertyName: "StartDateYear",
      },
      "Default",
    ); // Default value

    // --- Retrieve the SECOND value ---
    const StartDateMonth: string = dataViewObjects.getValue(
      persistedObjects,
      {
        objectName: "internalState",
        propertyName: "StartDateMonth", // <-- Use the new property name here
      },
      "Default", // <-- Provide a default for the new value
    );

    const EndDateYear: string = dataViewObjects.getValue(
      persistedObjects,
      {
        objectName: "internalState",
        propertyName: "EndDateYear",
      },
      "Default",
    ); // Default value

    // --- Retrieve the SECOND value ---
    const EndDateMonth: string = dataViewObjects.getValue(
      persistedObjects,
      {
        objectName: "internalState",
        propertyName: "EndDateMonth", // <-- Use the new property name here
      },
      "Default", // <-- Provide a default for the new value
    );

    if(StartDateYear != "Default") this.storedStartDateYear = StartDateYear;
    if(StartDateMonth != "Default") this.storedStartDateMonth = StartDateMonth;
    if(EndDateYear != "Default") this.storedEndDateYear = EndDateYear;
    if(EndDateMonth != "Default") this.storedEndDateMonth = EndDateMonth;
  }

  private setDateReportProperties(){
    this.host.persistProperties({
      merge: [
        {
          objectName: "internalState", // Must match the object name in capabilities.json
          selector: null, // Use null for visual-level properties
          properties: {
            StartDateYear: this.storedStartDateYear,
            StartDateMonth: this.storedStartDateMonth,
            EndDateYear: this.storedEndDateYear,            
            EndDateMonth: this.storedEndDateMonth,
          },
        },
      ],
    });
  }

  public async setDateLocalStorage(): Promise<void> {
    try {
      let status: PrivilegeStatus = await this.storageV2Service.status();
      if (status === PrivilegeStatus.Allowed) {
        await this.storageV2Service.set(this.startDateYearKey, this.storedStartDateYear);
        await this.storageV2Service.set(this.startDateMonthKey, this.storedStartDateMonth);
        await this.storageV2Service.set(this.endDateYearKey, this.storedEndDateYear);
        await this.storageV2Service.set(this.endDateMonthKey, this.storedEndDateMonth);
      }
    } catch (ex) {
      console.log("exception in setting local storage");
    }
  }

  public async getDateLocalStorage(): Promise<void> {
    try {
      let status: PrivilegeStatus = await this.storageV2Service.status();

      if (status === PrivilegeStatus.Allowed) {
        const startDateYear = await this.storageV2Service.get(this.startDateYearKey);
        const startDateMonth = await this.storageV2Service.get(this.startDateMonthKey);
        const endDateYear = await this.storageV2Service.get(this.endDateYearKey);
        const endDateMonth = await this.storageV2Service.get(this.endDateMonthKey);        

        if (startDateYear != "") this.storedStartDateYear = startDateYear;
        if (startDateMonth != "") this.storedStartDateMonth = startDateMonth;
        if (endDateYear != "") this.storedEndDateYear = endDateYear;
        if (endDateMonth != "") this.storedEndDateMonth = endDateMonth;

        if (this.oldData == false) {
          this.dateInputsContainer.dispatchEvent(new Event("change"));
        }
      }
    } catch (ex) {
      console.log("exception in getting local storage");
    }
  }

  public getDateForDay(date: string) {
    const tempDate = new Date(date);
    const endDateOnlyForTemp = new Date(parseInt(tempDate.getFullYear().toString()), parseInt((tempDate.getMonth() + 1).toString()), 0);
    return endDateOnlyForTemp;
  }

  /**
   * Returns properties pane formatting model content hierarchies, properties and latest formatting values, Then populate properties pane.
   * This method is called once every time we open properties pane or when the user edit any format property.
   */
  public getFormattingModel(): powerbi.visuals.FormattingModel {
    return this.formattingSettingsService.buildFormattingModel(this.formattingSettings);
  }
}