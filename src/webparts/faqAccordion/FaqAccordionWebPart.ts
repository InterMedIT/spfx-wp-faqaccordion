import * as React from 'react';
import * as ReactDom from 'react-dom';
import { DisplayMode, Version } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart } from "@microsoft/sp-webpart-base";
import {
  IPropertyPaneConfiguration,
  IPropertyPaneDropdownOption,
  PropertyPaneDropdown,
  PropertyPaneTextField,
  PropertyPaneToggle
} from "@microsoft/sp-property-pane";
import { IFieldInfo } from "@pnp/sp/fields";
import "@pnp/sp/search";
import { ISearchQuery, SearchResults } from "@pnp/sp/search";
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { Web } from '@pnp/sp/webs';
import { SPFI } from "@pnp/sp";

import * as strings from "FaqAccordionWebPartStrings";
import { getSP } from '../../utils/pnpjs-config';
import FaqAccordion, { IFaqAccordionProps } from './components/FaqAccordion';

export interface IFaqAccordionWebPartProps {
  webpartTitle: string;
  siteName: string;
  listName: string;
  categoryChoice: string;
  listQuestionColumn: string;
  listAnswerColumn: string;
  listSortColumn: string;
  isSortDescending: boolean;
  allowZeroExpanded: boolean;
  allowMultipleExpanded: boolean;
  displayMode: DisplayMode;
}

export interface ISPSite {
  SPSiteUrl: string;
  Title: string;
}
export interface ISPLists {
  value: ISPList[];
}
export interface ISPList {
  Title: string;
}

export default class FaqAccordionWebPart extends BaseClientSideWebPart<IFaqAccordionWebPartProps> {

  private _sp: SPFI;

  protected async onInit(): Promise<void> {
    await super.onInit();
    this._sp = getSP(this.context);
  }

  public render(): void {
    const element: React.ReactElement<IFaqAccordionProps> = React.createElement(FaqAccordion,
      {
        webpartTitle: this.properties.webpartTitle,
        siteName: this.properties.siteName,
        listName: this.properties.listName,
        categoryChoice: this.properties.categoryChoice,
        listQuestionColumn: this.properties.listQuestionColumn,
        listAnswerColumn: this.properties.listAnswerColumn,
        listSortColumn: this.properties.listSortColumn,
        isSortDescending: this.properties.isSortDescending ?? false,
        allowZeroExpanded: this.properties.allowZeroExpanded ?? true,
        allowMultipleExpanded: this.properties.allowMultipleExpanded ?? false,
        displayMode: this.displayMode,
        updateProperty: (value: string) => {
          this.properties.webpartTitle = value;
        },
        onConfigure: () => {
          this.context.propertyPane.open();
        },
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  /* ------------------------------------------------------------ */
  /* --------------------- Properties Panel --------------------- */
  /* ------------------------------------------------------------ */

  private _siteSelectOptions: IPropertyPaneDropdownOption[] = [];
  private _listSelectOptions: IPropertyPaneDropdownOption[] = [];
  private _categoryChoices: IPropertyPaneDropdownOption[];
  private _columnChoices: IPropertyPaneDropdownOption[];

  private _siteNameDropdownDisabled = true;
  private _listNameDropdownDisabled = true;
  private _categoryDropdownDisabled = true;
  private _columnsDropdownDisabled = true;

  private async _getSiteNames(): Promise<boolean> {
    //const queryPath: string = "path:" + window.location.hostname + "/sites/ ";
    const results: SearchResults = await this._sp.search(<ISearchQuery>{
      Querytext: "contentclass:STS_Site",
      RowLimit: 500,
      SelectProperties: ["SPSiteUrl", "Title"]
    });
    const sites = results.PrimarySearchResults;
    this._siteSelectOptions = sites.map((item: ISPSite) => {
      return {
        key: item.SPSiteUrl,
        text: item.Title
      };
    }).sort((a, b) => a.text.localeCompare(b.text));
    this._siteNameDropdownDisabled = false;
    return true;
  }

  private async _getListNames(): Promise<boolean> {
    if (this.properties.siteName) {
      const response: SPHttpClientResponse = await this.context.spHttpClient.get(`${this.properties.siteName}/_api/web/lists?$filter=Hidden eq false`, SPHttpClient.configurations.v1);
      const lists: Promise<ISPLists> = await response.json();
      this._listSelectOptions = (await lists).value.map((item: ISPList) => {
        return {
          key: item.Title,
          text: item.Title
        };
      });
      this._listNameDropdownDisabled = false;
      return true;
    }
    return false;
  }

  private async _loadCategoryChoices(): Promise<boolean> {
    if (!this.properties.listName) {
      throw new Error("List was not found.");
    }
    const web = Web([this._sp.web, this.properties.siteName]);
    const list = web.lists.getByTitle(this.properties.listName);
    if (!list) {
      throw new Error("List was not found.");
    }
    const field = await list.fields.getByTitle("Category")();
    if (!field) {
      throw new Error("Field with name 'Category' was not found in the selected list.");
    }
    const choices = field.Choices;
    if (!choices) {
      throw new Error("The 'Category' column is not of type 'Choice'.");
    }
    this._categoryChoices = choices.map((choice: string) => {
      return {
        key: choice,
        text: choice
      }
    });
    this._categoryDropdownDisabled = false;
    return true;
  }

  private async _loadColumnChoices(): Promise<boolean> {

    if (!this.properties.listName) {
      throw new Error("List was not found.");
    }
    const web = Web([this._sp.web, this.properties.siteName]);
    const columns = await web.lists.getByTitle(this.properties.listName).fields.filter("ReadOnlyField eq false and Hidden eq false")();
    if (!columns) {
      throw new Error("A list with columns and the name you specified was not found.")
    }
    this._columnChoices = columns.map((column: IFieldInfo) => {
      return {
        key: column.InternalName,
        text: column.Title
      };
    });
    this._columnsDropdownDisabled = false;
    return true;
  }

  protected resetFields(resetAll: boolean = false): void {
    if (resetAll) {
      this.properties.listName = "";
      this._listSelectOptions = [];
      this._listNameDropdownDisabled = true;
    }
    this.properties.categoryChoice = "";
    this._categoryChoices = [];
    this._categoryDropdownDisabled = true;

    this.properties.listQuestionColumn = "";
    this.properties.listAnswerColumn = "";
    this._columnChoices = [];
    this._columnsDropdownDisabled = true;

    this.properties.listSortColumn = "";
    this.context.propertyPane.refresh();
    this.render();
  }

  // fired when the properties panel is opened
  protected onPropertyPaneConfigurationStart(): void {
    //site name/URL defaults to the current site if none has been provided
    if (this.properties.siteName.trim().length < 1) {
      this.properties.siteName = this.context.pageContext.web.absoluteUrl;
    }
    this._getSiteNames().then((result1) => {
      if (result1) {
        this.context.propertyPane.refresh();
        this._getListNames().then((result2) => {
          if (result2) {
            this.context.propertyPane.refresh();
            if (this.properties.listName) {
              this._loadCategoryChoices().then((result3) => {
                if (result3) {
                  this.context.propertyPane.refresh();
                  this._loadColumnChoices().then((result4) => {
                    if (result4) {
                      this.context.propertyPane.refresh();
                    }
                  }).catch((error: Error) => {console.log(error.name)});
                }
              }).catch((error: Error) => {console.log(error.name)});
            }
          }
        }).catch((error: Error) => {console.log(error.name)});
      }
    }).catch((error: Error) => {console.log(error.name)})
  }

  // fired when the properties panel is changed (see https://learn.microsoft.com/en-us/sharepoint/dev/spfx/web-parts/guidance/use-cascading-dropdowns-in-web-part-properties)
  protected async onPropertyPaneFieldChanged(propertyPath: string, oldValue: string | undefined, newValue: string | undefined): Promise<void> {

    if (propertyPath === "siteName") {

      this.resetFields(true);
      await this._getListNames();
      this.context.propertyPane.refresh();
      this.render();

    } else if (propertyPath === "listName") {

      this.resetFields();
      await Promise.all([
        this._loadCategoryChoices(),
        this._loadColumnChoices()
      ]).then(() => {
        if (this.properties.listQuestionColumn.trim().length < 1) {
          this.properties.listQuestionColumn = "Title";
        }
        if (this.properties.listAnswerColumn.trim().length < 1) {
          this.properties.listAnswerColumn = "Answer";
        }
        if (this.properties.listSortColumn.trim().length < 1) {
          this.properties.listSortColumn = "SortOrder";
        }
        this.context.propertyPane.refresh();
        this.render();
      }).catch((error: Error) => {console.log(error.name)});

    } else {
      super.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);
      this.context.propertyPane.refresh();
      this.render();
    }
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [

                PropertyPaneTextField("webpartTitle", {
                  label: strings.WebpartTitleFieldLabel
              }),
                PropertyPaneDropdown('siteName', {
                  label: strings.SiteNameFieldLabel,
                  options: this._siteSelectOptions,
                  selectedKey: this.properties.siteName,
                  disabled: this._siteNameDropdownDisabled,
                }),
                PropertyPaneDropdown('listName', {
                  label: strings.ListNameFieldLabel,
                  options: this._listSelectOptions,
                  selectedKey: this.properties.listName,
                  disabled: this._listNameDropdownDisabled,
                }),
                PropertyPaneDropdown("categoryChoice", {
                  label: strings.CategoryChoiceFieldLabel,
                  options: this._categoryChoices,
                  selectedKey: this.properties.categoryChoice,
                  disabled: this._categoryDropdownDisabled,
                }),
                PropertyPaneDropdown("listQuestionColumn", {
                  label: strings.QuestionColumnFieldLabel,
                  options: this._columnChoices,
                  selectedKey: this.properties.listAnswerColumn,
                  disabled: this._columnsDropdownDisabled,
                }),
                PropertyPaneDropdown("listAnswerColumn", {
                  label: strings.AnswerColumnFieldLabel,
                  options: this._columnChoices,
                  selectedKey: this.properties.listAnswerColumn,
                  disabled: this._columnsDropdownDisabled,
                }),
                PropertyPaneDropdown("listSortColumn", {
                  label: strings.SortColumnFieldLabel,
                  options: this._columnChoices,
                  selectedKey: this.properties.listSortColumn,
                  disabled: this._columnsDropdownDisabled,
                }),
                PropertyPaneToggle("isSortDescending", {
                  label: strings.SortDirectionFieldLabel,
                  onText: "Descending",
                  offText: "Ascending",
                  disabled: !this.properties.listSortColumn
                }),
                PropertyPaneToggle("allowZeroExpanded", {
                  label: strings.AllowZeroExpandFieldLabel,
                  checked: this.properties.allowZeroExpanded,
                  key: "allowZeroExpanded",
                }),
                PropertyPaneToggle("allowMultipleExpanded", {
                  label: strings.AllowMultiExpandFieldLabel,
                  checked: this.properties.allowMultipleExpanded,
                  key: "allowMultipleExpanded",
                })

              ]
            }
          ]
        }
      ]
    };
  }
}
