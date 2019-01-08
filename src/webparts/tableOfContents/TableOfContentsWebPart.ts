import * as React from "react";
import * as ReactDom from "react-dom";
import { Version } from "@microsoft/sp-core-library";
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneToggle,
  PropertyPaneDropdown,
  PropertyPaneHorizontalRule,
  PropertyPaneButton
} from "@microsoft/sp-webpart-base";

import * as strings from "TableOfContentsWebPartStrings";
import TableOfContents from "./components/TableOfContents";
import { ITableOfContentsWebPartProps, ITableOfContentsProps } from "./interfaces/interfaces";

export default class TableOfContentsWebPart extends BaseClientSideWebPart<ITableOfContentsWebPartProps> {

  public render(): void {

    // render the main web part
    const element: React.ReactElement<ITableOfContentsProps> = React.createElement(
      TableOfContents, {
        context: this.context,
        baseTag: this.properties.baseTag,
        htmlTag: this.properties.htmlTag,
        showBackToTop: this.properties.showBackToTop,
        backToTopText: this.properties.backToTopText,
        showTOCHeading: this.properties.showTOCHeading,
        headingText: this.properties.headingText,
        showBackToPreviousPage: this.properties.showBackToPreviousPage,
        backToPreviousText: this.properties.backToPreviousText,
        floatTOC: this.properties.floatTOC,
        iconTOCItem: this.properties.iconTOCItem,
        iconPreviousPage: this.properties.iconPreviousPage,
        tocBackgroundColor: this.properties.tocBackgroundColor
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse("1.0");
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
              isCollapsed: false,
              groupFields: [
                PropertyPaneDropdown("htmlTag", {
                  label: strings.htmlTag, options: [
                    { key: "Heading 1", text: "Heading 1" },
                    { key: "Heading 2", text: "Heading 2" },
                    { key: "Heading 3", text: "Heading 3" }
                  ]
                }),
                PropertyPaneToggle("showTOCHeading", {
                  label: strings.showTOCHeading
                }),
                PropertyPaneTextField("headingText", {
                  label: strings.headingText, disabled: !this.properties.showTOCHeading,
                  onGetErrorMessage: this._checkToggleField
                }),
                PropertyPaneToggle("showBackToPreviousPage", {
                  label: strings.showBackToPreviousPage
                }),
                PropertyPaneTextField("backToPreviousText", {
                  label: strings.backToPreviousText, description: strings.backToPreviousFieldDescription,
                  disabled: !this.properties.showBackToPreviousPage,
                  onGetErrorMessage: this._checkToggleField
                })
              ]
            },
            {
              groupName: strings.backToTopGroupName,
              isCollapsed: false,
              groupFields: [
                PropertyPaneToggle("showBackToTop", {
                  label: strings.showBackToTop
                }),
                PropertyPaneTextField("backToTopText", {
                  label: strings.backToTopText, description: strings.backToTopFieldDescription,
                  disabled: !this.properties.showBackToTop,
                  onGetErrorMessage: this._checkToggleField
                })
              ]
            }
          ]
        },
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.iconGroup, isCollapsed: false,
              groupFields: [
                PropertyPaneTextField("iconTOCItem", {
                  label: strings.iconTOCItem, description: strings.iconDescription,
                  onGetErrorMessage: this._checkIconField
                }),
                PropertyPaneTextField("iconPreviousPage", {
                  label: strings.iconPreviousPage, description: strings.iconDescription,
                  onGetErrorMessage: this._checkIconField
                }),
                PropertyPaneTextField("tocBackgroundColor", {
                  label: strings.tocBackgroundColor, description: strings.tocBackgroundColorDescription
                })
              ]
            },
            {
              groupName: strings.pinGroup, isCollapsed: false,
              groupFields: [
                PropertyPaneToggle("floatTOC", {
                  label: strings.floatTOC
                })
              ]
            },
            {
              groupName: strings.copyGroup, isCollapsed: false,
              groupFields: [
                PropertyPaneButton("buttonCopySettings", {
                  text: strings.buttonCopySettingsLabel, icon: "Copy",
                  onClick: this._onClickCopySettings
                }),
                PropertyPaneButton("buttonPasteSettings", {
                  text: strings.buttonPasteSettingsLabel, icon: "Paste",
                  onClick: this._onClickPasteSettings
                })
              ]
            }
          ]
        }
      ]
    };
  }

  private _onClickCopySettings = (): void => {
    // copy settings to storage to retrieve in other web part
    const propsJSONString: string = JSON.stringify(this.properties);
    // save in localstorage
    window.localStorage.setItem("TOCSettings", propsJSONString);
  }

  private _onClickPasteSettings = (): void => {
    // paste settings to storage to retrieve in other web part
    const propsJSONString: string = window.localStorage.getItem("TOCSettings");
    // check item
    if (propsJSONString !== "") {
      // convert to JSON
      const propsJSONObject: any = JSON.parse(propsJSONString);
      if (propsJSONObject !== null) {
        // save the properties
        Object.keys(propsJSONObject).forEach((key: string) => {
          // copy key-value pairs to web part properties
          this.properties[key] = propsJSONObject[key];
        });
      }
    }
  }

  private _checkToggleField = (value: string): string => {
    // called from textbox linked to toggle fields, extra message
    if (value === "") {
      return strings.errorToggleFieldEmpty;
    } else {
      return "";
    }
  }

  private _checkIconField = (value: string): string => {
    // check if icon is filled in
    if (value === "") {
      return strings.errorFieldCannotBeEmpty;
    } else {
      return "";
    }
  }

}
