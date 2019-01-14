declare interface ITableOfContentsWebPartStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  htmlTag: string;
  textStyleHeading1Text: string;
  textStyleHeading2Text: string;
  textStyleHeading3Text: string;
  headingText: string;
  headingTextDefaultValue: string;
  showTOCHeading: string;
  backToTopGroupName: string;
  showBackToTop: string;
  backToTopText: string;
  backToTopDefaultValue: string;
  backToTopFieldDescription: string;
  showBackToPreviousPage: string;
  backToPreviousText: string;
  backToPreviousDefaultValue: string;
  backToPreviousFieldDescription: string;
  pageInEditMode: string;
  floatTOC: string;
  iconDescription: string;
  iconGroup: string;
  iconTOCItem: string;
  iconPreviousPage: string;
  tocBackgroundColor: string;
  tocBackgroundColorDescription: string;
  pinGroup: string;
  copyGroup: string;
  buttonCopySettingsLabel: string;
  buttonPasteSettingsLabel: string;
  sampleItemLabel: string;

  // errors
  errorFieldCannotBeEmpty: string;
  errorToggleFieldEmpty: string;
}

declare module 'TableOfContentsWebPartStrings' {
  const strings: ITableOfContentsWebPartStrings;
  export = strings;
}