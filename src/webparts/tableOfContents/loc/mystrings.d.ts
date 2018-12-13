declare interface ITableOfContentsWebPartStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  htmlTag: string;
  headingText: string;
  showTOCHeading: string;
  backToTopGroupName: string;
  showBackToTop: string;
  backToTopText: string;
  backToTopFieldDescription: string;
  showBackToPreviousPage: string;
  backToPreviousText: string;
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
  // errors
  errorFieldCannotBeEmpty: string;
  errorToggleFieldEmpty: string;
}

declare module 'TableOfContentsWebPartStrings' {
  const strings: ITableOfContentsWebPartStrings;
  export = strings;
}
