import { DisplayMode } from "@microsoft/sp-core-library";

export interface ITableOfContentsWebPartProps {
  baseTag: string;
  htmlTag: string;
  showBackToTop: boolean;
  backToTopText: string;
  showTOCHeading: boolean;
  headingText: string;
  showBackToPreviousPage: boolean;
  backToPreviousText: string;
  floatTOC: boolean;
  iconTOCItem: string;
  iconPreviousPage: string;
  tocBackgroundColor: string;
  buttonCopySetting?: string;
  buttonPasteSettings?: string;
  displayMode?: DisplayMode;
}

export interface ITableOfContentsProps extends ITableOfContentsWebPartProps { }

export interface ITOCItem {
  text: string;
  icon: string;
  onClickAction?: any;
  anchorID?: string;
  isBackToPreviousPage?: boolean;
}
