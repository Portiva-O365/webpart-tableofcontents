import * as React from "react";
import styles from "./TableOfContents.module.scss";
import { ITableOfContentsProps, ITOCItem } from "../interfaces/interfaces";
import * as strings from "TableOfContentsWebPartStrings";
import { DisplayMode } from "@microsoft/sp-core-library";

export default class TableOfContents extends React.Component<ITableOfContentsProps, { isTextPresent: boolean }> {

  constructor(props: any) {
    super(props);
    // set state
    this.state = { isTextPresent: false };
  }

  public render(): React.ReactElement<ITableOfContentsProps> {
    // set background style of the container
    const containerStyle: any = {
      backgroundColor: this.props.tocBackgroundColor === "" ? "Transparent" : this.props.tocBackgroundColor,
      borderRadius: "3px"
    };

    // render the component
    if (this.state.isTextPresent) {
      return (
        <div className={styles.tableOfContents}>
          <div
            className={`ms-fadeIn400 ${styles.container} ${this.props.floatTOC && DisplayMode.Read ? styles.fixedContainer : ""}`}
            style={containerStyle}>
            <div className={styles.row}>
              <div className={styles.column}>
                {this._renderTOC()}
              </div>
            </div>
          </div>
        </div>
      );
    } else {
      // do nothing
      return null;
    }
  } // end: render

  public componentDidMount(): void {
    // component loaded, wait for other text components to load
    if (!this.state.isTextPresent) {
      window.setTimeout(() => {
        this.setState({ isTextPresent: true });
      }, 1000);
    }
  }

  private _renderTOC = (): JSX.Element => {
    // return the TOC
    return (
      <div>
        {this._renderTitle()}
        {
          document.location.href.indexOf("Mode=Edit") !== -1 ? this._renderTOCItemsInEditMode() : this._renderTOCItems()
        }
        {this._renderBackToPreviousPage()}
      </div>
    );
  } // end: _renderTOC

  private _renderTOCItems = (): JSX.Element => {

    // get all tags from text on this page
    const items: any = document.querySelectorAll(`${this.props.baseTag} ${this._getTagToGenerateTOCFrom()}`);

    // iterate over each found item
    if (items && items.length > 0) {
      // get the item for TOC
      let itemJSX: JSX.Element[] = [];

      for (let index: number = 0; index < items.length; index++) {

        // get anchor ID for source and target
        let anchorID: string = this._randomKey();

        // get #top anchor id
        if (index === 0) {
          anchorID = "TOCTop";
        }

        // text: add anchor ref to item, nb: use style instead of class since classes are compiled at runtime!
        items[index].innerHTML = `<a style="text-decoration:none; color: inherit;" id="${anchorID}">${items[index].innerText}</a>`;

        // text: add back to top item only for paragraphes after the first
        if (index > 0 && this.props.showBackToTop) {
          // add item
          items[index].outerHTML = `
            <div style="border-bottom: 1px solid #f1f1f1; padding-bottom: 10px; text-align: right; font-size: small;">
              <a href="#TOCTop" style="text-decoration: none; cursor: pointer;">
                ${this.props.backToTopText.trim() === "" ? strings.backToTopDefaultValue : this.props.backToTopText }
              </a>
            </div>
            ${items[index].outerHTML}
            `;
        }

        // toc: add item to JSX for rendering later
        itemJSX.push(
          this._renderTOCItem(
            { text: items[index].innerText, icon: this.props.iconTOCItem, anchorID: anchorID }
          ));
      }

      // return the items
      return (
        <div>
          {
            itemJSX.map((item, idx) => {
              return item;
            })
          }
        </div>
      );
    } else {
      // nothing on screen
      return null;
    }
  } // end: _renderTOCItems

  private _renderTOCItemsInEditMode = (): JSX.Element => {
    // called when page is in edit mode
    return (
      <div>
        <div className={styles.tocInEditModeDescription}>
          {strings.pageInEditMode}
        </div>
        {this._renderTOCItem({ text: strings.sampleItemLabel, icon: this.props.iconTOCItem })}
        {this._renderTOCItem({ text: strings.sampleItemLabel, icon: this.props.iconTOCItem })}
        {this._renderTOCItem({ text: strings.sampleItemLabel, icon: this.props.iconTOCItem })}
      </div>
    );
  } // end: _renderTOCItemsInEditMode

  private _renderTOCItem = (tocItemProps: ITOCItem): JSX.Element => {

    // return HTML for single TOC item with parameters filled in
    return (
      <div className={`${styles.tocItem} ${tocItemProps.isBackToPreviousPage ? styles.tocItemBackToPreviousPage : ""}`}
        onClick={tocItemProps.onClickAction} >
        <span className={styles.tocIcon}>
          <i className={`ms-Icon ms-Icon--${tocItemProps.icon}`} aria-hidden="true"></i>
        </span>
        {tocItemProps.anchorID ?
          <span className={styles.tocItemText}>
            <a href={`#${tocItemProps.anchorID}`}>{tocItemProps.text}</a>
          </span>
          :
          <span className={styles.tocItemText}>
            {tocItemProps.text}
          </span>
        }
      </div>
    );
  } // end: _renderTOCItem

  private _renderTitle = (): JSX.Element => {
    if (this.props.showTOCHeading) {
      return (
        <div className={styles.title}>
          {this.props.headingText.trim() === "" ? strings.headingTextDefaultValue : this.props.headingText}
        </div>
      );
    } else {
      return null;
    }
  } // end: _renderTitle

  private _renderBackToPreviousPage = (): JSX.Element => {
    if (this.props.showBackToPreviousPage) {
      return this._renderTOCItem({
        text: this.props.backToPreviousText.trim() === "" ? strings.backToPreviousDefaultValue : this.props.backToPreviousText,
        icon: this.props.iconPreviousPage,
        onClickAction: this._onClickBackToPreviousPage,
        isBackToPreviousPage: true
      });
    } else {
      // do not render to 'Back to Previous Page' link
      return null;
    }

  } // end: _renderBackToPreviousPage

  private _getTagToGenerateTOCFrom = (): string => {
    // return the HTML tag used to generate TOC items for, depends in user settings
    return this.props.htmlTag.toLowerCase();
    
  } // end: _getTagToGenerateTOCFrom

  private _randomKey = (): string => {
    return "L" + Math.random().toString(36).substr(2, 9).toUpperCase();
  }

  // ------------------------------------------------------------------------------------
  // click actions defined for each specific TOC item if applicable
  // ------------------------------------------------------------------------------------

  private _onClickBackToPreviousPage = (): void => {
    if (DisplayMode.Read) {
      // go back to previous page ( = page before this page, not the anchor tag page)
      window.history.back();
    }
  }
}
