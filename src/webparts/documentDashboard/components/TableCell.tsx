import * as React from "react";

import {
  ITableCell,
  ITableCellHeader
} from "../classes/Interfaces";

import {
  css
} from "office-ui-fabric-react";

import styles from "../DocumentDashboard.module.scss";

export class TableCellHeader extends React.Component<ITableCellHeader, {}> {
  protected static tableCellHeaderSortClasses: string = css(styles.msTableHeaderCell, styles["ms-Link"]);
  protected static tableCellChevronDownClasses: string = css(styles["ms-Icon"], styles["ms-Icon--chevronThinDown"]);
  protected static tableCellChevronUpClasses: string = css(styles["ms-Icon"], styles["ms-Icon--chevronThinUp"]);

  public render(): JSX.Element {
    return (
      <td className={TableCell.tableCellClasses} onClick={this.props.onClick}>
        <a className={TableCellHeader.tableCellHeaderSortClasses} href="#">
          {this.props.displayData}
          {(() => {
            if (this.props.isSorted) {
              if (!this.props.sortDirDesc) {
                return <i className={TableCellHeader.tableCellChevronUpClasses}></i>;
              }
              else if (this.props.sortDirDesc) {
                return <i className={TableCellHeader.tableCellChevronDownClasses}></i>;
              }
            }
          })()}
        </a>
      </td>
    );
  }
}

export class TableCell extends React.Component<ITableCell<any>, {}> {
  public static tableCellClasses: string = css(styles.msTableCellNoWrap, styles.autoCursor, styles["ms-Table-cell"]);
  protected static tableCellHyperlinkClasses: string = css(styles["ms-Link"]);

  public render(): JSX.Element {
    if (this.props.href) {
      return (
        <td className={TableCell.tableCellClasses}>
          <a className={TableCell.tableCellHyperlinkClasses} href={this.props.href} target="_blank">
            {this.props.displayData}
          </a>
        </td>
      );
    }
    else {
      return (
        <td className={TableCell.tableCellClasses}>
          {this.props.displayData}
        </td>
      );
    }
  }
}