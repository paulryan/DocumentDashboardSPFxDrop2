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
  protected static tableCellHeaderSortClasses: string = css(styles.msTableHeaderCell, "ms-Link");

  public render(): JSX.Element {
    return (
      <td className={TableCell.tableCellClasses} onClick={this.props.onClick}>
        <a className={TableCellHeader.tableCellHeaderSortClasses} href="#">
          {this.props.displayData}
          {(() => {
            if (this.props.isSorted) {
              if (!this.props.sortDirDesc) {
                return <i className="ms-Icon ms-Icon--chevronThinUp"></i>;
              }
              else if (this.props.sortDirDesc) {
                return <i className="ms-Icon ms-Icon--chevronThinDown"></i>;
              }
            }
          })()}
        </a>
      </td>
    );
  }
}

export class TableCell extends React.Component<ITableCell<any>, {}> {
  public static tableCellClasses: string = css(styles.msTableCellNoWrap, styles.autoCursor, "ms-Table-cell");
  protected static tableCellHyperlinkClasses: string = css("ms-Link");

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