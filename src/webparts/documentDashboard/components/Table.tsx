import * as React from "react";

import {
  ITable,
  ITableCell,
  ITableCellHeader,
  ITableRow
} from "../classes/Interfaces";

import {
  TableCell,
  TableCellHeader
} from "./TableCell";

import {
  FocusZone,
  FocusZoneDirection,
  KeyCodes,
  css
} from "office-ui-fabric-react";

import styles from "../DocumentDashboard.module.scss";

export default class Table extends React.Component<ITable, ITable> {
  private static tableClasses: string = css("ms-Table");
  private static tableBodyClasses: string = css("ms-bgColor-white");
  private static tableRowClasses: string = css("ms-Table-row");
  private static tablePagerWrapperClasses: string = css(styles.msTablePagerWrapper);
  private static tablePagerClasses: string = css(styles.msTablePager, "ms-font-l");

  private maxPageStartIndex: number = 0;
  private minPageStartIndex: number = 0;
  private pageCount: number = 1;
  private columnCount: number = 0;
  private hasNextPage: boolean;
  private hasPrevPage: boolean;

  constructor() {
    super();
    this.nextPage = this.nextPage.bind(this);
    this.prevPage = this.prevPage.bind(this);
    this.sortOnColumn = this.sortOnColumn.bind(this);
  }

  public componentWillMount(): void {
    // Calcuate constants
    const rowCount: number = this.props.rows.length;
    this.maxPageStartIndex = rowCount - 1; // rowCount - this.props.pageSize;
    if (this.maxPageStartIndex < this.minPageStartIndex) {
      this.maxPageStartIndex = this.minPageStartIndex;
    }
    this.pageCount = Math.ceil(rowCount / this.props.pageSize);
    if (rowCount > 0) {
      this.columnCount = this.props.rows[0].cells.length;
    }

    // Support initial sort
    if (this.props.currentSort >= 0 && this.props.currentSort < this.columnCount) {
      this.sortOnColumnInternal(this.props);
    }

    this.setState(this.props);
  }

  public render(): JSX.Element {
      this.hasNextPage = (this.state.pageStartIndex + this.state.pageSize <= this.maxPageStartIndex);
      this.hasPrevPage = (this.state.pageStartIndex - this.state.pageSize >= this.minPageStartIndex);

      return (
        <FocusZone
          direction={ FocusZoneDirection.vertical }
          isInnerZoneKeystroke={ (ev: KeyboardEvent) => ev.which === KeyCodes.right }>
            <div className={styles.msTableOverflow}>
              <table className={Table.tableClasses}>
                <tbody className={Table.tableBodyClasses}>
                  <tr className={Table.tableRowClasses}>
                    {this.state.columns.cells.map((c, i) => {
                        const hc: ITableCellHeader = {
                          sortableData: c.sortableData,
                          displayData: c.displayData,
                          href: c.href,
                          key: c.key,
                          onClick: (): void => { this.sortOnColumn(i); },
                          isSorted: this.state.currentSort === i,
                          sortDirDesc: this.state.currentSortDescending
                        };
                        return (
                          <TableCellHeader {...hc} />
                        );
                      })}
                  </tr>

                    {this.state.rows.slice(this.state.pageStartIndex, this.state.pageStartIndex + this.state.pageSize).map(r => {
                      return (
                        <tr key={r.key} className={Table.tableRowClasses}>
                          {r.cells.map(c => {
                            return (
                              <TableCell {...c} />
                            );
                          })}
                        </tr>
                      );
                    })}
                </tbody>
              </table>
              <div className={Table.tablePagerWrapperClasses}>
              {(() => {
                if (this.hasPrevPage) {
                  return (
                    <a href="#" onClick={this.prevPage} className={Table.tablePagerClasses}>
                      <i className="ms-Icon ms-Icon--triangleLeft" aria-action="Previous page of results"></i>
                    </a>
                  );
                }
              })()}
                <span>{Math.round(this.state.pageStartIndex / this.state.pageSize) + 1} of {this.pageCount}</span>
              {(() => {
                if (this.hasNextPage) {
                  return (
                    <a href="#" onClick={this.nextPage} className={Table.tablePagerClasses}>
                      <i className="ms-Icon ms-Icon--triangleRight" aria-action="Next page of results"></i>
                    </a>
                  );
                }
              })()}
                <div>About {this.props.rows.length} results</div>
              </div>
            </div>
        </FocusZone>
      );
  }

  private nextPage(): void {
    if (this.hasNextPage) {
      this.state.pageStartIndex = this.state.pageStartIndex + this.state.pageSize;
      this.setState(this.state);
    }
  }

  private prevPage(): void {
    if (this.hasPrevPage) {
      this.state.pageStartIndex = this.state.pageStartIndex - this.state.pageSize;
      this.setState(this.state);
    }
  }

  private sortOnColumn(columnIndex: number): void {
      if (columnIndex >= 0 && columnIndex < this.columnCount) {

        // if the column is already sorted, sort in opposite direction, else sort in constant direction
        this.state.currentSortDescending = this.state.currentSort === columnIndex ? !this.state.currentSortDescending : false;
        this.state.currentSort = columnIndex;

        // do the sorting
        this.sortOnColumnInternal(this.state);

        // trigger the update
        this.setState(this.state);
      }
  }

  private sortOnColumnInternal(data: ITable): void {
    if (data.currentSort >= 0 && data.currentSort < this.columnCount) {
      data.rows.sort((rowA, rowB) => {
        const cmpr: number = this.compareTableRow(data.currentSort, rowA, rowB);
        return (data.currentSortDescending ? cmpr * -1 : cmpr);
      });
    }
  }

  private compareTableRow(columnIndex: number, rowA: ITableRow, rowB: ITableRow): number {
    const cellA: ITableCell<any> = rowA.cells[columnIndex];
    const cellB: ITableCell<any> = rowB.cells[columnIndex];
    const cellDataA: any = cellA.sortableData;
    const cellDataB: any = cellB.sortableData;

    let compareValue: number = 0;
    if (cellDataA && cellDataB) {
      if (typeof cellDataA === "string" && typeof cellDataB === "string") {
        compareValue = cellDataA.localeCompare(cellDataB);
      }
      else if (cellDataA instanceof Date && cellDataB instanceof Date) {
        compareValue = (cellDataA > cellDataB) ? 1
                        : (cellDataA < cellDataB) ? -1 : 0;
      }
      else if (cellDataA instanceof Array && cellDataB instanceof Array) {
        // sort on the display values... not best solution
        compareValue = cellA.displayData.localeCompare(cellB.displayData);
      }
      else if (typeof cellDataA === "object" && typeof cellDataB === "object") {
        // IOwsUser
        if (cellDataA.preferredName && cellDataB.preferredName) {
          compareValue = cellDataA.preferredName.localeCompare(cellDataB.preferredName);
        }
        // TODO: Other types....?
      }
    }
    else {
      if (!cellDataA && !cellDataB) {
        compareValue = 0;
      }
      else {
        compareValue = cellDataA ? 1 : -1;
      }
    }
    return compareValue;
  }
}
