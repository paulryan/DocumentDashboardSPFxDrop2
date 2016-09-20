import * as React from "react";

import {
  ChartAxis,
  ControlMode,
  DisplayType,
  GetDisplayTermForEnumChartAxis,
  GetDisplayTermForEnumDisplayType,
  GetDisplayTermForEnumMode,
  GetDisplayTermForEnumSPScope
} from "../classes/Enums";

import {
  IChart,
  IChartItem,
  IDocumentDashboardProps,
  IDocumentDashboardState,
  ISecurableObject,
  ISecurableObjectProperty,
  ITable,
  ITableCell,
  ITableRow
} from "../classes/Interfaces";

import {
  ToVeryShortDateString
} from "../classes/Utilities";

import {
  Logger
} from "../classes/Logger";

import ChartistBar from "./ChartistBar";
import ChartistLine from "./ChartistLine";
import ChartistPie from "./ChartistPie";
import Table from "./Table";

import importableModuleLoader from "@microsoft/sp-module-loader";

interface ITableProps {
  items: ISecurableObject[];
}

interface ITableRowProps {
  item: ISecurableObject;
}

export default class DocumentDashboard extends React.Component<IDocumentDashboardProps, IDocumentDashboardState> {
  private log: Logger;
  private isUpdateStateInProgress: boolean = false;
  private hasContentBeenFetched: boolean = false;

  constructor() {
    super();
    this.log = new Logger("DocumentDashboard");

    importableModuleLoader.loadCss("//cdn.jsdelivr.net/chartist.js/latest/chartist.min.css");
  }

  public componentWillMount(): void {
    this.log.logInfo("componentWillMount");
    this.setStateWrapper([], ControlMode.Loading, "Working on it...");
  }

  public componentDidMount(): void {
    this.log.logInfo("componentDidMount");
    this.updateState();
  }

  public componentWillReceiveProps(): void {
    this.log.logInfo("componentWillReceiveProps");
    this.setStateWrapper(this.state.results, ControlMode.Loading, "Working on it...");
  }

  public shouldComponentUpdate(nextProps: IDocumentDashboardProps, nextState: IDocumentDashboardState): boolean {
    this.log.logInfo("shouldComponentUpdate");
    return !this.state || this.state.controlMode !== nextState.controlMode;
      // || this.state.mode !== nextState.mode
      // || this.state.scope !== nextState.scope
      // || this.state.displayType !== nextState.displayType;
  }

  public componentDidUpdate(): void {
    this.log.logInfo("componentDidUpdate");
    this.updateState();
  }

  public render(): JSX.Element {
    // Reusable components
    const headerControls: JSX.Element = (
      <div>
        {(() => {
          if (this.props.showHeading) {
            return (<div className="ms-font-xxl">Document Dashboard</div>);
          }
        })()}
        {(() => {
          if (this.props.showSubHeading) {
            let subheading: string = GetDisplayTermForEnumMode(this.props.mode) + " "
                                      + GetDisplayTermForEnumSPScope(this.props.scope).toLowerCase() + " "
                                      + GetDisplayTermForEnumDisplayType(this.props.displayType).toLowerCase();
            if (this.props.displayType !== DisplayType.Table) {
              subheading += ` (by ${GetDisplayTermForEnumChartAxis(this.props.chartAxis).toLowerCase()})`;

            }
            return (<div className="ms-font-l">{subheading}</div>);
          }
        })()}
      </div>
    );

    // Select the appropriate comnponent
    let component: JSX.Element = null;

    // Render according to the control mode
    if (this.state && this.state.controlMode === ControlMode.Loading) {
      this.log.logInfo("render (Loading)");
      component = (
        <div className="ms-font-l">{this.state.message}</div>
      );
    }
    else if (this.state && this.state.controlMode === ControlMode.Message) {
      this.log.logInfo("render (Message)");
      component = (
        <div className="ms-font-l">{this.state.message}</div>
      );
    }
    else if (this.state && this.state.controlMode === ControlMode.Content) {
      this.log.logInfo("render (Content)");
      if (this.props.displayType === DisplayType.Table) {
        const params: ITable = this.getStateAsITable();
        component = (
          <Table {...params} />
        );
      }
      else if (this.props.displayType === DisplayType.PieChart) {
        const params: IChart = this.getStateAsIChart(this.props.chartAxis, this.props.limitPieChartSegments);
        component = (
          <ChartistPie {...params} />
        );
      }
      else if (this.props.displayType === DisplayType.BarChart) {
        const params: IChart = this.getStateAsIChart(this.props.chartAxis, this.props.limitXAxisPlots);
        component = (
          <ChartistBar {...params} />
        );
      }
      else if (this.props.displayType === DisplayType.LineChart) {
        const params: IChart = this.getStateAsIChart(this.props.chartAxis, this.props.limitXAxisPlots);
        component = (
          <ChartistLine {...params} />
        );
      }
      else {
        this.log.logError("Unsupported display type: " + this.props.displayType);
        component = (
          <div className="ms-font-l">Error: Unsupported display type</div>
        );
      }
    }
    else if (this.state && this.state.controlMode) {
      this.log.logError(`ControlMode is not supported ${this.state.controlMode}`);
      component = (
        <div className="ms-font-l">Error: Unsupported control mode</div>
      );
    }
    else {
      this.log.logError(`State is undefined`);
      component = (
        <div className="ms-font-l">Error: State is undefined</div>
      );
    }

    return (
      <div>
        {headerControls}
        {component}
      </div>
    );
  }

  private shouldFetchContent(): boolean {
    return !this.state || !this.hasContentBeenFetched
      || this.props.mode !== this.state.mode
      || this.props.scope !== this.state.scope;
  }

  private updateState(): void {
    if (!this.isUpdateStateInProgress) {
      this.isUpdateStateInProgress = true;
      if (this.shouldFetchContent()) {
        this.props.store.getContent()
          .then((r) => {
            const controlMode: ControlMode = r.isError || r.results.length < 1 ? ControlMode.Message : ControlMode.Content;
            this.setStateWrapper(r.results, controlMode, r.message);
            this.hasContentBeenFetched = true;
            this.isUpdateStateInProgress = false;
          })
          .catch((e) => {
            this.log.logError("Failed to get content", e.message ? e.message : e.toString());
            this.setStateWrapper(this.state.results, ControlMode.Message, "Failed to get content");
            this.isUpdateStateInProgress = false;
          });
      }
      else {
        this.log.logInfo("New content has not been fetched as only the display mode has changed");
        // This sets the mode, scope, display mode as per props without changing the data
        const noResults: boolean = this.state.results.length < 1;
        if (noResults) {
          this.setStateWrapper(this.state.results, ControlMode.Message, "No content was found");
        }
        else {
          this.setStateWrapper(this.state.results, ControlMode.Content, "Using cached content");
        }
        this.isUpdateStateInProgress = false;
      }
    }
    else {
      this.log.logInfo("Update state ignored as request is already in progress");
    }
  }

  private setStateWrapper(results: ISecurableObject[], controlMode: ControlMode, message: string): void {
    this.setState({
      results: results,
      controlMode: controlMode,
      message: message,
      mode: this.props.mode,
      scope: this.props.scope
    });
  }

  private getStateAsIChart(chartAxis: ChartAxis, maxGroups: number): IChart {
    const dataPoints: IChartItem[] = [];
    this.state.results.forEach((securableObj) => {
      if (chartAxis === ChartAxis.User) {
        // Get all users associated with item.
        // Only count each user once per item.
        let users: IChartItem[] = [];
        // Modified by
        users.push({
          label: securableObj.modifiedBy.data.preferredName,
          data: securableObj.modifiedBy.data.email,
          weight: 1
        });
        // Created by
        users.push({
          label: securableObj.createdBy.data.preferredName,
          data: securableObj.createdBy.data.email,
          weight: 1
        });
        // Shared by
        securableObj.sharedBy.data.forEach((d) => {
          users.push({
            label: d,
            data: d,
            weight: 1
          });
        });
        // Shared with
        securableObj.sharedWith.data.forEach((d) => {
          users.push({
            label: d,
            data: d,
            weight: 1
          });
        });

        // Ensure unique values
        const userData: string[] = users.map(u => u.data);
        users = users.filter((user, index, self) => {
          return userData.indexOf(user.data) === index;
        });

        // Add to data points array
        dataPoints.push(...users);
      }
      else if (chartAxis === ChartAxis.Site) {
        dataPoints.push({
          label: securableObj.siteTitle.displayValue,
          data: securableObj.siteID.data,
          weight: 1
        });
      }
      else if (chartAxis === ChartAxis.Time) {
        if (securableObj.lastModifiedTime.data) {
          const d: Date = securableObj.lastModifiedTime.data;
          const roundedDate: Date = new Date(d.getFullYear(), d.getMonth()); // , d.getDate()
          dataPoints.push({
            label: ToVeryShortDateString(roundedDate, false, true, true),
            data: roundedDate.getTime().toString(),
            weight: 1,
            xAxis: roundedDate.getTime()
          });
        }
      }
    });

    return {
      items: dataPoints,
      maxGroups: maxGroups,
      chartAxisOrder: this.props.chartAxisOrder,
      domElement: this.props.domElement,
      columnIndexToGroupUpon: 0 // As there is only a single column in the data we return
    };
  }

  private getStateAsITable(): ITable {
    // TODO : In cases with lots of data it will not be okay to process all data
    // upfront - only the current page should be processed?
    const columnWithHref: string = "title";
    const columns: ITableCell<string>[] = [
      { sortableData: "title", displayData: "Title", href: null, key: "headerCellTitle" },
      { sortableData: "lastModifiedTime", displayData: "Last modified", href: null, key: "headerCellModified" }
    ];

    // Select only the correct columns as configured
    if (this.props.tableColumnsShowSiteTitle) {
      columns.push({ sortableData: "siteTitle", displayData: "Site title", href: null, key: "headerCellSiteTitle" });
    }
    if (this.props.tableColumnsShowCreatedByModifiedBy) {
      columns.push({ sortableData: "modifiedBy", displayData: "Modified by", href: null, key: "headerCellModifiedBy" });
      columns.push({ sortableData: "createdBy", displayData: "Created by", href: null, key: "headerCellCreatedBy" });
    }
    if (this.props.tableColumnsShowSharedWith) {
      columns.push({ sortableData: "sharedWith", displayData: "Shared With", href: null, key: "headerCellSharedWith" });
      columns.push({ sortableData: "sharedBy", displayData: "Shared By", href: null, key: "headerCellSharedBy" });
    }
    if (this.props.tableColumnsShowCrawledTime) {
      columns.push({ sortableData: "crawlTime", displayData: "Last crawled", href: null, key: "headerCellCrawlTime" });
    }

    const rows: ITableRow[] = [];
    this.state.results.forEach((securableObj) => {
      const newRow: ITableRow = { cells: [], key: securableObj.key };
      columns.forEach((columnName) => {
        const key: string = columnName.sortableData;
        const cellSortableData: ISecurableObjectProperty<any> = securableObj[columnName.sortableData];
        if (cellSortableData) {
          const href: string = (columnName.sortableData === columnWithHref ? securableObj.url.data : null);
          newRow.cells.push({
            sortableData: cellSortableData.data,
            displayData: cellSortableData.displayValue,
            href: href,
            key: key
          });
        }
        else {
          this.log.logError("Column value not present on row: " + columnName.sortableData);
          // Still add a cell to so that the rows do go out of line
          newRow.cells.push({
            sortableData: "?",
            displayData: "",
            href: null,
            key: key
          });
        }
      });
      rows.push(newRow);
    });

    return {
      columns: { cells: columns, key: "headerRow" },
      rows: rows,
      currentSort: -1,
      currentSortDescending: false,
      pageSize: this.props.tablePageSize,
      pageStartIndex: 0
    };
  }
}
