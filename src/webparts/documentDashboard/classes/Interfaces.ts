import {
  IWebPartContext
} from "@microsoft/sp-client-preview";

import {
  ChartAxis,
  ChartAxisOrder,
  ControlMode,
  DisplayType,
  Mode,
  SPScope,
  SecurableObjectType
} from "./Enums";

export interface IContentFetcherProps {
  context: IWebPartContext;
  scope: SPScope;
  mode: Mode;
  sharedWithManagedPropertyName: string;
  crawlTimeManagedPropertyName: string;
  noResultsString: string;
  limitRowsFetched: number;
}

export interface IGetContentFuncResponse {
  results: ISecurableObject[];
  message: string;
  isError: boolean;
}

export interface IDocumentDashboardProps {
  store: ISecurableObjectStore;
  scope: SPScope;
  mode: Mode;
  displayType: DisplayType;
  limitPieChartSegments: number;
  limitXAxisPlots: number;
  tableColumnsShowSharedWith: boolean;
  tableColumnsShowCrawledTime: boolean;
  tableColumnsShowSiteTitle: boolean;
  tableColumnsShowCreatedByModifiedBy: boolean;

  chartAxis: ChartAxis;
  tablePageSize: number;
  chartAxisOrder: ChartAxisOrder;
}

export interface IDocumentDashboardState {
  results: ISecurableObject[];
  message: string;
  controlMode: ControlMode;

  scope: SPScope;
  mode: Mode;
}

export interface ISecurableObject {
  title: ISecurableObjectProperty<string>;
  fileExtension: ISecurableObjectProperty<string>;
  lastModifiedTime: ISecurableObjectProperty<Date>;
  siteTitle: ISecurableObjectProperty<string>;
  siteID: ISecurableObjectProperty<string>;
  crawlTime: ISecurableObjectProperty<Date>;
  url: ISecurableObjectProperty<string>;
  type: ISecurableObjectProperty<SecurableObjectType>;
  sharedWith: ISecurableObjectProperty<string[]>;
  sharedBy: ISecurableObjectProperty<string[]>;
  modifiedBy: ISecurableObjectProperty<IOwsUser>;
  createdBy: ISecurableObjectProperty<IOwsUser>;
  key: string;
}

export interface ISecurableObjectProperty<Type> {
  data: Type;
  displayValue: string;
}

export interface IGetContentFunc {
    (): Promise<IGetContentFuncResponse>;
}

export interface ISecurableObjectStore {
  getContent: IGetContentFunc;
}

export interface IDocumentDashboardWebPartProps {
  scope: SPScope;
  mode: Mode;
  displayType: DisplayType;
  noResultsString: string;
  sharedWithManagedPropertyName: string;
  crawlTimeManagedPropertyName: string;
  limitRowsFetched: number;
  limitPieChartSegments: number;
  limitXAxisPlots: number;
  tableColumnsShowSharedWith: boolean;
  tableColumnsShowCrawledTime: boolean;
  tableColumnsShowSiteTitle: boolean;
  tableColumnsShowCreatedByModifiedBy: boolean;
  chartAxis: ChartAxis;
  tablePageSize: number;
  chartAxisOrder: ChartAxisOrder;
}

export interface ISearchResponse {
  results: any[];
  rowCount: number;
  totalRows: number;
  totalRowsIncludingDuplicates: number;
  isSuccess: boolean;
  message: string;
}

export interface IOwsUser {
  preferredName: string;
  accountName: string;
  email: string;
}

export interface IChart {
  items: IChartItem[];
  columnIndexToGroupUpon: number;
  maxGroups: number;
  chartAxisOrder: ChartAxisOrder;
}

export interface IChartItem {
  label: string;
  data: string;
  weight: number;
}

export interface ITable {
  columns: ITableRow;
  rows: ITableRow[];
  pageSize: number;
  pageStartIndex: number;
  currentSort: number;
  currentSortDescending: boolean;
}

export interface ITableRow {
  cells: ITableCell<any>[];
  key: string;
}

export interface ITableCell<Type> {
  sortableData: Type;
  displayData: string;
  href: string;
  key: string;
}

export interface ITableCellHeader extends ITableCell<string> {
  onClick?: any;
  isSorted: boolean;
  sortDirDesc: boolean;
}
