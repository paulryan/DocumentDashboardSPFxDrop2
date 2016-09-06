import {
  EnvironmentType
} from "@microsoft/sp-client-base";

import {
  BaseClientSideWebPart,
  IPropertyPaneSettings,
  IWebPartContext,
  PropertyPaneCheckbox,
  PropertyPaneChoiceGroup,
  PropertyPaneDropdown,
  PropertyPaneLabel,
  PropertyPaneLink,
  PropertyPaneSlider,
  PropertyPaneTextField
} from "@microsoft/sp-client-preview";

// import * as strings from "mystrings";
import * as React from "react";
import * as ReactDom from "react-dom";

import DocumentDashboard from "./components/DocumentDashboard";

import {
  ChartAxis,
  ChartAxisOrder,
  DisplayType,
  GetDisplayTermForEnumChartAxis,
  GetDisplayTermForEnumChartAxisOrder,
  GetDisplayTermForEnumDisplayType,
  GetDisplayTermForEnumMode,
  GetDisplayTermForEnumSPScope,
  Mode,
  SPScope
} from "./classes/Enums";

import {
  IContentFetcherProps,
  IDocumentDashboardProps,
  IDocumentDashboardWebPartProps,
  ISecurableObjectStore
} from "./classes/Interfaces";

import ContentFetcher from "./classes/ContentFetcher";
import MockContentFetcher from "./tests/MockContentFetcher";

import {
  Logger
} from "./classes/Logger";

export default class DocumentDashboardWebPart extends BaseClientSideWebPart<IDocumentDashboardWebPartProps> {
  private log: Logger;

  public constructor(context: IWebPartContext) {
    super(context);
    this.log = new Logger("DocumentDashboardWebPart");
  }

  public render(): void {
    // Define properties for the Content Fetcher
    const contentFecherProps: IContentFetcherProps = {
      context: this.context,
      scope: this.properties.scope,
      mode: this.properties.mode,
      sharedWithManagedPropertyName: this.properties.sharedWithManagedPropertyName,
      crawlTimeManagedPropertyName: this.properties.crawlTimeManagedPropertyName,
      noResultsString: this.properties.noResultsString,
      limitRowsFetched: this.properties.limitRowsFetched
    };

    // Create appropriate Content Fectcher class for getting content
    let extContentStore: ISecurableObjectStore;
    if (this.context.environment.type === EnvironmentType.Local || this.context.environment.type === EnvironmentType.Test) {
      extContentStore = new MockContentFetcher(contentFecherProps);
    }
    else {
      extContentStore = new ContentFetcher(contentFecherProps);
    }

    // Create manager ReactElement
    const element: React.ReactElement<IDocumentDashboardProps> = React.createElement(DocumentDashboard, {
      store: extContentStore,
      mode: this.properties.mode,
      scope: this.properties.scope,
      displayType: this.properties.displayType,
      limitXAxisPlots: this.properties.limitXAxisPlots,
      limitPieChartSegments: this.properties.limitPieChartSegments,
      tableColumnsShowSharedWith: this.properties.tableColumnsShowSharedWith,
      tableColumnsShowCrawledTime: this.properties.tableColumnsShowCrawledTime,
      tableColumnsShowSiteTitle: this.properties.tableColumnsShowSiteTitle,
      tableColumnsShowCreatedByModifiedBy: this.properties.tableColumnsShowCreatedByModifiedBy,
      chartAxis: this.properties.chartAxis,
      tablePageSize: this.properties.tablePageSize,
      chartAxisOrder: this.properties.chartAxisOrder,
      domElement: this.domElement
    });

    // Build the control!
    ReactDom.render(element, this.domElement);
  }

  protected get propertyPaneSettings(): IPropertyPaneSettings {
    return {

      pages: [
        {
          header: {
            description: "Settings on this page determine which content you will see"
          },
          displayGroupsAsAccordion: false,
          groups: [
            {
              groupName: "What should we look for?",
              groupFields: [
                PropertyPaneChoiceGroup("mode", {
                  label: "What type content do you want to see?",
                  options: [
                    { key: Mode.AllDocuments, text: GetDisplayTermForEnumMode(Mode.AllDocuments) },
                    { key: Mode.MyDocuments, text: GetDisplayTermForEnumMode(Mode.MyDocuments) },
                    { key: Mode.AllExtSharedDocuments, text: GetDisplayTermForEnumMode(Mode.AllExtSharedDocuments) },
                    { key: Mode.MyExtSharedDocuments, text: GetDisplayTermForEnumMode(Mode.MyExtSharedDocuments) },
                    { key: Mode.AllAnonSharedDocuments, text: GetDisplayTermForEnumMode(Mode.AllAnonSharedDocuments) },
                    { key: Mode.MyAnonSharedDocuments, text: GetDisplayTermForEnumMode(Mode.MyAnonSharedDocuments) },
                    { key: Mode.RecentlyModifiedDocuments, text: GetDisplayTermForEnumMode(Mode.RecentlyModifiedDocuments) },
                    { key: Mode.InactiveDocuments, text: GetDisplayTermForEnumMode(Mode.InactiveDocuments) },
                    { key: Mode.Delve, text: GetDisplayTermForEnumMode(Mode.Delve) }
                  ]
                }),
                PropertyPaneDropdown("scope", {
                  label: "Where should we look for content?",
                  options: [
                    { key: SPScope.Tenant, text: GetDisplayTermForEnumSPScope(SPScope.Tenant) },
                    { key: SPScope.SiteCollection, text: GetDisplayTermForEnumSPScope(SPScope.SiteCollection) },
                    { key: SPScope.Site, text: GetDisplayTermForEnumSPScope(SPScope.Site) }
                  ]
                })
              ]
            },
            {
              groupName: "How should we handle empty or large data sets?",
              groupFields: [
                PropertyPaneTextField("noResultsString", {
                  label: "What message should we display when there are no results?"
                }),
                PropertyPaneSlider("limitRowsFetched", {
                  label: "What is the maximum number of items that we should fetch?",
                  min: 500,
                  max: 50000,
                  step: 500
                })
              ]
            }
          ]
        },
        {
          header: {
            description: "Settings on this page determine how content is displayed"
          },
          displayGroupsAsAccordion: true,
          groups: [
            {
              groupName: "How do you want content to be displayed?",
              groupFields: [
                PropertyPaneDropdown("displayType", {
                  label: "Which visual model should we use?",
                  options: [
                    { key: DisplayType.Table, text: GetDisplayTermForEnumDisplayType(DisplayType.Table) },
                    { key: DisplayType.PieChart, text: GetDisplayTermForEnumDisplayType(DisplayType.PieChart) },
                    { key: DisplayType.BarChart, text: GetDisplayTermForEnumDisplayType(DisplayType.BarChart) },
                    { key: DisplayType.LineChart, text: GetDisplayTermForEnumDisplayType(DisplayType.LineChart) }
                  ]
                }),
                PropertyPaneDropdown("chartAxis", {
                  label: "For charts only, on which property should we group results?",
                  options: [
                    { key: ChartAxis.Site, text: GetDisplayTermForEnumChartAxis(ChartAxis.Site) },
                    { key: ChartAxis.User, text: GetDisplayTermForEnumChartAxis(ChartAxis.User) },
                    { key: ChartAxis.Time, text: GetDisplayTermForEnumChartAxis(ChartAxis.Time) }
                  ]
                })
              ]
            },
            {
              groupName: "How do you want a table to look?",
              groupFields: [
                PropertyPaneCheckbox("tableColumnsShowSharedWith", {
                  text: "Should we display 'Shared with' and 'Shared by' columns?"
                }),
                PropertyPaneCheckbox("tableColumnsShowCrawledTime", {
                  text: "Should we display 'Last crawled' column?"
                }),
                PropertyPaneCheckbox("tableColumnsShowSiteTitle", {
                  text: "Should we display 'Site title' column?"
                }),
                PropertyPaneCheckbox("tableColumnsShowCreatedByModifiedBy", {
                  text: "Should we display 'Modified by' and 'Created by' columns?"
                }),
                PropertyPaneDropdown("tablePageSize", {
                  label: "How large should a page be?",
                  options: [
                    { key: 5, text: "5" },
                    { key: 10, text: "10" },
                    { key: 20, text: "20" },
                    { key: 50, text: "50" },
                    { key: 100, text: "100" }
                  ]
                })
              ]
            },
            {
              groupName: "How do you want charts to look?",
              groupFields: [
                PropertyPaneSlider("limitPieChartSegments", {
                  label: "What is the maximum number of segments that should be displayed on a pie chart?",
                  min: 2,
                  max: 50
                }),
                PropertyPaneSlider("limitXAxisPlots", {
                  label: "What is the maximum number of x-axis plots that should be displayed on a chart?",
                  min: 2,
                  max: 50
                }),
                PropertyPaneDropdown("chartAxisOrder", {
                  label: "When we can't show everything should we prioritise the largest groups or the smallest groups?",
                  options: [
                    { key: ChartAxisOrder.PrioritiseLargestGroups, text: GetDisplayTermForEnumChartAxisOrder(ChartAxisOrder.PrioritiseLargestGroups) },
                    { key: ChartAxisOrder.PrioritiseSmallestGroups, text: GetDisplayTermForEnumChartAxisOrder(ChartAxisOrder.PrioritiseSmallestGroups) }
                  ]
                })
              ]
            }
          ]
        },
        {
          header: {
            description: "Settings on this page are for advanced users only"
          },
          displayGroupsAsAccordion: false,
          groups: [
            {
              groupName: "Search schema configuration",
              groupFields: [
                PropertyPaneTextField("sharedWithManagedPropertyName", {
                  label: "What is the name of the queryable Managed Property containing shared with details?",
                  description: `This property must be configured as such:
                                Text, Multi, Queryable, Retrievable, and be mapped to 'ows_SharedWithDetails'`
                }),
                PropertyPaneTextField("crawlTimeManagedPropertyName", {
                  label: "What is the name of the Managed Property containing crawl time details?",
                  description: `This property must be configured as such:
                                Text, Retrievable, and be mapped to 'Internal:323'`
                }),
                PropertyPaneLabel("labelproperty01", {
                  text: "Use the following link to download a search schema file to import the above managed properties:"
                }),
                PropertyPaneLink("linkproperty01", {
                  href: "https://www.bing.com",
                  text: "Search schema"
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
