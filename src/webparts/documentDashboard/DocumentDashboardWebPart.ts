import {
  EnvironmentType
} from "@microsoft/sp-client-base";

import {
  BaseClientSideWebPart,
  IPropertyPaneSettings,
  IWebPartContext,
  PropertyPaneCheckbox,
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
  DisplayType,
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
      limitBarChartBars: this.properties.limitBarChartBars,
      limitPieChartSegments: this.properties.limitPieChartSegments,
      tableColumnsShowSharedWith: this.properties.tableColumnsShowSharedWith,
      tableColumnsShowCrawledTime: this.properties.tableColumnsShowCrawledTime,
      tableColumnsShowSiteTitle: this.properties.tableColumnsShowSiteTitle,
      tableColumnsShowCreatedByModifiedBy: this.properties.tableColumnsShowCreatedByModifiedBy
    });

    // Build the control!
    ReactDom.render(element, this.domElement);
  }

  protected get propertyPaneSettings(): IPropertyPaneSettings {
    return {
      pages: [
        {
          header: {
            description: "Standard settings"
          },
          displayGroupsAsAccordion: true,
          groups: [
            {
              groupName: "Content",
              groupFields: [
                PropertyPaneDropdown("mode", {
                  label: "What type content do you want to see?",
                  options: [
                    { key: Mode.AllDocuments, text: GetDisplayTermForEnumMode(Mode.AllDocuments) },
                    { key: Mode.MyDocuments, text: GetDisplayTermForEnumMode(Mode.MyDocuments) },
                    { key: Mode.AllExtSharedDocuments, text: GetDisplayTermForEnumMode(Mode.AllExtSharedDocuments) },
                    { key: Mode.MyExtSharedDocuments, text: GetDisplayTermForEnumMode(Mode.MyExtSharedDocuments) },
                    { key: Mode.AllAnonSharedDocuments, text: GetDisplayTermForEnumMode(Mode.AllAnonSharedDocuments) },
                    { key: Mode.MyAnonSharedDocuments, text: GetDisplayTermForEnumMode(Mode.MyAnonSharedDocuments) },
                    { key: Mode.RecentlyModifiedDocuments, text: GetDisplayTermForEnumMode(Mode.RecentlyModifiedDocuments) },
                    { key: Mode.InactiveDocuments, text: GetDisplayTermForEnumMode(Mode.InactiveDocuments) }
                  ]
                }),
                PropertyPaneDropdown("scope", {
                  label: "Where should we look for content?",
                  options: [
                    { key: SPScope.Tenant, text: GetDisplayTermForEnumSPScope(SPScope.Tenant) },
                    { key: SPScope.SiteCollection, text: GetDisplayTermForEnumSPScope(SPScope.SiteCollection) },
                    { key: SPScope.Site, text: GetDisplayTermForEnumSPScope(SPScope.Site) }
                  ]
                }),
                PropertyPaneDropdown("displayType", {
                  label: "How do you want the results rendered?",
                  options: [
                    { key: DisplayType.Table, text: GetDisplayTermForEnumDisplayType(DisplayType.Table) },
                    { key: DisplayType.BySite, text: GetDisplayTermForEnumDisplayType(DisplayType.BySite) },
                    { key: DisplayType.ByUser, text: GetDisplayTermForEnumDisplayType(DisplayType.ByUser) },
                    { key: DisplayType.OverTime, text: GetDisplayTermForEnumDisplayType(DisplayType.OverTime) }
                  ]
                })
              ]
            },
            {
            groupName: "Table columns",
              groupFields: [
                PropertyPaneCheckbox("tableColumnsShowSharedWith", {
                  //label: "Display 'Shared with' and 'Shared by' columns?"
                  text: "Display 'Shared with' and 'Shared by' columns?"
                }),
                PropertyPaneCheckbox("tableColumnsShowCrawledTime", {
                  //label: "Display 'Last crawled' column?"
                  text: "Display 'Last crawled' column?"
                }),
                PropertyPaneCheckbox("tableColumnsShowSiteTitle", {
                  //label: "Display 'Site title' column?"
                  text: "Display 'Site title' column?"
                }),
                PropertyPaneCheckbox("tableColumnsShowCreatedByModifiedBy", {
                  //label: "Display 'Modified by' and 'Created by' columns?"
                  text: "Display 'Modified by' and 'Created by' columns?"
                })
              ]
            },
            {
            groupName: "Misc.",
              groupFields: [
                PropertyPaneTextField("noResultsString", {
                  label: "What message should we display when there are no results?"
                })
              ]
            }
          ]
        },
        {
          header: {
            description: "Advanced settings"
          },
          displayGroupsAsAccordion: true,
          groups: [
            {
            groupName: "Limits",
              groupFields: [
                PropertyPaneSlider("limitRowsFetched", {
                  label: "What is the maximum number of items that should be fetched?",
                  min: 500,
                  max: 50000,
                  step: 500
                }),
                PropertyPaneSlider("limitPieChartSegments", {
                  label: "What is the maximum number of segments to display on a pie graph?",
                  min: 2,
                  max: 50
                }),
                PropertyPaneSlider("limitBarChartBars", {
                  label: "What is the maximum number of bars to display on a bar graph?",
                  min: 2,
                  max: 50
                })
              ]
            },
            {
            groupName: "Search schema",
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
