import {
  ControlMode,
  Mode,
  SPScope,
  SecurableObjectType
} from "./Enums";

import {
  IContentFetcherProps,
  IGetContentFuncResponse,
  IOwsUser,
  ISearchResponse,
  ISecurableObject,
  ISecurableObjectStore
} from "./Interfaces";

import {
  EnsureBracesOnGuidString,
  GetDateFqlString,
  ParseDisplayNameFromExtUserAccountName,
  ParseOWSUSER,
  ToColloquialDateString,
  TransformSearchResponse
} from "./Utilities";

import {
  Logger
} from "./Logger";

export default class ContentFetcher implements ISecurableObjectStore {

  public props: IContentFetcherProps;
  private log: Logger;

  public constructor(props: IContentFetcherProps) {
    this.props = props;
    this.log = new Logger("ContentFetcher");
  }

  public getContent(): Promise<IGetContentFuncResponse> {
    const self: ContentFetcher = this;
    self.log.logInfo("getContent()");

    const rowLimit: number = 500;
    const baseUri: string = self.props.context.pageContext.web.absoluteUrl + "/_api/search/query";

    // "MY" should represent things I have created or edited or shared.
    // TODO: Can we get Graphy in a meaningful way?
    const un: string = self.props.context.pageContext.user.loginName;
    const me: string = un.substring(0, un.indexOf("@")); // TODO: get the query working with @ symbol.. .replace("@", "%40");
    let myFql: string = `EditorOWSUSER:${me} OR AuthorOWSUSER:${me}`;
    if (self.props.sharedWithManagedPropertyName) {
      myFql += ` OR ${self.props.sharedWithManagedPropertyName}:${me}`;
    }
    myFql = `(${myFql})`;

    const documentsFql: string = "ContentClass:STS_ListItem_DocumentLibrary";
    const extSharedFql: string = "ViewableByExternalUsers:1";
    const anonSharedFql: string = "ViewableByAnonymousUsers:1";

    let graphFql: string = "";
    let modeFql: string = "";
    if (self.props.mode === Mode.AllDocuments) {
      //modeFql = `${documentsFql}`;
      modeFql = "*";
    }
    else if (self.props.mode === Mode.MyDocuments) {
      modeFql = `${myFql} ${documentsFql}`;
    }
    else if (self.props.mode === Mode.AllExtSharedDocuments) {
      modeFql = `${extSharedFql}`;
    }
    else if (self.props.mode === Mode.MyExtSharedDocuments) {
      modeFql = `${myFql} ${extSharedFql}`;
    }
    else if (self.props.mode === Mode.AllAnonSharedDocuments) {
      modeFql = `${anonSharedFql}`;
    }
    else if (self.props.mode === Mode.MyAnonSharedDocuments) {
      modeFql = `${myFql} ${anonSharedFql}`;
    }
    else if (self.props.mode === Mode.RecentlyModifiedDocuments) {
      const now: Date = new Date();
      const earlier: Date = new Date(now.getFullYear(), now.getMonth(), now.getDate() - 14);
      modeFql = `${documentsFql} Write>${GetDateFqlString(earlier)}`;
    }
    else if (self.props.mode === Mode.InactiveDocuments) {
      const now: Date = new Date();
      const earlier: Date = new Date(now.getFullYear(), now.getMonth() - 6, now.getDate());
      modeFql = `${documentsFql} Write<${GetDateFqlString(earlier)}`;
    }
    else if (self.props.mode === Mode.Delve) {
      modeFql = `${documentsFql}`;
      graphFql = "properties='GraphQuery:ACTOR(ME)'";
    }
    else {
      self.log.logError("Unsupported mode: " + self.props.mode);
      return null;
    }

    let scopeFql: string = "";
    if (self.props.scope === SPScope.SiteCollection) {
      scopeFql = " SiteId:" + EnsureBracesOnGuidString(self.props.context.pageContext.site.id.toString());
    }
    else if (self.props.scope === SPScope.Site) {
      scopeFql = " WebId:" + EnsureBracesOnGuidString(self.props.context.pageContext.web.id.toString());
    }
    else if (self.props.scope === SPScope.Tenant) {
      // do nothing
    }
    else {
      self.log.logError("Unsupported scope: " + self.props.scope);
      return null;
    }

    const selectPropsArray: string[] = ["LastModifiedTime", "Title", "Filename", "ServerRedirectedURL", "Path",
                                        "FileExtension", "UniqueID", "SharedWithDetails", "SiteTitle", "SiteID",
                                        "EditorOWSUSER", "AuthorOWSUSER"];
    if (this.props.crawlTimeManagedPropertyName) {
      selectPropsArray.push(this.props.crawlTimeManagedPropertyName);
    }

    const queryText: string = "querytext='" + modeFql + scopeFql + "'";
    const selectProps: string = "selectproperties='" + selectPropsArray.join(",") + "'";
    let finalUri: string = baseUri + "?" + queryText + "&" + selectProps;
    if (graphFql) {
      finalUri += "&" + graphFql;
    }

    // TODO: support parallel requests
    const prom: Promise<IGetContentFuncResponse> = new Promise<IGetContentFuncResponse>((resolve: any, reject: any) => {
      self.queryForAllItems(finalUri, 0, rowLimit, null, resolve, reject);
    });
    return prom;
  }

  private queryForAllItems(uri: string, startIndex: number, rowlimit: number, response: IGetContentFuncResponse, resolve: any, reject: any): void {
    // Don't fetch more items than allowed
    if (startIndex + rowlimit > this.props.limitRowsFetched) {
      rowlimit = this.props.limitRowsFetched - startIndex;
    }
    const pagedUri: string = uri + "&startRow=" + startIndex + "&rowLimit=" + rowlimit;
    this.log.logInfo("Submitting request to " + pagedUri);

    const headers: Headers = new Headers();
    headers.append("odata-version", "3.0");
    headers.append("accept", "application/json;odata=nometadata");
    this.props.context.httpClient.get(pagedUri, { headers: headers })
      .then((r1: Response) => {
        this.log.logInfo("Recieved response from " + pagedUri);
        if (r1.ok) {
          r1.json().then((r) => {
            const searchResponse: ISearchResponse = TransformSearchResponse(r);
            const currentResponse: IGetContentFuncResponse = this.transformSearchResultsToResponseObject(searchResponse, this.props.noResultsString);
            if (response) {
              response.results.push(...currentResponse.results);
            }
            else {
              response = currentResponse;
              // for (let i=0; i<10; i++) {
              //   const newResults = response.results.slice();
              //   for (let j=0; j<newResults.length; j++) {
              //     debugger;
              //     const r: ISecurableObject = newResults[j];
              //     response.results.push({
              //       crawlTime: r.crawlTime,
              //       createdBy: r.createdBy,
              //       fileExtension: r.fileExtension,
              //       key: r.key + i.toString(),
              //       lastModifiedTime: r.lastModifiedTime,
              //       modifiedBy: r.modifiedBy,
              //       sharedBy: r.sharedBy,
              //       sharedWith: r.sharedWith,
              //       siteID: r.siteID,
              //       siteTitle: r.siteTitle,
              //       title: r.title,
              //       type: r.type,
              //       url: r.url
              //     });
              //   }
              // }
            }
            let getAnotherPage: boolean = false;
            if (!currentResponse.isError && response.results.length < this.props.limitRowsFetched) {
              // Get the next page if results === rowlimit
              if (searchResponse.rowCount === rowlimit && startIndex + rowlimit < searchResponse.totalRows) {
                getAnotherPage = true;
              }
            }
            if (getAnotherPage) {
              // Recursive call
              this.log.logInfo("Fetching an additional page of results");
              this.queryForAllItems(uri, startIndex + rowlimit, rowlimit, response, resolve, reject);
            }
            else {
              resolve(response);
            }
          });
        }
        else {
          reject({
            extContent: [],
            controlMode: ControlMode.Message,
            message: r1.statusText
          });
        }
      }, (error: any) => {
        reject({
          extContent: [],
          controlMode: ControlMode.Message,
          message: "Sorry, there was an error submitting the request"
        });
      });
  }

  private transformSearchResultsToResponseObject(searchResponse: ISearchResponse, noResultsString: string): IGetContentFuncResponse {
    let isError: boolean = false;
    let message: string = "";
    const securableObjects: ISecurableObject[] = [];
    if (searchResponse.isSuccess) {
      if (searchResponse && searchResponse.isSuccess && searchResponse.results.length > 0) {
        try {
          searchResponse.results.forEach((doc: any) => {
            // Parse out SharedWithDetails
            const sharedWith: string[] = [];
            const sharedBy: string[] = [];
            if (doc.SharedWithDetails) {
              try {
                const sharedWithDetails: any = JSON.parse(doc.SharedWithDetails);
                for (const sharedWithUser in sharedWithDetails) {
                  if (sharedWithDetails.hasOwnProperty(sharedWithUser)) {
                    const sharedWithUserDisplayName: string = ParseDisplayNameFromExtUserAccountName(sharedWithUser);
                    sharedWith.push(sharedWithUserDisplayName);
                    const sharedByUser: string = sharedWithDetails[sharedWithUser]["LoginName"];
                    sharedBy.push(sharedByUser);
                  }
                }
                sharedWith.sort();
                sharedBy.sort();
              }
              catch (e) {
                this.log.logError("Failed to parse SharedWithDetails", e.message);
              }
            }

            // Parse CrawlTime if managed property provided and populated
            let crawlTime: Date = null;
            let isCrawlTimeInvalid: boolean = true;
            if (this.props.crawlTimeManagedPropertyName) {
              const crawlTimeString: string = doc[this.props.crawlTimeManagedPropertyName];
              if (crawlTimeString) {
                crawlTime = new Date(crawlTimeString);
                const now: Date = new Date();
                const old: Date = new Date(now.getFullYear(), now.getMonth() - 1, now.getDate());
                isCrawlTimeInvalid = (crawlTime > now || crawlTime < old);
              }
            }

            // Parse last modified date
            let lastModifedTime: Date = null;
            let isLastModifiedTimeInvalid: boolean = true;
            if (doc.LastModifiedTime) {
              lastModifedTime = new Date(doc.LastModifiedTime);
              isLastModifiedTimeInvalid = false;
            }

            // Parse OWSUSER fields
            const editor: IOwsUser = ParseOWSUSER(doc.EditorOWSUSER);
            const author: IOwsUser = ParseOWSUSER(doc.AuthorOWSUSER);

            // Create ISecurableObject from search results
            securableObjects.push({
              title: { data: doc.Filename, displayValue: doc.Filename },
              fileExtension: { data: doc.FileExtension, displayValue: doc.FileExtension },
              lastModifiedTime: { data: lastModifedTime, displayValue: isLastModifiedTimeInvalid ? "" : ToColloquialDateString(lastModifedTime) },
              siteID: { data: doc.SiteID, displayValue: doc.SiteID },
              siteTitle: { data: doc.SiteTitle, displayValue: doc.SiteTitle },
              url: { data: doc.ServerRedirectedURL || doc.Path, displayValue: doc.ServerRedirectedURL || doc.Path },
              type: { data: SecurableObjectType.Document, displayValue: "Document" },
              sharedBy: { data: sharedBy, displayValue: sharedBy.join(", ") },
              sharedWith: { data: sharedWith, displayValue: sharedWith.join(", ") },
              crawlTime: { data: crawlTime, displayValue: isCrawlTimeInvalid ? "" : ToColloquialDateString(crawlTime) },
              modifiedBy: { data: editor, displayValue: editor.preferredName },
              createdBy: { data: author, displayValue: author.preferredName },
              key: doc.UniqueID || doc.Path
            });
          });
        }
        catch (e) {
          isError = true;
          message = "Sorry, there was an error parsing the response.";
          this.log.logError(message, e.message);
        }
      }

      if (!isError && securableObjects.length < 1) {
        message = noResultsString;
      }
    }
    else {
      isError = true;
      message = searchResponse.message;
    }

    return {
      results: securableObjects,
      isError: isError,
      message: message
    };
  }
}
