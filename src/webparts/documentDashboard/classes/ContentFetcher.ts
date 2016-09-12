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
    self.log.logTime("getContent() entered");

    const rowLimit: number = 500;
    const finalUri: string = this.getSearchQueryUri();

    const prom: Promise<IGetContentFuncResponse> = new Promise<IGetContentFuncResponse>((resolve: any, reject: any) => {
      // Single request to determine rowCount
      self.getPageOfContent(finalUri, 0, rowLimit)
      .then((r1: IGetContentFuncResponse) => {
        // Handle fetching additional pages in parallel
        const rowsToFetch: number = this.props.limitRowsFetched < r1.totalRows ? this.props.limitRowsFetched : r1.totalRows;
        const getMorePages: boolean = (!r1.isError && r1.rowCount < rowsToFetch && r1.rowCount === rowLimit);
        if (getMorePages) {
          self.getContentParallelBatchesRecursive(r1, finalUri, r1.rowCount, rowLimit, rowsToFetch)
          .then((r2: IGetContentFuncResponse) => {
            this.log.logTime("getContent() exited");
            resolve(r2);
          });
        }
        else {
          this.log.logTime("getContent() exited");
          resolve(r1);
        }
      });
    });
    return prom;
  }

  private getContentParallelBatchesRecursive(r1: IGetContentFuncResponse,
          finalUri: string, startIndex: number, rowLimit: number, rowsToFetch: number): Promise<IGetContentFuncResponse> {
    // In parallel
    const maxParallelism: number = 6; // TODO: optimise this for very large data sets. This should be optimised and not a parameter.
    const prom: Promise<IGetContentFuncResponse> = new Promise<IGetContentFuncResponse>((resolve: any, reject: any) => {
      const promArray: Promise<IGetContentFuncResponse>[] = [];
      let nextStartIndex = startIndex;
      for (; nextStartIndex < rowsToFetch && promArray.length < maxParallelism; nextStartIndex += rowLimit) {
        const p = this.getPageOfContent(finalUri, nextStartIndex, rowLimit);
        promArray.push(p);
      }
      Promise.all(promArray)
      .then((responses) => {
        responses.forEach(r2 => {
          if (!r1.isError) {
            if (r2.isError) {
              r1.isError = true;
              r1.message = r2.message;
            }
            else {
              r1.results.push(...r2.results);
              r1.rowCount += r2.rowCount;
            }
          }
        });
        const allResponsesHitRowLimit: boolean = responses.every(r => !r.isError && r.rowCount === rowLimit);

        // If more pages to fetch..
        const getMorePages: boolean = (allResponsesHitRowLimit && r1.rowCount < rowsToFetch);
        if (getMorePages) {
          this.getContentParallelBatchesRecursive(r1, finalUri, r1.rowCount, rowLimit, rowsToFetch)
          .then((r3: IGetContentFuncResponse) => {
            resolve(r1);
          });
        }
        else {
          resolve(r1);
        }
      });
    });
    return prom;
  }

  private getSearchQueryUri(): string {
    const baseUri: string = this.props.context.pageContext.web.absoluteUrl + "/_api/search/query";

    // "MY" should represent things I have created or edited or shared.
    // TODO: Can we get Graphy in a meaningful way?
    const un: string = this.props.context.pageContext.user.loginName;
    const me: string = un.substring(0, un.indexOf("@")); // TODO: get the query working with @ symbol.. .replace("@", "%40");
    let myFql: string = `EditorOWSUSER:${me} OR AuthorOWSUSER:${me}`;
    if (this.props.sharedWithManagedPropertyName) {
      myFql += ` OR ${this.props.sharedWithManagedPropertyName}:${me}`;
    }
    myFql = `(${myFql})`;

    const documentsFql: string = "ContentClass=STS_ListItem_DocumentLibrary";
    const extSharedFql: string = "ViewableByExternalUsers=1";
    const anonSharedFql: string = "ViewableByAnonymousUsers=1";

    let graphFql: string = "";
    let modeFql: string = "";
    if (this.props.mode === Mode.AllDocuments) {
      modeFql = `${documentsFql}`;
      //modeFql = "*";
    }
    else if (this.props.mode === Mode.MyDocuments) {
      modeFql = `${documentsFql} ${myFql} `;
    }
    else if (this.props.mode === Mode.AllExtSharedDocuments) {
      modeFql = `${extSharedFql}`;
    }
    else if (this.props.mode === Mode.MyExtSharedDocuments) {
      modeFql = `${extSharedFql} ${myFql}`;
    }
    else if (this.props.mode === Mode.AllAnonSharedDocuments) {
      modeFql = `${anonSharedFql}`;
    }
    else if (this.props.mode === Mode.MyAnonSharedDocuments) {
      modeFql = `${anonSharedFql} ${myFql}`;
    }
    else if (this.props.mode === Mode.RecentlyModifiedDocuments) {
      const now: Date = new Date();
      const earlier: Date = new Date(now.getFullYear(), now.getMonth(), now.getDate() - 14);
      modeFql = `${documentsFql} Write>${GetDateFqlString(earlier)}`;
    }
    else if (this.props.mode === Mode.InactiveDocuments) {
      const now: Date = new Date();
      const earlier: Date = new Date(now.getFullYear(), now.getMonth() - 6, now.getDate());
      modeFql = `${documentsFql} Write<${GetDateFqlString(earlier)}`;
    }
    else if (this.props.mode === Mode.Delve) {
      modeFql = `${documentsFql}`;
      graphFql = "properties='GraphQuery:ACTOR(ME)'";
    }
    else {
      this.log.logError("Unsupported mode: " + this.props.mode);
      return null;
    }

    let scopeFql: string = "";
    if (this.props.scope === SPScope.SiteCollection) {
      // HACK: site.id works in modern pages only, site.absoluteUrl works on classic pages only..
      const siteId: any = this.props.context.pageContext.site.id;
      const siteUrl: string = (this.props.context.pageContext.site as any).absoluteUrl;
      if (siteId) {
        const siteIdString: string = EnsureBracesOnGuidString(siteId.toString());
        scopeFql = ` SiteId:${siteIdString}`;
      }
      else if (siteUrl) {
        scopeFql = ` Path:${siteUrl}`;
      }
    }
    else if (this.props.scope === SPScope.Site) {
      // HACK: web.id works in modern pages only, web.absoluteUrl works on classic pages only..
      const webId: any = this.props.context.pageContext.web.id;
      const webUrl: string = (this.props.context.pageContext.web as any).absoluteUrl;
      if (webId) {
        const webIdString: string = EnsureBracesOnGuidString(webId.toString());
        scopeFql = ` WebId:${webIdString}`;
      }
      else if (webUrl) {
        scopeFql = ` Path:${webUrl}`;
      }
    }
    else if (this.props.scope === SPScope.Tenant) {
      // do nothing
    }
    else {
      this.log.logError("Unsupported scope: " + this.props.scope);
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
    return finalUri;
  }

  private getPageOfContent(uri: string, startIndex: number, rowLimit: number): Promise<IGetContentFuncResponse> {
    // Don't fetch more items than allowed
    if (startIndex + rowLimit > this.props.limitRowsFetched) {
      rowLimit = this.props.limitRowsFetched - startIndex;
    }
    const pagedUri: string = uri + "&startRow=" + startIndex + "&rowLimit=" + rowLimit;
    this.log.logInfo("Submitting request to " + pagedUri);

    const headers: Headers = new Headers();
    headers.append("odata-version", "3.0");
    headers.append("accept", "application/json;odata=nometadata");

    const prom: Promise<IGetContentFuncResponse> = new Promise<IGetContentFuncResponse>((resolve: any, reject: any) => {
      this.props.context.httpClient.get(pagedUri, { headers: headers })
      .then(
        (r1: Response) => {
          this.log.logInfo("Recieved response from " + pagedUri);
          if (r1.ok) {
            r1.json().then((r) => {
              const searchResponse: ISearchResponse = TransformSearchResponse(r);
              const currentResponse: IGetContentFuncResponse = this.transformSearchResultsToResponseObject(searchResponse, this.props.noResultsString);
              resolve(currentResponse);
            });
          }
          else {
            reject({
              extContent: [],
              controlMode: ControlMode.Message,
              message: r1.statusText
            });
          }
        },
        (error: any) => {
          reject({
            extContent: [],
            controlMode: ControlMode.Message,
            message: "Sorry, there was an error submitting the request"
          });
        }
      );
    });
    return prom;
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
      message: message,
      rowCount: searchResponse.rowCount,
      totalRows: searchResponse.totalRows,
      totalRowsIncludingDuplicates: searchResponse.totalRowsIncludingDuplicates
    };
  }
}
