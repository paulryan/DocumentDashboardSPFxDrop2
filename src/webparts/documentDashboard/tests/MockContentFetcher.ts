import {
  SecurableObjectType
} from "../classes/Enums";

import {
  IContentFetcherProps,
  IGetContentFuncResponse,
  ISecurableObject,
  ISecurableObjectStore
} from "../classes/Interfaces";

export default class MockContentFetcher implements ISecurableObjectStore {

  public props: IContentFetcherProps;

  private content: ISecurableObject[] = [
    {
      title: { data: "My first document", displayValue: "My first document" },
      url: { data: "https://www.google.com/01", displayValue: "https://www.google.com/01" },
      type: { data: SecurableObjectType.Document, displayValue: "Document" },
      fileExtension: { data: "docx", displayValue: "docx" },
      lastModifiedTime: { data: new Date(), displayValue: (new Date()).toDateString() },
      sharedWith: { data: ["Paul Ryan"], displayValue: "Paul Ryan" },
      sharedBy: { data: ["Chris O'Brien"], displayValue: "Chris O'Brien" },
      siteID: { data: "1", displayValue: "1" },
      siteTitle: { data: "Team Site", displayValue: "Team Site" },
      crawlTime: { data: new Date(), displayValue: "Never" },
      modifiedBy: { data: { accountName: "", preferredName: "Paul Ryan", email: "" }, displayValue: "Paul Ryan" },
      createdBy: { data: { accountName: "", preferredName: "Chris O'Brien", email: "" }, displayValue: "Chris O'Brien" },
      key: "1"
    },
    {
      title: { data: "My second document", displayValue: "My second document" },
      url: { data: "https://www.google.com/02", displayValue: "https://www.google.com/02" },
      type: { data: SecurableObjectType.Document, displayValue: "Document" },
      fileExtension: { data: "docx", displayValue: "docx" },
      lastModifiedTime: { data: new Date(), displayValue: (new Date()).toDateString() },
      sharedWith: { data: ["Paul Ryan"], displayValue: "Paul Ryan" },
      sharedBy: { data: ["Chris O'Brien"], displayValue: "Chris O'Brien" },
      siteID: { data: "1", displayValue: "1" },
      siteTitle: { data: "Team Site", displayValue: "Team Site" },
      crawlTime: { data: new Date(), displayValue: "Never" },
      modifiedBy: { data: { accountName: "", preferredName: "Paul Ryan", email: "" }, displayValue: "Paul Ryan" },
      createdBy: { data: { accountName: "", preferredName: "Chris O'Brien", email: "" }, displayValue: "Chris O'Brien" },
      key: "2"
    }
  ];

  public constructor (props: IContentFetcherProps) {
    this.props = props;
  }

  public getContent(): Promise<IGetContentFuncResponse> {
    return new Promise<IGetContentFuncResponse>((resolve) => {
        resolve({
          results: this.content,
          isError: false,
          message: "Mocked documents"
        });
    });
  }
}
