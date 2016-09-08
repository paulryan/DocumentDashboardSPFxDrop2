import {
  IOwsUser,
  ISearchResponse
} from "./Interfaces";

export function GetDateFqlString (date: Date): string {
  return `${date.getFullYear()}-${date.getMonth() + 1}-${date.getDate()}`; // 1900-01-01
}

export function ToShortDateString (date: Date): string {
  return `${ToVeryShortDateString(date)} ${date.getFullYear()}`;
}

export function ToVeryShortDateString (date: Date): string {
  // e.g. 18 Aug
  const months: string[] = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec" ];
  const ds: string = `${date.getDate()} ${months[date.getMonth()]}`;
  return ds;
}

export function ToColloquialDateString (then: Date): string {
  let returnString: string = ToShortDateString(then);
  const now: Date = new Date();
  const minsInHr: number = 60;

  const isSameDay: boolean = then.getFullYear() === now.getFullYear() && then.getMonth() === now.getMonth() && then.getDate() === now.getDate();
  if (isSameDay) {
      if (now > then) {
          const totalMinutesAgo: number = (now.getHours() * minsInHr) + now.getMinutes() - (then.getHours() * minsInHr) - then.getMinutes();
          const hoursAgo: number = Math.floor(totalMinutesAgo / minsInHr);
          const minsAgo: number = totalMinutesAgo % minsInHr;

          if (hoursAgo < 1) {
              returnString = "" + minsAgo + " minutes ago";
          }
          else if (hoursAgo === 1) {
              returnString = "" + hoursAgo + " hour and " + minsAgo + " minutes ago";
          }
          else {
              returnString = "" + hoursAgo + " hours and " + minsAgo + " minutes ago";
          }
      }
  }
  return returnString;
}

export function EnsureBracesOnGuidString(guidString: string): string {
  if (guidString) {
    guidString = "{" + guidString.trim().replace("{", "").replace("}", "") + "}";
  }
  return guidString;
}

export function ParseDisplayNameFromExtUserAccountName(extUserAccountName: string): string {
  let extUserDisplayName: string = "";
  if (extUserAccountName) {
    // We want the bit betwee the last index of | and the first index of #
    let startIndex: number = extUserAccountName.lastIndexOf("|");
    if (startIndex < 0) {
      startIndex = 0;
    }
    else {
      startIndex += "|".length;
    }
    let endIndex: number = extUserAccountName.indexOf("#", startIndex);
    if (endIndex < 0) {
      endIndex = extUserAccountName.length;
    }
    extUserDisplayName = extUserAccountName.substring(startIndex, endIndex);
  }
  return extUserDisplayName;
}

export function TransformSearchResponse(response: any): ISearchResponse {
  // Simplify the data strucutre
  const searchRowsSimplified: any[] = [];

  let rowCount: number = 0;
  let totalRows: number = 0;
  let totalRowsIncludingDuplicates: number = 0;
  let isSuccess: boolean = true;
  let message: string = "";

  if (response.PrimaryQueryResult && response.PrimaryQueryResult.RelevantResults) {
    try {
      const searchRows: any[] = response.PrimaryQueryResult.RelevantResults.Table.Rows;
      searchRows.forEach((d: any) => {
        const doc: any = {};
        d.Cells.forEach((c: any) => {
          doc[c.Key] = c.Value;
        });
        searchRowsSimplified.push(doc);
      });
      rowCount = response.PrimaryQueryResult.RelevantResults.RowCount;
      totalRows = response.PrimaryQueryResult.RelevantResults.TotalRows;
      totalRowsIncludingDuplicates = response.PrimaryQueryResult.RelevantResults.TotalRowsIncludingDuplicates;
    } catch (e) {
      isSuccess = false;
      message = e.toString();
    }
  }
  else {
    message = "There are no RelevantResults";
  }

  return {
    results: searchRowsSimplified,
    rowCount: rowCount,
    totalRows: totalRows,
    totalRowsIncludingDuplicates: totalRowsIncludingDuplicates,
    isSuccess: isSuccess,
    message: message
  };
}

export function ParseOWSUSER(owsUserString: string): IOwsUser {
  const user: IOwsUser = {
    preferredName: "",
    accountName: "",
    email: ""
  };
  if (owsUserString) {
    // | Paul Ryan | 693A30...36F6D i:0#.f|membership|paul.ryan@paulryan.onmicrosoft.com
    const owsUserArraySplitBar: string[] = owsUserString.split("|");
    if (owsUserArraySplitBar.length > 1) {
      user.preferredName = owsUserArraySplitBar[1].trim();
      user.email = owsUserArraySplitBar[owsUserArraySplitBar.length - 1].trim();
    }
    const owsUserArraySplitSpace: string[] = owsUserString.split(" ");
    if (owsUserArraySplitSpace.length > 1) {
      user.accountName = owsUserArraySplitSpace[owsUserArraySplitSpace.length - 1].trim();
    }
  }
  return user;
}
