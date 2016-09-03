export class Logger {
  private currentContext: string;
  constructor (currentContext: string) {
    this.currentContext = currentContext;
  }

  public logInfo (message: string): void {
    this._log("INFO ", message);
  }

  public logWarn (message: string): void {
    this._log("WARN ", message);
  }

  public logError (message: string, exception: string = null): void {
    const errMsg: string = (exception === null) ? `${message}` : `${message} ${exception}`;
    this._log("ERROR", errMsg, true);
  }

  private _log (level: string, message: string, isError: boolean = false): void {
    if (console && typeof console.log === "function") {
      const logMsg: string = `${level}:  ${this.currentContext}:  ${message}`;
      if (isError) {
        console.error(logMsg);
      }
      else {
        /* tslint:disable */
        console.log(logMsg);
        /* tslint:enable */
      }
    }
  }
}
