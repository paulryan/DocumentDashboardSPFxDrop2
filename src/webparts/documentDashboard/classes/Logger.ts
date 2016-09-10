export class Logger {
  private currentContext: string;
  constructor (currentContext: string) {
    this.currentContext = currentContext;
  }

  private lastTime: Date = new Date();

  public logTime (message: string): void {
    const thisTime: Date = new Date();
    const time: Date = new Date(thisTime.getTime() - this.lastTime.getTime());
    const timeString: string = `${time.getMinutes()}.${time.getSeconds()}:${time.getMilliseconds()}`;

    this.lastTime = thisTime;
    this._log("TIME ", `${timeString} ${message}`);
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
      /* tslint:disable */
      if (isError) {
        console.error(logMsg);
      }
      else {
        console.log(logMsg);
      }
      /* tslint:enable */
    }
  }
}
