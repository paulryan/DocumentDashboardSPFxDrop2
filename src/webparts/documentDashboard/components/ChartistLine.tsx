import * as Chartist from "chartist";

import "../DocumentDashboard.module.css";

import {
  ChartistBase
} from "./ChartistBase";

export default class ChartistLine extends ChartistBase {

  public renderChart(): void {
    const maxLabelLength: number = 20;
    const data: Chartist.IChartistData = this.getChartistData();

    const options: Chartist.ILineChartOptions = {
      axisY: {
        low: 0,
        onlyInteger: true,
        labelInterpolationFnc: (label: string, index: number): string => {
          if (label && label.length > maxLabelLength) {
            label = label.substr(0, maxLabelLength) + "...";
          }
          return label;
        }
      },
      chartPadding: {
        top: 30,
        right: 30,
        bottom: 30,
        left: 30
      }
    };

    // Line graphs take an array of series as they support many lines
    data.series = [ data.series ];

    let line: any = new Chartist.Line(this.getTargetElement(), data, options, this.responsiveOptions);
    // tslint ignore..
    line = line;
  }
}
