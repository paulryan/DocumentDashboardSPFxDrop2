import * as Chartist from "chartist";

import {
  ChartistBase
} from "./ChartistBase";

export default class ChartistBar extends ChartistBase {

  public renderChart(): void {
    const maxLabelLength: number = 20;
    const data: Chartist.IChartistData = this.getChartistData(false);

    const options: Chartist.IBarChartOptions = {
      axisY: {
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
        right: 5,
        bottom: 30,
        left: 5
      }
    };

    // Bar graphs take an array of series as they support many bars
    data.series = [ data.series ];

    let bar: any = new Chartist.Bar(this.getTargetElement(), data, options, this.responsiveOptions);
    // tslint ignore..
    bar = bar;
  }
}
