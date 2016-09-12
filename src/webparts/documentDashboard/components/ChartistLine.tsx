import * as Chartist from "chartist";

import {
  ChartistBase
} from "./ChartistBase";

export default class ChartistLine extends ChartistBase {

  public renderChart(): void {
    const maxLabelLength: number = 20;
    const data: Chartist.IChartistData = this.getChartistData(false);
    const labelByXAxisDict: any = {};
    (data.series as any[]).forEach((d, i) => {
      labelByXAxisDict[d.x.toString()] = (data.labels as string[])[i];
    });

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
      axisX: {
        // type: Chartist.AutoScaleAxis,
        type: Chartist.FixedScaleAxis,
        ticks: (data.series as any[]).map(d => d.x),
        onlyInteger: true,
        labelInterpolationFnc: (label: number | string, index: number): string => {
          if (labelByXAxisDict[label.toString()]) {
            return labelByXAxisDict[label.toString()];
            // let l: string = ToVeryShortDateString(new Date(label));
            // return l;
          }
          else {
            return "";
          }
        }
      },
      chartPadding: {
        top: 30,
        right: 30,
        bottom: 30,
        left: 30
      },
      plugins: [
        Chartist.plugins.tooltip({
          transformTooltipTextFnc: (xyLabel: string): string => {
            const coordsArray: string[] = xyLabel.split(",");
            return "Count: " + coordsArray[coordsArray.length-1];
          }
        })
      ]
    };

    // Line graphs take an array of series as they support many lines
    data.series = [ data.series ];

    let line: any = new Chartist.Line(this.getTargetElement(), data, options, this.responsiveOptions);
    // tslint ignore..
    line = line;
  }
}
