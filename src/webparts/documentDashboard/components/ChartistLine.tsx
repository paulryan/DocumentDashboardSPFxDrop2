import * as Chartist from "chartist";

import "../DocumentDashboard.module.css";

import {
  ChartistBase
} from "./ChartistBase";

export default class ChartistLine extends ChartistBase {

  public renderChart(): void {

    const data: Chartist.IChartistData = this.getChartistData(this.props.maxGroups);

    // Add data points for every day between earliest and laters item
    // const newLabels: string[] = [];
    // const newSeries: number[] = [];

    // const firstDay: Date = new Date(data.series[0] as number);
    // const finalDay: Date = new Date(data.series[data.series.length - 1] as number);

    // for (let day = firstDay; day <= finalDay; day.setDate(day.getDate() + 1)) {

    // }

    // data.labels = (data.labels as string[]).map(l => parseInt(l));

    // const lowLabel: number = data.labels[0] as number;
    // const highLabal: number = data.labels[data.labels.length - 1] as number;
    // const labelDelta: number = (highLabal - lowLabel); // * 0.1;

    const options: Chartist.ILineChartOptions = {
      // axisX: {
      //   type: Chartist.AutoScaleAxis,
      //   labelOffset: {
      //       x: 20,
      //       y: 0
      //     }
      // },
      axisY: {
        type: Chartist.FixedScaleAxis,
        low: 0,
        onlyInteger: true
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

    let line: any = new Chartist.Line("#chartist", data, options, this.responsiveOptions);
    // tslint ignore..
    line = line;
  }
}
