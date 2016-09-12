import * as Chartist from "chartist";

import {
  ChartistBase
} from "./ChartistBase";

export default class ChartistPie extends ChartistBase {

  public renderChart(): void {
    const maxLabelLength: number = 12;
    const data: Chartist.IChartistData = this.getChartistData(true);
    const seriesTotal: number = (data.series as Chartist.IChartistSeriesData[]).map(d => d.value).reduce(this.reduceSumNumber);

    const options: Chartist.IPieChartOptions = {
      chartPadding: {
        top: 30,
        right: 50,
        bottom: 30,
        left: 50
      },
      donut: true,
      labelOffset: 50,
      labelDirection: "explode",
      // labelInterpolationFnc: (label: string, index: number): string => {
      //   const valueAsNumber: number = (data.series[index] as Chartist.IChartistSeriesData).value as number;
      //   const valueAsPercentage: number = Math.round(valueAsNumber / seriesTotal * 1000) / 10;
      //   if (label && label.length > maxLabelLength) {
      //     label = label.substr(0, maxLabelLength) + "...";
      //   }
      //   return `${valueAsPercentage}% ${label}`;
      // },
      plugins: [
        Chartist.plugins.tooltip({
          transformTooltipTextFnc: (value: string): string => {
            const valueAsPercentage: number = Math.round(parseInt(value) / seriesTotal * 1000) / 10;
            return `Count: ${value} (${valueAsPercentage}%)`;
          }
        })
      ]
    };

    let pie: any = new Chartist.Pie(this.getTargetElement(), data, options, this.responsiveOptions);
    // tslint ignore..
    pie = pie;
  }
}
