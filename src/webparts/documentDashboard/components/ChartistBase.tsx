import * as Chartist from "chartist";
import * as React from "react";

import {
  IChart,
  IChartItem
} from "../classes/Interfaces";

export abstract class ChartistBase extends React.Component<IChart, IChart> {

  protected responsiveOptions: Chartist.IResponsiveOptionTuple<Chartist.IPieChartOptions>[] = [
    ["screen and (max-width: 1024px)", {
      labelOffset: 0,
      chartPadding: {
        top: 30,
        right: 5,
        bottom: 30,
        left: 5
      }
    }]
  ];

  public render(): JSX.Element {
      return (
        <div id="chartist" className="ct-chart ct-perfect-fourth"></div>
      );
      // ct-golden-section
      // ct-perfect-fourth
      // ct-major-twelfth
  }

  public componentDidMount(): void {
    this.renderChart();
  }

  public abstract renderChart(): void;

  protected getChartistData(maxGroups: number): Chartist.IChartistData {
    // Create a object of chart items
    const chartItemDatas: IChartItem[] = [];
    const chartItemsDict: any = {};

    this.props.items.forEach(dataPoint => {
      const dataPointFromDict: IChartItem = chartItemsDict[dataPoint.data];
      if (dataPointFromDict) {
        dataPointFromDict.weight++;
      }
      else {
        chartItemsDict[dataPoint.data] = dataPoint;
        chartItemDatas.push(dataPoint);
      }
    });

    // Find the top (maxGroups - 1). Then add all other groups together as an 'other' group
    let finalDataToChart: IChartItem[] = null;
    if (maxGroups < chartItemDatas.length) {
      chartItemDatas.sort((a, b) => b.weight - a.weight);
      const sliceIndex: number = maxGroups > 2 ? maxGroups - 1 : 1;
      const firstGroups: IChartItem[] = chartItemDatas.slice(0, sliceIndex);
      const otherGroup: IChartItem = chartItemDatas.slice(sliceIndex).reduce((p, c) => { p.weight += c.weight; return p; });
      otherGroup.label = "Other";
      otherGroup.data = "";
      firstGroups.push(otherGroup);
      finalDataToChart = firstGroups;
    }
    else {
      finalDataToChart = chartItemDatas;
    }

    // Sort on data to support timeline charts
    finalDataToChart.sort((a, b) => a.data.localeCompare(b.data));

    const data: Chartist.IChartistData = {
      labels: [],
      series: []
    };

    finalDataToChart.forEach(d => {
      (data.labels as string[]).push(d.label);
      (data.series as number[]).push(d.weight);
    });

    return data;
  }

  protected reduceSumIChartistSeriesData(prev: Chartist.IChartistSeriesData, curr: Chartist.IChartistSeriesData): number {
    return this.reduceSumNumber(prev.value, curr.value);
  }

  protected reduceSumNumber(prev: number, curr: number): number {
    return prev + curr;
  }
}
