import * as Chartist from "chartist";
import * as React from "react";

import {
  IChart,
  IChartItem
} from "../classes/Interfaces";

import {
  ChartAxisOrder
} from "../classes/Enums";

import styles from "../DocumentDashboard.module.scss";

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
      const chartContainerClasses: string = `${styles.chartistContainer} ct-chart ct-golden-section`;
      return (
        <div className={chartContainerClasses}></div>
      );

      // ct-perfect-fourth
      // ct-golden-section
      // ct-minor-seventh
      // ct-major-tenth
      // ct-major-twelfth
  }

  public componentDidMount(): void {
    this.renderChart();
  }

  public abstract renderChart(): void;

  protected getTargetElement(): Element {
    const els: NodeListOf<Element> = this.props.domElement.getElementsByClassName("ct-chart");
    if (els.length === 1) {
      return els[0];
    }
    else {
      throw new Error("Failed to find the Chartist container DOM element");
    }
  }

  protected getChartistData(suppressXAxisPlots: boolean): Chartist.IChartistData {
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
    if (this.props.maxGroups < chartItemDatas.length) {
      if (this.props.chartAxisOrder === ChartAxisOrder.PrioritiseSmallestGroups) {
        chartItemDatas.sort((a, b) => a.weight - b.weight);
      }
      else {
        chartItemDatas.sort((a, b) => b.weight - a.weight);
      }
      const sliceIndex: number = this.props.maxGroups > 2 ? this.props.maxGroups - 1 : 1;
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

    // Sort on data to support timeline charts (if xAxis points are defined)
    let isUsingXAxis: boolean = false;
    if (!suppressXAxisPlots) {
      isUsingXAxis = finalDataToChart.every(c => typeof c.xAxis === "number");
      if (isUsingXAxis) {
        finalDataToChart.sort((a, b) => a.xAxis - b.xAxis);
      }
    }
    const data: Chartist.IChartistData = {
      labels: [],
      series: []
    };

    finalDataToChart.forEach((d, i) => {
      (data.labels as string[]).push(d.label);
      if (suppressXAxisPlots) {
        (data.series as number[]).push(d.weight);
      }
      else {
        (data.series as any[]).push({
          x: isUsingXAxis ? d.xAxis : i,
          y: d.weight
        });
      }
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
