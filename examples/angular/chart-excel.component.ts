// Angular: a plain ViewChild + Highcharts.chart (no wrapper package needed). With highcharts-angular,
// capture the chart from its callbackFunction input instead and pass it to downloadHighchartsAsXlsx.
import { Component, ElementRef, ViewChild, type AfterViewInit, type OnDestroy } from '@angular/core';
import Highcharts from 'highcharts';
import 'highcharts/modules/exporting';
import { downloadHighchartsAsXlsx, installHighchartsExcelExport } from 'highcharts-editable-excel';

installHighchartsExcelExport(Highcharts);

@Component({
  selector: 'app-chart-excel',
  standalone: true,
  template: `<div #chartHost></div><button (click)="download()">Download editable Excel chart</button>`,
})
export class ChartExcelComponent implements AfterViewInit, OnDestroy {
  @ViewChild('chartHost', { static: true }) chartHost!: ElementRef<HTMLDivElement>;
  private chart?: Highcharts.Chart;

  ngAfterViewInit(): void {
    this.chart = Highcharts.chart(this.chartHost.nativeElement, {
      title: { text: 'Orders per region' },
      xAxis: { categories: ['North', 'South', 'East', 'West'] },
      series: [{ type: 'bar', name: 'Orders', data: [320, 210, 180, 260] }],
    });
  }

  async download(): Promise<void> {
    if (this.chart) await downloadHighchartsAsXlsx(this.chart, { filename: 'orders' });
  }

  ngOnDestroy(): void {
    this.chart?.destroy();
  }
}
