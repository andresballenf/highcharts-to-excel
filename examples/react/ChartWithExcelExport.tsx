// React + highcharts-react-official (install both in your app; they are not dependencies of this library).
import { useRef } from 'react';
import Highcharts from 'highcharts';
import 'highcharts/modules/exporting';
import HighchartsReact from 'highcharts-react-official';
import { downloadHighchartsAsXlsx, installHighchartsExcelExport } from 'highcharts-editable-excel';

installHighchartsExcelExport(Highcharts); // adds the context-menu item; safe to call more than once

const options: Highcharts.Options = {
  title: { text: 'Monthly revenue' },
  xAxis: { categories: ['Jan', 'Feb', 'Mar', 'Apr'] },
  series: [{ type: 'column', name: 'Revenue', data: [12, 19, 15, 22] }],
};

export function ChartWithExcelExport() {
  const chartComponentRef = useRef<HighchartsReact.RefObject>(null);
  const onDownload = async () => {
    const chart = chartComponentRef.current?.chart;
    if (chart) await downloadHighchartsAsXlsx(chart, { filename: 'monthly-revenue' });
  };
  return (
    <div>
      <HighchartsReact highcharts={Highcharts} options={options} ref={chartComponentRef} />
      <button onClick={onDownload}>Download editable Excel chart</button>
    </div>
  );
}
