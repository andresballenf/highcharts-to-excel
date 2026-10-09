import Highcharts from 'highcharts12';
import 'highcharts12/modules/exporting';
import 'highcharts12/highcharts-more';
import 'highcharts12/modules/stock';
import { runExtractorSuite } from './extract-chart.shared';

runExtractorSuite(
  Highcharts as unknown as Parameters<typeof runExtractorSuite>[0],
  `Highcharts ${(Highcharts as { version?: string }).version ?? '12'}`,
);
