import Highcharts from 'highcharts';
import 'highcharts/modules/exporting';
import 'highcharts/highcharts-more';
import 'highcharts/modules/stock';
import { runExtractorSuite } from './extract-chart.shared';

runExtractorSuite(
  Highcharts as unknown as Parameters<typeof runExtractorSuite>[0],
  `Highcharts ${(Highcharts as { version?: string }).version ?? '13'}`,
);
