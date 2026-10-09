import Highcharts from 'highcharts11';
import HighchartsMore from 'highcharts11/highcharts-more';
import Exporting from 'highcharts11/modules/exporting';
import Stock from 'highcharts11/modules/stock';
import { runExtractorSuite } from './extract-chart.shared';

// Highcharts 11 modules are UMD factories under CommonJS: they must be called with the namespace.
for (const init of [Exporting, HighchartsMore, Stock] as unknown as Array<(h: unknown) => void>) init(Highcharts);

runExtractorSuite(
  Highcharts as unknown as Parameters<typeof runExtractorSuite>[0],
  `Highcharts ${(Highcharts as { version?: string }).version ?? '11'}`,
);
