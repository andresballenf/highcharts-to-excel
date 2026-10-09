import Highcharts from 'highcharts11';
import HighchartsMore from 'highcharts11/highcharts-more';
import Exporting from 'highcharts11/modules/exporting';
import { runInstallSuite, type HighchartsNamespace } from './install-export-menu.shared';

// Highcharts 11 modules are UMD factories under CommonJS: they must be called with the namespace.
for (const init of [Exporting, HighchartsMore] as unknown as Array<(h: unknown) => void>) init(Highcharts);

runInstallSuite(
  Highcharts as unknown as HighchartsNamespace,
  `Highcharts ${(Highcharts as { version?: string }).version ?? '11'}`,
);
