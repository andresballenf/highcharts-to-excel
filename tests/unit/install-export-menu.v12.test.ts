import Highcharts from 'highcharts12';
import 'highcharts12/modules/exporting';
import 'highcharts12/highcharts-more';
import { runInstallSuite, type HighchartsNamespace } from './install-export-menu.shared';

runInstallSuite(
  Highcharts as unknown as HighchartsNamespace,
  `Highcharts ${(Highcharts as { version?: string }).version ?? '12'}`,
);
