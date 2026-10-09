import Highcharts from 'highcharts';
import 'highcharts/modules/exporting';
import 'highcharts/highcharts-more';
import { runInstallSuite, type HighchartsNamespace } from './install-export-menu.shared';

runInstallSuite(
  Highcharts as unknown as HighchartsNamespace,
  `Highcharts ${(Highcharts as { version?: string }).version ?? '13'}`,
);
