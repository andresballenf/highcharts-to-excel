import Highcharts from 'highcharts12';
import 'highcharts12/modules/exporting';
import 'highcharts12/highcharts-more';
import type { HighchartsLike } from '../helpers/render-chart';
import { runExportFixtureSuite } from './export-fixtures.shared';

runExportFixtureSuite(Highcharts as unknown as HighchartsLike, 'v12');
