import Highcharts from 'highcharts';
import 'highcharts/modules/exporting';
import 'highcharts/highcharts-more';
import type { HighchartsLike } from '../helpers/render-chart';
import { runExportFixtureSuite } from './export-fixtures.shared';

runExportFixtureSuite(Highcharts as unknown as HighchartsLike, 'v13');
