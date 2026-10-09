import Highcharts from 'highcharts12';
import 'highcharts12/modules/exporting';
import 'highcharts12/highcharts-more';
import type { HighchartsLike } from '../helpers/render-chart';
import { runCompatibilityMatrix } from './compatibility-matrix.shared';

runCompatibilityMatrix(Highcharts as unknown as HighchartsLike, 'v12');
