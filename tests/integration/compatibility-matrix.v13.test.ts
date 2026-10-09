import Highcharts from 'highcharts';
import 'highcharts/modules/exporting';
import 'highcharts/highcharts-more';
import type { HighchartsLike } from '../helpers/render-chart';
import { runCompatibilityMatrix } from './compatibility-matrix.shared';

runCompatibilityMatrix(Highcharts as unknown as HighchartsLike, 'v13');
