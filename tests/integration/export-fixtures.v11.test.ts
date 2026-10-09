import Highcharts from 'highcharts11';
import HighchartsMore from 'highcharts11/highcharts-more';
import Exporting from 'highcharts11/modules/exporting';
import type { HighchartsLike } from '../helpers/render-chart';
import { runExportFixtureSuite } from './export-fixtures.shared';

// Highcharts 11 modules are UMD factories under CommonJS: they must be called with the namespace.
for (const init of [Exporting, HighchartsMore] as unknown as Array<(h: unknown) => void>) init(Highcharts);

runExportFixtureSuite(Highcharts as unknown as HighchartsLike, 'v11');
