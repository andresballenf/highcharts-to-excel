/**
 * Highcharts 11 vs 12 vs 13 compatibility matrix. Each version file renders every fixture with its own
 * Highcharts build, runs `analyzeChartCompatibility` and (for editable charts) the full export,
 * and checks both against the SAME expected table below — so v11, v12 and v13 are proven identical in
 * chart type, Excel type and editability without depending on test-file order. Each run writes
 * `tests/output/compatibility-<version>.json`.
 */

import { mkdirSync, writeFileSync } from 'node:fs';
import { resolve } from 'node:path';
import type { Options } from 'highcharts';
import { afterAll, afterEach, describe, expect, it } from 'vitest';
import { analyzeChartCompatibility, ExportError, exportHighchartsToXlsx } from '../../src/index';
import type { DiagnosticCode } from '../../src/types/diagnostics';
import { allFixtures } from '../fixtures/highcharts-options';
import { inspectXlsx } from '../helpers/inspect-xlsx';
import { destroyAll, renderChart, type HighchartsLike } from '../helpers/render-chart';

export interface ExpectedCompat {
  sourceChartType: string;
  excelChartType: string | null;
  editable: boolean;
  /** Codes that must be present (subset check). */
  requiredCodes?: DiagnosticCode[];
}

export const EXPECTED_COMPATIBILITY: Readonly<Record<string, ExpectedCompat>> = {
  simpleLine: { sourceChartType: 'line', excelChartType: 'line', editable: true },
  multiLine: { sourceChartType: 'line', excelChartType: 'line', editable: true },
  splineChart: { sourceChartType: 'spline', excelChartType: 'line', editable: true },
  columnChart: { sourceChartType: 'column', excelChartType: 'column', editable: true },
  stackedColumn: {
    sourceChartType: 'column',
    excelChartType: 'stackedColumn',
    editable: true,
    requiredCodes: ['APPROXIMATED_CHART_TYPE'],
  },
  percentStackedColumn: { sourceChartType: 'column', excelChartType: 'percentStackedColumn', editable: true },
  barChart: { sourceChartType: 'bar', excelChartType: 'bar', editable: true },
  areaChart: { sourceChartType: 'area', excelChartType: 'area', editable: true },
  stackedArea: { sourceChartType: 'area', excelChartType: 'stackedArea', editable: true },
  pieChart: { sourceChartType: 'pie', excelChartType: 'pie', editable: true },
  doughnutChart: { sourceChartType: 'pie', excelChartType: 'doughnut', editable: true },
  scatterChart: { sourceChartType: 'scatter', excelChartType: 'scatter', editable: true },
  datetimeChart: { sourceChartType: 'line', excelChartType: 'line', editable: true },
  percentChart: { sourceChartType: 'column', excelChartType: 'column', editable: true },
  customColors: { sourceChartType: 'column', excelChartType: 'column', editable: true },
  styledMode: {
    sourceChartType: 'column',
    excelChartType: 'combo:column+line',
    editable: true,
    requiredCodes: ['STYLED_MODE_FALLBACK', 'MIXED_SERIES_TYPES'],
  },
  customAxes: { sourceChartType: 'line', excelChartType: 'line', editable: true },
  nullNegative: {
    sourceChartType: 'line',
    excelChartType: 'combo:line+column',
    editable: true,
    requiredCodes: ['NULL_VALUES', 'MIXED_SERIES_TYPES'],
  },
  hiddenSeries: {
    sourceChartType: 'line',
    excelChartType: 'line',
    editable: true,
    requiredCodes: ['HIDDEN_SERIES_EXCLUDED'],
  },
  comboChart: {
    sourceChartType: 'column',
    excelChartType: 'combo:column+line',
    editable: true,
    requiredCodes: ['MIXED_SERIES_TYPES'],
  },
  secondaryAxis: {
    sourceChartType: 'column',
    excelChartType: 'combo:column+line',
    editable: true,
    requiredCodes: ['SECONDARY_AXIS'],
  },
  customStyling: { sourceChartType: 'column', excelChartType: 'column', editable: true },
  bubbleChart: { sourceChartType: 'bubble', excelChartType: 'bubble', editable: true },
  unsupportedType: {
    sourceChartType: 'columnrange',
    excelChartType: null,
    editable: false,
    requiredCodes: ['UNSUPPORTED_CHART_TYPE'],
  },
  polarChart: { sourceChartType: 'line', excelChartType: null, editable: false, requiredCodes: ['UNSUPPORTED_POLAR'] },
  emptyChart: { sourceChartType: 'line', excelChartType: null, editable: false, requiredCodes: ['EMPTY_CHART'] },
  largeChart: { sourceChartType: 'line', excelChartType: 'line', editable: true },
};

export interface CompatRow {
  fixture: string;
  sourceChartType: string;
  excelChartType: string | null;
  editable: boolean;
  warningCodes: string[];
  /** Plot groups of the exported chart part (editable charts only). */
  plotGroups?: string[];
}

export function runCompatibilityMatrix(Highcharts: HighchartsLike, version: 'v11' | 'v12' | 'v13'): void {
  const rows: CompatRow[] = [];

  describe(`compatibility matrix (Highcharts ${Highcharts.version ?? version})`, () => {
    afterEach(() => destroyAll());

    afterAll(() => {
      const out = resolve(__dirname, '../output');
      mkdirSync(out, { recursive: true });
      writeFileSync(
        resolve(out, `compatibility-${version}.json`),
        `${JSON.stringify({ highchartsVersion: Highcharts.version ?? null, results: rows }, null, 2)}\n`,
      );
    });

    it('the expected table covers every fixture', () => {
      expect(Object.keys(EXPECTED_COMPATIBILITY).sort()).toEqual(Object.keys(allFixtures).sort());
    });

    it.each(Object.entries(allFixtures) as Array<[string, Options]>)('%s', async (name, fixture) => {
      const expected = EXPECTED_COMPATIBILITY[name]!;
      const chart = renderChart(Highcharts, fixture);
      const report = analyzeChartCompatibility(chart);
      const codes = report.warnings.map((w) => w.code);
      const row: CompatRow = {
        fixture: name,
        sourceChartType: report.sourceChartType,
        excelChartType: report.excelChartType,
        editable: report.editable,
        warningCodes: codes,
      };
      rows.push(row);

      expect({
        sourceChartType: report.sourceChartType,
        excelChartType: report.excelChartType,
        editable: report.editable,
      }).toEqual({
        sourceChartType: expected.sourceChartType,
        excelChartType: expected.excelChartType,
        editable: expected.editable,
      });
      for (const code of expected.requiredCodes ?? []) expect(codes, `${name} warning codes`).toContain(code);

      if (!report.editable) {
        expect(report.blocking.length).toBeGreaterThan(0);
        // The export refuses with CHART_NOT_EDITABLE rather than producing an image or an empty chart.
        const error = await exportHighchartsToXlsx(chart).then(
          () => null,
          (e: unknown) => e,
        );
        expect(error).toBeInstanceOf(ExportError);
        expect((error as ExportError).code).toBe('CHART_NOT_EDITABLE');
        return;
      }

      expect(report.blocking).toEqual([]);
      const result = await exportHighchartsToXlsx(chart);
      // Dry run and export agree.
      expect(result.report.excelChartType).toBe(report.excelChartType);
      expect(result.report.sourceChartType).toBe(report.sourceChartType);
      expect(result.report.editable).toBe(true);
      expect(result.warnings.map((w) => w.code)).toEqual(codes);
      const x = await inspectXlsx(result.bytes);
      x.assertWellFormed();
      expect(x.chartPaths()).toHaveLength(1);
      expect(x.hasImages()).toBe(false);
      row.plotGroups = x.plotGroupKinds(x.chartPaths()[0]!);
      expect(row.plotGroups.length).toBeGreaterThan(0);
    });
  });
}
