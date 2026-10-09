/**
 * Structured diagnostics describing how faithfully a chart was translated.
 * Codes are stable identifiers; consumers may switch on them.
 */

export type DiagnosticSeverity = 'info' | 'warning' | 'error';

/**
 * Translation outcome classes:
 * - translated: represented exactly (or as exactly as Excel allows) — rarely reported, only when noteworthy
 * - approximated: represented with a visible difference
 * - unsupported: dropped; Excel has no equivalent
 * - blocking: the chart cannot be exported as a native editable chart
 */
export type TranslationOutcome = 'translated' | 'approximated' | 'unsupported' | 'blocking';

export type DiagnosticCode =
  // chart type
  | 'UNSUPPORTED_CHART_TYPE'
  | 'APPROXIMATED_CHART_TYPE'
  | 'MIXED_SERIES_TYPES'
  | 'UNSUPPORTED_SERIES_TYPE'
  | 'UNSUPPORTED_POLAR'
  | 'UNSUPPORTED_3D'
  | 'EMPTY_CHART'
  | 'EMPTY_SERIES'
  // data
  | 'DATA_GROUPED'
  | 'DATA_CROPPED'
  | 'DATA_MODE_FALLBACK'
  | 'HIDDEN_SERIES_EXCLUDED'
  | 'HIDDEN_SERIES_INCLUDED'
  | 'HIDDEN_POINT'
  | 'UNALIGNED_X_VALUES'
  | 'ROW_LIMIT_EXCEEDED'
  | 'COLUMN_LIMIT_EXCEEDED'
  | 'NULL_VALUES'
  | 'NON_NUMERIC_VALUE'
  // axes
  | 'APPROXIMATED_AXIS_SCALE'
  | 'UNSUPPORTED_AXIS_FEATURE'
  | 'MULTIPLE_X_AXES'
  | 'SECONDARY_AXIS'
  | 'APPROXIMATED_DATETIME'
  // styling
  | 'APPROXIMATED_COLOR'
  | 'UNRESOLVED_COLOR'
  | 'UNSUPPORTED_GRADIENT'
  | 'APPROXIMATED_FONT'
  | 'APPROXIMATED_FONT_SIZE'
  | 'APPROXIMATED_LEGEND_POSITION'
  | 'APPROXIMATED_DASH_STYLE'
  | 'APPROXIMATED_MARKER'
  | 'APPROXIMATED_LAYOUT'
  | 'UNSUPPORTED_STYLE'
  | 'STYLED_MODE_FALLBACK'
  | 'UNSUPPORTED_NUMBER_FORMAT'
  | 'APPROXIMATED_NUMBER_FORMAT'
  | 'UNSUPPORTED_FORMATTER'
  | 'APPROXIMATED_DATA_LABELS'
  // interactivity (never representable)
  | 'UNSUPPORTED_TOOLTIP'
  | 'UNSUPPORTED_ANNOTATION'
  | 'UNSUPPORTED_PLOT_BAND'
  // environment / integration
  | 'FORMULA_LIKE_TEXT_ESCAPED'
  | 'SHEET_NAME_ADJUSTED'
  | 'WRITER_LIMITATION';

export interface Diagnostic {
  code: DiagnosticCode;
  severity: DiagnosticSeverity;
  outcome: TranslationOutcome;
  /** Dot path into the source options, e.g. "series[1].marker.symbol" or "title.style.fontFamily". */
  property: string;
  message: string;
  /** Series index when the diagnostic relates to a single series. */
  seriesIndex?: number;
  /** Free-form extra data useful for debugging (always JSON-serializable). */
  details?: Record<string, unknown>;
}

export interface CompatibilityReport {
  /** True when a native, editable Excel chart can be produced. */
  editable: boolean;
  sourceChartType: string;
  /** Excel chart type that will be / was produced, e.g. "line", "column", "barClustered". Null when blocked. */
  excelChartType: string | null;
  /** Property paths grouped by outcome. */
  supported: string[];
  approximated: string[];
  unsupported: string[];
  blocking: string[];
  /** Diagnostics specifically about data semantics (grouping, cropping, hidden series, nulls). */
  dataConcerns: Diagnostic[];
  /** Every diagnostic in emission order. */
  warnings: Diagnostic[];
}

export function createDiagnostic(
  code: DiagnosticCode,
  outcome: TranslationOutcome,
  property: string,
  message: string,
  extra: Partial<Pick<Diagnostic, 'severity' | 'seriesIndex' | 'details'>> = {},
): Diagnostic {
  const severity: DiagnosticSeverity =
    extra.severity ??
    (outcome === 'blocking'
      ? 'error'
      : outcome === 'translated'
        ? 'info'
        : outcome === 'unsupported'
          ? 'warning'
          : 'warning');
  const d: Diagnostic = { code, severity, outcome, property, message };
  if (extra.seriesIndex !== undefined) d.seriesIndex = extra.seriesIndex;
  if (extra.details !== undefined) d.details = extra.details;
  return d;
}

/** Collects diagnostics; shared by extractors, translators and the writer. */
export class DiagnosticCollector {
  readonly items: Diagnostic[] = [];
  private readonly seen = new Set<string>();

  constructor(private readonly onWarning?: (d: Diagnostic) => void) {}

  add(d: Diagnostic): void {
    const key = `${d.code}|${d.property}|${d.seriesIndex ?? ''}`;
    if (this.seen.has(key)) return;
    this.seen.add(key);
    this.items.push(d);
    this.onWarning?.(d);
  }

  report(
    code: DiagnosticCode,
    outcome: TranslationOutcome,
    property: string,
    message: string,
    extra?: Partial<Pick<Diagnostic, 'severity' | 'seriesIndex' | 'details'>>,
  ): void {
    this.add(createDiagnostic(code, outcome, property, message, extra));
  }

  addAll(items: Iterable<Diagnostic>): void {
    for (const d of items) this.add(d);
  }

  hasBlocking(): boolean {
    return this.items.some((d) => d.outcome === 'blocking');
  }
}

export function buildCompatibilityReport(
  sourceChartType: string,
  excelChartType: string | null,
  diagnostics: readonly Diagnostic[],
  supportedProperties: readonly string[],
): CompatibilityReport {
  const approximated: string[] = [];
  const unsupported: string[] = [];
  const blocking: string[] = [];
  const dataCodes: ReadonlySet<DiagnosticCode> = new Set<DiagnosticCode>([
    'DATA_GROUPED',
    'DATA_CROPPED',
    'DATA_MODE_FALLBACK',
    'HIDDEN_SERIES_EXCLUDED',
    'HIDDEN_SERIES_INCLUDED',
    'HIDDEN_POINT',
    'UNALIGNED_X_VALUES',
    'ROW_LIMIT_EXCEEDED',
    'COLUMN_LIMIT_EXCEEDED',
    'NULL_VALUES',
    'NON_NUMERIC_VALUE',
    'APPROXIMATED_DATETIME',
  ]);
  const dataConcerns: Diagnostic[] = [];
  for (const d of diagnostics) {
    if (d.outcome === 'approximated') approximated.push(d.property);
    else if (d.outcome === 'unsupported') unsupported.push(d.property);
    else if (d.outcome === 'blocking') blocking.push(d.property);
    if (dataCodes.has(d.code)) dataConcerns.push(d);
  }
  const flagged = new Set([...approximated, ...unsupported, ...blocking]);
  return {
    editable: blocking.length === 0 && excelChartType !== null,
    sourceChartType,
    excelChartType,
    supported: supportedProperties.filter((p) => !flagged.has(p)),
    approximated: dedupe(approximated),
    unsupported: dedupe(unsupported),
    blocking: dedupe(blocking),
    dataConcerns,
    warnings: [...diagnostics],
  };
}

function dedupe(list: string[]): string[] {
  return [...new Set(list)];
}
