/**
 * Minimal ambient declarations so `tsc -p examples/tsconfig.json` can check the framework examples
 * without installing React, Angular or the Highcharts wrappers in this repository.
 * In a real app, delete this file: the installed packages provide the real types.
 */

declare module 'react' {
  export interface RefObject<T> {
    current: T | null;
  }
  export function useRef<T>(initialValue: T | null): RefObject<T>;
  export function useMemo<T>(factory: () => T, deps: readonly unknown[]): T;
}

declare module 'react/jsx-runtime' {
  export namespace JSX {
    type Element = unknown;
    interface IntrinsicElements {
      [name: string]: Record<string, unknown>;
    }
  }
  export const jsx: unknown;
  export const jsxs: unknown;
  export const Fragment: unknown;
}

declare module 'highcharts-react-official' {
  import type * as Highcharts from 'highcharts';
  namespace HighchartsReact {
    interface RefObject {
      chart: Highcharts.Chart;
      container: { current: HTMLDivElement | null };
    }
    interface Props {
      highcharts?: typeof Highcharts;
      options?: Highcharts.Options;
      ref?: { current: RefObject | null };
      [key: string]: unknown;
    }
  }
  const HighchartsReact: (props: HighchartsReact.Props) => unknown;
  export default HighchartsReact;
}

declare module '@angular/core' {
  export function Component(metadata: { selector: string; template: string; standalone?: boolean }): ClassDecorator;
  export function ViewChild(selector: string, options?: { static?: boolean }): PropertyDecorator;
  export class ElementRef<T = unknown> {
    nativeElement: T;
  }
  export interface AfterViewInit {
    ngAfterViewInit(): void;
  }
  export interface OnDestroy {
    ngOnDestroy(): void;
  }
}
