// Highcharts 12/13 render under jsdom without stubs; getBBox is provided for stability only.
if (typeof SVGElement !== 'undefined' && !(SVGElement.prototype as { getBBox?: unknown }).getBBox) {
  (SVGElement.prototype as { getBBox?: () => DOMRect }).getBBox = () =>
    ({ x: 0, y: 0, width: 10, height: 10 }) as DOMRect;
}
