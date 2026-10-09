/**
 * Package-level parts: content types, root rels, doc props, workbook, workbook rels, theme.
 */
import { escapeAttr, escapeXml, xmlDocument } from './xml';

const NS_PKG_REL = 'http://schemas.openxmlformats.org/package/2006/relationships';
const NS_MAIN = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const NS_R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

export const CT = {
  rels: 'application/vnd.openxmlformats-package.relationships+xml',
  xml: 'application/xml',
  png: 'image/png',
  workbook: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml',
  worksheet: 'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml',
  styles: 'application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml',
  theme: 'application/vnd.openxmlformats-officedocument.theme+xml',
  drawing: 'application/vnd.openxmlformats-officedocument.drawing+xml',
  chart: 'application/vnd.openxmlformats-officedocument.drawingml.chart+xml',
  core: 'application/vnd.openxmlformats-package.core-properties+xml',
  app: 'application/vnd.openxmlformats-officedocument.extended-properties+xml',
} as const;

export interface ContentTypeOverride {
  partName: string;
  contentType: string;
}

export function buildContentTypesXml(overrides: readonly ContentTypeOverride[]): string {
  return xmlDocument(
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
      `<Default Extension="rels" ContentType="${CT.rels}"/>` +
      `<Default Extension="xml" ContentType="${CT.xml}"/>` +
      `<Default Extension="png" ContentType="${CT.png}"/>` +
      overrides.map((o) => `<Override PartName="${escapeAttr(o.partName)}" ContentType="${o.contentType}"/>`).join('') +
      '</Types>',
  );
}

export function buildRootRelsXml(): string {
  return xmlDocument(
    `<Relationships xmlns="${NS_PKG_REL}">` +
      `<Relationship Id="rId1" Type="${REL}/officeDocument" Target="xl/workbook.xml"/>` +
      '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>' +
      `<Relationship Id="rId3" Type="${REL}/extended-properties" Target="docProps/app.xml"/>` +
      '</Relationships>',
  );
}

function w3cdtf(d: Date | undefined): string {
  const date = d instanceof Date && Number.isFinite(d.getTime()) ? d : new Date();
  return date.toISOString().replace(/\.\d{3}Z$/, 'Z');
}

export function buildCoreXml(props: { title?: string; creator?: string; created?: Date }): string {
  const created = w3cdtf(props.created);
  return xmlDocument(
    '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"' +
      ' xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/"' +
      ' xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">' +
      (props.title ? `<dc:title>${escapeXml(props.title)}</dc:title>` : '') +
      (props.creator ? `<dc:creator>${escapeXml(props.creator)}</dc:creator>` : '') +
      `<dcterms:created xsi:type="dcterms:W3CDTF">${created}</dcterms:created>` +
      `<dcterms:modified xsi:type="dcterms:W3CDTF">${created}</dcterms:modified>` +
      '</cp:coreProperties>',
  );
}

export function buildAppXml(): string {
  return xmlDocument(
    '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"' +
      ' xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">' +
      '<Application>highcharts-editable-excel</Application>' +
      '<DocSecurity>0</DocSecurity><ScaleCrop>false</ScaleCrop>' +
      '<LinksUpToDate>false</LinksUpToDate><SharedDoc>false</SharedDoc><HyperlinksChanged>false</HyperlinksChanged>' +
      '<AppVersion>16.0300</AppVersion>' +
      '</Properties>',
  );
}

export interface WorkbookSheetEntry {
  name: string;
  hidden: boolean;
  relId: string;
}

export function buildWorkbookXml(sheets: readonly WorkbookSheetEntry[], activeTab: number): string {
  return xmlDocument(
    `<workbook xmlns="${NS_MAIN}" xmlns:r="${NS_R}">` +
      '<workbookPr defaultThemeVersion="164011"/>' +
      `<bookViews><workbookView xWindow="0" yWindow="0" windowWidth="28800" windowHeight="12300"${activeTab > 0 ? ` firstSheet="0" activeTab="${activeTab}"` : ''}/></bookViews>` +
      '<sheets>' +
      sheets
        .map(
          (s, i) =>
            `<sheet name="${escapeAttr(s.name)}" sheetId="${i + 1}"${s.hidden ? ' state="hidden"' : ''} r:id="${s.relId}"/>`,
        )
        .join('') +
      '</sheets>' +
      '<calcPr calcId="191029"/>' +
      '</workbook>',
  );
}

export function buildWorkbookRelsXml(sheetCount: number): { xml: string; sheetRelIds: string[] } {
  const sheetRelIds: string[] = [];
  const rels: string[] = [];
  for (let i = 0; i < sheetCount; i++) {
    const id = `rId${i + 1}`;
    sheetRelIds.push(id);
    rels.push(`<Relationship Id="${id}" Type="${REL}/worksheet" Target="worksheets/sheet${i + 1}.xml"/>`);
  }
  rels.push(`<Relationship Id="rId${sheetCount + 1}" Type="${REL}/styles" Target="styles.xml"/>`);
  rels.push(`<Relationship Id="rId${sheetCount + 2}" Type="${REL}/theme" Target="theme/theme1.xml"/>`);
  return { xml: xmlDocument(`<Relationships xmlns="${NS_PKG_REL}">${rels.join('')}</Relationships>`), sheetRelIds };
}

// ---------------------------------------------------------------------------
// Theme (Office 2013+ default colours and fonts)
// ---------------------------------------------------------------------------

const srgb = (hex: string, mods = ''): string => `<a:srgbClr val="${hex}">${mods}</a:srgbClr>`.replace('></a:srgbClr>', '/>');
const ph = (mods: string): string => `<a:schemeClr val="phClr">${mods}</a:schemeClr>`;

const CLR_SCHEME =
  '<a:clrScheme name="Office">' +
  '<a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1>' +
  '<a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>' +
  `<a:dk2>${srgb('44546A')}</a:dk2>` +
  `<a:lt2>${srgb('E7E6E6')}</a:lt2>` +
  `<a:accent1>${srgb('4472C4')}</a:accent1>` +
  `<a:accent2>${srgb('ED7D31')}</a:accent2>` +
  `<a:accent3>${srgb('A5A5A5')}</a:accent3>` +
  `<a:accent4>${srgb('FFC000')}</a:accent4>` +
  `<a:accent5>${srgb('5B9BD5')}</a:accent5>` +
  `<a:accent6>${srgb('70AD47')}</a:accent6>` +
  `<a:hlink>${srgb('0563C1')}</a:hlink>` +
  `<a:folHlink>${srgb('954F72')}</a:folHlink>` +
  '</a:clrScheme>';

const FONT_SCHEME =
  '<a:fontScheme name="Office">' +
  '<a:majorFont><a:latin typeface="Calibri Light" panose="020F0302020204030204"/><a:ea typeface=""/><a:cs typeface=""/></a:majorFont>' +
  '<a:minorFont><a:latin typeface="Calibri" panose="020F0502020204030204"/><a:ea typeface=""/><a:cs typeface=""/></a:minorFont>' +
  '</a:fontScheme>';

const gradFill = (stops: Array<[number, string]>, ang: number): string =>
  '<a:gradFill rotWithShape="1"><a:gsLst>' +
  stops.map(([pos, mods]) => `<a:gs pos="${pos}">${ph(mods)}</a:gs>`).join('') +
  `</a:gsLst><a:lin ang="${ang}" scaled="0"/></a:gradFill>`;

const FMT_SCHEME =
  '<a:fmtScheme name="Office">' +
  '<a:fillStyleLst>' +
  `<a:solidFill>${ph('')}</a:solidFill>` +
  gradFill(
    [
      [0, '<a:lumMod val="110000"/><a:satMod val="105000"/><a:tint val="67000"/>'],
      [50000, '<a:lumMod val="105000"/><a:satMod val="103000"/><a:tint val="73000"/>'],
      [100000, '<a:lumMod val="105000"/><a:satMod val="109000"/><a:tint val="81000"/>'],
    ],
    5400000,
  ) +
  gradFill(
    [
      [0, '<a:satMod val="103000"/><a:lumMod val="102000"/><a:tint val="94000"/>'],
      [50000, '<a:satMod val="110000"/><a:lumMod val="100000"/><a:shade val="100000"/>'],
      [100000, '<a:lumMod val="99000"/><a:satMod val="120000"/><a:shade val="78000"/>'],
    ],
    5400000,
  ) +
  '</a:fillStyleLst>' +
  '<a:lnStyleLst>' +
  [6350, 12700, 19050]
    .map(
      (w) =>
        `<a:ln w="${w}" cap="flat" cmpd="sng" algn="ctr"><a:solidFill>${ph('')}</a:solidFill><a:prstDash val="solid"/><a:miter lim="800000"/></a:ln>`,
    )
    .join('') +
  '</a:lnStyleLst>' +
  '<a:effectStyleLst>' +
  '<a:effectStyle><a:effectLst/></a:effectStyle>' +
  '<a:effectStyle><a:effectLst/></a:effectStyle>' +
  '<a:effectStyle><a:effectLst><a:outerShdw blurRad="57150" dist="19050" dir="5400000" algn="ctr" rotWithShape="0">' +
  '<a:srgbClr val="000000"><a:alpha val="63000"/></a:srgbClr></a:outerShdw></a:effectLst></a:effectStyle>' +
  '</a:effectStyleLst>' +
  '<a:bgFillStyleLst>' +
  `<a:solidFill>${ph('')}</a:solidFill>` +
  `<a:solidFill>${ph('<a:tint val="95000"/><a:satMod val="170000"/>')}</a:solidFill>` +
  gradFill(
    [
      [0, '<a:tint val="93000"/><a:satMod val="150000"/><a:shade val="98000"/><a:lumMod val="102000"/>'],
      [50000, '<a:tint val="98000"/><a:satMod val="130000"/><a:shade val="90000"/><a:lumMod val="103000"/>'],
      [100000, '<a:shade val="63000"/><a:satMod val="120000"/>'],
    ],
    5400000,
  ) +
  '</a:bgFillStyleLst>' +
  '</a:fmtScheme>';

export function buildThemeXml(): string {
  return xmlDocument(
    '<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office Theme">' +
      `<a:themeElements>${CLR_SCHEME}${FONT_SCHEME}${FMT_SCHEME}</a:themeElements>` +
      '<a:objectDefaults/><a:extraClrSchemeLst/>' +
      '</a:theme>',
  );
}
