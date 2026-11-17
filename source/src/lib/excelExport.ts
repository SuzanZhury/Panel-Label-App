import { unzipSync, zipSync } from 'fflate';

const LABEL_ROWS = [5, 13, 21, 29, 37] as const;
const TEMPLATE_COLUMN_COUNT = 19;
const EXCEL_MAX_COLUMN_COUNT = 16_384;
const LABEL_CELL_ROWS = [...new Set(
  LABEL_ROWS.flatMap(row => [row - 4, row - 3, row - 2, row, row + 2]),
)].sort((a, b) => a - b);
const LABEL_ANCHOR_ROWS = LABEL_ROWS.map(row => row - 3);

type LabelValue = { panelLabel: string };
type LabelMetadata = { project: string; release: string; finish: string };

const decoder = new TextDecoder();
const encoder = new TextEncoder();

function getAttribute(attributes: string, name: string) {
  return attributes.match(new RegExp(`(?:^|\\s)${name}="([^"]*)"`))?.[1];
}

function decodeXmlFile(files: Record<string, Uint8Array>, path: string) {
  const bytes = files[path];
  if (!bytes) throw new Error(`The Excel template is missing ${path}.`);
  return decoder.decode(bytes);
}

function resolvePackagePath(target: string, baseDirectory = 'xl') {
  const parts = target.startsWith('/') ? [] : baseDirectory.split('/').filter(Boolean);
  for (const part of target.split('/')) {
    if (!part || part === '.') continue;
    if (part === '..') parts.pop();
    else parts.push(part);
  }
  return parts.join('/');
}

function findRelatedPackagePath(
  files: Record<string, Uint8Array>,
  sourcePath: string,
  relationshipId: string,
) {
  const sourceDirectory = sourcePath.slice(0, sourcePath.lastIndexOf('/'));
  const sourceFile = sourcePath.slice(sourcePath.lastIndexOf('/') + 1);
  const relationshipsPath = `${sourceDirectory}/_rels/${sourceFile}.rels`;
  const relationshipsXml = decodeXmlFile(files, relationshipsPath);
  const relationship = [...relationshipsXml.matchAll(/<Relationship\b([^>]*)\/>/g)]
    .find(match => getAttribute(match[1], 'Id') === relationshipId);
  const target = relationship && getAttribute(relationship[1], 'Target');
  if (!target) throw new Error(`The Excel template has no relationship ${relationshipId} for ${sourcePath}.`);
  return resolvePackagePath(target, sourceDirectory);
}

function findLabelsWorksheetPath(files: Record<string, Uint8Array>) {
  const workbookXml = decodeXmlFile(files, 'xl/workbook.xml');
  const labelsSheet = [...workbookXml.matchAll(/<sheet\b([^>]*)\/>/g)]
    .find(match => getAttribute(match[1], 'name') === 'LABLES');
  if (!labelsSheet) throw new Error('The original Excel template does not contain a LABLES worksheet.');

  const relationshipId = getAttribute(labelsSheet[1], 'r:id');
  if (!relationshipId) throw new Error('The LABLES worksheet has no workbook relationship.');
  return findRelatedPackagePath(files, 'xl/workbook.xml', relationshipId);
}

function findLabelsDrawingPath(files: Record<string, Uint8Array>, worksheetPath: string) {
  const worksheetXml = decodeXmlFile(files, worksheetPath);
  const drawing = /<drawing\b([^>]*)\/>/.exec(worksheetXml);
  const relationshipId = drawing && getAttribute(drawing[1], 'r:id');
  if (!relationshipId) throw new Error('The LABLES worksheet has no original artwork drawing.');
  return findRelatedPackagePath(files, worksheetPath, relationshipId);
}

function escapeXml(value: string) {
  return value
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;');
}

function setCellText(xml: string, address: string, value: string) {
  const escapedAddress = address.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  const selfClosingPattern = new RegExp(
    `<c\\b(?=[^>]*\\br="${escapedAddress}")([^>]*?)\\s*\\/>`,
  );
  const pairedPattern = new RegExp(
    `<c\\b(?=[^>]*\\br="${escapedAddress}")([^>]*)>([\\s\\S]*?)<\\/c>`,
  );
  const selfClosingMatch = selfClosingPattern.exec(xml);
  const pairedMatch = selfClosingMatch ? undefined : pairedPattern.exec(xml);
  const match = selfClosingMatch ?? pairedMatch;
  if (!match) {
    if (!value) return xml;
    throw new Error(`The original Excel template is missing label cell ${address}.`);
  }

  const originalAttributes = match[1] ?? '';
  const attributes = originalAttributes
    .replace(/\s+t="[^"]*"/g, '')
    .replace(/\s+vm="[^"]*"/g, '');
  const replacement = value
    ? `<c${attributes} t="inlineStr"><is><t${/^\s|\s$/.test(value) ? ' xml:space="preserve"' : ''}>${escapeXml(value)}</t></is></c>`
    : `<c${attributes}/>`;
  return `${xml.slice(0, match.index)}${replacement}${xml.slice(match.index + match[0].length)}`;
}

function columnName(columnIndex: number) {
  let value = columnIndex + 1;
  let name = '';
  while (value > 0) {
    const remainder = (value - 1) % 26;
    name = String.fromCharCode(65 + remainder) + name;
    value = Math.floor((value - 1) / 26);
  }
  return name;
}

function cellMatch(xml: string, address: string) {
  const escapedAddress = address.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  const selfClosingPattern = new RegExp(
    `<c\\b(?=[^>]*\\br="${escapedAddress}")([^>]*?)\\s*\\/>`,
  );
  const pairedPattern = new RegExp(
    `<c\\b(?=[^>]*\\br="${escapedAddress}")([^>]*)>([\\s\\S]*?)<\\/c>`,
  );
  const selfClosingMatch = selfClosingPattern.exec(xml);
  return selfClosingMatch ?? pairedPattern.exec(xml);
}

function replaceOrInsertCell(xml: string, address: string, replacement: string) {
  const existing = cellMatch(xml, address);
  if (existing) {
    return `${xml.slice(0, existing.index)}${replacement}${xml.slice(existing.index + existing[0].length)}`;
  }

  const rowNumber = Number(address.match(/\d+$/)?.[0]);
  const rowPattern = new RegExp(
    `(<row\\b(?=[^>]*\\br="${rowNumber}")[^>]*>)([\\s\\S]*?)(<\\/row>)`,
  );
  const row = rowPattern.exec(xml);
  if (!row) throw new Error(`The original Excel template is missing row ${rowNumber}.`);
  const contentStart = row.index + row[1].length;
  const contentEnd = contentStart + row[2].length;
  return `${xml.slice(0, contentEnd)}${replacement}${xml.slice(contentEnd)}`;
}

function templateCellForAddress(templateXml: string, sourceAddress: string, targetAddress: string) {
  const sourceCell = cellMatch(templateXml, sourceAddress)?.[0];
  if (!sourceCell) throw new Error(`The original Excel template is missing label cell ${sourceAddress}.`);
  const escapedSource = sourceAddress.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  return sourceCell.replace(new RegExp(`(\\br=")${escapedSource}(")`), `$1${targetAddress}$2`);
}

function addTemplateColumnCells(xml: string, templateXml: string, columnIndex: number) {
  const column = columnName(columnIndex);
  for (const row of LABEL_CELL_ROWS) {
    const address = `${column}${row}`;
    xml = replaceOrInsertCell(xml, address, templateCellForAddress(templateXml, `A${row}`, address));
  }
  return xml;
}

function ensureLabelLogoCell(xml: string, templateXml: string, address: string, sourceAddress: string) {
  const existing = cellMatch(xml, address);
  if (existing?.[1]?.includes('vm="1"')) return xml;
  return replaceOrInsertCell(xml, address, templateCellForAddress(templateXml, sourceAddress, address));
}

function updateWorksheetColumnRange(xml: string, columnCount: number) {
  const lastColumn = columnName(columnCount - 1);
  let updated = xml.replace(
    /(<dimension\b[^>]*\bref=")A1:[^"]+(")/,
    `$1A1:${lastColumn}40$2`,
  );
  const columnsMatch = /<cols\b[^>]*>[\s\S]*?<\/cols>/.exec(updated);
  if (!columnsMatch) throw new Error('The LABLES worksheet has no column sizing information.');
  const extendedColumns = columnsMatch[0].replace(
    /<col\b([^>]*)\/>/g,
    (cell, attributes: string) => {
      const min = Number(getAttribute(attributes, 'min'));
      const max = Number(getAttribute(attributes, 'max'));
      if (min !== 1 || max < TEMPLATE_COLUMN_COUNT || max >= columnCount) return cell;
      const changedAttributes = attributes.replace(/\s+max="[^"]*"/, ` max="${columnCount}"`);
      return `<col${changedAttributes}/>`;
    },
  );
  if (extendedColumns === columnsMatch[0] && columnCount > TEMPLATE_COLUMN_COUNT) {
    throw new Error('The LABLES worksheet column sizing could not be extended.');
  }
  return updated.replace(columnsMatch[0], extendedColumns);
}

function setLabelsPrintArea(workbookXml: string, columnCount: number) {
  const sheets = [...workbookXml.matchAll(/<sheet\b([^>]*)\/>/g)];
  const sheetIndex = sheets.findIndex(match => getAttribute(match[1], 'name') === 'LABLES');
  if (sheetIndex < 0) throw new Error('The original Excel template does not contain a LABLES worksheet.');

  const finalColumn = columnName(columnCount - 1);
  const area = escapeXml(`'LABLES'!$A$1:$${finalColumn}$40`);
  const localSheetId = String(sheetIndex);
  const definitionPattern = /<definedName\b([^>]*)>([\s\S]*?)<\/definedName>/g;
  const definedNamesPattern = /<definedNames\b[^>]*>[\s\S]*?<\/definedNames>/;
  const definedNamesMatch = definedNamesPattern.exec(workbookXml);

  if (definedNamesMatch) {
    let printAreaFound = false;
    const updatedDefinedNames = definedNamesMatch[0].replace(
      definitionPattern,
      (definition, attributes: string) => {
        if (
          getAttribute(attributes, 'name') === '_xlnm.Print_Area' &&
          getAttribute(attributes, 'localSheetId') === localSheetId
        ) {
          printAreaFound = true;
          return `<definedName${attributes}>${area}</definedName>`;
        }
        return definition;
      },
    );
    const completedDefinedNames = printAreaFound
      ? updatedDefinedNames
      : updatedDefinedNames.replace(
          '</definedNames>',
          `<definedName name="_xlnm.Print_Area" localSheetId="${localSheetId}">${area}</definedName></definedNames>`,
        );
    return workbookXml.replace(definedNamesMatch[0], completedDefinedNames);
  }

  const sheetsClose = workbookXml.indexOf('</sheets>');
  if (sheetsClose < 0) throw new Error('The Excel template workbook has no sheets section.');
  const insertAt = sheetsClose + '</sheets>'.length;
  const definition = `<definedNames><definedName name="_xlnm.Print_Area" localSheetId="${localSheetId}">${area}</definedName></definedNames>`;
  return `${workbookXml.slice(0, insertAt)}${definition}${workbookXml.slice(insertAt)}`;
}

function getAnchorPosition(anchor: string) {
  const from = /<(?:[A-Za-z_][\w.-]*:)?from\b[^>]*>([\s\S]*?)<\/(?:[A-Za-z_][\w.-]*:)?from>/.exec(anchor)?.[1];
  if (!from) return null;
  const column = Number(/<(?:[A-Za-z_][\w.-]*:)?col\b[^>]*>(\d+)<\/(?:[A-Za-z_][\w.-]*:)?col>/.exec(from)?.[1]);
  const row = Number(/<(?:[A-Za-z_][\w.-]*:)?row\b[^>]*>(\d+)<\/(?:[A-Za-z_][\w.-]*:)?row>/.exec(from)?.[1]);
  const rowIndex = LABEL_ANCHOR_ROWS.indexOf(row);
  return Number.isInteger(column) && rowIndex >= 0
    ? { column, rowIndex, labelIndex: column * LABEL_ROWS.length + rowIndex }
    : null;
}

function updateAnchorColumn(anchor: string, columnIndex: number, columnWidthEmu: number) {
  const from = /<(?:[A-Za-z_][\w.-]*:)?from\b[^>]*>([\s\S]*?)<\/(?:[A-Za-z_][\w.-]*:)?from>/.exec(anchor)?.[1];
  const fromColumnOffset = Number(from?.match(/<(?:[A-Za-z_][\w.-]*:)?colOff\b[^>]*>(\d+)<\/(?:[A-Za-z_][\w.-]*:)?colOff>/)?.[1]);
  let nextId = 0;
  const updated = anchor
    .replace(
      /(<(?:[A-Za-z_][\w.-]*:)?col\b[^>]*>)(\d+)(<\/(?:[A-Za-z_][\w.-]*:)?col>)/g,
      (_match, open: string, _current: string, close: string) => `${open}${columnIndex}${close}`,
    )
    .replace(
      /(<(?:[A-Za-z_][\w.-]*:)?off\b[^>]*\bx=")(\d+)(")/,
      (_match, open: string, _x: string, close: string) => `${open}${columnIndex * columnWidthEmu + fromColumnOffset}${close}`,
    )
    .replace(
      /(<(?:[A-Za-z_][\w.-]*:)?cNvPr\b[^>]*\bid=")\d+("[^>]*\bname=")([^"]*)(")/,
      (_match, idOpen: string, nameOpen: string, name: string, close: string) => {
        nextId = 1;
        return `${idOpen}0${nameOpen}${name}${close}`;
      },
    );
  if (!nextId) throw new Error('The original Excel arrow artwork could not be copied safely.');
  return updated;
}

function updateLabelArrows(drawingXml: string, labelCount: number) {
  const anchorPattern = /<(?:[A-Za-z_][\w.-]*:)?twoCellAnchor\b[\s\S]*?<\/(?:[A-Za-z_][\w.-]*:)?twoCellAnchor>/g;
  const anchors = [...drawingXml.matchAll(anchorPattern)].map(match => match[0]);
  const anchorsByLabel = new Map<number, string[]>();
  const canonicalAnchorsByRow = new Map<number, string[]>();
  let maxShapeId = 0;

  for (const anchor of anchors) {
    const shapeId = Number(/<(?:[A-Za-z_][\w.-]*:)?cNvPr\b[^>]*\bid="(\d+)"/.exec(anchor)?.[1]);
    if (Number.isInteger(shapeId)) maxShapeId = Math.max(maxShapeId, shapeId);
    const position = getAnchorPosition(anchor);
    if (!position) continue;
    const byLabel = anchorsByLabel.get(position.labelIndex) ?? [];
    byLabel.push(anchor);
    anchorsByLabel.set(position.labelIndex, byLabel);
    if (position.column === 0) {
      const byRow = canonicalAnchorsByRow.get(position.rowIndex) ?? [];
      byRow.push(anchor);
      canonicalAnchorsByRow.set(position.rowIndex, byRow);
    }
  }

  const canonicalFirst = anchorsByLabel.get(0)?.[0];
  const canonicalSecond = anchorsByLabel.get(LABEL_ROWS.length)?.[0];
  const firstX = Number(canonicalFirst?.match(/<(?:[A-Za-z_][\w.-]*:)?off\b[^>]*\bx="(\d+)"/)?.[1]);
  const secondX = Number(canonicalSecond?.match(/<(?:[A-Za-z_][\w.-]*:)?off\b[^>]*\bx="(\d+)"/)?.[1]);
  const columnWidthEmu = secondX - firstX;
  if (!Number.isFinite(columnWidthEmu) || columnWidthEmu <= 0) {
    throw new Error('The original Excel column width could not be read from the label artwork.');
  }

  const keptAnchors: string[] = [];
  const unpositionedAnchors: string[] = [];
  for (const anchor of anchors) {
    const position = getAnchorPosition(anchor);
    if (!position) unpositionedAnchors.push(anchor);
    else if (position.labelIndex < labelCount) keptAnchors.push(anchor);
  }

  const addedAnchors: string[] = [];
  for (let labelIndex = 0; labelIndex < labelCount; labelIndex++) {
    if (anchorsByLabel.has(labelIndex)) continue;
    const rowIndex = labelIndex % LABEL_ROWS.length;
    const targetColumn = Math.floor(labelIndex / LABEL_ROWS.length);
    const templates = canonicalAnchorsByRow.get(rowIndex);
    if (!templates?.length) throw new Error(`The Excel template has no arrow artwork for label row ${rowIndex + 1}.`);
    for (const template of templates) {
      maxShapeId++;
      const creationId = `{00000000-0000-4000-8000-${maxShapeId.toString(16).padStart(12, '0')}}`;
      let clone = updateAnchorColumn(template, targetColumn, columnWidthEmu);
      clone = clone.replace(
        /(<(?:[A-Za-z_][\w.-]*:)?cNvPr\b[^>]*\bid=")0("[^>]*\bname=")[^"]*(")/,
        `$1${maxShapeId}$2Arrow: Up ${maxShapeId}$3`,
      );
      clone = clone.replace(
        /(<(?:[A-Za-z_][\w.-]*:)?creationId\b[^>]*\bid=")\{[^"]+\}(")/,
        `$1${creationId}$2`,
      );
      addedAnchors.push(clone);
    }
  }

  const closingRoot = /<\/(?:[A-Za-z_][\w.-]*:)?wsDr>/;
  if (!closingRoot.test(drawingXml)) throw new Error('The original Excel drawing has no closing root element.');
  const drawingWithoutOldAnchors = drawingXml.replace(anchorPattern, '');
  return drawingWithoutOldAnchors.replace(
    closingRoot,
    `${[...keptAnchors, ...unpositionedAnchors, ...addedAnchors].join('')}</xdr:wsDr>`,
  );
}

/**
 * Populate the existing LABLES cells in-place inside a copy of the supplied
 * workbook archive. Remove unused logo cells and arrow shapes, and limit the
 * print area to the pages needed. Original print margins and artwork remain.
 */
export function createOriginalExcelLabelWorkbook(
  templateBytes: ArrayBuffer | Uint8Array,
  labels: readonly LabelValue[],
  metadata: LabelMetadata,
) {
  const sourceBytes = templateBytes instanceof Uint8Array
    ? templateBytes
    : new Uint8Array(templateBytes);
  const files = unzipSync(sourceBytes);
  const labelsWorksheetPath = findLabelsWorksheetPath(files);
  const templateWorksheetXml = decodeXmlFile(files, labelsWorksheetPath);
  let worksheetXml = decodeXmlFile(files, labelsWorksheetPath);
  const requiredColumnCount = Math.max(1, Math.ceil(labels.length / LABEL_ROWS.length));
  if (requiredColumnCount > EXCEL_MAX_COLUMN_COUNT) {
    throw new Error(`This workbook exceeds Excel's ${EXCEL_MAX_COLUMN_COUNT.toLocaleString()}-column worksheet limit.`);
  }

  for (let columnIndex = TEMPLATE_COLUMN_COUNT; columnIndex < requiredColumnCount; columnIndex++) {
    worksheetXml = addTemplateColumnCells(worksheetXml, templateWorksheetXml, columnIndex);
  }

  const columnsToProcess = Math.max(TEMPLATE_COLUMN_COUNT, requiredColumnCount);
  for (let columnIndex = 0; columnIndex < columnsToProcess; columnIndex++) {
    const columnLetter = columnName(columnIndex);
    for (let slotIndex = 0; slotIndex < LABEL_ROWS.length; slotIndex++) {
      const labelIndex = columnIndex * LABEL_ROWS.length + slotIndex;
      const panelLabel = labels[labelIndex]?.panelLabel ?? '';
      const row = LABEL_ROWS[slotIndex];
      if (!panelLabel) {
        worksheetXml = setCellText(worksheetXml, `${columnLetter}${row - 4}`, '');
      } else {
        worksheetXml = ensureLabelLogoCell(
          worksheetXml,
          templateWorksheetXml,
          `${columnLetter}${row - 4}`,
          `A${row - 4}`,
        );
      }
      const project = panelLabel ? metadata.project : '';
      const release = panelLabel ? metadata.release : '';
      const finish = panelLabel ? metadata.finish : '';

      worksheetXml = setCellText(worksheetXml, `${columnLetter}${row - 3}`, project);
      worksheetXml = setCellText(worksheetXml, `${columnLetter}${row - 2}`, release);
      worksheetXml = setCellText(worksheetXml, `${columnLetter}${row}`, panelLabel);
      worksheetXml = setCellText(worksheetXml, `${columnLetter}${row + 2}`, finish);
    }
  }

  worksheetXml = updateWorksheetColumnRange(worksheetXml, requiredColumnCount);
  const drawingPath = findLabelsDrawingPath(files, labelsWorksheetPath);
  files[drawingPath] = encoder.encode(updateLabelArrows(decodeXmlFile(files, drawingPath), labels.length));
  files['xl/workbook.xml'] = encoder.encode(
    setLabelsPrintArea(decodeXmlFile(files, 'xl/workbook.xml'), requiredColumnCount),
  );
  files[labelsWorksheetPath] = encoder.encode(worksheetXml);
  return zipSync(files, { level: 6 });
}
