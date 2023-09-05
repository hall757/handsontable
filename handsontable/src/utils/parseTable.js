import { isDefined, isEmpty } from './../helpers/mixed';

const ESCAPED_HTML_CHARS = {
  '&nbsp;': '\x20',
  '&amp;': '&',
  '&lt;': '<',
  '&gt;': '>',
};
const regEscapedChars = new RegExp(Object.keys(ESCAPED_HTML_CHARS).map(key => `(${key})`).join('|'), 'gi');

/**
 * Verifies if node is an HTMLTable element.
 *
 * @param {Node} element Node to verify if it's an HTMLTable.
 * @returns {boolean}
 */
function isHTMLTable(element) {
  return (element && element.nodeName || '') === 'TABLE';
}

/**
 * Converts Handsontable into HTMLTableElement.
 *
 * @param {Core} instance The Handsontable instance.
 * @returns {string} OuterHTML of the HTMLTableElement.
 */
export function instanceToHTML(instance) {
  const includeRowHeaders = instance.hasRowHeaders();
  const includeColumnHeaders = instance.hasColHeaders();
  const startRow = includeColumnHeaders ? -1 : 0;
  const startColumn = includeRowHeaders ? -1 : 0;
  const endRow = instance.countRows() - 1;
  const endColumn = instance.countCols() - 1;

  return getTableByCoords(instance, startRow, startColumn, endRow, endColumn, includeRowHeaders, includeColumnHeaders);
}

export function getConfig(instance, metaInfoAndModifiers) {
  const { withCells, withColumnHeaders, withRowHeaders, onlyFirstLevel, columnHeadersCount, rowsLimit, columnsLimit,
    ignoredRows, ignoredColumns } = metaInfoAndModifiers;
  const selection = instance.getSelectedLast();
  const [startRow, startColumn, endRow, endColumn] = [
    Math.min(selection[0], selection[2]),
    Math.min(selection[1], selection[3]),
    Math.max(selection[0], selection[2]),
    Math.max(selection[1], selection[3])
  ];
  const config = {
    ignoredRows,
    ignoredColumns,
    startColumn: Math.max(startColumn, 0),
    endColumn: Math.max(endColumn, 0),
  };

  if (withColumnHeaders === true && columnHeadersCount > 0) {
    if (onlyFirstLevel === true) {
      config.startColumnHeader = -1;

    } else {
      config.startColumnHeader = -1 * columnHeadersCount;
    }
  }

  if (withRowHeaders === true && startRow === -1) {
    config.startRowHeader = startRow;
  }

  if (withCells === false) {
    return config;
  }

  if (endRow >= 0) {
    config.startRow = Math.max(startRow, 0);

    if (rowsLimit !== Infinity) {
      config.endRow = endRow;

    } else {
      config.endRow = Math.min(endRow, config.startRow + rowsLimit);
    }
  }

  return config;
}

/**
 * Converts Handsontable's selection into HTMLTableElement.
 *
 * @param {Core} instance The Handsontable instance.
 * @param {object} metaInfoAndModifiers
 * @returns {string} OuterHTML of the HTMLTableElement.
 */
export function selectionToHTML(instance, metaInfoAndModifiers) {
  return getTableByCoords(instance, getConfig(instance, metaInfoAndModifiers));
}

export function selectionToData(instance, metaInfoAndModifiers) {
  return [
    ...getHeadersDataByCoords(instance, metaInfoAndModifiers),
    ...getBodyDataByCoords(instance, metaInfoAndModifiers),
  ];
}

/**
 * Encode text to HTML.
 *
 * @param {string} text Text to prepare.
 * @returns {string}
 */
function encodeHTMLEntities(text) {
  return `${text}`.replace('<', '&lt;')
    .replace('>', '&gt;')
    .replace(/(<br(\s*|\/)>(\r\n|\n)?|\r\n|\n)/g, '<br>\r\n')
    .replace(/\x20{2,}/gi, (substring) => {
      // The way how Excel serializes data with at least two spaces.
      return `<span style="mso-spacerun: yes">${'&nbsp;'.repeat(substring.length - 1)} </span>`;
    })
    .replace(/\t/gi, '&#9;');
}

function getHeadersHTMLByCoords(instance, config) {
  const { startColumn, endColumn, startColumnHeader, columnsLimit } = config;

  if (isDefined(startColumnHeader)) {
    const headers = [];

    for (let columnHeaderLevel = startColumnHeader; columnHeaderLevel < 0; columnHeaderLevel += 1) {
      const tr = ['<tr>'];

      if (config.startRowHeader === -1) {
        tr.push(`<th>${encodeHTMLEntities(instance.getColHeader(-1, columnHeaderLevel))}</th>`);
      }

      const lastCopiedColumnIndex = Math.min(startColumn + columnsLimit - 1, endColumn);

      for (let columnIndex = startColumn; columnIndex <= lastCopiedColumnIndex; columnIndex += 1) {
        const header = instance.getCell(columnHeaderLevel, columnIndex);
        const colspan = header?.getAttribute('colspan');
        let colspanAttribute = '';

        if (colspan) {
          const parsedColspan = parseInt(colspan, 10);

          colspanAttribute = ` colspan=${parsedColspan}`;
          columnIndex += parsedColspan - 1;
        }

        tr.push(`<th${colspanAttribute}>${encodeHTMLEntities(instance.getColHeader(
          columnIndex, columnHeaderLevel))}</th>`);
      }

      tr.push('</tr>');
      headers.push(...tr);
    }

    if (headers.length > 0) {
      return ['<thead>', ...headers, '</thead>'];
    }
  }

  return [];
}

function getHeadersDataByCoords(instance, metaInfoAndModifiers) {
  const { startRowHeader, startColumn, endColumn, startColumnHeader, columnsLimit } = getConfig(instance, metaInfoAndModifiers);
  const headers = [];

  if (isDefined(startColumnHeader)) {
    for (let columnHeaderLevel = startColumnHeader; columnHeaderLevel < 0; columnHeaderLevel += 1) {
      const tr = [];

      if (startRowHeader === -1) {
        tr.push(instance.getColHeader(-1, columnHeaderLevel));
      }

      const lastCopiedColumnIndex = Math.min(startColumn + columnsLimit - 1, endColumn);

      for (let columnIndex = startColumn; columnIndex <= lastCopiedColumnIndex; columnIndex += 1) {
        tr.push(instance.getColHeader(columnIndex, columnHeaderLevel));
      }

      headers.push(tr);
    }
  }

  return headers;
}

function getBodyHTMLByCoords(instance, config) {
  const { startRow, startColumn, endRow, endColumn, rowsLimit, columnsLimit } = config;
  const ignoredRows = new Set(config.ignoredRows);
  const ignoredColumns = new Set(config.ignoredColumns);

  if (isDefined(startRow)) {
    const cells = [];
    const data = instance.getData(startRow, startColumn, endRow, endColumn);
    const countRows = Math.min(endRow - startRow + 1, rowsLimit);
    const countColumns = Math.min(endColumn - startColumn + 1, columnsLimit);

    for (let rowIndex = 0; rowIndex < countRows && ignoredRows.has(rowIndex) === false; rowIndex += 1) {
      const tr = ['<tr>'];

      if (config.startRowHeader === -1) {
        tr.push(`<th>${encodeHTMLEntities(instance.getRowHeader(startRow + rowIndex))}</th>`);
      }

      for (let columnIndex = 0; columnIndex < countColumns && ignoredColumns.has(columnIndex) === false;
        columnIndex += 1) {
        const cellValue = data[rowIndex][columnIndex];
        const cellValueParsed = isEmpty(cellValue) ? '' : encodeHTMLEntities(cellValue);
        const { hidden, rowspan, colspan } =
          instance.getCellMeta(rowIndex + startRow, columnIndex + startColumn);

        if (!hidden) {
          const attrs = [];

          if (rowspan) {
            const recalculatedRowSpan = Math.min(rowspan, countRows - rowIndex);

            if (recalculatedRowSpan > 1) {
              attrs.push(` rowspan="${recalculatedRowSpan}"`);
            }
          }

          if (colspan) {
            const recalculatedColumnSpan = Math.min(colspan, countColumns - columnIndex);

            if (recalculatedColumnSpan > 1) {
              attrs.push(` colspan="${recalculatedColumnSpan}"`);
            }
          }

          tr.push(`<td${attrs.join('')}>${cellValueParsed}</td>`);
        }
      }

      tr.push('</tr>');
      cells.push(...tr);
    }

    if (cells.length > 0) {
      return ['<tbody>', ...cells, '</tbody>'];
    }
  }

  return [];
}

function getBodyDataByCoords(instance, config) {
  const { startRow, startColumn, endRow, endColumn, rowsLimit, columnsLimit, startRowHeader } = getConfig(instance, config);
  const ignoredRows = new Set(config.ignoredRows);
  const ignoredColumns = new Set(config.ignoredColumns);
  const cells = [];

  if (isDefined(startRow)) {
    const data = instance.getData(startRow, startColumn, endRow, endColumn);
    const countRows = Math.min(endRow - startRow + 1, rowsLimit);
    const countColumns = Math.min(endColumn - startColumn + 1, columnsLimit);

    for (let rowIndex = 0; rowIndex < countRows && ignoredRows.has(rowIndex) === false; rowIndex += 1) {
      const tr = [];

      if (startRowHeader === -1) {
        tr.push(instance.getRowHeader(startRow + rowIndex));
      }

      for (let columnIndex = 0; columnIndex < countColumns && ignoredColumns.has(columnIndex) === false;
        columnIndex += 1) {
        const cellValue = data[rowIndex][columnIndex];
        const cellValueParsed = isEmpty(cellValue) ? '' : cellValue;

        tr.push(cellValueParsed);
      }

      cells.push(tr);
    }
  }

  return cells;
}

/**
 * Creates OuterHTML of the HTMLTableElement from instance of Handsontable basing on handled coordinates.
 *
 * @private
 * @param {Core} instance The Handsontable instance.
 * @param {object} config
 * @returns {string}
 */
function getTableByCoords(instance, config) {
  return ['<table>', ...getHeadersHTMLByCoords(instance, config), ...getBodyHTMLByCoords(instance, config),
    '</table>'].join('');
}

/**
 * Converts 2D array into HTMLTableElement.
 *
 * @param {Array} input Input array which will be converted to HTMLTable.
 * @returns {string} OuterHTML of the HTMLTableElement.
 */
// eslint-disable-next-line no-restricted-globals
export function _dataToHTML(input) {
  const inputLen = input.length;
  const result = ['<table>'];

  for (let row = 0; row < inputLen; row += 1) {
    const rowData = input[row];
    const columnsLen = rowData.length;
    const columnsResult = [];

    if (row === 0) {
      result.push('<tbody>');
    }

    for (let column = 0; column < columnsLen; column += 1) {
      const cellData = rowData[column];
      const encodeHTMLEntitiesdCellData = isEmpty(cellData) ?
        '' :
        cellData.toString()
          .replace(/</g, '&lt;')
          .replace(/>/g, '&gt;')
          .replace(/(<br(\s*|\/)>(\r\n|\n)?|\r\n|\n)/g, '<br>\r\n')
          .replace(/\x20{2,}/gi, (substring) => {
            // The way how Excel serializes data with at least two spaces.
            return `<span style="mso-spacerun: yes">${'&nbsp;'.repeat(substring.length - 1)} </span>`;
          })
          .replace(/\t/gi, '&#9;');

      columnsResult.push(`<td>${encodeHTMLEntitiesdCellData}</td>`);
    }

    result.push('<tr>', ...columnsResult, '</tr>');

    if (row + 1 === inputLen) {
      result.push('</tbody>');
    }
  }

  result.push('</table>');

  return result.join('');
}

/**
 * Converts HTMLTable or string into Handsontable configuration object.
 *
 * @param {Element|string} element Node element which should contain `<table>...</table>`.
 * @param {Document} [rootDocument] The document window owner.
 * @returns {object} Return configuration object. Contains keys as DefaultSettings.
 */
// eslint-disable-next-line no-restricted-globals
export function htmlToGridSettings(element, rootDocument = document) {
  const settingsObj = {};
  const fragment = rootDocument.createDocumentFragment();
  const tempElem = rootDocument.createElement('div');

  fragment.appendChild(tempElem);

  let checkElement = element;

  if (typeof checkElement === 'string') {
    const escapedAdjacentHTML = checkElement.replace(/<td\b[^>]*?>([\s\S]*?)<\/\s*td>/g, (cellFragment) => {
      const openingTag = cellFragment.match(/<td\b[^>]*?>/g)[0];
      const cellValue = cellFragment
        .substring(openingTag.length, cellFragment.lastIndexOf('<')).replace(/(<(?!br)([^>]+)>)/gi, '');
      const closingTag = '</td>';

      return `${openingTag}${cellValue}${closingTag}`;
    });

    tempElem.insertAdjacentHTML('afterbegin', `${escapedAdjacentHTML}`);
    checkElement = tempElem.querySelector('table');
  }

  if (!checkElement || !isHTMLTable(checkElement)) {
    return;
  }

  const generator = tempElem.querySelector('meta[name$="enerator"]');
  const hasRowHeaders = checkElement.querySelector('tbody th') !== null;
  const trElement = checkElement.querySelector('tr');
  const countCols = !trElement ? 0 : Array.from(trElement.cells)
    .reduce((cols, cell) => cols + cell.colSpan, 0) - (hasRowHeaders ? 1 : 0);
  const fixedRowsBottom = checkElement.tFoot && Array.from(checkElement.tFoot.rows) || [];
  const fixedRowsTop = [];
  let hasColHeaders = false;
  let thRowsLen = 0;
  let countRows = 0;

  if (checkElement.tHead) {
    const thRows = Array.from(checkElement.tHead.rows).filter((tr) => {
      const isDataRow = tr.querySelector('td') !== null;

      if (isDataRow) {
        fixedRowsTop.push(tr);
      }

      return !isDataRow;
    });

    thRowsLen = thRows.length;
    hasColHeaders = thRowsLen > 0;

    if (thRowsLen > 1) {
      settingsObj.nestedHeaders = Array.from(thRows).reduce((rows, row) => {
        const headersRow = Array.from(row.cells).reduce((headers, header, currentIndex) => {
          if (hasRowHeaders && currentIndex === 0) {
            return headers;
          }

          const {
            colSpan: colspan,
            innerHTML,
          } = header;
          const nextHeader = colspan > 1 ? { label: innerHTML, colspan } : innerHTML;

          headers.push(nextHeader);

          return headers;
        }, []);

        rows.push(headersRow);

        return rows;
      }, []);

    } else if (hasColHeaders) {
      settingsObj.colHeaders = Array.from(thRows[0].children).reduce((headers, header, index) => {
        if (hasRowHeaders && index === 0) {
          return headers;
        }

        headers.push(header.innerHTML);

        return headers;
      }, []);
    }
  }

  if (fixedRowsTop.length) {
    settingsObj.fixedRowsTop = fixedRowsTop.length;
  }
  if (fixedRowsBottom.length) {
    settingsObj.fixedRowsBottom = fixedRowsBottom.length;
  }

  const dataRows = [
    ...fixedRowsTop,
    ...Array.from(checkElement.tBodies).reduce((sections, section) => {
      sections.push(...Array.from(section.rows));

      return sections;
    }, []),
    ...fixedRowsBottom];

  countRows = dataRows.length;

  const dataArr = new Array(countRows);

  for (let r = 0; r < countRows; r++) {
    dataArr[r] = new Array(countCols);
  }

  const mergeCells = [];
  const rowHeaders = [];

  for (let row = 0; row < countRows; row++) {
    const tr = dataRows[row];
    const cells = Array.from(tr.cells);
    const cellsLen = cells.length;

    for (let cellId = 0; cellId < cellsLen; cellId++) {
      const cell = cells[cellId];
      const {
        nodeName,
        innerHTML,
        rowSpan: rowspan,
        colSpan: colspan,
      } = cell;
      const col = dataArr[row].findIndex(value => value === void 0);

      if (nodeName === 'TD') {
        if (rowspan > 1 || colspan > 1) {
          for (let rstart = row; rstart < row + rowspan; rstart++) {
            if (rstart < countRows) {
              for (let cstart = col; cstart < col + colspan; cstart++) {
                dataArr[rstart][cstart] = null;
              }
            }
          }

          const styleAttr = cell.getAttribute('style');
          const ignoreMerge = styleAttr && styleAttr.includes('mso-ignore:colspan');

          if (!ignoreMerge) {
            mergeCells.push({ col, row, rowspan, colspan });
          }
        }

        let cellValue = '';

        if (generator && /excel/gi.test(generator.content)) {
          cellValue = innerHTML.replace(/[\r\n][\x20]{0,2}/g, '\x20')
            .replace(/<br(\s*|\/)>[\r\n]?[\x20]{0,3}/gim, '\r\n');

        } else {
          cellValue = innerHTML.replace(/<br(\s*|\/)>[\r\n]?/gim, '\r\n');
        }

        dataArr[row][col] = cellValue.replace(regEscapedChars, match => ESCAPED_HTML_CHARS[match]);

      } else {
        rowHeaders.push(innerHTML);
      }
    }
  }

  if (mergeCells.length) {
    settingsObj.mergeCells = mergeCells;
  }
  if (rowHeaders.length) {
    settingsObj.rowHeaders = rowHeaders;
  }

  if (dataArr.length) {
    settingsObj.data = dataArr;
  }

  return settingsObj;
}
