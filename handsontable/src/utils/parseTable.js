import { isEmpty } from './../helpers/mixed';

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

/**
 * Converts Handsontable's selection into HTMLTableElement.
 *
 * @param {Core} instance The Handsontable instance.
 * @param {object} metaInfoAndModifiers
 * @returns {string} OuterHTML of the HTMLTableElement.
 */
export function selectionToHTML(instance, metaInfoAndModifiers) {
  const { withCells, withColumnHeaders, withRowHeaders, onlyFirstLevel } = metaInfoAndModifiers;

  const selection = instance.getSelectedLast();
  const endColumn = Math.max(selection[1], selection[3]);
  let [startRow, startColumn, endRow] = [
    Math.min(selection[0], selection[2]),
    Math.min(selection[1], selection[3]),
    Math.max(selection[0], selection[2]),
  ];

  if (withCells === false) {
    startRow = Math.min(startRow, -1);
    endRow = Math.min(endRow, -1);

    return getTableByCoords(instance, startRow, startColumn, endRow, endColumn, withRowHeaders, true);
  }

  if (withColumnHeaders === false) {
    startRow = Math.max(startRow, 0);

  } else if (onlyFirstLevel === true) {
    startRow = Math.max(startRow, -1);
  }

  if (withRowHeaders === false) {
    startColumn = Math.max(startColumn, 0);

  } else if (onlyFirstLevel === true) {
    startColumn = Math.max(startColumn, -1);
  }

  return getTableByCoords(instance, startRow, startColumn, endRow, endColumn, withRowHeaders, withColumnHeaders);
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

/**
 * Creates OuterHTML of the HTMLTableElement from instance of Handsontable basing on handled coordinates.
 *
 * @private
 * @param {Core} instance The Handsontable instance.
 * @param {number} startRow Starting row for creating a HTML table.
 * @param {number} startColumn Starting row for creating a HTML table.
 * @param {number} endRow Ending row for creating a HTML table.
 * @param {number} endColumn Ending row for creating a HTML table.
 * @param {boolean} includeRowHeaders Defines whether row headers should be included in created HTML table.
 * @param {boolean} includeColumnHeaders Defines whether row headers should be included in created HTML table.
 * @returns {string}
 */
function getTableByCoords(instance, startRow, startColumn, endRow, endColumn, includeRowHeaders, includeColumnHeaders) {
  const data = instance.getData(startRow, startColumn, endRow, endColumn);
  const countRows = endRow - startRow + 1;
  const countCols = endColumn - startColumn + 1;
  const TABLE = ['<table>', '</table>'];
  const THEAD = includeColumnHeaders ? ['<thead>', '</thead>'] : [];
  const TBODY = ['<tbody>', '</tbody>'];

  if (data.length === 0) {
    return '<table></table>';
  }

  for (let row = 0; row < countRows; row += 1) {
    const isColumnHeadersRow = includeColumnHeaders && startRow + row < 0;
    const CELLS = [];

    for (let column = 0; column < countCols; column += 1) {
      const isRowHeadersColumn = !isColumnHeadersRow && includeRowHeaders && startColumn + column < 0;
      let cell = '';

      if (isColumnHeadersRow) {
        cell = `<th>${encodeHTMLEntities(instance.getColHeader(startColumn + column, startRow + row))}</th>`;

      } else if (isRowHeadersColumn) {
        cell = `<th>${encodeHTMLEntities(instance.getRowHeader(row))}</th>`;

      } else {
        const cellData = data[row][column];
        const { hidden, rowspan, colspan } =
          instance.getCellMeta(startRow + row, startColumn + column);

        if (!hidden) {
          const attrs = [];

          if (rowspan) {
            const recalculatedRowSpan = Math.min(rowspan, countRows - row);

            if (recalculatedRowSpan > 1) {
              attrs.push(` rowspan="${recalculatedRowSpan}"`);
            }
          }
          if (colspan) {
            const recalculatedColumnSpan = Math.min(colspan, countCols - column);

            if (recalculatedColumnSpan > 1) {
              attrs.push(` colspan="${recalculatedColumnSpan}"`);
            }
          }
          if (isEmpty(cellData)) {
            cell = `<td${attrs.join('')}></td>`;
          } else {
            const value = cellData.toString();

            cell = `<td${attrs.join('')}>${encodeHTMLEntities(value)}</td>`;
          }
        }
      }

      CELLS.push(cell);
    }

    const TR = ['<tr>', ...CELLS, '</tr>'].join('');

    if (isColumnHeadersRow) {
      THEAD.splice(-1, 0, TR);
    } else {
      TBODY.splice(-1, 0, TR);
    }
  }

  TABLE.splice(1, 0, THEAD.join(''), TBODY.join(''));

  return TABLE.join('');
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
