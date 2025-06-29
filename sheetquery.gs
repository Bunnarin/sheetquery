/**
 * Run new sheet query
 *
 * @param {Spreadsheet} activeSpreadsheet Specific spreadsheet to use, or will use SpreadsheetApp.getActiveSpreadsheet() if undefined\
 * @return {SheetQueryBuilder}
 */
function sheetQuery(activeSpreadsheet) {
  return new SheetQueryBuilder(activeSpreadsheet);
}

/**
 * SheetQueryBuilder class - Kind of an ORM for Google Sheets
 */
class SheetQueryBuilder {
  constructor(activeSpreadsheet) {
    this.columnNames = [];
    this.headingRow = 1;
    this._sheetHeadings = [];
    this.activeSpreadsheet = activeSpreadsheet || SpreadsheetApp.getActiveSpreadsheet();
  }
  select(columnNames) {
    this.columnNames = Array.isArray(columnNames) ? columnNames : [columnNames];
    return this;
  }
  /**
   * Name of spreadsheet to perform operations on
   *
   * @param {string} sheetName
   * @param {number} headingRow
   * @return {SheetQueryBuilder}
   */
  from(sheetName, headingRow = 1) {
    this.sheetName = sheetName;
    this.headingRow = headingRow;
    return this;
  }
  /**
   * Apply a filtering function on rows in a spreadsheet before performing an operation on them
   *
   * @param {Function} fn
   * @return {SheetQueryBuilder}
   */
  where(fn) {
    fn = (typeof fn == 'object') ? filter_fn(fn) : fn;
    this.whereFn = fn;
    return this;
  }
  /**
   * Get Sheet object that is referenced by the current query from() method
   *
   * @return {Sheet}
   */
  getSheet() {
    if (!this.sheetName) {
      throw new Error('SheetQuery: No sheet selected. Select sheet with .from(sheetName) method');
    }
    if (!this._sheet) {
      this._sheet = this.activeSpreadsheet.getSheetByName(this.sheetName);
    }
    return this._sheet;
  }
  /**
   * Get values in sheet from current query + where condition
   */
  getValues() {
    if (!this._sheetValues) {
      const zh = this.headingRow - 1;
      const sheet = this.getSheet();
      if (!sheet) 
        return [];
      const rowValues = [];
      const sheetValues = sheet.getDataRange().getValues();
      const numCols = sheetValues[0] ? sheetValues[0].length : 0;
      const numRows = sheetValues.length;
      const headings = (this._sheetHeadings = sheetValues[zh] || []);
      for (let r = 0; r < numRows; r++) {
        const obj = { __meta: { row: r + 1, cols: numCols } };
        for (let c = 0; c < numCols; c++) 
          obj[headings[c]] = sheetValues[r][c]; // @ts-ignore
        rowValues.push(obj);
      }
      this._sheetValues = rowValues;
    }
    return this._sheetValues;
  }
  /**
   * Return matching rows from sheet query excluding the header row
   *
   * @return {RowObject[]}
   */
  getRows() {
    const sheetValues = this.getValues().slice(1, -1);
    return this.whereFn ? sheetValues.filter(this.whereFn) : sheetValues;
  }
  /**
   * Return matching rows from sheet query including the header row
   *
   * @return {RowObject[]}
   */
  getTable() {
    const sheetValues = this.getValues();
    return this.whereFn ? sheetValues.filter(this.whereFn) : sheetValues;
  }
  /**
   * Get array of headings in current sheet from()
   *
   * @return {string[]}
   */
  getHeadings() {
    if (!this._sheetHeadings || !this._sheetHeadings.length) {
      const zh = this.headingRow - 1;
      const sheet = this.getSheet();
      const numCols = sheet.getLastColumn();
      this._sheetHeadings = sheet.getSheetValues(1, 1, this.headingRow, numCols)[zh] || [];
      this._sheetHeadings = this._sheetHeadings
        .map((s) => (typeof s === 'string' ? s.trim() : ''))
        .filter(Boolean);
    }
    return this._sheetHeadings || [];
  }
  /**
   * Get all cells from a query + where condition
   * @returns {any[]}
   */
  getCells() {
    const rows = this.getTable();
    const cellArray = [];
    rows.forEach((row) => {
      cellArray.push(this._sheet.getRange(row.__meta.row, 1, 1, row.__meta.cols));
    });
    return cellArray;
  }
  /**
   * Get cells in sheet from current query + where condition and from specific header
   * @param {string} key name of the column
   * @param {Array<string>} [keys] optionnal names of columns use to select more columns than one
   * @returns {any[]} all the colum cells from the query's rows
   */
  getCellsWithHeadings(key, headings) {
    let rows = this.getTable();
    let indexColumn = 1;
    const arrayCells = [];
    for (const elem of this._sheetHeadings) {
      if (elem == key) break;
      indexColumn++;
    }
    rows.forEach((row) => {
      arrayCells.push(this._sheet.getRange(row.__meta.row, indexColumn));
    });
    //If we got more thant one param
    headings.forEach((col) => {
      let indexColumn = 1;
      for (const elem of this._sheetHeadings) {
        if (elem == col) break;
        indexColumn++;
      }
      rows.forEach((row) => {
        arrayCells.push(this._sheet.getRange(row.__meta.row, indexColumn));
      });
    });
    return arrayCells;
  }
  /**
   * Insert new rows into the spreadsheet
   * Arrays of objects like { Heading: Value }
   *
   * @param {DictObject[]} newRows - Array of row objects to insert
   * @return {SheetQueryBuilder}
   */
  insertRows(newRows) {
    const sheet = this.getSheet();
    const headings = this.getHeadings();
    newRows.forEach((row) => {
      if (!row) {
        return;
      }
      const rowValues = headings.map((heading) => {
        const val = row[heading];
        return val === undefined || val === null || val === false ? '' : val;
      });
      // appendRow() will throw if array is empty, so we check to prevent that
      if (rowValues && rowValues.length !== 0) {
        sheet.appendRow(rowValues);
      }
    });
    return this;
  }
  /**
   * Delete matched rows from spreadsheet
   *
   * @return {SheetQueryBuilder}
   */
  deleteRows() {
    const rows = this.getTable();
    let i = 0;
    rows.forEach((row) => {
      const deleteRowRange = this._sheet.getRange(row.__meta.row - i, 1, 1, row.__meta.cols);
      deleteRowRange.deleteCells(SpreadsheetApp.Dimension.ROWS);
      i++;
    });
    this.clearCache();
    return this;
  }
  /**
   * Update matched rows in spreadsheet with provided function
   *
   * @param {UpdateFn} updateFn
   * @return {SheetQueryBuilder}
   */
  updateRows(updateFn) {
    updateFn = (typeof updateFn == 'object') ? update_fn(updateFn) : updateFn;
    const rows = this.getTable();
    for (let i = 0; i < rows.length; i++) {
      this.updateRow(rows[i], updateFn);
    }
    this.clearCache();
    return this;
  }
  /**
   * Update single row
   */
  updateRow(row, updateFn) {
    updateFn = (typeof updateFn == 'object') ? update_fn(updateFn) : updateFn;
    const updatedRow = updateFn(row) || row;
    const rowMeta = updatedRow.__meta;
    const headings = this.getHeadings();
    delete updatedRow.__meta;
    // Put new array data in order of headings in sheet
    const arrayValues = headings.map((heading) => {
      const val = updatedRow[heading];
      return val === undefined || val === null || val === false ? '' : val;
    });
    const maxCols = Math.max(rowMeta.cols, arrayValues.length);
    const updateRowRange = this.getSheet().getRange(rowMeta.row, 1, 1, maxCols);
    const rangeData = updateRowRange.getValues()[0] || [];
    // Map over old data in same index order to update it and ensure array length always matches
    const newValues = rangeData.map((value, index) => {
      const val = arrayValues[index];
      return val === undefined || val === null || val === false ? '' : val;
    });
    updateRowRange.setValues([newValues]);
    return this;
  }
  /**
   * Clear cached values, headings, and flush all operations to sheet
   *
   * @return {SheetQueryBuilder}
   */
  clearCache() {
    this._sheetValues = null;
    this._sheetHeadings = [];
    SpreadsheetApp.flush();
    return this;
  }
}

// util
function update_fn(updateHash) {
  let functionBody = '';

  // Iterate over the keys in the updateHash
  for (const key in updateHash) {
    // Ensure it's an own property, not inherited
    if (Object.prototype.hasOwnProperty.call(updateHash, key)) {
      const value = updateHash[key];

      // Dynamically format the value correctly for inclusion in the function string.
      // Strings need to be quoted, numbers/booleans/null/undefined can be inserted directly.
      let formattedValue;
      if (typeof value === 'string') {
        // Escape quotes within the string to prevent breaking the generated code
        formattedValue = JSON.stringify(value);
      } else if (typeof value === 'object' && value !== null) {
        // For objects/arrays, JSON.stringify is usually the safest way to embed them,
        // but be aware of complexity for deeply nested objects.
        formattedValue = JSON.stringify(value);
      } else {
        // For numbers, booleans, null, undefined
        formattedValue = String(value);
      }

      // Add the assignment statement to the function body
      functionBody += `  row["${key}"] = ${formattedValue};\n`;
    }
  }

  // Construct the anonymous function using new Function().
  return new Function('row', functionBody);
}

function filter_fn(filterHash) {
  const conditions = [];
  for (const key in filterHash) {
    if (Object.prototype.hasOwnProperty.call(filterHash, key)) {
      const value = filterHash[key];
      // Escape string values for proper inclusion in the function string
      const escapedValue = typeof value === 'string' ? `'${value.replace(/'/g, "\\'")}'` : value;
      conditions.push(`row["${key}"] === ${escapedValue}`);
    }
  }

  const functionBody = `return ${conditions.join(' && ')};`;
  return new Function('row', functionBody);
}