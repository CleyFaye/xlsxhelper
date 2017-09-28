/** Helper to use js-xlsx. */
/* @license
* Copyright 2017 Gabriel Paul 'Cley Faye' Risterucci
*
* Permission is hereby granted, free of charge, to any person obtaining a copy 
* of this software and associated documentation files (the "Software"), to deal
* in the Software without restriction, including without limitation the rights
* to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
* copies of the Software, and to permit persons to whom the Software is 
* furnished to do so, subject to the following conditions:
*
* The above copyright notice and this permission notice shall be included in all
* copies or substantial portions of the Software.
*
* THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR 
* IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, 
* FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE 
* AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER 
* LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
* OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
* SOFTWARE.
* 
* A lot taken from https://github.com/SheetJS/js-xlsx/blob/master/tests/write.js
* 
* js-xlsx: https://github.com/SheetJS/js-xlsx
*/
var xlsxhelper = {};
/** A single workbook that can contain many sheets. */
xlsxhelper.Workbook = class Workbook {
    constructor() {
        this.SheetNames = [];
        this.Sheets = {};
    }
    /** Add a sheet to the workbook.
    *
    * Parameters
    * ----------
    * sheet : Sheet
    *     The sheet to add to the workbook. It can be linked to many workbooks.
    *
    *
    * Notes
    * -----
    * If a sheet with the same name already exists in this workbook, it will be
    * replaced by the newly provided one.
    */
    addSheet(sheet) {
        var previousSheetIndex = this.SheetNames.indexOf(sheet.name);
        if (previousSheetIndex == -1) {
            this.SheetNames.push(sheet.name);
        }
        this.Sheets[sheet.name] = sheet;
    }
    /** Create a file from the workbook content.
    *
    * Parameters
    * ----------
    * format : string
    *     A suitable file format for js-xlsx. Common ones are 'odt', 'csv',
    *     'xlsx', 'xlsm', 'xlsb', 'fods'.
    *     This method also support 'xls' as a file format (which use the
    *     underlying biff2 format from js-xlsx).
    * fileName : string (optional)
    *     The base name for the output file. This is only used as a convenient
    *     way to manage the xls->biff2 thing, and have no effect on the actual
    *     data buffer.
    * downloader : function (optional)
    *     A runnable that will be called to to the actual downloading.
    *     It will receive three parameters: the final filename, an ArrayBuffer
    *     containing the data, and a boolean indicating if the data are binary.
    *     If not binary, the encoding will be UTF-8.
    *
    *
    * Returns
    * -------
    * The created buffer.
    */
    createFile(format, fileName, downloader) {
        var exportFormat = (format == 'xls') ? 'biff2'
                                             : format;
        var fileExtension = format;
        var options = {
            bookType: exportFormat,
            bookSST: false,
            type: 'binary'
        };
        var finalFileName = fileName ? (fileName + '.' + fileExtension)
                                     : null;
        var outputData = XLSX.write(this, options);
        if (format != 'csv') {
            var buf = new ArrayBuffer(outputData.length);
            var view = new Uint8Array(buf);
            for (var i = 0; i < outputData.length; ++i) {
                view[i] = outputData.charCodeAt(i) & 0xFF;
            }
            outputData = buf;
        }
        if (downloader) {
            downloader(finalFileName, outputData, format != 'csv');
        }
        return outputData;
    }
};

/** A workbook sheet.
*
* Parameters
* ----------
* name : string
*     Sheet name
* data : Array
*     Array of rows
* options : Object
*     Parsing options. Support the following:
*     - keepDateWithOffset : boolean
*       Indicate that dates that have a timezone offset should be kept as-is and
*       not converted into date objects.
*
*
* Notes
* -----
* Sheet.setCell() is used for every cells, so the transformation rules applies.
*/
xlsxhelper.Sheet = class Sheet {
    constructor(name, data, options) {
        this.name = name;
        this._disableRangeCalculation = true;
        for (var y = 0; y < data.length; ++y) {
            var row = data[y];
            for (var x = 0; x < row.length; ++x) {
                var cellValue = row[x];
                if (cellValue == null) {
                    continue;
                }
                this.setCell(x, y, cellValue, options);
            }
        }
        this._disableRangeCalculation = false;
        this._updateRange();
    }
    /** Create a sheet from a CSV file.
    *
    * Parameters
    * ----------
    * csvData : string
    *     The raw CSV data
    * name : string
    *     The sheet name
    * options : object
    *     Parser options. Possible properties:
    *     - stringDelimiter : string (optional)
    *         Optional delimiter for strings (default to '"')
    *     - cellDelimiter : string (optional)
    *         Delimiter for cells on a row (default to ',')
    *     - lineDelimiter : string (optional)
    *         Line delimiter (default to '\n')
    *     - keepDateWithOffset : boolean (optional)
    *         If true, date with timezone specifier are kept as text instead of
    *         being converted to date objects.
    *
    *
    * Notes
    * -----
    * The CSV format supported here is restricted to the basics and is in no way
    * a full CSV parser.
    *
    * Data interpretation rules are the same as when creating a Sheet instance.
    * The "keep date with offset" rule is required because some spreadsheet
    * formats don't handle timezones.
    */
    static fromCSV(
        csvData,
        name,
        options) {
        if (options === undefined) {
            options = {};
        }
        var stringDelimiter = options.stringDelimiter || '"';
        var cellDelimiter = options.cellDelimiter || ',';
        var lineDelimiter = options.lineDelimiter || '\n';
        var keepDateWithOffset = options.keepDateWithOffset || true;
        var csvLines = csvData.split(lineDelimiter);
        var dataRows = [];
        $(csvLines).each(function() {
            var csvLine = this.toString();
            var dataRow = [];
            let position = 0;
            while (position < csvLine.length) {
                var startingChar = csvLine[position];
                ++position;
                if (startingChar == stringDelimiter) {
                    // Parse a delimited string
                    let closingQuotePosition = csvLine.indexOf(stringDelimiter,
                                                               position);
                    do {
                        if (closingQuotePosition == -1) {
                            throw new Error('No closing quote found; line: "' 
                                + csvLine + '"');
                        }
                        if (csvLine[closingQuotePosition - 1] != '\\') {
                            break;
                        }
                        closingQuotePosition = csvLine.indexOf(
                            stringDelimiter,
                            closingQuotePosition + 1);
                    } while (true);
                    dataRow.push(csvLine.substring(
                        position, 
                        closingQuotePosition));
                    position = closingQuotePosition + 1;
                } else if (startingChar == cellDelimiter) {
                    // Parse an empty cell
                    dataRow.push(null);
                } else {
                    // Parse a non-delimited string
                    let commaPosition = csvLine.indexOf(
                        cellDelimiter, 
                        position);
                    if (commaPosition == -1) {
                        commaPosition = csvLine.length;
                    }
                    dataRow.push(csvLine.substring(
                        position - 1,
                        commaPosition));
                    position = commaPosition;
                }
                if (position < csvLine.length 
                    && csvLine[position] != cellDelimiter) {
                    throw new Error('Unknown cell delimiter: "' 
                        + csvLine[position] + '"');
                }
                ++position;
            }
            dataRows.push(dataRow);
        });
        return new xlsxhelper.Sheet(name, dataRows, {
            keepDateWithOffset: keepDateWithOffset
        });
    }
    /** Get the value from a cell. */
    getCell(col, row) {
        var cellRef = XLSX.utils.encode_cell({c: col, r: row});
        return this[cellRef];
    }
    /** Set the value in a cell.
    *
    * Value can be one of the following:
    *
    * - A Date object
    * - An ISO8601 date string
    * - a boolean
    * - a number
    * - a string
    *
    * Calling setCell() with an undefined value will empty the cell.
    */
    setCell(col, row, value, options) {
        if (options === undefined) {
            options = {};
        }
        var keepDateWithOffset = options.keepDateWithOffset || true;
        const ISO8601_DATE_REGEX = /^(\d{4})-(\d{2})-(\d{2})T(\d{2}):(\d{2}):(\d{2})(\.\d{3})?([+-]\d{2}:?\d{2}|Z)?$/;
        var match;
        var cellRef = XLSX.utils.encode_cell({c: col, r: row});
        if (value === undefined) {
            delete this[cellRef];
        } else {
            if (value instanceof Date) {
                if (value.getTimezoneOffset() != 0 && keepDateWithOffset) {
                    var cellType = 's';
                    var cellText = value.toLocaleString();
                } else {
                    var cellType = 'd';
                    var cellText = value.toISOString();
                }
            } else if ((match = ISO8601_DATE_REGEX.exec(value)) !== null) {
                if (keepDateWithOffset &&
                    (match[8] != 'Z' &&
                     match[8] != '+00:00' &&
                     match[8] != '-00:00' &&
                     match[8] != '+0000' &&
                     match[8] != '-0000')) {
                    var cellType = 's';
                    var cellText = value;
                } else {
                    var cellType = 'd';
                    var cellText = value;
                    value = new Date(value);
                }
            } else if (typeof(value) == 'boolean') {
                var cellType = 'b';
                var cellText = undefined;
            } else if (typeof(value) == 'number' 
                || parseFloat(value) == value) {
                var cellType = 'n';
                value = parseFloat(value);
                var cellText = undefined;
            } else {
                var cellType = 's';
                var cellText = value;
            }
            var cell = {
                v: value,
                t: cellType,
                w: cellText
            };
            this[cellRef] = cell;
        }
        this._updateRange();
    }
    /** Update the !ref property, containing the active range. */
    _updateRange() {
        if (this._disableRangeCalculation) {
            return;
        }
        var cMin = null;
        var cMax = null;
        var rMin = null;
        var rMax = null;
        for (let key in this) {
            if (this[key].t == undefined) {
                continue;
            }
            var cellCoordinates = XLSX.utils.decode_cell(key);
            if (cellCoordinates.c == NaN || cellCoordinates.r == NaN) {
                continue;
            }
            if (cMin === null) {
                cMin = cellCoordinates.c;
                cMax = cMin;
                rMin = cellCoordinates.r;
                rMax = rMin;
            } else {
                if (cellCoordinates.c < cMin) {
                    cMin = cellCoordinates.c;
                } else if (cellCoordinates.c > cMax) {
                    cMax = cellCoordinates.c;
                }
                if (cellCoordinates.r < rMin) {
                    rMin = cellCoordinates.r;
                } else if (cellCoordinates.r > rMax) {
                    rMax = cellCoordinates.r;
                }
            }
        }
        if (cMin === null) {
            cMin = 0;
            cMax = 0;
            rMin = 0;
            rMax = 0;
        }
        var range = {
            s: {
                c: cMin,
                r: rMin
            },
            e: {
                c: cMax,
                r: rMax
            }
        };
        this['!ref'] = XLSX.utils.encode_range(range);
    }
};
