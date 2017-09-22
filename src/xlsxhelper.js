/** Helper to use js-xlsx.
*
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
        if (previousIndex == -1) {
            this.SheetNames.push(sheet.name);
        }
        this.Sheets[sheetName] = sheet;
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
*/
xlsxhelper.Sheet = class Sheet {
    constructor(name, data) {
        this.name = name;
        this._disableRangeCalculation = true;
        for (var y = 0; y < inputArray.length; ++y) {
            var row = inputArray[y];
            for (var x = 0; x < row.length; ++x) {
                var cellValue = row[x];
                if (cellValue == null) {
                    continue;
                }
                this.setCell(x, y, cellValue);
            }
        }
        this._disableRangeCalculation = false;
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
    setCell(col, row, value) {
        const ISO8601_DATE_REGEX = /(\d{4})-(\d{2})-(\d{2})T(\d{2}):(\d{2}):(\d{2})(\.\d{3})?([+-]\d{2}:?\d{2}|Z)?/;
        var match;
        var cellRef = XLSX.utils.encode_cell({c: col, r: row});
        if (value === undefined) {
            delete this[cellRef];
        } else {
            if (value instanceof Date) {
                var cellType = 'd';
            } else if ((match = ISO8601_DATE_REGEX.exec(value)) !== null) {
                var cellType = 'd';
                value = new Date(value);
            } else if (typeof(value) == 'boolean') {
                var cellType = 'b';
            } else if (typeof(value) == 'number') {
                var cellType = 'n';
            } else {
                var cellType = 's';
            }
            var cell = {
                v: value,
                t: cellType
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
        for (key in this) {
            if (this[key].t == undefined) {
                return;
            }
            var cellCoordinates = XLSX.utils.decode_cell(key);
            if (cellCoordinates.c == NaN || cellCoordinates.r == Nan) {
                return;
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
