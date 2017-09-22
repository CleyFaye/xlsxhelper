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
var poly_xlsx = poly_xlsx || {
    /** Create an empty workbook */
    Workbook: function() {
        if (!(this instanceof poly_xlsx.Workbook)) {
            return new poly_xlsx.Workbook();
        }
        this.SheetNames = [];
        this.Sheets = {};
        this.addSheet = poly_xlsx.addSheet;
        this.download = poly_xlsx.download;
    },
    /** Create a sheet from a CSV file */
    sheetFromCSV: function(data, name) {
        var rows = data.split('\n');
        var array = [];
        $(rows).each(function() {
            var rowData = this.split(',');
            var row = [];
            $(rowData).each(function() {
                var cellData = this.toString();
                var value;
                if (cellData == parseInt(cellData)) {
                    value = parseInt(cellData);
                } else if (cellData == parseFloat(cellData)) {
                    value = parseFloat(cellData);
                } else {
                    value = cellData;
                }
                row.push(value);
            });
            array.push(row);
        });
        var sheet = new poly_xlsx.Sheet(name, array);
        return sheet;
    },
    /** Create a sheet.
     *
     * The initial content will be taken from inputArray, which is expected to be an array of arrays, rows (y) first.
     * inputArray[0][0] == A1
     * inputArray[0][1] == B1
     * inputArray[1][0] == A2
     */
    Sheet: function(name, inputArray) {
        if (!(this instanceof poly_xlsx.Sheet)) {
            return new poly_xlsx.Sheet(name, inputArray);
        }
        this.name = name;
        var range = {s: {c: 10000000, r: 10000000}, e: {c: 0, r: 0}};
        for (var y = 0; y < inputArray.length; ++y) {
            var row = inputArray[y];
            for (var x = 0; x < row.length; ++x) {
                if (range.s.c > x) {
                    range.s.c = x;
                }
                if (range.s.r > y) {
                    range.s.r = y;
                }
                if (range.e.c < x) {
                    range.e.c = x;
                }
                if (range.e.r < y) {
                    range.e.r = y;
                }
                var cell = {v: row[x]};
                if (cell.v == null) {
                    continue;
                }
                var cell_ref = XLSX.utils.encode_cell({c:x,r:y});
                if (typeof cell.v === 'number') {
                    cell.t = 'n';
                } else if (typeof cell.v === 'boolean') {
                    cell.t = 'b';
                } else if (typeof cell.v instanceof Date) {
                    cell.t = 'n';
                    cell.z =XLSX.SSF._table[14]; // ?
                    cell.v = poly_xlsx.datenum(cell.v);
                } else {
                    cell.t = 's';
                }
                this[cell_ref] = cell;
            }
        }
        if (range.s.c < 10000000) {
            this['!ref'] = XLSX.utils.encode_range(range);
        }
    },
    datenum: function(v, date1904) {
        var epoch = v.getTime();
        if (date1904) {
            epoch += 1462*24*60*60*1000;
        }
        return (epoch + 2209161600000) / (24 * 60 * 60 * 1000);
    },
    addSheet: function(sheet) {
        var sheetName = sheet.name;
        var previousIndex = this.SheetNames.indexOf(sheetName);
        if (previousIndex == -1) {
            this.SheetNames.push(sheetName);
        }
        this.Sheets[sheetName] = sheet;
    },
    /** Download a workbook.
     *
     * Supported formats:
     * - xlsx
     * - xlsm (!)
     * - xlsb (!)
     * - ods
     * - biff2 (xls)
     * - fods
     * - csv
     */
    download: function(format, fileName) {
        if (format == 'xls') {
            format = 'biff2';
        }
        var opt = {bookType:format, bookSST:false, type:'binary'};
        var output = XLSX.write(this, opt);
        var ext = format;
        if (format == 'biff2') {
            ext = 'xls';
        }
        var finalFileName = fileName + '.' + ext;
        if (format == 'csv') {
            polydata.utils.downloadData(finalFileName, output, false, 'text/csv', 'UTF-8');
        } else {
            var buf = new ArrayBuffer(output.length);
            var view = new Uint8Array(buf);
            for (var i = 0; i < output.length; ++i) {
                view[i] = output.charCodeAt(i) & 0xFF;
            }
            polydata.utils.downloadData(finalFileName, buf);
        }
    },
};
