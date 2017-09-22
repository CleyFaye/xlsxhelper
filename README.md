xlsxhelper - Utility code for js-xlsx
=====================================


Description
-----------
These classes makes it easier to create workbooks from scratch.
The Workbook class is used to gather spreadsheets, while the Sheet class contain the actual data.


Use case
--------
The primary use case here is to format data, coming from various source and processed by JavaScript, into CSV/ODT/XLSX files.


Usage
-----
Instances of Workbook can be created without any parameter.
Instances of Sheet are created by providing their initial content and name as an argument.
Adding Sheet instance to a Workbook is done by calling Workbook.addSheet().


Example
-------
``` JavaScript
var data = [
    [1, 2, 3],
    [4, 5, 6]];
var wb = new xlsxhelper.Workbook();
var sheet = new xlsxhelper.Sheet('SheetName', data);
wb.addSheet(sheet);
wb.download('odt', 'basefilename');
```


Installation
------------
Include the 'dist/xlsxhelper.js' file somewhere before you use it.
