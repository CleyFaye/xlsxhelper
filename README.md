xlsx/pptx helper - Utility code for spreadsheets and presentations
==================================================================


xlxshelper
----------


### Description
These classes makes it easier to create workbooks from scratch.
The Workbook class is used to gather spreadsheets, while the Sheet class contain the actual data.


### Use case
The primary use case here is to format data, coming from various source and processed by JavaScript, into CSV/ODT/XLSX files.


### Usage
Instances of Workbook can be created without any parameter.
Instances of Sheet are created by providing their initial content and name as an argument.
Adding Sheet instance to a Workbook is done by calling Workbook.addSheet().


### Example
``` JavaScript
var data = [
    [1, 2, 3],
    [4, 5, 6]];
var wb = new xlsxhelper.Workbook();
var sheet = new xlsxhelper.Sheet('SheetName', data);
wb.addSheet(sheet);
wb.createFile('odt', 'basefilename', downloader);
```


pptxhelper
----------


### Description
Classes to create PPTX presentations with images.
This helper main goal is to semi-transparently allow the download of images for integration into the output.


### Use case
Dynamic generation of a presentation slide depending on user-specific actions.


### Usage
Create an instance of Presentation to start.
Add slides using Presentation.addSlide().
Add text and images to a slide, or pre-populate it by using a short json description.
To download the presentation, use Presentation.download().


### Example
``` JavaScript
var pptx = new pptxhelper.Presentation();
var slide = pptx.addSlide();
slide.addText(0, 0, "Hello");
slide.addImage(5, 2, 1, 1, '/img/carret.jpg');
pptx.download('sample');
```


Installation
------------
Include the 'dist/xlsxhelper.js' file somewhere before you use it.
