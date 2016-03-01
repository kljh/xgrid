// npm install win32ole
var win32ole = require('win32ole');

// var xl = new ActiveXObject('Excel.Application'); // You may write it as: 
var xl = win32ole.client.Dispatch('Excel.Application');
xl.Visible = true;

var book = xl.Workbooks.Add();

var sheet = book.Worksheets(1);
sheet.Name = 'sheetnameA utf8';
sheet.Cells(1, 2).Value = 'test utf8';

var rg = sheet.Range(sheet.Cells(2, 2), sheet.Cells(4, 4));
rg.RowHeight = 5.18;
rg.ColumnWidth = 0.58;
rg.Interior.ColorIndex = 6; // Yellow 

//var result = book.SaveAs('testfileutf8.xls');
console.log(result);

xl.ScreenUpdating = true;
xl.Workbooks.Close();
xl.Quit();
