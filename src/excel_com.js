function test() {
	// npm install -g win32ole
	var win32ole = require('win32ole');

	// var xl = new ActiveXObject('Excel.Application'); // You may write it as: 
	var xl = win32ole.client.Dispatch('Excel.Application');
	xl.Visible = true;

	var book = xl.Workbooks.Add();

	var sheet = book.Worksheets(1);
	sheet.Name = 'sheetnameA utf8';
	sheet.Cells(1, 2).Value = 'test utf8';

	var rng = sheet.Range(sheet.Cells(2, 2), sheet.Cells(4, 4));
	rng.RowHeight = 5.18;
	rng.ColumnWidth = 0.58;
	rng.Interior.ColorIndex = 6; // Yellow 

	//var result = book.SaveAs('testfileutf8.xls');
	console.log(result);

	xl.ScreenUpdating = true;
	xl.Workbooks.Close();
	xl.Quit();
}

var folder="..\\..\\..\\vision\\Bin\\xls";
//C:\Code\vision\Bin\xls";

var fs = require('fs');
var win32ole = require('win32ole');
var xl = win32ole.client.Dispatch('Excel.Application');
xl.Visible = true;

fs.readdir(folder, function(err, files) {
	if (err) throw err;
	console.log(files);
	for (var f=0; f<files.length; f++) {
		var path = folder+"\\"+files[f];
		//fs.stat(), fs.lstat() and fs.fstat() and their synchronous counterparts are of this type.
		//stats.isDirectory()
		if (path.indexOf(".xls")!=path.length-4)
			continue;
		
		var book = xl.Workbooks.Open(path);
		var nb_sheet = book.Worksheets.length;
		
		console.log(path);
		console.log(nb_sheet);
	
	}
});