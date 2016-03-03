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

var path = require('path');
var fs = require('fs');
var win32ole = require('win32ole');
var xl = win32ole.client.Dispatch('Excel.Application');
xl.Visible = true;

fs.readdir(folder, function(err, files) {
	if (err) throw err;
	console.log(files);
	for (var f=0; f<files.length; f++) {
		var filepath = path.resolve(folder+"\\"+files[f]);
		//fs.stat(), fs.lstat() and fs.fstat() and their synchronous counterparts are of this type.
		//stats.isDirectory()
		if (filepath.indexOf(".xls")!=filepath.length-4)
			continue;
		
		console.log("filepath: "+filepath);
		var book = xl.Workbooks.Open(filepath);
		var nb_sheet = book.Worksheets.Count;
		console.log("nb_sheet: "+nb_sheet);
		for (var s=0; s<nb_sheet; s++) {
			console.log("sheet: "+(s+1)+"/"+nb_sheet);
			var sheet = book.Worksheets(s+1);
			console.log("sheet: "+sheet.Name);
			// Range("B4").End(xlUp).Select  go to end of contiguous range
			
			console.log(""+sheet.Cells.Rows.Count);
			var xlCellTypeLastCell =11;
			var rng_bottom_right = sheet.Cells(1,1) .SpecialCells(xlCellTypeLastCell)
			var n = rng_bottom_right.Row, 
				m= rng_bottom_right.Column;
			console.log(n+" x "+m);
			//
			for (var i=0; i<n; i++) {
				for (var j=0; j<m; j++) {
					console.log(i+" x "+j+" "+sheet.Cells(i+1,j+1).FormulaR1C1);
				}
			}
			
			//
		}
	
		book.Close();
		break;
	}
});

/*
var win32com = require('win32com');
var xl = win32com.client.Dispatch('Excel.Application', 'C'); // locale 
xl.Visible = true;
var book = xl.Workbooks.Add();
var sheet = book.Worksheets(1);
sheet.Name = 'sheetnameA utf8';
sheet.Cells(1, 2).Value = 'test utf8';
var rg = sheet.Range(sheet.Cells(2, 2), sheet.Cells(4, 4));
rg.RowHeight = 5.18;
rg.ColumnWidth = 0.58;
rg.Interior.ColorIndex = 6; // Yellow 
book.SaveAs('testfileutf8.xls');
xl.ScreenUpdating = true;
xl.Workbooks.Close();
xl.Quit();

Application.ScreenUpdating = False
        Cancel = True
        LastRow = Cells(Rows.Count, keyColumn).End(xlUp).Row
        Set SortRange = Target.CurrentRegion
        
*/