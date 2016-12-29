function excel_automation_test() {
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
	//console.log(result);

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
		
		console.log("filepath: "+filepath+"\n");
		var book = xl.Workbooks.Open(filepath);
		var nb_sheet = book.Worksheets.Count;
		console.log("nb_sheet: "+nb_sheet);
		for (var s=0; s<nb_sheet; s++) {
			var sheet = book.Worksheets(s+1);
			console.log("sheet: "+(s+1)+"/"+nb_sheet+" "+sheet.Name);
			
			// find total span of content
			var xlCellTypeLastCell =11;
			var rng_bottom_right = sheet.Cells(1,1) .SpecialCells(xlCellTypeLastCell)
			var n = rng_bottom_right.Row, 
				m= rng_bottom_right.Column;
			console.log(n+" x "+m);
			// Range("B4").End(xlUp).Select  go to end of contiguous range
			
			for (var j=0; j<m; j++) {
				for (var i=0; i<n; i++) {
					var rng = sheet.Cells(i+1,j+1);
					var addr = ""+rng.Address;
					if (rng.HasArray()==true) {
						//console.log(rng.Address+" AF "+JSON.stringify(rng.HasArray())+" "+rng.CurrentArray().Cells(1,1).Address );
						var array_addr = ""+rng.CurrentArray().Cells(1,1).Address
						if (addr==array_addr)
						//if (rng.CurrentArray().Cells(1,1).Address==rng.Address)
							console.log(rng.CurrentArray().Address+" {} "+rng.FormulaArray);
						//rng.FormulaArray
					} else if (rng.HasFormula()==true) {
						var txt = ""+rng.FormulaR1C1;
						var rng0 = rng;
						if (i>0 && (""+sheet.Cells(i, j+1).FormulaR1C1)==txt)
							continue; // already used
						while ((""+sheet.Cells(i+1+1, j+1).FormulaR1C1)==txt)
							i++;
						var rng1 = sheet.Cells(i+1, j+1);
						var rng = sheet.Range(rng0, rng1);
						if (""+rng0.Address != ""+rng1.Address) {
							console.log(rng.Address+" [] "+rng0.FormulaR1C1);
						} else {
							console.log(rng.Address+"    "+rng.FormulaR1C1);
						}
					} else if (rng.Text!="") {
						// Value2 property doesn't use the Currency and Date 
						console.log(rng.Address+" VAL "+rng.Text);
					}
				}
			}
		}
	
		book.Close();
		break;
	}
});

/*
var book = xl.Workbooks.Add();
var sheet = book.Worksheets(1);
sheet.Cells(1, 2).Value = 'test utf8';
rg.RowHeight = 5.18;
rg.ColumnWidth = 0.58;
rg.Interior.ColorIndex = 6; // Yellow 
book.SaveAs('testfileutf8.xls');
xl.ScreenUpdating = true;
xl.Workbooks.Close();
xl.Quit();

Set SortRange = Target.CurrentRegion

*/