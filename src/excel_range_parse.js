var assert = require('assert');

function parse_range_test() {
	var ranges = [
		'$A1',
		'$B$2',
		'AA$3',
		'Sheet1!Z1',
		'\'Another sheet\'!A1',
		'[Book1]Sheet1!$A$1',
		'[data.xls]sheet1!$A$1',
		'\'[Book1]Another sheet\'!$A$1',
		'\'[Book No2.xlsx]Sheet1\'!$A$1',
		'\'[Book No2.xlsx]Another sheet\'!$A$1',
		'[data.xls]sheet1!$A$1',
		'\'Let\'\'s test\'!A1',
		'B5:B15',
		'sheet1!$A$1:$B$2',
		'A:A',
		'1:1',
		'R2',
		'R[-2]',
		'C5',
		'C[3]',
		'R[41]C[2]',
		'R13C3',
		'R[4]C10',
		'R23C[11]',
		'named_cell'
		];

	for (var r=0; r<ranges.length; r++) {
		var tmp = parse_range(ranges[r]);
		info_msg(JSON.stringify(tmp,null,4));
	}
}

function range_regression_tests() {
	assert.equal(JSON.stringify(parse_range('R1')),
		'{"arg":"R1","cell0":{"col":18,"col_abs":false,"row":1,"row_abs":false}}');
	assert.equal(JSON.stringify(parse_range('R[1]')),
		'{"arg":"R[1]","cell0":{"row":"1","row_abs":false,"col_abs":true}}');
	assert.equal(JSON.stringify(parse_range('R[1]C[1]')),
		'{"arg":"R[1]C[1]","cell0":{"row":"1","row_abs":false,"col":"1","col_abs":false}}');
}

function parse_range(arg, mode) {

	var rex = /((.*!)?)([^!]*)/.exec(arg);
	var book_and_sheet = rex[2];
	var range = rex[3];

	var cells = range.split(":");
	if (cells>2) throw new Error(arguments.callee.name+": "+arg)
	var cell0 = parse_cell(cells[0], mode);
	var cell1 = parse_cell(cells[1], mode);
	// named cell
	cell0 = cell0 || cells[0];

	if (book_and_sheet) {
		var quoted = (book_and_sheet[0]=='\'');
		var len = book_and_sheet.length;
		book_and_sheet = quoted ? book_and_sheet.substr(1,len-2) : book_and_sheet.substr(0,len-1); 
		var rex = /(\[(.*)\])?(.*)/.exec(book_and_sheet);
		var book = rex[2];
		var sheet = rex[3];

		book = book ? book.replace(/''/g, '\'') : book;
		sheet = sheet ? sheet.replace(/''/g, '\'') : sheet;
	}

	var res = cell0;
	res.end = cell1;
	res.book = book;
	res.sheet = sheet;

	return res;
}

function parse_cell(cell, mode) {
	if (!cell) return cell;

	var r1c1 = /(R(\[?(-?[0-9]*)\]?))?(C(\[?(-?[0-9]*)\]?))?/.exec(cell);
	var r1c1 = /(R(\[(-?[0-9]*)\]))?(C(\[(-?[0-9]*)\]))?/.exec(cell);
	if ( (r1c1[3] && r1c1[6]) || mode=="R1C1" ) {
		return { 
			row : r1c1[3],
			abs_row : r1c1[2]==r1c1[3],
			col : r1c1[6],
			abs_col : r1c1[5]==r1c1[6],
			//style : "R1C1"
		};
	}
	//
	var a1 = /(\$?([A-Z]*))?(\$?([0-9]*))?/.exec(cell);
	if ( (a1[2] && a1[4]) || mode=="A1" ) {
		return { 
			col : parse_column_name_to_index(a1[2]),
			abs_col : a1[1]!=a1[2],
			row : parseFloat(a1[4]),
			abs_row : a1[3]!=a1[4],
		};
	}
}

function parse_column_name_to_index(col) {
	if (!col) return col;

	var pos = 0;
	var mult = 1;
	var A = "A".charCodeAt(0);
	for (var i=col.length-1; i>=0; i--) {
		var v = col.charCodeAt(i) - A + 1;
		pos += mult*v;
		mult *= 26;
	}
	return pos;
}

function stringify_range(rng, mode) {
	if (rng.sheet || rng.book)
		throw new Error(arguments.callee.name+": external ranges not supported");

	return (rng.abs_col ? "$" : "") + parse_column_index_to_name(rng.col)
		+ (rng.abs_row ? "$" : "") + rng.row
		+ (rng.end ? ":"+stringify_range(rng.end, mode) : ""); 
	
	return "R"+(rng.abs_row ? rng.row : "["+rng.row+"]") 
		+ "C"+(rng.abs_col ? rng.col : "["+rng.col+"]")
		+ (rng.end ? ":"+stringify_range(rng.end, mode) : ""); 
}

function parse_column_index_to_name(pos) {
	var col = "";
	var A = "A".charCodeAt(0);
	do {
		var u = pos%26;
		var pos = (pos-u)/26;
		col = String.fromCharCode(A+u) + col;
	} while (pos!=0)
	return col;
}

function info_msg(msg) {
	if (typeof console != "undefined")
		console.log(msg);
	else 
		print(msg);
}


// Nodejs stuff
if (typeof module!="undefined") {

	var parser = require("./excel_formula_parse")

	// run tests if this file is called directly
	if (require.main === module) {
		range_regression_tests();
		parse_range_test();
	}

	module.exports.parse_range = parse_range;
	module.exports.stringify_range = stringify_range;

}
