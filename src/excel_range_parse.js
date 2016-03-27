var assert = require('assert');

function parse_range_bijection_test() {
	var ranges_a1 = [
		'A1',
		'A$1',
		'$A1',
		'$A$1',
		
		'Z2',
		'AA3',
		'AZ4',
		'BA5',
		'ZZ6',
		'AAA7',

		'B5:B15',
		'sheet1!$A$1:$B$2',
		
		'Sheet1!Z1',
		'\'Another sheet\'!A1',
		'[Book1]Sheet1!$A$1',
		'[data.xls]sheet1!$A$1',
		'\'[Book1]Another sheet\'!$A$1',
		'\'[Book No2.xlsx]Sheet1\'!$A$1',
		'\'[Book No2.xlsx]Another sheet\'!$A$1',
		'[data.xls]sheet1!$A$1',
		'\'Let\'\'s test\'!A1',
		
		'A:A',
		'1:1',
		
		'named_cell',
		'alpha0',

		'R1C1',
		'R1',
		'C1',
		
		];

	var ranges_r1c1 = [
	
		// R1C1 mode 

		'R[41]C[2]',
		'R13C3',
		'R[4]C10',
		'R23C[11]',
		
		//'R2',  // ambiguous
		//'R[-2]',
		//'C5',  // ambiguous
		//'C[3]',
		
		'named_cell',
		'radio_control_c1',

		];

	function test(ranges, mode) {
		for (var r=0; r<ranges.length; r++) {
			var rng_in = ranges[r]; 
			var tmp = parse_range(rng_in, mode);
			var rng_out = stringify_range(tmp, mode);

			//info_msg(JSON.stringify(rng_in));
			//info_msg(JSON.stringify(rng_out));
			//info_msg(JSON.stringify(tmp,null,4));
			
			// make sure everything is parsed
			assert(typeof tmp == "object" || rng_in=="named_cell" || rng_in=="alpha0" || rng_in=="radio_control_c1" || (rng_in=="R1C1"&&mode=="A1")); 
			// and bijective
			assert.equal(rng_in,rng_out); 
		}
	}
	test(ranges_a1, "A1");
	test(ranges_r1c1, "R1C1");
}

function parse_range_regression_tests() {
	assert.equal(JSON.stringify(parse_range('R1')),
		'{"col":18,"abs_col":false,"row":1,"abs_row":false}');
	assert.equal(JSON.stringify(parse_range('R[1]C[1]')),
		'{"row":"1","abs_row":false,"col":"1","abs_col":false}');
	//assert.equal(JSON.stringify(parse_range('R[1]')),
	//	'{"row":"1","abs_row":false,"abs_col":true,"end":{"row":"1","abs_row":false,"abs_col":true}');
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
		book_and_sheet = quoted ? book_and_sheet.substr(1,len-3) : book_and_sheet.substr(0,len-1); 
		var rex = /(\[(.*)\])?(.*)/.exec(book_and_sheet);
		var book = rex[2];
		var sheet = rex[3];

		book = book ? book.replace(/''/g, '\'') : book;
		sheet = sheet ? sheet.replace(/''/g, '\'') : sheet;
	}

	if (cell0 && !isNaN(cell0.row) && !isNaN(cell0.col)
		|| cell0 && cell1 && !isNaN(cell0.row) && !isNaN(cell1.row)
		|| cell0 && cell1 && !isNaN(cell0.col) && !isNaN(cell1.col) )
	{
		// properly formated range
		var res = cell0;
		res.end = cell1;
		res.book = book;
		res.sheet = sheet;
		return res;
	}
	else 
	{
		// named cell
		return arg;
	}
}

function parse_cell(cell, mode) {
	if (!cell) return cell;

	if (mode!="A1") {
		var r1c1 = /^(R(\[?(-?[0-9]*)\]?))?(C(\[?(-?[0-9]*)\]?))?$/.exec(cell);
		if (r1c1)
		if ( (r1c1[3] && r1c1[6]) || mode=="R1C1" ) {
			return { 
				row : r1c1[3],
				abs_row : r1c1[2]==r1c1[3],
				col : r1c1[6],
				abs_col : r1c1[5]==r1c1[6],
				//style : "R1C1"
			};
		}
	}

	if (mode!="R1C1") {
		var a1 = /^(\$?([A-Z]{1,3}))?(\$?([0-9]*))?$/.exec(cell);
		if (a1)
		if ( (a1[2] && a1[4]) || mode=="A1" ) {
			return { 
				col : parse_column_name_to_index(a1[2]),
				abs_col : a1[1]!=a1[2],
				row : parseFloat(a1[4]),
				abs_row : a1[3]!=a1[4],
			};
		}
	}

	return cell;
}

function parse_column_name_to_index(col) {
	if (!col) return col;

	var pos = 0;
	var mult = 1;
	var A = "A".charCodeAt(0);
	for (var k=0; k<col.length; k++) {
		var v = col.charCodeAt(col.length-(k+1)) - A;
		
		if (k==0) pos += v + 1; // units
		if (k==1) pos += v*26 + 26; // tens
		if (k==2) pos += v*26*26 + 26*26; // hundreds
		if (k>2) throw new Error("column on more than 3 letters")
	}
	return pos;
}

function stringify_range(rng, mode) {
	if (rng.substr) return rng; // named range

	var scope = "";
	var bNeedQuotes = false;
	var reNeedsQuotes = /^[A-Za-z0-9.]*$/;
	var reSubstQuotes = /'/g;
	if (rng.book) {
		bNeedQuotes = bNeedQuotes || !reNeedsQuotes.test(rng.book);
		scope += "["+rng.book.replace(reSubstQuotes,"''")+"]";
	}
	if (rng.sheet) {
		bNeedQuotes = bNeedQuotes || !reNeedsQuotes.test(rng.sheet);
		scope += rng.sheet.replace(reSubstQuotes,"''")

		if (bNeedQuotes) scope = "'"+scope+"'";
		scope += "!";
	}

	if (mode==="R1C1")
		return scope
			+ (rng.row ? "R"+(rng.abs_row ? rng.row : "["+rng.row+"]") : "")
			+ (rng.col ? "C"+(rng.abs_col ? rng.col : "["+rng.col+"]") : "")
			+ (rng.end ? ":"+stringify_range(rng.end, mode) : ""); 
	else
		return scope
			+ (rng.abs_col ? "$" : "") + parse_column_index_to_name(rng.col)
			+ (rng.abs_row ? "$" : "") + ( rng.row || "" )
			+ (rng.end ? ":"+stringify_range(rng.end, mode) : ""); 
	
}

function parse_column_index_to_name(pos) {
	if (pos===undefined || pos===null)
		return ""; // for instance for 1:1 
	
	var col = "";
	var A = "A".charCodeAt(0);
	var k = 0; // number of characters used to encode

	do {
		if (k==0) pos -= 1;
		if (k==1) pos -= 1; // effectively -26 because we have divieded by 26 once
		if (k==2) pos -= 1; // effectively -26*26 because we have divieded by 26 twice
		if (k>2) throw new Error("column on more than 3 letters. "+pos);
		k++;

		var u = pos%26;
		col = String.fromCharCode(A+u) + col;

		var pos = (pos-u)/26;
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
		parse_range_regression_tests();
		parse_range_bijection_test();
	}

	module.exports.parse_range = parse_range;
	module.exports.stringify_range = stringify_range;

}
