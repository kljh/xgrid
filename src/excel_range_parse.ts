
export enum ReferenceStyle { Auto, A1, R1C1 };

interface CellAddress {
	col? : number;
	abs_col? : boolean;
	row? : number;
	abs_row? : boolean;
}

export interface RangeAddress extends CellAddress {
	reference_to_named_range? : string;
	end? : CellAddress;
	book? : string;
	sheet? : string;
}

export function parse_range_bijection_test() {
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
		'A:C',
		'1:1',
		'1:5',
		
		'named_cell',
		'alpha0',

		'R1',
		'C1',
		
		];

	var ranges_r1c1 = [
	
		// R1C1 mode 

		'RC',
		'R[1]C',
		'RC[2]',
		'R[1]C[2]',
		'R13C3',
		'R[4]C7',
		'R23C[11]',
		'RC13', 	// ambiguous : cell at current row and absolute column can be interpreted as A1 cell

		'R1C2:R[1]C[2]',
		
		// Not supported :
		//'R2', 	// ambiguous : full R1C1 row can be interpreted as A1 cell
		//'R[-2]',
		//'C5', 	// ambiguous : full R1C1 column can be interpreted as A1 cell
		//'C[3]',
		//'R3:R5', 	// ambiguous : full R1C1 rows can be interpreted as A1 range
		
		'named_cell',
		'radio_control_c1',

		];

	let nb_errors = 0;
	function test(ranges, mode : ReferenceStyle, rng_ref? : RangeAddress) {
		for (var r=0; r<ranges.length; r++) {
			var rng_in = ranges[r]; 
			var tmp = parse_range(rng_in, mode, rng_ref);
			var rng_out = stringify_range(tmp, mode, rng_ref);

			//info_msg(JSON.stringify(rng_in));
			//info_msg(JSON.stringify(rng_out));
			//info_msg(JSON.stringify(tmp,null,4));
			
			// make sure everything is parsed
			var bOk = typeof tmp == "object" || rng_in=="named_cell" || rng_in=="alpha0" || rng_in=="radio_control_c1" || (rng_in=="R1C1"&&mode==ReferenceStyle.A1);
			if (!bOk) throw new Error("parse_range_bijection_test"); 
			// and bijective
			// assert.equal(rng_in,rng_out);
			if (rng_in!=rng_out) {
				console.log(rng_in, " <> ", rng_out);
				nb_errors++;
			} 
		}
	}
	var rng_ref = parse_range("A1", ReferenceStyle.A1);
	test(ranges_a1, ReferenceStyle.A1, rng_ref);
	test(ranges_r1c1, ReferenceStyle.R1C1, rng_ref);
	if (nb_errors) throw new Error("parse_range_bijection_test");
	
	function assert_equal(a, b) { if (a!==b) throw new Error("parse_range_bijection_test"); }
	var assert = { equal: assert_equal };

	assert.equal(JSON.stringify(parse_range('R1', ReferenceStyle.A1, rng_ref)),
		'{"col":18,"abs_col":false,"row":1,"abs_row":false}');
	assert.equal(JSON.stringify(parse_range('R[1]C[1]', ReferenceStyle.R1C1, rng_ref)),
		'{"row":2,"abs_row":false,"col":2,"abs_col":false}');
	//assert.equal(JSON.stringify(parse_range('R[1]', ReferenceStyle.R1C1, rng_ref)),
	//	'{"row":1,"abs_row":false,"abs_col":false,"end":{"row":1,"abs_row":false,"abs_col":false}');

	assert.equal(parse_column_name_to_index("A"),1);
	assert.equal(parse_column_name_to_index("Z"),26);
	assert.equal(parse_column_name_to_index("AA"),27);
	assert.equal(parse_column_name_to_index("AZ"),52);
	assert.equal(parse_column_name_to_index("BA"),53);
	assert.equal(parse_column_name_to_index("AAA"), parse_column_name_to_index("ZZ")+1);
	assert.equal(parse_column_name_to_index("ABA"), parse_column_name_to_index("AAZ")+1);
	assert.equal(parse_column_name_to_index("BAA"), parse_column_name_to_index("AZZ")+1);
	assert.equal(parse_column_index_to_name(parse_column_name_to_index("A")),"A");
	assert.equal(parse_column_index_to_name(parse_column_name_to_index("Z")),"Z");
	assert.equal(parse_column_index_to_name(parse_column_name_to_index("AA")),"AA");
	assert.equal(parse_column_index_to_name(parse_column_name_to_index("AZ")),"AZ");
	assert.equal(parse_column_index_to_name(parse_column_name_to_index("AAA")),"AAA");
	assert.equal(parse_column_index_to_name(parse_column_name_to_index("AZZ")),"AZZ");

}


export function parse_range(arg, mode : ReferenceStyle, rng_ref? : RangeAddress) {

	var rex = /((.*!)?)([^!]*)/.exec(arg);
	var book_and_sheet = rex[2];
	var range = rex[3];

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

	var res = parse_cell_range(range, mode, rng_ref);
	res.book = book;
	res.sheet = sheet;
	return res;
}


function parse_cell_range(range : string, mode : ReferenceStyle, rng_ref? : RangeAddress) : RangeAddress {
	let cells = range.split(":");

	if (cells.length>2) throw new Error("range address invalid, contains more than on colons. "+range);
	
	var cell0 = parse_cell(cells[0], mode, rng_ref);
	var cell1 = parse_cell(cells[1], mode, rng_ref);
	
	if (cell0 && !isNaN(cell0.row) && !isNaN(cell0.col)
		|| cell0 && cell1 && !isNaN(cell0.row) && !isNaN(cell1.row)
		|| cell0 && cell1 && !isNaN(cell0.col) && !isNaN(cell1.col) )
	{
		// properly formated range
		var res = cell0;
		res.end = cell1;
		return res;
	}
	else 
	{
		// named range
		return { reference_to_named_range: range };
	}
		
}


function parse_cell(cell : string, mode : ReferenceStyle, rng_ref? : RangeAddress) : RangeAddress {
	if (!cell) return;
	switch (mode) {
		case ReferenceStyle.Auto:
			var isA1 = A1.test(cell),
				isR1C1 = R1C1.test(cell) || cell=="RC";
			if (isA1 && !isR1C1)
				return parse_cell_A1(cell, rng_ref);
			if (!isA1 && isR1C1)
				return parse_cell_R1C1(cell, rng_ref);
			if (isA1 && isR1C1)
				throw new Error("parse_cell: ambiguous range address '"+cell+"' in 'Auto' mode.");
			return { reference_to_named_range: cell };
		case ReferenceStyle.A1:
			return parse_cell_A1(cell, rng_ref);
		case ReferenceStyle.R1C1:
			return parse_cell_R1C1(cell, rng_ref);
		default:
			throw new Error("parse_cell: unknown reference style "+cell+" "+mode);
	}
}

const R1C1 = /^(R(\[?(-?[0-9]*)\]?))?(C(\[?(-?[0-9]*)\]?))?$/;
function parse_cell_R1C1(cell : string, rng_ref? : RangeAddress) : RangeAddress {
	const r1c1 = R1C1.exec(cell);
	if (r1c1) {
		if (r1c1[3] || r1c1[6] || cell=="RC") {
			var res = { 
				row : parseFloat(r1c1[3]),
				abs_row : r1c1[2]==r1c1[3] && !!r1c1[2],
				col : parseFloat(r1c1[6]),
				abs_col : r1c1[5]==r1c1[6] && !!r1c1[5]
			};
			if (!res.abs_row) res.row = rng_ref.row + (isNaN(res.row)?0:res.row);
			if (!res.abs_col) res.col = rng_ref.col + (isNaN(res.col)?0:res.col);
			return res;
		}
	}
}

const A1 = /^(\$?([A-Z]{0,3}))?(\$?([0-9]*))?$/;
function parse_cell_A1(cell : string, rng_ref? : RangeAddress) : RangeAddress {
	if (!cell) return;

	const a1 = A1.exec(cell);
	if (a1) {
		if (a1[2] || a1[4]) {
			return { 
				col : parse_column_name_to_index(a1[2]),
				abs_col : a1[1]!=a1[2],
				row : parseFloat(a1[4]),
				abs_row : a1[3]!=a1[4],
			};
		}
	}
}

function parse_column_name_to_index(col : string) : number|undefined {
	if (!col) 
		return parseFloat("NaN");

	// we are NOT counting in base 25/26 : A is both 0 (when not the leading number) and 1 (when the leading number)
	// we are using symbols [A-Z] rather than [0-9A-P]
	//return parseInt(col, 26);

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

export function stringify_range(rng, mode : ReferenceStyle, rng_ref? : RangeAddress) {
	if (rng.reference_to_named_range) return rng.reference_to_named_range; // named range

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

	if (mode===ReferenceStyle.R1C1)
		return scope
			+ (rng.row ? "R"+(rng.abs_row ? rng.row : (rng_ref.row==rng.row ? "" : "["+(rng.row-rng_ref.row)+"]")) : "")
			+ (rng.col ? "C"+(rng.abs_col ? rng.col : (rng_ref.col==rng.col ? "" : "["+(rng.col-rng_ref.col)+"]")) : "")
			+ (rng.end ? ":"+stringify_range(rng.end, mode, rng_ref) : ""); 
	else
		return scope
			+ (rng.abs_col ? "$" : "") + parse_column_index_to_name(rng.col)
			+ (rng.abs_row ? "$" : "") + ( rng.row || "" )
			+ (rng.end ? ":"+stringify_range(rng.end, mode, rng_ref) : ""); 
	
}

function parse_column_index_to_name(pos: number):string {
	if (pos===undefined || pos===null || isNaN(pos))
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


// Nodejs stuff
if (typeof module!="undefined") {

	var parser = require("./excel_formula_parse")

	// run tests if this file is called directly
	if (require.main === module) {
		parse_range_bijection_test();
	}

	module.exports.parse_range = parse_range;
	module.exports.stringify_range = stringify_range;

}
