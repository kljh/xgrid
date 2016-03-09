
function parse_range_test() {
	var ranges = [
		'$A1',
		'$B$2',
		'Sheet1!A1',
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
		];
	
	for (var r=0; r<ranges.length; r++) {
		var tmp = parse_range(ranges[r]);
		info_msg(JSON.stringify(tmp,null,4));
	}
}

function parse_cell(cell) {
	if (!cell) return cell;
	
	var r1c1 = /(R(\[?([0-9]*)\]))?(C(\[?([0-9]*)\]))?/.exec(cell);
	info_msg(JSON.stringify(r1c1,null,4));
	if (r1c1[3] || r1c1[6] ) {
		return { 
			row : r1c1[3],
			row_abs : r1c1[2]==r1c1[3],
			col : r1c1[6],
			col_abs : r1c1[5]==r1c1[6],
			style : "R1C1"
		};
	}
	//
	var a1 = /(\$?([A-Z]*))?(\$?([0-9]*))?/.exec(cell);
	if (a1[2] || a1[4] ) {
		return { 
			row : a1[2],
			row_abs : a1[1]!=a1[2],
			col : a1[4],
			col_abs : a1[3]!=a1[4],
		};
	}
}
	
function parse_range(arg) {
	
	var rex = /((.*!)?)([^!]*)/.exec(arg);
	if (rex.length!=4) {
		return { name : arg };
	} 
	var book_and_sheet = rex[2];
	var range = rex[3];
	var cells = range.split(":");
	if (cells>2) throw new Error(arguments.callee.name+": "+arg)
	var cell0 = parse_cell(cells[0]);
	var cell1 = parse_cell(cells[1]);
	
	if (book_and_sheet) {
		var quoted = (book_and_sheet[0]=='\'');
		var len = book_and_sheet.length;
		book_and_sheet = quoted ? book_and_sheet.substr(1,len-2) : book_and_sheet.substr(0,len-1); 
		var rex = /(\[(.*)\])?(.*)/.exec(book_and_sheet);
		var book = rex[2];
		var sheet = rex[3];
	}
	
	return { 
		arg : arg,
		range: range, 
		cell0 : cell0,
		cell1 : cell1,
		sheet : sheet ? sheet.replace(/''/g, '\'') : sheet,
		book : book ? book.replace(/''/g, '\'') : book,
		 };
	//info_msg(rex.length + " " + JSON.stringify(rex));
	
}

function info_msg(msg) {
	console.log(msg);
}

parse_range_test();