/**
 * Created by Charles on 13/03/2016.
 */

var transf  = require('./excel_formula_transform');
var parser = require('./excel_formula_parse');
var rng	= require('./excel_range_parse');
var fs 	= require('fs');
var assert = require('assert');
var exec 	= require('child_process').exec;
var path	= require('path');


function info_msg(msg) {
	console.log(msg);
}

function dbg(name, variable) {
	info_msg(name+":\n"+JSON.stringify(variable,null,4));
}

function deps_graph_test() {
	
	var no_edges = [
		['A2', '="a"&"b"'],
		['A3', '=50'],
		['A4', '=50%']
	];
	var adjacency_list = formulas_list_to_adjacency_list(no_edges);
	adjacency_list_to_dot(adjacency_list, 'C:/temp/graph1.svg');
	assert.equal(JSON.stringify(adjacency_list),
	JSON.stringify({
		'{"cell0":{"col":1,"row":2}}': [],
		'{"cell0":{"col":1,"row":3}}': [],
		'{"cell0":{"col":1,"row":4}}': []
	}));
	
	var simple_directed_cells_only = [
		['W1', '=$A1'],
		['U1', '=$B$2'],
		['A1', '=$B$5']
	];
	var adjacency_list = formulas_list_to_adjacency_list(simple_directed_cells_only);
	adjacency_list_to_dot(adjacency_list, 'C:/temp/graph2.svg');
	assert.equal(JSON.stringify(adjacency_list),
	JSON.stringify({
		'{"cell0":{"col":23,"row":1}}': [ '{"cell0":{"col":1,"row":1}}' ],
		'{"cell0":{"col":21,"row":1}}': [ '{"cell0":{"col":2,"row":2}}' ],
		'{"cell0":{"col":1,"row":1}}': [ '{"cell0":{"col":2,"row":5}}' ]
	}));

	/*var all_formulas = [
		['A1', '="a"+1.2+TRUE+A1+xyz+#VALUE!'], // all operands
		['A2', '="a"&"b"'],	// & operator
		['A3', '=50'],
		['A4', '=50%'], 	// % operator
		['A5', '=3.1E-24-2.1E-24'],
		['A6', '=1+3+5'],
		['A7', '=3 * 4 + \n\t 5'],
		['A8', '=1-(3+5) '],
		['A9', '=1=5'], 	// = operator
		['B1', '=1<>5'],	// <> operator
		['C1', '=2^8'], 	// ^ operator
		['D1', '=TRUE'],
		['E1', '=FALSE'],
		['Q1', '={1.2,"s";TRUE,FALSE}'],	// array notation
		['W1', '=$A1'],
		['R1', '=$B$2'],
		['T1', '=Sheet1!A1'],
		['Y1', '=\'Another sheet\'!A1'],
		['U1', '=[Book1]Sheet1!$A$1'],
		['I1', '=[data.xls]sheet1!$A$1'],
		['O1', '=\'[Book1]Another sheet\'!$A$1'],
		['P1', '=\'[Book No2.xlsx]Sheet1\'!$A$1'],
		['S1', '=\'[Book No2.xlsx]Another sheet\'!$A$1'],
		['F1', '=[data.xls]sheet1!$A$1'],
		//['A1', '=[\'my data.xls\'],\'my sheet\'!$A$1'],
		['G1', '=SUM(B5:B15)'],
		['H1', '=SUM(B5:B15,D5:D15)'],
		['J1', '=SUM(B5:B15 A7:D7)'],
		['K1', '=SUM(sheet1!$A$1:$B$2)'],
		['L1', '=SUM((A:A 1:1))'],
		['Z1', '=SUM((A:A,1:1))'],
		['X1', '=SUM((A:A A1:B1))'],
		['V1', '=SUM(D9:D11,E9:E11,F9:F11)'],
		['N1', '=SUM((D9:D11,(E9:E11,F9:F11)))'],
		['M1', '=IF(P5=1.0,"NA",IF(P5=2.0,"A",IF(P5=3.0,"B",IF(P5=4.0,"C",IF(P5=5.0,"D",IF(P5=6.0,"E",IF(P5=7.0,"F",IF(P5=8.0,"G"))))))))'],
		['M2', '=SUM(B2:D2*B3:D3)'],  // array formula
		['M3', '=SUM( 123 + SUM(456) + IF(DATE(2002,1,6), 0, IF(ISERROR(R[41]C[2]),0, IF(R13C3>=R[41]C[2],0, '
			+'  IF(AND(R[23]C[11]>=55,R[24]C[11]>=20), R53C3, 0))))'],
		['M4', '=IF("a"={"a","b";"c",#N/A;-1,TRUE}, "yes", "no") &   "  more ""test"" text"'],
		['M5', '=+ AName- (-+-+-2^6) = {"A","B"} + @SUM(R1C1) + (@ERROR.TYPE(#VALUE!) = 2)'],
		['M6', '=IF(R13C3>DATE(2002,1,6),0,IF(ISERROR(R[41]C[2]),0,IF(R13C3>=R[41]C[2],0, IF(AND(R[23]C[11]>=55,R[24]C[11]>=20),R53C3,0))))'],
		['M7', '=IF(R[39]C[11]>65,R[25]C[42],ROUND((R[11]C[11]*IF(OR(AND(R[39]C[11]>=55, R[40]C[11]>=20),AND(R[40]C[11]>=20,R11C3="YES")),R[44]C[11],R[43]C[11]))+(R[14]C[11] *IF(OR(AND(R[39]C[11]>=55,R[40]C[11]>=20),AND(R[40]C[11]>=20,R11C3="YES")), R[45]C[11],R[43]C[11])),0))']
	];
	
	var adjacency_list = formulas_list_to_adjacency_list(all_formulas);

	adjacency_list_to_dot(adjacency_list, 'C:/temp/graph1.dot');*/
}

function formulas_list_to_adjacency_list(formulas /* [ [cell, formula_in_cell], ...]*/) {
	var adjacency_list = {};
	for (var f=0; f<formulas.length; f++) {
		var cell = formulas[f][0];
		var xlformula = formulas[f][1];
		try {
			var tokens = parser.getTokens(xlformula);
			// info_msg("TOKENS:\n"+JSON.stringify(tokens,null,4));
			var cell_deps = tokens_to_adjacent_list(tokens["items"]);
			var key = range_to_key(rng.parse_range(cell));
			if (adjacency_list.hasOwnProperty(key))
				throw "Dependencies of cell " + key + " read more than once !";
			adjacency_list[key] = cell_deps;
		} catch (e) {$
			info_msg("ERROR: "+e+"\n"+e.stack);
		}
	}
	//info_msg("adjacency_list:\n"+JSON.stringify(adjacency_list,null,4));
	return adjacency_list;
}


function tokens_to_adjacent_list(tokens) {
	var deps = []
	for (var i=0; i<tokens.length; ++i) {
		var token = tokens[i];
		if (token.subtype === "range") {
			var range_as_key = range_to_key(rng.parse_range(token.value));
			if (!contains(deps, range_as_key))
				deps.push(range_as_key);
		}
	}
	return deps;
}

/*function eq_range_obj(left, right) {
	// no need for  (left.hasOwnProperty("workbook") && right.hasOwnProperty("workbook"))
	// undefined !== undefined --> false
	if (left["workbook"] !== right["workbook"])
		return false;
	if (left["sheet"] !== right["sheet"])
		return false;
	if (left["cell1"] !== right["cell1"])
		return false;
	if (left["cell0"] !== right["cell0"])
		return false;
	return true;
}*/

function range_to_key(range) {
	var dict_to_print = range;
	// remove key arg
	dict_to_print.arg = undefined;
	// remove keys col_abs and row_abs
	if (dict_to_print["cell0"] !== undefined) {
		dict_to_print["cell0"].row_abs = undefined;
		dict_to_print["cell0"].col_abs = undefined;
	}
	if (dict_to_print["cell1"] !== undefined) {
		dict_to_print["cell1"].row_abs = undefined;
		dict_to_print["cell1"].col_abs = undefined;
	}
	return JSON.stringify(dict_to_print);
}

function adjacency_list_to_dot(adj_list, output_path) {
	var output = 'digraph G {\n'
	for (var vertex in adj_list) {
		//dbg("vertex", vertex);
		var parents = adj_list[vertex];
		output += '\t"' + vertex.replace(/"/g, "'") + '"';
		for (var i=0; i<parents.length; ++i)
			output += ' -> "' + parents[i].replace(/"/g, "'") + '"';
		output += ';\n';
	}
	output += '}';
	
	var output_dir = path.dirname(output_path);
	var filename = path.basename(output_path);
	var extension = path.extname(output_path).replace(/\./g, "");
	var dot_file = path.join(output_dir, filename.replace(extension, 'dot'));
	fs.writeFile(dot_file, output, function (err) {
		if (err)  throw err;
	});
	var path_to_dot = '"C:\\Program Files (x86)\\Graphviz2.38\\bin\\dot.exe"'
	var cmd = path_to_dot + " -T" + extension + ' ' + dot_file + ' -o ' + output_path;
	exec(cmd, function(error, stdout, stderr) {
		//dbg('stdout', stdout);
		//dbg('stderr', stderr);
		if (stderr) {
			throw stderr;
		} else {
			info_msg(output_path + " successfully created !");
		}
	});
}

function contains(array, obj) {
    var i = this.length;
    while (i--) {
        if (this[i] === obj) {
            return true;
        }
    }
    return false;
}

if (module == require.main) {
	deps_graph_test();
}

