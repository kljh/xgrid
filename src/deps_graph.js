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

var info_msg = rng.info_msg;
var dbg = rng.dbg;

function deps_graph_test() {
	
	var no_edges = [
		['A2', '="a"&"b"'],
		['A3', '=50'],
		['A4', '=50%']
	];
	var g1 = new Graph();
	g1.create_from_formulas(no_edges);
	g1.dot_display('C:/temp/graph1.svg');
	assert.equal(JSON.stringify(g1.vertices),
	JSON.stringify({
	    "{\"cell0\":{\"col\":1,\"row\":2}}": [],
	    "{\"cell0\":{\"col\":1,\"row\":3}}": [],
	    "{\"cell0\":{\"col\":1,\"row\":4}}": []
	}));
	
	var simple_directed_cells_only = [
		['W1', '=$A1'],
		['U1', '=$B$2'],
		['A1', '=$B$5']
	];
	var g2 = new Graph();
	g2.create_from_formulas(simple_directed_cells_only);
	g2.dot_display('C:/temp/graph2.svg');
	assert.equal(JSON.stringify(g2.vertices),
	JSON.stringify({
	    "{\"cell0\":{\"col\":23,\"row\":1}}": [],
	    "{\"cell0\":{\"col\":1,\"row\":1}}": [
		"{\"cell0\":{\"col\":23,\"row\":1}}"
	    ],
	    "{\"cell0\":{\"col\":21,\"row\":1}}": [],
	    "{\"cell0\":{\"col\":2,\"row\":2}}": [
		"{\"cell0\":{\"col\":21,\"row\":1}}"
	    ],
	    "{\"cell0\":{\"col\":2,\"row\":5}}": [
		"{\"cell0\":{\"col\":1,\"row\":1}}"
	    ]
	}));
	
	var circular_deps = [
		['A1', '=$A2'],
		['A2', '=$A$3'],
		['A3', '=A1']
	];
	var g3 = new Graph();
	g3.create_from_formulas(circular_deps);
	g3.dot_display('C:/temp/graph3.svg');
	assert.equal(JSON.stringify(g3.vertices),
	JSON.stringify({
	    "{\"cell0\":{\"col\":1,\"row\":1}}": [
		"{\"cell0\":{\"col\":1,\"row\":3}}"
	    ],
	    "{\"cell0\":{\"col\":1,\"row\":2}}": [
		"{\"cell0\":{\"col\":1,\"row\":1}}"
	    ],
	    "{\"cell0\":{\"col\":1,\"row\":3}}": [
		"{\"cell0\":{\"col\":1,\"row\":2}}"
	    ]
	}));
}

function Graph() {
	this.vertices = {};
	this.size = function() {return Object.keys(this.vertices).length;}
	this.create_from_formulas = function(formulas) {
		this.vertices = formulas_list_to_adjacency_list(formulas)
	}
	this.dot_display = function(output_path) {adjacency_list_to_dot(this.vertices, output_path);}
}

function formulas_list_to_adjacency_list(formulas /* [ [cell, formula_in_cell], ...]*/) {
	var adjacency_dict = {};
	for (var f=0; f<formulas.length; f++) {
		var cell = formulas[f][0];
		var xlformula = formulas[f][1];
		try {
			var child = range_to_key(rng.parse_range(cell));
			if (adjacency_dict[child] === undefined) {
				adjacency_dict[child] = [ ];
			}
			var tokens = parser.getTokens(xlformula);
			var parents = tokens_to_range_list(tokens["items"]);
			//dbg('child', child);
			//dbg('parents', parents);
			for (var iParent=0; iParent<parents.length; ++iParent) {
				var parent = parents[iParent];
				if (adjacency_dict[parent] !== undefined) {
					adjacency_dict[parent].push(child);
				} else {
					adjacency_dict[parent] = [child];
				}
			}
		} catch (e) {
			info_msg("ERROR: "+e+"\n"+e.stack);
		}
	}
	//dbg("adjacency_dict", adjacency_dict);
	return adjacency_dict;
}


function tokens_to_range_list(tokens) {
	var ranges = []
	for (var i=0; i<tokens.length; ++i) {
		var token = tokens[i];
		if (token.subtype === "range") {
			var range_as_key = range_to_key(rng.parse_range(token.value));
			if (!contains(ranges, range_as_key))
				ranges.push(range_as_key);
		}
	}
	return ranges;
}

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

module.exports.Graph = Graph;
