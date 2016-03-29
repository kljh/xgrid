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
	dbg("g1.vertices", g1.vertices);
	assert.equal(JSON.stringify(g1.vertices),
	JSON.stringify({
	    "{\"col\":1,\"row\":2}": [],
	    "{\"col\":1,\"row\":3}": [],
	    "{\"col\":1,\"row\":4}": []
	}));
	g1.begin();
	var i=1;
	while (!g1.isEnd()) {
		g1.next();
		++i;
	}
	assert.equal(i, g1.size());
	var deps = circular_deps_check(g1);
	assert.equal(deps, null);
	
	var simple_directed_cells_only = [
		['W1', '=$A1'],
		['U1', '=$B$2'],
		['A1', '=$B$5']
	];
	var g2 = new Graph();
	g2.create_from_formulas(simple_directed_cells_only);
	g2.dot_display('C:/temp/graph2.svg');
	dbg("g2.vertices", g2.vertices);
	assert.equal(JSON.stringify(g2.vertices),
	JSON.stringify({
	    "{\"col\":23,\"row\":1}": [],
	    "{\"col\":1,\"row\":1}": [
		"{\"col\":23,\"row\":1}"
	    ],
	    "{\"col\":21,\"row\":1}": [],
	    "{\"col\":2,\"row\":2}": [
		"{\"col\":21,\"row\":1}"
	    ],
	    "{\"col\":2,\"row\":5}": [
		"{\"col\":1,\"row\":1}"
	    ]
	}));
	g2.beginAt("{\"col\":2,\"row\":5}");
	var i=1;
	while (!g2.isEnd()) {
		g2.next();
		++i;
	}
	assert.equal(i, g2.size()+1);
	var deps = circular_deps_check(g2);
	assert.equal(deps, null);
	
	var circular_deps = [
		['A1', '=$A2'],
		['A2', '=$A$3'],
		['A3', '=A1'],
		['AX5', '=3*A1'],
		['A2', '=2-T21']
	];
	var g3 = new Graph();
	g3.create_from_formulas(circular_deps);
	g3.dot_display('C:/temp/graph3.svg');
	dbg("g3.vertices", g3.vertices);
	assert.equal(JSON.stringify(g3.vertices), JSON.stringify({
	    "{\"col\":1,\"row\":1}": [
		"{\"col\":1,\"row\":3}",
		"{\"col\":50,\"row\":5}"
	    ],
	    "{\"col\":1,\"row\":2}": [
		"{\"col\":1,\"row\":1}"
	    ],
	    "{\"col\":1,\"row\":3}": [
		"{\"col\":1,\"row\":2}"
	    ],
	    "{\"col\":50,\"row\":5}": [],
	    "{\"col\":20,\"row\":21}": [
		"{\"col\":1,\"row\":2}"
	    ]
	}));
	var deps = circular_deps_check(g3);
	dbg('deps', deps);
	assert.equal(JSON.stringify(deps), JSON.stringify(
		[
		    "{\"col\":1,\"row\":1}",
		    "{\"col\":1,\"row\":3}",
		    "{\"col\":1,\"row\":2}",
		    "{\"col\":1,\"row\":1}"
		]
	));
	
	var diamond = [
		['A1', '=B1'],
		['W1', '=A1+A2'],
		['A2', '=B2'],
		['B2', '=C1'],
		['B1', '=C1']
	];
	var g4 = new Graph();
	g4.create_from_formulas(diamond);
	g4.dot_display('C:/temp/graph4.svg');
	dbg("g4.vertices", g4.vertices);
	assert.equal(JSON.stringify(g4.vertices),
		JSON.stringify({
		    "{\"col\":1,\"row\":1}": [
			"{\"col\":23,\"row\":1}"
		    ],
		    "{\"col\":2,\"row\":1}": [
			"{\"col\":1,\"row\":1}"
		    ],
		    "{\"col\":23,\"row\":1}": [],
		    "{\"col\":1,\"row\":2}": [
			"{\"col\":23,\"row\":1}"
		    ],
		    "{\"col\":2,\"row\":2}": [
			"{\"col\":1,\"row\":2}"
		    ],
		    "{\"col\":3,\"row\":1}": [
			"{\"col\":2,\"row\":2}",
			"{\"col\":2,\"row\":1}"
		    ]
	}));
	g4.beginAt("{\"col\":3,\"row\":1}");
	var i=1;
	while (!g4.isEnd()) {
		g4.next();
		++i;
	}
	assert.equal(i, g4.size());
	var deps = circular_deps_check(g4);
	assert.equal(deps, null);
}

function Graph() {
	this.vertices = {};
	this.size = function() {return Object.keys(this.vertices).length;}
	this.create_from_formulas = function(formulas) {
		this.vertices = formulas_list_to_adjacency_list(formulas)
	}
	this.dot_display = function(output_path) {adjacency_list_to_dot(this.vertices, output_path);}
	
	//
	//graph iteration functions
	//
	// current is the index of the current node in Object.keys(vertices)
	this.current = null;
	// use of a hashmap to check for missing vertices
	// retrieval O(1), worst-case (when collisions happen) O(N) (to be verified)
	this.visited = {};
	this.queue = new Queue();
	this.begin = function() {
		this.current = 0;
		this.visited = {};
		this.queue.flush();
		var vertex = Object.keys(this.vertices)[0];
		if (this.vertices[vertex].length > 0) {
			for (var i=0; i<this.vertices[vertex].length; ++i) {
				var child_vertex = this.vertices[vertex][i];
				if (!contains(this.queue.elements, child_vertex))
					this.queue.enqueue(child_vertex);
			}
		}
		this.visited[vertex] = true;
		return vertex;
	}
	this.beginAt = function(root) {
		if (this.vertices[root] === undefined)
			throw root + " not in this graph";
		this.current = Object.keys(this.vertices).indexOf(root);
		this.visited = {};
		this.queue.flush();
		if (this.vertices[root].length > 0) {
			for (var i=0; i<this.vertices[root].length; ++i) {
				var child_vertex = this.vertices[root][i];
				if (!contains(this.queue.elements, child_vertex))
					this.queue.enqueue(child_vertex);
			}
		}
		this.visited[root] = true;
		return root;
	}
	this.next = function() {
		var debug = false;
		// new 'loosely' connected vertex
		if (this.queue.empty()) {
			this.current = (this.current + 1) % this.size();
			if (debug)	dbg('current', this.current);
			var vertex = Object.keys(this.vertices)[this.current];
			// already visited
			if (this.visited[vertex] !== undefined) {
				if (debug)	info_msg("Already visited " + vertex);
				return this.next();
			}
		} else {		
			// we need to traverse some connected vertices first
			if (debug)	dbg("State of the queue", this.queue.elements);
			var vertex = this.queue.dequeue();
		}

		if (this.vertices[vertex].length === 0) {
			if (debug)	dbg("Childless vertex", vertex);
		} else {
			// new sub-graph
			if (debug)	dbg('Starting search from root', vertex);
			for (var i=0; i<this.vertices[vertex].length; ++i) {
				var child_vertex = this.vertices[vertex][i];
				if (!contains(this.queue.elements, child_vertex))
					this.queue.enqueue(child_vertex);
			}
		}

		this.visited[vertex] = true;
		return vertex;
	}
	this.isEnd = function() {
		return this.queue.empty() && Object.keys(this.visited).length === this.size();
	}
	
}

// FIFO Queue class
function Queue() {
	this.elements = [];
	this.enqueue = function(elem) {this.elements.push(elem);}
	this.dequeue = function() {return this.elements.shift();}
	this.empty = function() {return this.elements.length === 0;}
	this.repr = function() {return JSON.stringify(this.elements);}
	this.flush = function() {this.elements = [];}
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
	dict_to_print["abs_row"] = undefined;
	dict_to_print["abs_col"] = undefined;
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

function circular_deps_check(graph) {
	var visited = {};
	var vertex = graph.begin();
	var index = 0;
	var journey = [vertex];
	visited[vertex] = index;
	while (!graph.isEnd()) {
		vertex = graph.next();
		// new strongly-connected graph
		if (graph.queue.empty()) {
			visited = {};
			index = 0;
			journey = [vertex];
		} else {
			++index;
			journey.push(vertex);
			if (visited[vertex] !== undefined) {
				var full_journey = journey.slice(visited[vertex], index+1);
				// direct_journey is the direct path of the circular dependency
				var direct_journey = [];
				var current_node = full_journey[0];
				for (var i=0; i<full_journey.length-1; ++i) {
					var next_node = full_journey[i+1];
					var children = graph.vertices[current_node];
					if (contains(children, next_node)) {
						direct_journey.push(current_node);
						current_node = next_node;
					}
				}
				direct_journey.push(full_journey[full_journey.length-1]);
				return direct_journey;
			}
		}
		visited[vertex] = index;
	}
	return null;
}

function contains(array, obj) {
	var i = array.length;
	while (i--)
		if (array[i] === obj)
			return true;
	return false;
}

if (module == require.main) {
	deps_graph_test();
}

module.exports.Graph = Graph;
