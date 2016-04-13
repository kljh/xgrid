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
	
	/*var file_path = path.resolve(__dirname+'/../work/gearshape.xjson');
	var g = new Graph();
	g.from_xjson_file(file_path);
	g.dot_display('C:/temp/graph_gearshape.svg');
	var deps = circular_deps_check(g);
	assert.equal(deps, null);*/
	
	/*var no_edges = {
		'A2': '="a"&"b"',
		'A3': '=50',
		'A4': '=50%'
	};
	var g1 = new Graph();
	g1.create_from_formulas(no_edges);
	g1.dot_display('C:/temp/graph1.svg');
	assert.equal(JSON.stringify(g1.vertices),
		'{"A2":[],"A3":[],"A4":[]}');
	g1.begin();
	var i=1;
	while (!g1.isEnd()) {
		g1.next();
		++i;
	}
	assert.equal(i, g1.size());
	var deps = circular_deps_check(g1);
	assert.equal(deps, null);*/
	
	var simple_directed_cells_only = {
		'W1': '=$A1',
		'U1': '=$B$2',
		'A1': '=$B$5'
	};
	var g2 = new Graph();
	g2.create_from_formulas(simple_directed_cells_only);
	g2.dot_display('C:/temp/graph2.svg');
	assert.equal(JSON.stringify(g2.vertices),
		'{"W1":[],"A1":["W1"],"U1":[],"B2":["U1"],"B5":["A1"]}');
	// iterate from the top of a directed sub-graph thanks to prior knowledge
	g2.beginAt("B5");
	var i=1;
	while (!g2.isEnd()) {
		g2.next();
		++i;
	}
	assert.equal(i, g2.size()+1);
	var deps = circular_deps_check(g2);
	assert.equal(deps, null);
	g2.refresh_vertex('B5');
	
	/*var circular_deps = [
		['A1', '=$A2'],
		['A2', '=$A$3'],
		['A3', '=A1'],
		['AB5', '=3*A1'],
		['A2', '=2-T21']
	];
	var g3 = new Graph();
	g3.create_from_formulas(circular_deps);
	g3.dot_display('C:/temp/graph3.svg');
	assert.equal(JSON.stringify(g3.vertices),
		'{"A1":["A3","BB5"],"A2":["A1"],"A3":["A2"],"BB5":[],"T21":["A2"]}');
	var deps = circular_deps_check(g3);
	assert.equal(JSON.stringify(deps),
		'["A1","A3","A2","A1"]');
	
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
	assert.equal(JSON.stringify(g4.vertices),
		'{"A1":["W1"],"B1":["A1"],"W1":[],"A2":["W1"],"B2":["A2"],"C1":["B2","B1"]}');
	g4.beginAt("C1");
	var i=1;
	while (!g4.isEnd()) {
		g4.next();
		++i;
	}
	assert.equal(i, g4.size());
	var deps = circular_deps_check(g4);
	assert.equal(deps, null);*/
}

function Graph() {
	this.vertices = {};
	this.formulas_dict = {};
	this.size = function() {return Object.keys(this.vertices).length;}
	
	//
	//constructors
	//
	this.create_from_formulas = function(formulas) {
		this.vertices = formulas_dict_to_adjacency_list(formulas);
		this.formulas_dict = formulas;
	}
	this.dot_display = function(output_path) {adjacency_list_to_dot(this.vertices, output_path);}
	this.from_xjson_file = function(file_path) {
		var txt = fs.readFileSync(file_path, 'utf8');
		var data = JSON.parse(txt);
		var sheet = data.input;
		var adjacency_dict = {};
		
		for (var cell in sheet) {
			var xlformula = sheet[cell];
			var isFormula = xlformula.substr && xlformula[0]=="=";
			// no dependency whatsoever if no formula
			if (!isFormula)
				continue;
			var child = range_to_key(rng.parse_range(cell));
			//dbg('cell', cell);
			//dbg('xlformula', xlformula);
			if (adjacency_dict[child] === undefined) {
				adjacency_dict[child] = [];
			}

			var tokens = parser.getTokens(xlformula);
			var parents = tokens_to_range_list(tokens["items"]);
			for (var iParent=0; iParent<parents.length; ++iParent) {
				var parent = parents[iParent];
				if (adjacency_dict[parent] !== undefined) {
					adjacency_dict[parent].push(child);
				} else {
					adjacency_dict[parent] = [child];
				}
			}
		}
		this.vertices = adjacency_dict;
	}
	
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
	
	this.refresh_vertex = function(vertex) {
		var root = this.beginAt(vertex);
		dbg('root', root);
		while (!this.isEnd() && !this.queue.empty()) {
			var current_cell = this.next();
			dbg('refreshing cell', current_cell);
			var formula = this.formulas_dict[current_cell];
			if (formula === undefined)
				continue
			formula = formula.substr(1);
			dbg('eval_formula', eval_formula(formula));
		}
	}
	
	
}

function eval_formula(formula) {
	var vars =  {}, fcts= {};
	var expr = transf.parse_and_transfrom(formula,vars,fcts);
	dbg('expr', expr);
	var ids = []; 
	var args = []; 
	for (var addr in vars) {
		ids.push(vars[addr]);
		args.push(addr);
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

function formulas_dict_to_adjacency_list(formulas_dict) {
	this.formulas_dict = formulas_dict;
	var adjacency_dict = {};
	for (var cell in formulas_dict) {
		var xlformula = formulas_dict[cell];
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
	// temporary fix
	if (range === 'Math.PI')
		return range;
	var key =	rng.stringify_range(range);
	key = key.replace(/\$/g, "");
	return key;
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
	var debug = false;
	
	var visited = {};
	var vertex = graph.begin();
	var index = 0;
	var journey = [vertex];
	while (!graph.isEnd()) {
		if (debug)	dbg('vertex', vertex);
		if (debug)	dbg('queue', graph.queue.elements);
		if (debug)	dbg('visited', visited);
		// new strongly-connected graph
		if (graph.queue.empty()) {
			visited = {};
			index = 0;
			vertex = graph.next();
			journey = [vertex];
		} else {
			journey.push(vertex);
			++index;
			if (visited[vertex] !== undefined) {
				if (debug)	dbg('visited[vertex] !== undefined', visited);
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
			visited[vertex] = index;
			vertex = graph.next();
		}
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
