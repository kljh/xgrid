import * as rp from "./excel_range_parse"
var glbl = require("./global_scope.js");

/*
function evaluate_formula(src, vars) {
	var var_names=[], var_values=[];
	for (var v in vars) {
		var_names.push(v);
		var_values.push(vars[v]);
	}

	var var_names_and_function_src = var_names.concat(["return "+src+";"]);
	var fct = Function.prototype.constructor.apply(undefined, var_names_and_function_src);

	var res = fct.apply(fct, var_values);
	return res;
}
*/

export function spreadsheet_exec_test() {
	var fs = require('fs');
	var path = require('path');

	var folder_path = __dirname+'/../work/';
	var filepaths = fs.readdirSync(folder_path);
	for (var i=0; i<filepaths.length; i++) {
		var filepath = path.resolve(folder_path+filepaths[i]);
		if (filepath.indexOf(".xjson")==-1) continue;

		info_msg("Executing "+filepath);
		
		var txt = fs.readFileSync(filepath, 'utf8');
		var data = JSON.parse(txt);
		var sheet = data.input;
		if (data.named_ranges)
			for (var k in data.named_ranges) 
				sheet[k] = data.named_ranges[k];
		
		//try {
			var node_values = sheet_exec(sheet, { 
				"input_syntax" : data.input_syntax || "excel",
				"xjs": true, // transform Excel operators into javascript
			 });

			var tmp = filepath.split('/');
			tmp[tmp.length-2] = "ref";
			var filepath_out = tmp.join('/');
			tmp.pop();
			var folder_out = tmp.join('/');

			//if (!fs.existsSync(folder_out)) fs.mkdirSync(folder_out);
			//fs.writeFileSync(filepath_out, JSON.stringify({input: node_values},null,4));

			data.ref_output = node_values;
			fs.writeFileSync(filepath, JSON.stringify(data,null,2));

		//} catch (e) {
		//	info_msg("sheet "+JSON.stringify(sheet, null, 4));
		//	throw new Error("Error evaluating "+filepath+"\n"+(e.stack||e)+"\n");
		//}
	}
	info_msg("\nspreadsheet_exec_test DONE.")
}

export interface NodeInfo  {
	value? : any; // result of node evaluation
	target : rp.RangeAddress; // result's target range (spreadsheet specific
	
	expr?: string;  // a valid Javascript expression (using A1 or RC or anonymous notations ? does it mater ?)
	fct?: Function; // the corresponding javascript code
	args?: Array<string>; // source range for the arguments (will have to be in the nodes map at some stage)
	
}

//export type NodeMap = Dictionary<string, NodeInfo>;
export type  NodeMap = any;

function sheet_exec(sheet, prms) {
	var indent = "";

	info_msg("sheet range and formular parsing...")
	
	var node_eval:NodeMap = {};
	var node_values = {};
	for (var r in sheet) {
		var val = sheet[r];
		if (val==="") continue;

		var target = range_parser.parse_range(r, rp.RangeAddressStyle.A1);
		
		var isFormula = val.substr && val[0]=="=";
		var isFormulaArray = val.substr && val[0]=="{" && val[1]=="=" && val[val.length-1]=="}";
		if (isFormula || isFormulaArray) {
			//try {
				var formula = isFormula ? val.substr(1) : val.substr(2, val.length-3);
				let expr:string;

				let vars =  {}, fcts= {};
				switch (prms.input_syntax) {
					case "excel":
						expr = xlfx.parse_and_transfrom(formula,vars,fcts,prms);
						break;
					case "javascript":
						let formula2 = lexer.rebuild_source_text(lexer.handle_at_operator(lexer.lexer(formula)));
						expr = jsfx.parse_and_transfrom(formula2,vars,fcts,prms);
						break;
					default:
						throw new Error("unhandled input syntax.");
				}
				
				var ids = []; 
				var args = []; 
				for (var addr in vars) {
					if ([ "Array", "Date", "Math", "Number", "Object", "String", "arguments", "caller", "this", "undefined", "JSON", "window" ].indexOf(addr)!=-1) {
						warn_msg("WARN: reserved keyword '"+addr+"' used in "+r);
						continue;
					}
					ids.push(vars[addr]);
					args.push(addr);
				}

				//var func = new Function(ids, "try { var res="+expr+"; return res; } catch(e) { throw new Error('"+r+" '+(e.stack||e)+'\\n'); }");
				//var func = new Function(ids, "try { var res="+expr+"; return res; } catch(e) { throw new Error('"+r+" '+(e)+'\\n'); }");
				//var func = Function(ids, "try { var res="+expr+"; return res; } catch(e) { throw new Error('"+r+" '+(e)+'\\n'); }");
				var func = new Function(ids.join(","), "var res="+expr+"; return res;");
				info_msg("expr: "+expr);

				if (target.end && !isFormulaArray) {
					for (var i=target.row; i<=target.end.row; i++)
						for (var j=target.col; j<=target.end.col; j++) {

							var tmp_args = new Array(args.length);
							for (var a=0; a<args.length; a++) 
								tmp_args[a] = range_parser.stringify_range(move_range(range_parser.parse_range(args[a], rp.RangeAddressStyle.A1), i-target.row, j-target.col));

							let tmp_target = { row: i, col: j }
							let tmp_id = range_parser.stringify_range(tmp_target, rp.RangeAddressStyle.A1);
							node_eval[tmp_id] = {
								target: tmp_target,
								expr: expr,
								fct: func,
								args: tmp_args
								};
						}
				} else {
					node_eval[r] = {
						target: target,
						expr: expr,
						fct: func,
						args: args
						};
				}
			/*
			} catch (e) {
				node_eval[r] = {
					target: target,
					expr: val,
					fct: function(x) { return function() { throw new Error("ill formated expression: "+x); }; },
					error: e.stack || (""+e)
					};
			}
			*/
		} else {

			node_eval[r] = {
				target: target,
				value: val,
				};
		}
	}

	info_msg("sheet static dependency tree ...")
	let formula_targets = Object.keys(node_eval);
	let formula_missing = sheet_static_dependencies(node_eval, true);
	info_msg("  formula_targets : #"+formula_targets.length+" "+formula_targets);
	info_msg("  formula_missing : #"+formula_missing.length+" "+formula_missing);
	
	info_msg("sheet evaluate...")
	/*
	info_msg("evaluation: "+JSON.stringify(node_eval, null, 4));
	*/

	for (var rng in node_eval) {
		var res = evaluate_node(rng, node_eval, indent+" ");
		node_values[rng] = res;
		try {
			info_msg(rng+" value is "+JSON.stringify(res));
		} catch(e) {
			info_msg(rng+" value is "+res+" (not JSON stringifiable).");
		}
	}

	let adjacency_list = extract_adjacency_list(node_eval);
	let calculation_order = static_calculation_order(adjacency_list);

	return node_values;
}

function extract_adjacency_list(nodes) {
	let adjacency_list = {};
	for (let n in nodes) {
		if (!nodes[n]) 
			continue;
		info_msg("extract_adjacency_list node "+JSON.stringify(n));
		let args = nodes[n].args || [];
		let deps = nodes[n].deps ? Object.keys(nodes[n].deps) : [];
		if (JSON.stringify(args)!=JSON.stringify(deps) && args.length && deps.length)
			throw new Error("args (for formula) and deps (for range aggregation) both present and inconsistent");
		adjacency_list[n] = deps.length ? deps : args;
	}
	return adjacency_list;
}

function static_calculation_order(adjacency_list) {
	let calculation_order = [];
	let adjacency_list_length = Object.keys(adjacency_list).length;
	for (let it=0; ; it++) {
		let nb_added_at_this_iteration = 0;		
		for (let n in adjacency_list) {
			if (calculation_order.indexOf(n)!=-1)
				continue;

			var deps = adjacency_list[n];
			let depsOk = true;
			for (let d=0; d<deps.length; d++) {
				if (calculation_order.indexOf(deps[d])==-1) {
					depsOk = false;
					break;
				}
			}
			if (!depsOk)
				continue;
			calculation_order.push(n);
			nb_added_at_this_iteration++;
		}
		if (calculation_order.length==adjacency_list_length)
			return calculation_order;
		if (nb_added_at_this_iteration===0)
			throw new Error("circular reference detected");
	}
}

function sheet_static_dependencies(nodes, bCreateMissing) : string[] {
	var missing_deps = {};
	for (let n in nodes) {
		const args = nodes[n].args;
		if (args)
		for (var i=0; i<args.length; i++) {
			if (!(args[i] in nodes)) 
				missing_deps[args[i]] = true;
		}
	}

	return Object.keys(missing_deps);
}

function create_on_the_fly_node(id, indent: string) {
	info_msg(indent+"create_on_the_fly_node: id="+id);
	var rng = range_parser.parse_range(id, rp.RangeAddressStyle.A1);
	//if (rng.end && (rng.end.row!=rng.row || rng.end.col!=rng.col || rng.end.row_abs || rng.col_abs)) {
 		return { rng: rng };
	//} else {
	//	info_msg(indent+"create_on_the_fly_node: id="+id+" not needed");
	//	return; // throw new Error(id+" is not a multi-cell range");
	//}
}

function evaluate_node(id: string, nodes, indent: string) {
	info_msg(indent+"evaluate_node: id="+id);

	if (id.split('.').shift()=="Math")
		return (Function("return "+id+";"))();

	if (!(id in nodes)) 
		nodes[id] = create_on_the_fly_node(id, indent+" ");
	
	var node = nodes[id];
	var isNulNode = node===undefined || node==null;
	var isValNode = node && node.fct===undefined && node.rng===undefined;
	var isRngNode = node && node.fct===undefined && node.rng!==undefined;
	var isFctNode = node && node.fct!==undefined && node.rng===undefined;

	if (isNulNode) 
		return null;
	else if (isValNode) 
		return node.value;
	else if (isRngNode)
		return evaluate_range(id, nodes, indent+" ");
	else if (isFctNode)
		return evaluate_depth_first(id, nodes, indent+" ");
	else 
		throw new Error(id+" unknown node type. "+JSON.stringify(node));
}

function get_range_dependencies(id:string, nodes) {
	var rng = nodes[id].rng;

	if (rng.reference_to_named_range) 
		throw new Error("get_range_dependencies called on '"+id+"' named range (should only be called on A1:B2 style addresses)");

	// build list of dependencies
	var deps = nodes[id].deps;
	if (!deps) {
		deps = {};
		for (var n in nodes) {
			var node = nodes[n];
			if (!node.target)
				continue;
			
			var inter  = ranges_intersection(rng, node.target);
			if (inter)
				deps[n] = inter;
		}
	}
	return deps;
}

function ranges_intersection(rng1:rp.RangeAddress, rng2:rp.RangeAddress) : rp.RangeAddress {
	if (rng1.reference_to_named_range || rng2.reference_to_named_range) { 
		warn_msg("ranges_intersection called on named range (should only be called on A1:B2 style addresses)");
		return undefined;
	}

	var beg1 = rng1;
	var end1 = rng1.end || rng1;
	
	var beg2 = rng2;
	var end2 = rng2.end || rng2;

	var inter  : rp.RangeAddress = {
		row : Math.max(beg1.row, beg2.row),
		col : Math.max(beg1.col, beg2.col) };
	var inter_end = {
		row : Math.min(end1.row, end2.row),
		col : Math.min(end1.col, end2.col) };
	
	if ( inter.row>inter_end.row || inter.col>inter_end.col)
		// no interection 
		return undefined;

	if (inter.row==inter_end.row || inter.col==inter_end.col)
		// single cell intersection
		return inter;

	inter.end = inter_end;
	return inter;
}

function evaluate_range(id: string, nodes, indent: string) {
	info_msg(indent+"evaluate_range: id="+id);

	// cached results
	if (nodes[id].value) {
		info_msg(indent+"evaluate_range: "+id+" value (cached) is "+JSON .stringify(res));
		return nodes[id].value;
	}

	var deps = get_range_dependencies(id, nodes);
	nodes[id].deps = deps;

	// evaluate dependencies
	for (var n in deps) 
		evaluate_node(n, nodes, indent+" ");

	// assemble range 
	var rng = nodes[id].rng;
	var beg1 = rng;
	var end1 = rng.end || rng;
	var nb_row = end1.row - beg1.row + 1;
	var nb_col = end1.col - beg1.col + 1;

	// create an empty range
	var res = new Array(nb_row);
	for (var i=0; i<nb_row; i++) {
		res[i] = new Array(nb_col);
		for (var j=0; j<nb_col; j++) 
			res[i][j] = null;
	}
		
	for (var n in nodes[id].deps) {
		// !! compatible with cells containing arrays ??
		var tmp = to_array_2d(evaluate_node(n, nodes, indent+" "));
		info_msg(indent+"evaluate_range: "+id+" dependency "+n+" value is "+JSON.stringify(tmp));
		
		var tmp_row = nodes[n].target.row;
		var tmp_col = nodes[n].target.col;

		var inter : rp.RangeAddress = nodes[id].deps[n];
		info_msg(indent+"evaluate_range: "+id+" intersection with "+n+" is "+JSON .stringify(inter));
		
		if (!inter.end) inter.end = { row: inter.row, col: inter.col };
		for (var i=inter.row; i<=inter.end.row; i++) 
			for (var j=inter.col; j<=inter.end.col; j++) 
				res[i-rng.row][j-rng.col] = tmp[i-tmp_row][j-tmp_col];

	}

	// if asked for a scalar, then return a scalar
	if (!rng.end && res.length==1 && res[0].length==1)
		res = res[0][0];

	// cache results 
	nodes[id].value = res;

	info_msg(indent+"evaluate_range: "+id+" value (calculated) is "+JSON .stringify(res));
	return res;
}

function evaluate_depth_first(id: string, nodes, indent: string) {
	info_msg(indent+"evaluate_depth_first: id="+id);
	var node = nodes[id];

	// cached results
	if (node.value) {
		info_msg(indent+"evaluate_depth_first: "+id+" value (cached) is "+JSON .stringify(node.value));
		return node.value;
	}

	// evaluate arguments first
	var args = []; 
	var arg_map = {}; // for info only
	if (node.args) {
		args = new Array(node.args.length);
		for (var a=0; a<node.args.length; a++) {
			args[a] = evaluate_node(node.args[a], nodes, indent+" ")
			arg_map[node.args[a]] = args[a];
		}
	}
	
	info_msg(indent+"evaluate_depth_first: "+id+" evaluating '"+node.expr+"' with args "+JSON.stringify(arg_map));
	node.t0 = new Date();
	try {
		var res = node.fct.apply(this, args);
		// cache reults
		node.value = res;
		node.error = undefined;
	} catch(e) {
		node.value = undefined;
		node.error = ""+e;
	}
	node.t1 = new Date();
	
	//try {
	info_msg(indent+"evaluate_depth_first: "+id+" value (calculated) is "+JSON.stringify(res));
	//} catch (e) {	info_msg(indent+"evaluate_depth_first: "+id+" value (calculated) is "+res); }


	return res;
}

function move_range(rng, i, j) {
	var tmp = JSON.parse(JSON.stringify(rng));
	if (tmp.row!==undefined && !tmp.abs_row) tmp.row += i;
	if (tmp.col!==undefined && !tmp.abs_col) tmp.col += j;
	if (tmp.end)
		tmp.end = move_range(tmp.end, i, j)
	return tmp;
}

function to_array_2d(v) {
	if (!Array.isArray(v))
		return to_array_2d([v]);
	if (v.length==0 || !Array.isArray(v[0]))
		return [v];
	return v;		
}

function info_msg(msg) {
	console.log(msg);
}
function warn_msg(msg) {
	console.warn(msg);
}

// Nodejs stuff
if (typeof module!="undefined") {
	var fs = require("fs");
	var parser = require("./excel_formula_parse");
	var range_parser = require("./excel_range_parse");
	var xlfx = require("./excel_formula_transform");
	var jsfx = require("./js_formula_transform");
	var global_scope = require("./global_scope");
	var lexer = require("./lexer");
	//var moment = require('moment');
	
	// run tests if this file is called directly
	if (require.main === module)
		spreadsheet_exec_test();

	global["spreadsheet_exec_test"] = spreadsheet_exec_test;
}
