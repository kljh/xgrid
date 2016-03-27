


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

function spreadsheet_exec_test() {
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
		info_msg("sheet "+JSON.stringify(sheet, null, 4));

		try {
			var node_values = sheet_exec(sheet, { xjs: true });

			var tmp = filepath.split('/');
			tmp[tmp.length-2] = "ref";
			var filepath_out = tmp.join('/');
			tmp.pop();
			var folder_out = tmp.join('/');

			if (!fs.existsSync(folder_out)) fs.mkdirSync(folder_out);
			fs.writeFileSync(filepath_out, JSON.stringify({input: node_values},null,4));

		} catch (e) {
			throw new Error("evaluating "+filepath+"\n"+(e.stack||e)+"\n");
		}
	}
}

function sheet_exec(sheet, prms) {
	var eval = {};
	var node_values = {};
	for (var r in sheet) {
		var val = sheet[r];
		if (val==="") continue;

		var target = range_parser.parse_range(r);
		
		var isFormula = val.substr && val[0]=="=";
		var isFormulaArray = val.substr && val[0]=="{" && val[1]=="=" && val[val.length-1]=="}";
		if (isFormula || isFormulaArray) {
			try {
				var formula = isFormula ? val.substr(1) : val.substr(2, val.length-3);

				var vars =  {}, fcts= {};
				var expr = xlexpr.parse_and_transfrom(formula,vars,fcts,prms);
				
				var ids = []; 
				var args = []; 
				for (var addr in vars) {
					ids.push(vars[addr]);
					args.push(addr);
				}

				//var func = new Function(ids, "try { var res="+expr+"; return res; } catch(e) { throw new Error('"+r+" '+(e.stack||e)+'\\n'); }");
				//var func = new Function(ids, "try { var res="+expr+"; return res; } catch(e) { throw new Error('"+r+" '+(e)+'\\n'); }");
				//var func = Function(ids, "try { var res="+expr+"; return res; } catch(e) { throw new Error('"+r+" '+(e)+'\\n'); }");
				var func = new Function(ids, "var res="+expr+"; return res;");
				info_msg("expr: "+expr);

				if (target.end && !isFormulaArray) {
					for (var i=target.row; i<=target.end.row; i++)
						for (var j=target.col; j<=target.end.col; j++) {

							var tmp_args = new Array(args.length);
							for (var a=0; a<args.length; a++) 
								tmp_args[a] = range_parser.stringify_range(move_range(range_parser.parse_range(args[a]), i-target.row, j-target.col));

							eval[r+"_"+i+"_"+j] = {
								target: { row: i, col: j },
								expr: expr,
								fct: func,
								args: tmp_args
								};
						}
				} else {
					eval[r] = {
						target: target,
						expr: expr,
						fct: func,
						args: args
						};
				}
			} catch (e) {
				eval[r] = {
					target: target,
					expr: val,
					fct: function(x) { return function() { throw new Error("ill formated expression"+x); }; },
					error: e.stack || (""+e)
					};
			}
		} else {

			eval[r] = {
				target: target,
				value: val,
				};
		}
	}

	info_msg("evaluation: "+JSON.stringify(eval, null, 4));

	//return ;

	var final_ranges_to_evaluate = [ "G32", "G30", "H32" ];
	var final_ranges_to_evaluate = Object.keys(eval);
	for (var r=0; r<final_ranges_to_evaluate.length; r++) {
		var rng = final_ranges_to_evaluate[r];
		var res = evaluate_node(rng, eval);
		node_values[rng] = res;
		info_msg(rng+" value is "+JSON.stringify(res));
	}

	return node_values;
}

function create_on_the_fly_node(id) {
	info_msg(arguments.callee.name+": id: "+id);
	var rng = range_parser.parse_range(id);
	///if (rng.end) {
		return { rng: rng };
	//} else {
	//	throw new Error(arguments.callee.name+": "+id+" is not a multi-cell range");
	//}
}

function evaluate_node(id, nodes) {
	info_msg(arguments.callee.name+": "+id);

	if (id.split('.').shift()=="Math")
		return (Function("return "+id+";"))();

	if (!(id in nodes)) 
		nodes[id] = create_on_the_fly_node(id);
	
	var node = nodes[id];
	var isValNode = node.fct===undefined && node.rng===undefined;
	var isRngNode = node.fct===undefined && node.rng!==undefined;
	var isFctNode = node.fct!==undefined && node.rng===undefined;

	if (isValNode) 
		return node.value;
	else if (isRngNode)
		return evaluate_range(id, nodes);
	else if (isFctNode)
		return evaluate_depth_first(id, nodes);
	else 
		throw new Error(arguments.callee.name+": "+id+" unknown node type. "+JSON.stringify(node));
}

function evaluate_range(id, nodes) {
	info_msg(arguments.callee.name+": "+id);

	// cached results
	if (nodes[id].value)
		return nodes[id].value;

	var rng = nodes[id].rng;
	var beg1 = rng;
	var end1 = rng.end || rng;

	// build list of dependencies
	var deps = nodes[id].deps;
	if (!deps) {
		deps = [];
		for (var n in nodes) {
			var node = nodes[n];
			if (!node.target)
				continue;
			
			var beg2 = node.target;
			var end2 = node.target.end || node.target;

			var inter = {
				row : Math.max(beg1.row, beg2.row),
				col : Math.max(beg1.col, beg2.col),
				end :{
					row : Math.min(end1.row, end2.row),
					col : Math.min(end1.col, end2.col),
				}};

			if ( inter.row>inter.end.row || inter.col>inter.end.col)
				inter = undefined;
			else
				deps[n] = inter;
		}
		nodes[id].deps = deps;
	}

	// evaluate dependencies
	//for (var n in deps) 
	//	evaluate_node(n, nodes);

	// assemble range 
	var nb_row = end1.row - beg1.row + 1;
	var nb_col = end1.col - beg1.col + 1;
	var res = new Array(nb_row);
	for (var i=0; i<nb_row; i++) {
		res[i] = new Array(nb_col);
	}
		
	for (var n in deps) {
		var tmp = to_array_2d(evaluate_node(n, nodes));
		info_msg(arguments.callee.name+": "+id+" subrange "+n+" value is "+JSON.stringify(tmp));
		
		var tmp_row = nodes[n].target.row;
		var tmp_col = nodes[n].target.col;

		var inter = deps[n];
		info_msg(arguments.callee.name+": inter "+JSON .stringify(inter));
		
		for (var i=inter.row; i<=inter.end.row; i++) 
			for (var j=inter.col; j<=inter.end.col; j++) 
				res[i-rng.row][j-rng.col] = tmp[i-tmp_row][j-tmp_col];

		// if asked for a scalar, then return a scalar
		if (!rng.end && res.length==1 && res[0].length==1)
			res = res[0][0];
	}

	// cache reuls 
	nodes[id].value = res;

	info_msg(arguments.callee.name+": "+id+" value is "+JSON .stringify(res));
	return res;
}

function evaluate_depth_first(id, nodes) {
	info_msg(arguments.callee.name+": "+id);
	var node = nodes[id];

	// cached results
	if (node.value) {
		info_msg(arguments.callee.name+": "+id+" cached "+JSON .stringify(node.value))
		return node.value;
	}

	// evaluate arguments first
	var args = []; 
	var arg_map = {}; // for info only
	if (node.args) {
		args = new Array(node.args.length);
		for (var a=0; a<node.args.length; a++) {
			args[a] = evaluate_node(node.args[a], nodes)
			arg_map[node.args[a]] = args[a];
		}
	}
	
	info_msg(arguments.callee.name+": evaluating "+id+": '"+node.expr+"' with args "+JSON .stringify(arg_map));
	var res = node.fct.apply(this, args);

	// cache reuls 
	node.value = res;

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

// Nodejs stuff
if (typeof module!="undefined") {
	var fs = require("fs");
	var parser = require("./excel_formula_parse");
	var range_parser = require("./excel_range_parse");
	var xlexpr = require("./excel_formula_transform");
	var global_scope = require("./global_scope");

	// run tests if this file is called directly
	//if (require.main === module)
		spreadsheet_exec_test();

}
