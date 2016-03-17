
function spreadsheet_scope() {
	this.A1 = 13
	this.f = function() { info_msg(this.A1); }
}
spreadsheet_scope.prototype.fun = function() { info_msg(this.A1); }
spreadsheet_scope.prototype.fun2 = function() { info_msg(A1); }

function scope_test() {
	var obj = {
		get A1 () {
			return 11;
		}
	}
	
	function f() { info_msg(this.A1); }
	f.call({ A1: 7 })
	f.call(obj)

	var ss = new spreadsheet_scope;
	ss.f();
	ss.fun();
	ss.fun2();

	ss.A2 = 17;
	ss.g = function() { info_msg(this.A2); }
	ss.g();

	ss.A3 = 23;
	ss.h = new Function("return this.A3;");
	info_msg(ss.h());
	
	var A5 = 5
	function f5() { info_msg(A5); }
	f5();

	var A6 = 6
	f6 = function() { info_msg(A6); }
	f6();
	
	var A7 = 7
	f7a = new Function("info_msg(A7);");
	f7 = f7a.bind(this);
	f7();
	
}


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


	var path = require('path');
	var filepath = path.resolve(__dirname+'/../work/music_theory.xjson');
	info_msg("Executing "+filepath);

	var txt = fs.readFileSync(filepath, 'utf8');
	var data = JSON.parse(txt);
	var input = data.input;
	info_msg("input "+JSON.stringify(input, null, 4));

	var eval = {};
	for (var r in input) {
		if (input[r].substr && input[r][0]=="=") {
			if (input[r]==="") continue;

			append_depth_first_eval_function(r, input[r], eval);
			

		} else {
			eval[r] = {
				f : new Function("return "+JSON.stringify(input[r])) };
		}
	}

	info_msg("evaluation: "+JSON.stringify(eval, null, 4));

	var final_ranges_to_evaluate = [ "G32", "G30", "H32" ];
	for (var r=0; r<final_ranges_to_evaluate.length; r++) {
		var rng = final_ranges_to_evaluate[r];
		info_msg(rng+" value is "+eval[rng].f());
	}
}

function append_depth_first_eval_function(input_id, input_expr, eval) {
	info_msg("input_id: "+input_id);
	info_msg("input_expr: "+JSON.stringify(input_expr));

	var vars =  {}, fcts= {};
	var expr = xlexpr.parse_and_transfrom(input_expr,vars,fcts);
	var args = []; for (var addr in vars) args.push(vars[addr]);
	var func = new Function(args, "var res="+expr+"; return res;");
	info_msg("expr: "+expr);

	var f = function() {
		// evaluate arguments first
		var arg_vals = [];
		for (var addr in vars) {
			addr = addr.replace(/\$/g, "");
			if (!(addr in eval))
				throw new Error("evaluating "+input_id+": '"+expr+"', input '"+addr+"' unknown in "+Object.keys(eval));
			if (!eval[addr].f)
				throw new Error("evaluating "+input_id+": '"+expr+"', input '"+addr+"' has no menber 'f()'");
			
			var v = eval[addr].f();
			arg_vals.push(v);
		}
			
		info_msg("evaluating "+input_id+": '"+expr+"'");
		var res = func.apply(this, arg_vals);
		return res;
	}

	eval[input_id] = {
		f: f,
		expr: expr,
		vars: vars };
}

function info_msg(msg) {
	console.log(msg);
}

// Nodejs stuff
if (typeof module!="undefined") {
	var fs = require("fs");
	var parser = require("./excel_formula_parse");
	var xlexpr = require("./excel_formula_transform");

	// run tests if this file is called directly
	if (require.main === module)
		spreadsheet_exec_test();

}
