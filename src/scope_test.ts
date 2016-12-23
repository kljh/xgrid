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

function info_msg(msg) {
	console.log(msg);
}
scope_test();