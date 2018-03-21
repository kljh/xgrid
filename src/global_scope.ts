
export function sum(rng) {
	var total = 0;
	for (var i=0; i<rng.length; i++) 
		for (var j=0; j<rng[i].length; j++) 
			total += rng[i][j] || 0;
	return total;
}

export function trsp(rng) {
	var n = rng.length;
	var m = n==0 ? 0 : rng[0].length;

	var res = new Array(m);
	for (var j=0; j<m; j++) {
		res[j] = new Array(n);
		for (var i=0; i<n; i++) 
			res[j][i] = rng[i][j];
	}
	return res;
}

function op_gen(f, a, b) {
	if (!Array.isArray(a) && !Array.isArray(b)) {
		// scalar-scalar operation
		return f(a,b);
	} else if (Array.isArray(a) && Array.isArray(b)) {
		// array-array operation
		throw new Error("op_gen: array-array operation not implemented");
	} else {
		throw new Error("op_gen: array-scalar operation not implemented");
	}	
}

/*
function op_add(a,b)  { return op_gen((a,b) => a+b, a, b); };     global["op_add"] = op_add;
function op_sub(a,b)  { return op_gen((a,b) => a-b, a, b); };     global["op_sub"] = op_sub;
function op_mult(a,b) { return op_gen((a,b) => a*b, a, b); };     global["op_mult"] = op_mult;
function op_div(a,b)  { return op_gen((a,b) => a/b, a, b); };     global["op_div"] = op_div;
function op_pow(a,b)  { return op_gen((a,b) => a^b, a, b); };     global["op_pow"] = op_pow;
function op_eq(a,b)   { return op_gen((a,b) => a===b, a, b); };   global["op_eq"] = op_eq;
function op_neq(a,b)  { return op_gen((a,b) => a!==b, a, b); };   global["op_neq"] = op_neq;
function op_gt(a,b)   { return op_gen((a,b) => a>b, a, b); };     global["op_gt"] = op_gt;
function op_lt(a,b)   { return op_gen((a,b) => a<b, a, b); };     global["op_lt"] = op_lt;
function op_gte(a,b)  { return op_gen((a,b) => a>=b, a, b); };    global["op_gte"] = op_gte;
function op_lte(a,b)  { return op_gen((a,b) => a<=b, a, b); };    global["op_lte"] = op_lte;
*/
 