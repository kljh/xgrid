
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
	} else if (!Array.isArray(a) && Array.isArray(b) && Array.isArray(b[0])) {
		// scalar-range operation
		var n=b.length, m=b[0].length;
		var res = new Array(n);
		for (var i=0; i<n; i++) {
			res[i] = new Array(m);
			for (var j=0; j<m; j++) 
				res[i][j] = f(a, b[i][j]);
		}
		return res;
	} else if (Array.isArray(a) && Array.isArray(a[0]) && !Array.isArray(b)) {
		// range-scalar operation
		var n=a.length, m=a[0].length;
		var res = new Array(n);
		for (var i=0; i<n; i++) {
			res[i] = new Array(m);
			for (var j=0; j<m; j++) 
				res[i][j] = f(a[i][j], b);
		}
		return res;
	} else if (Array.isArray(a) && Array.isArray(a[0]) && Array.isArray(b) && Array.isArray(b[0])) {
		// range-range operation
		var na=a.length, ma=a[0].length, xna=1, xma=1,
			nb=b.length, mb=b[0].length, xnb=1, xmb=1;
		if (na!=nb) { 
			if (na==1) xna=0;
			else if (nb==1) xnb=0;
			else new Error("op_gen: array-array operation nb_rows mismatch: "+na+" vs "+nb); }
		if (ma!=mb) { 
			if (ma==1) xma=0;
			else if (mb==1) xmb=0;
			else new Error("op_gen: array-array operation nb_rows mismatch: "+ma+" vs "+mb); }
		var n = Math.max(na, nb), m = Math.max(ma, mb);
		var res = new Array(n);
		for (var i=0; i<n; i++) {
			res[i] = new Array(m);
			for (var j=0; j<m; j++) 
				res[i][j] = f(a[i*xna][j*xma], b[i*xnb][j*xmb]);
		}
		return res;
	} else {
		throw new Error("op_gen: operation not implemented on this type combination");
	}	
}

export var op = {
	add: function op_add(a,b)  { return op_gen((a,b) => a+b, a, b); }, 
	sub: function op_sub(a,b)  { return op_gen((a,b) => a-b, a, b); },
	mult: function op_mult(a,b) { return op_gen((a,b) => a*b, a, b); },
	div: function op_div(a,b)  { return op_gen((a,b) => a/b, a, b); },
	pow: function op_pow(a,b)  { return op_gen((a,b) => a^b, a, b); },
	eq: function op_eq(a,b)   { return op_gen((a,b) => a===b, a, b); },
	neq: function op_neq(a,b)  { return op_gen((a,b) => a!==b, a, b); },
	gt: function op_gt(a,b)   { return op_gen((a,b) => a>b, a, b); },
	lt: function op_lt(a,b)   { return op_gen((a,b) => a<b, a, b); },
	gte: function op_gte(a,b)  { return op_gen((a,b) => a>=b, a, b); },
	lte: function op_lte(a,b)  { return op_gen((a,b) => a<=b, a, b); },
};
 