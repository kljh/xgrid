
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

function op_scalar(f, a, b) {
	// resolve arguments and apply scalar operation

	if (Promise.prototype.isPrototypeOf(a)
		|| Promise.prototype.isPrototypeOf(b))

		return Promise.all([a,b])
			.then((args) => op_scalar(f, args[0], args[1]));
	
	if (a.is_range_view)
		a = a.valueOf();
	else
		a = top_left(a);
	
	if (b.is_range_view)
		b = b.valueOf();
	else
		b = top_left(b);

	return f(a, b);
}

function op_array(f, a, b) {

	if (Promise.prototype.isPrototypeOf(a)
		|| Promise.prototype.isPrototypeOf(b))

		return Promise.all([a,b])
			.then((args) => op_array(f, args[0], args[1]));

	var left_array = Array.isArray(a) || a.is_range_view,
		right_array = Array.isArray(b) || b.is_range_view;
	var n: number, res;
	if (!left_array && !right_array) {
		// scalar-scalar operation
		return f(a,b);
	} else if (!left_array && right_array) {
		// scalar-range operation
		n = b.length;
		res = new Array(n);
		for (var i=0; i<n; i++) 
			res[i] = op_array(f, a, b[i]);
		return res;
	} else if (left_array && !right_array) {
		// range-scalar operation
		n = a.length;
		res = new Array(n);
		for (var i=0; i<n; i++)
			res[i] = op_array(f, a[i], b);
		return res;
	} else if (left_array && right_array) {
		// range-range operation
		var na=a.length, xna=1,
			nb=b.length, xnb=1;
		if (na!=nb) { 
			if (na==1) xna=0;
			else if (nb==1) xnb=0;
			else new Error("op_array: array-array operation nb_rows mismatch: "+na+" vs "+nb); }
		n = Math.max(na, nb);
		res = new Array(n);
		for (var i=0; i<n; i++)
			res[i] = op_array(f, a[i*xna], b[i*xnb]);
		return res;
	} else {
		throw new Error("op_array: operation not implemented on this type combination");
	}	
}


function top_left(a) {
	if (Array.isArray(a))
		return top_left(a[0]);
	else 
		return a;
}

export var op = {
	add: function add(a,b)  { return op_scalar((a,b) => a+b, a, b); }, 
	sub: function sub(a,b)  { return op_scalar((a,b) => a-b, a, b); },
	mult: function mult(a,b) { return op_scalar((a,b) => a*b, a, b); },
	div: function div(a,b)  { return op_scalar((a,b) => a/b, a, b); },
	pow: function pow(a,b)  { return op_scalar((a,b) => a^b, a, b); },
	eq: function eq(a,b)   { return op_scalar((a,b) => a===b, a, b); },
	neq: function neq(a,b)  { return op_scalar((a,b) => a!==b, a, b); },
	gt: function gt(a,b)   { return op_scalar((a,b) => a>b, a, b); },
	lt: function lt(a,b)   { return op_scalar((a,b) => a<b, a, b); },
	gte: function gte(a,b)  { return op_scalar((a,b) => a>=b, a, b); },
	lte: function lte(a,b)  { return op_scalar((a,b) => a<=b, a, b); },
};

export var ops = {
	add: function add(a,b)  { return op_array((a,b) => a+b, a, b); }, 
	sub: function sub(a,b)  { return op_array((a,b) => a-b, a, b); },
	mult: function mult(a,b) { return op_array((a,b) => a*b, a, b); },
	div: function div(a,b)  { return op_array((a,b) => a/b, a, b); },
	pow: function pow(a,b)  { return op_array((a,b) => a^b, a, b); },
	eq: function eq(a,b)   { return op_array((a,b) => a===b, a, b); },
	neq: function neq(a,b)  { return op_array((a,b) => a!==b, a, b); },
	gt: function gt(a,b)   { return op_array((a,b) => a>b, a, b); },
	lt: function lt(a,b)   { return op_array((a,b) => a<b, a, b); },
	gte: function gte(a,b)  { return op_array((a,b) => a>=b, a, b); },
	lte: function lte(a,b)  { return op_array((a,b) => a<=b, a, b); },
};
