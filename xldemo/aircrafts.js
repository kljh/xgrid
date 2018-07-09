
function XlValue(rng) {
	if (rng instanceof Promise) 
		return rng.then(XlValue);
	
	var n=rng.length, 
		m=rng[0].length;
	for (var i=0; i<n; i++) {
		for (var j=0; j<m; j++) {
			try {
				var tmp = JSON.parse(rng[i][j]);
				if (typeof tmp == "number" || typeof tmp == "boolean")
					rng[i][j] = tmp;
			} catch(e) {}
		}
	}
	return rng;
}

async function XlReadTabulatedFile(path) {
	// quick and dirty implementatoin
	return $.get(path)
		.then(data => data.split("\n").map(line => line.split("\t")));
}

function AVERAGE(rng) {
	var n=rng.length, 
		m=rng[0].length;
	var sum=0, count=0;
	for (var i=0; i<n; i++) {
		for (var j=0; j<m; j++) {
			sum += rng[i][j];
			count++
		}
	}
	var avg = sum/count;
	return avg;
}

var COVARIANCE = { S: sample_covariance };
function sample_covariance(va, vb) {
	// quick and dirty implementatoin
	var n=va.length, 
		m=va[0].length;
	var ab=0, a=0, b=0;
	for (var i=0; i<n; i++) {
		for (var j=0; j<m; j++) {
			ab += va[i][j] * vb[i][j];
			a += va[i][j];
			b += vb[i][j];
		}
	}
	ab /= n;
	a /= n;
	b /= n;
	var cov = ab - a*b;
	return cov;
}

var STDEV = { S: sample_stddev };
function sample_stddev(rng) {
	var v = sample_covariance(rng, rng);
	var s = Math.sqrt(v);
	return s;
}
function TRANSPOSE(rng) {
	return trsp(rng);
}

function ABS(arg) {
	if (!Array.isArray(arg))
		// scalar
		return Math.abs(arg);
	
	var rng = deep_copy(arg);
	var n = rng.length,
		m = rng[0].length;
	for (var i=0; i<n; i++)
		for (var j=0; j<m; j++)
			rng[i][j] = Math.abs(rng[i][j]);
	return rng;
}

function MAX(arg, arg2) {
	if (!Array.isArray(arg) || arg2!==undefined)
		return Math.abs(arg);
	
	var res = Math.max.apply(this, arg.map(row => Math.max.apply(this, row)));
	return res;
}

function MIN(arg, arg2) {
	if (!Array.isArray(arg) || arg2!==undefined)
		return Math.abs(arg);
	
	var res = Math.min.apply(this, arg.map(row => Math.min.apply(this, row)));
	return res;
}

function SUM(rng) {
	/*if (rng2) throw new Error("SUM only implemented to handle a single argument")
	
	var n = rng1.length,
		m = rng1[0].length;
	var sum = 0.0;
	for (var i=0; i<n; i++)
		for (var j=0; j<m; j++)
			sum += rng1[i][j];
	return sum;
	*/

	// call with multiple args
	if (arguments.length>1)
		return SUM(Array.from(arguments));
	
	// array
	if (Array.isArray(rng) || rng.is_range_view) {
		var n = rng.length;
		var sum = 0.0;
		for (var i=0; i<n; i++) 
			sum += rng[i];
		return sum;
	}

	// scalar
	return rng;
}


function SUMPRODUCT(rng1, rng2, rng3) {
	if (!Array.isArray(rng1))
		return "FIRST ARG IS NOT CALCULATION BECAUSE IT IS : A12:D12 - K$8:N$8"; // !!
		 
	var n = rng1.length,
		m = rng1[0].length;
	var sum = 0.0;
	for (var i=0; i<n; i++)
		for (var j=0; j<m; j++)
			sum += rng1[i][j] * rng2[i][j] * (rng3?rng3[i][j]:1.0);
	return sum;
}

function INDEX(rng, row, col) {
	if (row!==undefined && col!==undefined)
		return rng[row-1][col-1];
	if (row!==undefined)
		return [rng[row-1]];
	if (col!==undefined)
		return rng.map(row => row[col-1]);
	return rng;
	
}

function MATCH(val, rng, cmp) {
	if (cmp!=0)	throw new Error("MATCH only implemented to handle exact comparison");

	var n = rng.length;
	if (typeof val == "number") {
		for (var i=0; i<n; i++) 
			if (val===rng[i][0])
				return i+1;		
	} else {
		var ibest = 0, dbest = Math.abs(val-rng[0][0]);
		for (var i=0; i<n; i++) 
			if (Math.abs(val-rng[i][0]))
				ibest = i+1;
		if (dbest<1e-14 && dbest<1e-12*rng[ibest-1][0])
			return ibest;
	}

	return sum;
	
}
function XlSvd(mtx) {
	var tmp = diag_jacobi(mtx);
	return { W: tmp.diag, U: tmp.P, V: trsp(tmp.P) };
	return "SVD (or simply diagonalisation of sym def positive matrix) TO DO"
}

function XlGet(obj, path) { 
	//var obj = wsh.objects[tkr.split(":").shift()];
	
	if (obj && obj[path])
		return obj[path];
			
	return "XlGet error: "+path+" not found in "+JSON.stringify(obj);
}
