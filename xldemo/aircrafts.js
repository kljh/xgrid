
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
			rng = Math.abs(rng[i][j]);
	return rng;
}

function SUM(rng1, rng2, rng3) {
	if (rng2) throw new Error("SUM only implemented to handle a single argument")
	
	var n = rng1.length,
		m = rng1[0].length;
	var sum = 0.0;
	for (var i=0; i<n; i++)
		for (var j=0; j<m; j++)
			sum += rng1[i][j];
	return sum;
}


function SUMPRODUCT(rng1, rng2, rng3) {
	if (!Array.isArray(rng1))
		return -1; // !!
		 
	var n = rng1.length,
		m = rng1[0].length;
	var sum = 0.0;
	for (var i=0; i<n; i++)
		for (var j=0; j<m; j++)
			sum += rng1[i][j] * rng2[i][j] * rng2[i][j];
	return sum;
}

function XlSvd(mtx) {
	var tmp = diag_jacobi(mtx);
	return { W: tmp.diag, U: tmp.P, V: tmp.P };
	return "SVD (or simply diagonalisation of sym def positive matrix) TO DO"
}

function XlGet(obj, path) { 
	//var obj = wsh.objects[tkr.split(":").shift()];
	
	if (obj && obj[path])
		return obj[path];
		
	if (path=="U" || path=="V") 
		return [ 
			[ 1,  1,  0,  0 ],
			[ 1, -1, -1,  0 ],
			[ 1,  1,  1,  1 ],
			[ 1,  1,  1, -1 ]];
			
	return "XlGet";
}
