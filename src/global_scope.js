function sum(rng) {
	var total = 0;
	for (var i=0; i<rng.length; i++) 
		for (var j=0; j<rng[i].length; j++) 
			total += rng[i][j] || 0;
	return total;
}

function trsp(rng) {
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

// Nodejs stuff
if (typeof module!="undefined") {
	global.sum = sum;
	global.trsp = trsp;
}
