
function trsp(mtx) {
	var m = mtx.length, 
		n = mtx[0].length;
	
	var res = new Array(n);
	for (var i=0; i<n; i++) {
		res[i] = new Array(m);
		for (var j=0; j<m; j++)
			res[i][j] = mtx[j][i];
	}
	return res;
}

function mmult(mtx1, mtx2) {
	var n1 = mtx1.length, 
		m1 = mtx1[0].length,
		n2 = mtx2.length, 
		m2 = mtx2[0].length;
	
	if (m1!=n2) throw new Error("mmult: inconsistent matrices size");
	
	var res = new Array(n1);
	for (var i=0; i<n1; i++) {
		res[i] = new Array(m2);
		for (var j=0; j<m2; j++) {
			var tmp = 0;
			for (var k=0; k<m1; k++) 
				tmp += mtx1[i][k] * mtx2[k][j];
			res[i][j] = tmp;
		}
	}
	return res;
}


function deep_copy(x) {
	return JSON.parse(JSON.stringify(x));
}

function identity_mtx(n) {
	var id = new Array(n).fill(null).map(row => new Array(n).fill(0));
	for (var i=0; i<n; i++) id[i][i] = 1.0;
	return id;
}
	

function diag_jacobi(A) {
	// https://cel.archives-ouvertes.fr/cel-01066570/file/AnalyseNumerique.pdf
	
	var n = A.length;
	var Bk = deep_copy(A);
	var P = identity_mtx(n);
	for (var iter=0; iter<(n*n); iter++) {
		// find greatest off-diag term
		var i0, j0, v0=-1;
		for (var i=0; i<n; i++) 
			for (var j=0; j<i; j++) 
				if (Math.abs(Bk[i][j])>v0) {
					i0 =i; j0=j; v0=Math.abs(Bk[i][j]); }
		console.log("diag_jacobi: greatest off-diag term ", v0, "in", i0, j0);
		if (v0<1e-14) break;
		
		var phi = 0.5 * Math.atan( 2*Bk[i0][j0] / ( Bk[j0][j0] - Bk[i0][i0] ));
		var cp = Math.cos(phi), sp = Math.sin(phi);
					
		var Pk = identity_mtx(n);
		Pk[i0][i0] = cp;
		Pk[j0][j0] = cp;
		Pk[i0][j0] = sp;	
		Pk[j0][i0] = -sp;
		
		Bk = mmult(trsp(Pk), mmult(Bk, Pk));
		P = mmult(P, Pk);
	}
	
	var D = Bk;
	var diag = D.map((row, i) => row[i]);
	// var checkA = mmult(P, mmult(D, trsp(P)));
	
	return  { diag: diag, D: D, P: P };
}

/*
console.log(diag_jacobi([
	[ 1,   	0.6,	 0.9, 	0.89 ],
	[ 0.6,	1, 	 0.63,	0.69 ],
	[ 0.9,	0.63,	 1,    	0.96 ],
	[ 0.89,	0.69,	 0.96,	1     ]]));
*/