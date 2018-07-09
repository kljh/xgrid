/* 
	Some formula require some context to be able to evaluate:
	ROW(), COLUMN(), ADDRESS()
*/

var formula_this = {};

var ROW = (function (opt_ref) {
	if (opt_ref && !opt_ref.is_range_view) 
		throw new Error("ROW([ref]) called with a non ref argument. ref="+JSON.stringify(ref,null,4)+" this="+JSON.stringify(this.caller,null,4));
	
	var rng = (opt_ref && opt_ref.rng) || (this.node && this.node.target);
	return rng.row; // refer to doc where ref and target are range of cells
}).bind(formula_this);

var COLUMN = (function (opt_ref) {
	if (opt_ref && !opt_ref.is_range_view) 
		throw new Error("COLUMN([ref]) called with a non ref argument. ref="+JSON.stringify(ref,null,4)+" this="+JSON.stringify(this.caller,null,4));
	
	var ref = opt_ref || this;
	return ref.rng.col; // refer to doc where ref and target are range of cells
}).bind(formula_this);

var CELL = (function (info_type, opt_ref) {
	if (opt_ref && !opt_ref.is_range_view) 
		throw new Error("CELL(info, [ref]) called with a non ref argument. ref="+JSON.stringify(ref,null,4)+" this="+JSON.stringify(this.caller,null,4));

	var ref = opt_ref || this;
	switch(info_type) {
		case "address":
			var abs_ref = clone(ref);
			abs_ref.rng.abs_row = true;
			abs_ref.rng.abs_col = true;
			return xl_range_parse.stringify_range(abs_ref.rng);
		case "filename":
			return "folder_path\[file_name]Sheet1";
		default:
			throw new Error("CELL: unsupported '"+info_type+"' info_type.");
	}
}).bind(formula_this);

function OFFSET(rng_view_data, offset_row, offset_col, num_row, num_col) {
	if (rng_view_data.is_range_view) {
		// a range reference
		var rng = rng_view_data.rng;
		var caller = rng_view_data.caller;
		var new_rng = { 
			caller: caller,
			rng : { 
				row: rng.row + (offset_row||0), 
				col: rng.col + (offset_col||0) 
			}};
		
		var end = rng.end || rng;
		if (end || num_row || num_col) 
			new_rng.rng.end = {
				row: num_row ? rng.row + num_row - 1 : end.row + (offset_row||0), 
				col: num_col ? rng.col + num_col - 1 : end.col + (offset_col||0) };
				
		return range_view(new_rng);
	} else {
		// a data  array
		var rng = rng_view_data;
		var res = Array.isArray(rng) ? rng : [[rng]];
		if (offset_row)
			res = res.slice(offset_row);
		if (offset_col)
			res = res.map(row = row.slice(offset_col));
		return res;
	}
}

function clone(o) {
	return JSON.parse(JSON.stringify(o));
}