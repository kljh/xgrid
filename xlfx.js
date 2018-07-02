
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

function CELL(info_type, reference) {
	switch(info_type) {
		case "address":
			if (!reference || !reference.rng)
				throw new Error("CELL('address') called without reference range.");
			return xl_range_parse.stringify_range(reference.rng);
		case "filename":
			return "folder_path\[file_name]Sheet1";
		default:
			throw new Error("CELL: unsupported '"+info_type+"' info_type.");
	}
}