//$(init_amd);
//$(init_no_amd);
//$(load_workbook);

function init_no_amd() { 
    // check modules (with no AMD/requirejs)
    console.log("acorn", acorn);
    console.log("escodegen", escodegen);
    //console.log("grid", grid);
}

function get_and_load_workbook() {
    var query = parse_query();

	xll_http_init()
	.catch(err => {
		console.warn("xll_http_init() failed. ", err);
		return  null;
	})
	.then(_ => {
		var wbk = query.wbk || "xljs.json"; 
		$.get(wbk)
		.then(load_workbook)
		.fail(function (err) { 
			console.error(err);
			alert("get_and_load_workbook error. "+(err.stack||err.message||err)); 
		});
	});
}
		
function load_workbook(wbk) {
        console.log("wbk", wbk); 
        var ref_style = xl_range_parse.ReferenceStyle[wbk.ReferenceStyle || "A1"];
        if (!ref_style) {
        	var msg = "missing WorkBook.ReferenceStyle";
        	throw new Error(msg);
        }

        var wsht = wbk.Worksheets[0];
        //for (var style of wsht.formats.styles) {
        //    $('head').append('<style type="text/css">body {margin:0;}</style>');
        //}
        var fmt = wsht.formats;
        if (!fmt) { 
            console.warn("load_workbook: no formatting in loaded workbook.")
        }

        // add CSS styles defintions
        var style_text = "";
        {
            style_text += "table { border-collapse: collapse; }\n";
            //style_text += "td { min-width: 64px; overflow: hidden; white-space: nowrap; text-overflow: ellipsis; }\n";
            style_text += "tr:first-child > td { text-align: center; }\n";
            style_text += "td:not(:first-child) { min-width: 64px; max-width: 128px; overflow: visible; white-space: nowrap; text-overflow: ellipsis; }\n";
            //style_text += "td + td:empty {  background-color: lightgray; }\n";
            //style_text += "td:not([class]):empty  {  background-color: lightgray; }\n";
            style_text += ".bold { font-weight: bold; }\n";
            style_text += ".italic { font-style: italic; }\n";
            style_text += ".center { text-align: center; }\n";
            style_text += ".justify { text-align: justify; }\n";
            style_text += ".left { text-align: left; }\n";
            style_text += ".right { text-align: right; }\n";
            style_text += ".border-top  { border-top:  1px solid gray; }\n";
            style_text += ".border-left { border-left: 1px solid gray; }\n";
            
        }
        if (wsht.formats && wsht.formats.styles)
        for (var style in wsht.formats.styles) 
            style_text += "."+style+" { "+wsht.formats.styles[style]+" }\n"
        $('head').append('<style type="text/css">\n'+style_text+'</style>');


        var formulas = wsht.formulas;

        build_table();
        
        // add CSS classes per cell
        if (wsht.formats && wsht.formats.cells)
        for (var rng_id in wsht.formats.cells) {
            var css_classes = wsht.formats.cells[rng_id];
            var rng = xl_range_parse.parse_range(rng_id, ref_style);
            for (var cell of range_cell_iterator(rng)) {
                var cell_id = xl_range_parse.stringify_range(cell);
                $('#'+cell_id).addClass(css_classes);
                if (css_classes.indexOf("input")!=-1)
                    $('#'+cell_id).attr("contenteditable", true);   
            }
        }
        
        // add CSS borders per cell
        if (wsht.formats && wsht.formats.borders)
        for (var rng_id in wsht.formats.borders) {
            var rng = xl_range_parse.parse_range(rng_id, ref_style);
            for (var cell of range_cell_iterator(rng)) {
                var cell_id = xl_range_parse.stringify_range(cell);
                $('#'+cell_id).addClass(wsht.formats.borders[rng_id]);
            }
        }
        
        // add values
        if (wsht.values) {
            var n = wsht.values.length, m = n==0 ? 0 : wsht.values[0].length;
            for (var i=0; i<n; i++) {
                for (var j=0; j<m; j++) {
                    var v = wsht.values[i][j];
                    if (v!==undefined && v!==null) {
                        cell_id = xl_range_parse.stringify_range({ row: i, col: j });
                        $('#'+cell_id).text(v);
                    }
                }
            }
        }
        // add formulas
        for (var rng_id in wsht.formulas) {
            var rng = xl_range_parse.parse_range(rng_id, ref_style);
            var val = wsht.formulas[rng_id];

            if (val[0]=="=" || (val[0]=="{" && val[1]=="=")) {
                add_formula(rng, val, ref_style);
                continue;
            }

            // set values
            set_range_values(rng, val);

            // display value
            for (var cell of range_cell_iterator(rng)) {
                var cell_id = xl_range_parse.stringify_range(cell);
                $('#'+cell_id).text(val);
            }
        }

        // add validation
        if (wsht.formats && wsht.formats.number_format)
        for (var rng_id in wsht.formats.number_format) {
            var fmt =  wsht.formats.number_format[rng_id];
            if (fmt.indexOf("xlValidateList")!=-1) {
                var lst = fmt.match(/{ xlValidateList: 1, showError: 1, list: '([^']*)' }/);
                var items = lst[1].split(',');
            }
        }

		// check function usage
		console.log("function_usage", function_usage);
		var global_object = window || global;
		for (var fct_name in function_usage) {
			var fct_obj = Function('if (typeof '+fct_name+' != "undefined") return '+fct_name+';')();
			if (typeof fct_obj == "function")
				continue;
			else 
				global_object[fct_name] = xll_http_stub(fct_name);
		}


        pop_all_formulas(wbk)
        .then(_ => render_graphs(wbk))
        .then(_ => console.log("all formulas popped"))
        .catch(err => console.warn(""+(err.stack||err)));
}

function build_table() {
    var n = 120, m = 32;
    var table = $('<table>').attr("id", "tab0");
    for (var i=0; i<n; i++) {
        var row = $('<tr>');
        for (var j=0; j<m; j++) {
            var cell_id="", cell_txt="";
            if (i==0 && j==0) {
                cell_id = "top_left_header_cell";
            } else if (i==0) {
                // header row
                cell_txt = xl_range_parse.stringify_range({ row: 0, col: j }).replace("1", "") + " ("+j+")";
            } else if (j==0) {
                cell_txt = ""+i;
            } else {
                cell_id = xl_range_parse.stringify_range({ row: i, col: j });
            }
            var cell = row.append('<td' + (cell_id?' id="'+cell_id+'"':'') +'>'+cell_txt+'</td>');
        }
        table.append(row);
    }
    $('body').append(table)
}

function range_cell_iterator(rng) {
    var resIterable = { rng: rng };

    // make res an iterable 
    resIterable[Symbol.iterator] = function* () {
        if (!this.rng.end) {
            yield { row: this.rng.row, col: this.rng.col };
        } else {
            for (var i=this.rng.row; i<=this.rng.end.row; i++)
                for (var j=this.rng.col; j<=this.rng.end.col; j++)
                    yield { row: i, col: j };
        }
    }
    return resIterable;
}

// --------------------------
// formulas

var formula_queue = [];
var formula_queue_next = [];

var function_usage = {};

//var async_fct = new AsyncFunction("return new Promise("+formula+");");
function add_formula(rng, formula_text, ref_style) {
	var is_array_formula = formula_text[0]=="{";

    var args = {}, vars_out = {}, fcts_out = function_usage;
    var formula0 = xl_parse_and_transfrom(formula_text, args);
    var formula = formula0;
    try {
        formula = js_parse_and_transfrom(formula0, vars_out, fcts_out, { is_array_formula: is_array_formula });
    } catch (e) {
        console.warn("acorn("+formula0+"): "+e);
        throw e;
    }
    var array_formula = formula_text[0]=="{";
                    
    var fct_arg_names = Object.values(args);
    var fct_body = "return "+formula+";";
    var fct = new Function(fct_arg_names, fct_body);

    var fct_arg_ranges = Object.keys(args).map(name => xl_range_parse.parse_range(name, ref_style, rng));
	if (!is_array_formula) {
		fct_arg_ranges.forEach(arg => arg.caller = rng);
	} 

    update_range_values(rng, undefined, -1); 
            
    if (array_formula||!rng.end) {
        formula_queue.push({
            formula_text: formula_text,
            target: rng, 
            fct: fct, 
            arg_ranges: fct_arg_ranges, 
            offset: { i:0, j:0 } });
    } else {
        var k = 0
        for (var i=0, io=rng.row; io<=rng.end.row; i++, io++) {
            for (var j=0, jo=rng.col; jo<=rng.end.col; j++, jo++, k++) {
                //if (k>4) continue;    // !!
                formula_queue.push({
                    formula_text: formula_text,
                    target: { row: io, col: jo }, 
                    fct: fct, 
                    arg_ranges: fct_arg_ranges, 
                    offset: { i:i, j:j } });
            }
        }
    }
}

function pop_formula(wbk) {
    var node = formula_queue.shift();

    //if (node.formula_text.substr(0,10)=="=MATCH(K81") debugger;
	if (node.target.row==2 && node.target.col==8) debugger;

    var is_array_formula = node.formula_text[0]=="{";
    var caller = { wbk: wbk, wsh: wsh, rng: node.target, is_array_formula: is_array_formula };
    var arg_values = node.arg_ranges.map(rng => range_view(absolute_range_ref(rng, node.offset, caller)));
    /*
    try {
        var arg_values = node.arg_ranges.map(rng => 
            coerce_range_values(wbk, wsh, rng, node.offset));
    } catch (e) {
        if (e instanceof TooOldToCorceError)
            return formula_queue_next.push(node);
        else 
            throw e;
    }
    */

    
    // if (node.formula_text.substr(0,8)=="=SUMPROD") debugger; // !!

    return Promise.all(arg_values)
        .then(arg_values => {
        	var that = node;
            var res = node.fct.apply(node, arg_values);
            //if ((""+res)=="NaN")  // !!
            //    console.warn("  NaN with args", arg_values, node.arg_ranges);
            return res;
        })
        .then(data => { 
            // console.log("ok", xl_range_parse.stringify_range(node.target), node.formula_text, node.target, data); 
            if (data.is_range_view) {
            	if (data.caller.is_array_formula)
            		data = [...data].map(row => [...row]);
            	else 
            		data = data.valueOf();
            }
            update_range_values(node.target, data, wsh.last_changed); // , node.offset 
            add_range_to_html(node.target); 
        })
        .catch(e => {
            if (e instanceof TooOldToCorceError) {
				console.warn("deps", xl_range_parse.stringify_range(node.target), node.formula_text, "\n", e);
                return formula_queue_next.push(node);
			} else {
				console.error("err", xl_range_parse.stringify_range(node.target), node.formula_text, "\n", e);
			}
        });
}

async function pop_all_formulas(wbk) {
    console.log("wsh", wsh);
    console.log("wsh.last_changed", wsh.last_changed);
    var maxDepth = formula_queue.length;
    for (var depth=0; depth<maxDepth; depth++) {
        var nbTasksBegin = formula_queue.length;
        console.log("depth", depth, "#tasks", nbTasksBegin);
        for (var i=0; i<nbTasksBegin; i++)
            await pop_formula(wbk);
        var nbTasksEnd = formula_queue_next.length;

        if (nbTasksBegin==nbTasksEnd) {
            console.warn("#tasks in the queue stuck on dependencies", nbTasksEnd);
            break;
        } else if (nbTasksEnd>0) {
            formula_queue = formula_queue_next;
            formula_queue_next = [];
        } else {
            console.warn("max depth found", depth);
        }
    }
}

// -------------------------- 
// values

var wsh = {
    last_changed: -1,
    objects: {},
    values: new Array(400).fill(null).map(row => new Array(26).fill(null)),
    updated: new Array(400).fill(null).map(row => new Array(26).fill(null)),
    };

function set_range_values(rng, data) {
    wsh.last_changed = (new Date())*1;
    //console.log("set_range_values now", rng, wsh.last_changed)
    update_range_values(rng, data, null);
} 

function update_range_values(rng, data, opt_update_time) {
    var now = opt_update_time!==undefined ? opt_update_time : (new Date())*1;
    //console.log("update_range_values now", now, opt_update_time<0 ? "" : ((new Date())*1-opt_update_time))
    var end = rng.end || rng;
    for (var ii=0, io=rng.row; io<=end.row; ii++, io++) {
        for (var ji=0, jo=rng.col; jo<=end.col; ji++, jo++) {
            wsh.values[io][jo] = data_to_range_elnt(data, ii, ji);
            wsh.updated[io][jo] = now;
        }
    }
    return now;
}

function data_to_range_elnt(data,i,j) {
    if (data===undefined)
        return undefined;
        
    if (data.is_range_view)
        return data[i][j];

    if (!Array.isArray(data))
        // scalar
        return data;
    if (Array.isArray(data) && !Array.isArray(data[0]))
        // array 1D
        return data[i];
    return data[i][j];
}

function absolute_range_ref(rng, rng_offset, caller) {
    var abs_rng = {
    	row: rng.row + ( rng.abs_row ? 0 : rng_offset.i),
        col: rng.col + ( rng.abs_col ? 0 : rng_offset.j) };
    if (rng.end) {
    	var end = rng.end;
    	abs_rng.end = {
    		row: end.row + ( end.abs_row ? 0 : end_offset.i),
        	col: end.col + ( end.abs_col ? 0 : end_offset.j) };	
    }
    return { rng: abs_rng, caller: caller };
}

function TooOldToCorceError(rng) {
	this.addr =  xl_range_parse.stringify_range(rng);
    this.rng = rng;
}

function coerce_range_values(wbk, wsh, rng, rng_offset) {
    if (rng.reference_to_named_range) {
        rng = xl_range_parse.parse_range(
            wbk.Names["Sheet1"][rng.reference_to_named_range], 
            xl_range_parse.ReferenceStyle.A1);
    }

    var end = rng.end || rng;
    var is = rng.row + ( rng.abs_row ? 0 : rng_offset.i),
        js = rng.col + ( rng.abs_col ? 0 : rng_offset.j),
        ie = end.row + ( end.abs_row ? 0 : rng_offset.i),
        je = end.col + ( end.abs_col ? 0 : rng_offset.j);
     
    var n = ie - is,
        m = je - js; 
    var data = new Array(n);
    for (var i=0; i<=n; i++) {
        data[i] = new Array(m);
        for (var j=0; j<=m; j++) {
            if (wsh.updated[is+i][js+j]!==null && wsh.updated[is+i][js+j] < wsh.last_changed)
                throw new TooOldToCorceError({row: is+i, col: js+j });
            data[i][j] = wsh.values[is+i][js+j];
        }
    }
    return rng.end ? data : data[0][0];
}

function add_range_to_html(rng) {
    var end = rng.end || rng;
    for (var i=rng.row; i<=end.row; i++) {
        for (var j=rng.col; j<=end.col; j++) {
            var cell_id = xl_range_parse.stringify_range({ row: i, col: j });
            var val = wsh.values[i][j];
            // !! use NumberFormat 
            if (typeof val == "object" && !Array.isArray(val)) {
                wsh.objects[cell_id] = val;
                $('#'+cell_id).text(cell_id+": {...}");
            } else if (typeof val == "number") {
                $('#'+cell_id).text(Math.round(val*100)/100);
            } else {
                $('#'+cell_id).text(val);
            }
        }
    }
}

// -------------------------- 
// graphs

async function render_graphs(wbk) {
    console.log("wbk", wbk);
    var ref_style = xl_range_parse.ReferenceStyle[wbk.ReferenceStyle || "A1"];
    var wsh = wbk.Worksheets[0];
    var charts = wsh.charts;

    if (charts)
    for (var chart of charts) {
        console.log("chart", chart);
        var formula = chart.formula1.replace(/^=SERIES\(/, '=SERIES("'+JSON.stringify(chart).replace(/"/g,'""')+'", ');
        console.log("chart formula", formula);
        add_formula({ graph: true}, formula, ref_style);
    }
    var _ = await pop_all_formulas();
}

function SERIES(chart_json, title, arg1, arg2) {
    console.log(title);

    var chart = JSON.parse(chart_json);

    var chart_topleft = chart.range.split(":").shift().replace(/\$/g,"");
    var chart_id = chart_topleft + "_graph";

    var chart_type;
    if (chart.type.indexOf("xlXYScatter")===0) chart_type = 'scatter';

    var nb_series = arg2[0].length;
    var data = arg2.map((x,i) => [ arg1[i][0], x[0] ]);

    var container = $("#"+chart_topleft+" > #graph");
    if (container.length===0) {
    $("#"+chart_topleft).append('<div id="'+chart_id+'" style="position:absolute; width: 400px; height: 300px;"></div>');
        container = $("#"+chart_topleft+" > #graph");
    }
    Highcharts.chart(chart_id, {
        chart: {
            type: chart_type,
            zoomType: 'xy'
        },
        title: {
            text: title
        },
        xAxis: {
            title: {
                enabled: false,
                text: 'abscisa'
            },
        },
        yAxis: {
            title: {
                enabled: false,
                text: 'yyy'
            }
        },
        legend: {
            enabled: false
        },
        credits: { enabled: false },
        plotOptions: {
            marker: {
                enable: true
            }
        },
        series: [{
            // name: "wingspan",
            data: data
        }]
    });
}

function parse_query(query_string)
{
	// parse the query
	var qs = query_string || location.search;
	if (!qs) return {};
	
	var x = qs.substring(1).replace(/\+/g, ' ').replace(/;/g, '&').split('&');
	
	// q changes from string version of query to object
	var q={};
	for (var i=0; i<x.length; i++)
	{
		var t = x[i].split('=', 2);
		var name = unescape(t[0]);
		if (!q[name])
			q[name] = [];
		if (t.length>1)
			q[name][q[name].length] = unescape(t[1]);
		// next two lines are nonstandard 
		else
			q[name][q[name].length] = true;
	}
	return q;
}
