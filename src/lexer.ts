
import * as parser from "./excel_formula_parse";
import * as range_parser from "./excel_range_parse";
import * as xlexpr from "./excel_formula_transform";
import * as global_scope from "./global_scope";

var dflt_token_re = {
	string: /^("([^"\\]|""|\\.)*"|'([^'\\]|''|\\.)*')/,
	number: /^[+-]?[0-9]+\.?[0-9]*(e[+-]?[0-9]+)?/,
	id: /^[@$a-zA-Z~][\w_:.]*/,
	op: /^(=|\+|-|\*|\/|\(|\)|\{|\}|\[|\]|<|>|,|;|:|!|%|&|.)/,
	ws: /^\s+/,
}


export function lexer(txt_input : string, opt_token_re?)
{
	var txt = txt_input;
	var token_re = opt_token_re || dflt_token_re;
	var tokens = new Array();
	for (; txt.length>0;)
	{
		//info_msg("\ntxt: "+txt);

		var token = undefined;
		for (var t in token_re) {
			let re = token_re[t];
			let res = re.exec(txt);
			//info_msg(t+": "+JSON.stringify(res));
			if (res!=null) {
				token = { type: t, txt: res[0] };
				break;
			}
		}

		if (token) {
			//info_msg(JSON.stringify(token));
			tokens.push(token);
			txt = txt.substring(token.txt.length);			
		} else {
			if (txt!="")
				throw "unexpected token when hitting: "+txt+"\nfull expression: "+txt_input+"";
			// done !!
		}
	}

	return tokens;
}

export function handle_at_operator(tokens) {
	// use '@' prefix operator
	// to reference a graph node as opposed to a sheet range 
	// (it means we can manipulate objects or functions, not just scalars)

	for (var token of tokens)
		if (token.txt[0]=='@')
			token.txt = /* "graph_nodes." + */ token.txt.substr(1);  
			
	return tokens;
}

export function rebuild_source_text(tokens) {
	var txt = "";
	for (var t of tokens)
		txt += t.txt;
	return txt;
}

export function lexer_test() {
	info_msg("lexer_test");
	var expressions = [
		'="B1 is not" & A1',
		'="this formula countains \\" and "" double quotes"',
		"='this formula countains \\' and '' single quotes'",
		"window.setViews=function(v) { window.views = v; }",
		];

	var nb_errors = 0;
	for (var i=0;i<expressions.length; i++) {
		var expr_in = expressions[i];
		var expr_tok = lexer(expr_in);
		var expr_out = rebuild_source_text(expr_tok);
		if (expr_out!=expr_in) {
			nb_errors++;
			info_msg(JSON.stringify(expr_in));
			info_msg(JSON.stringify(expr_tok));
		}
	}
	if (nb_errors) throw new Error("lexer_test: "+nb_errors+" errors.");
}

function info_msg(msg) {
	console.log(msg);
}

lexer_test();
