var dflt_token_re = {
	string: /^"[^"]*"|'[^']*'/,
	number: /^[+-]?[0-9]+\.?[0-9]*(e[+-]?[0-9]+)?/,
	id: /^[a-zA-Z~][\w_:.]*/,
	op: /^=|\+|-|\*|\/|\(|\)|\{|\}|\[|\]|<|>|,|;|:|!|%|&/,
	ws: /^\s+/,
}


function lexer(txt : string, opt_token_re?)
{
	var token_re = opt_token_re || dflt_token_re;
	var tokens = new Array();
	for (; txt.length>0;)
	{
		info_msg("\ntxt: "+txt);

		var token = undefined;
		for (var t in token_re) {
			let re = token_re[t];
			let res = re.exec(txt);
			info_msg(t+": "+JSON.stringify(res));
			if (res!=null) {
				token = { type: t, txt: res[0] };
				break;
			}
		}

		if (token) {
			info_msg(JSON.stringify(token));
			tokens.push(token);
			txt = txt.substring(token.txt.length);			
		} else {
			if (txt!="")
				throw "unexpected token "+txt;
			// done !!
		}
	}

	return tokens;
}

function lexer_test() {
	info_msg("lexer_test");
	var expressions = [
		'="B1 is not" &A1'
		];
	for (var i=0;i<expressions.length; i++) {
		info_msg(JSON.stringify(lexer(expressions[i])));
	}
}

function info_msg(msg) {
	console.log(msg);
}

// Nodejs stuff
if (typeof module!="undefined") {
	var fs = require("fs");
	var parser = require("./excel_formula_parse");
	var range_parser = require("./excel_range_parse");
	var xlexpr = require("./excel_formula_transform");
	var global_scope = require("./global_scope");

	// run tests if this file is called directly
	if (require.main === module)
		lexer_test();

	module.exports.lexer = lexer;
}
