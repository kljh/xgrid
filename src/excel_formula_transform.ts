/*
	Purpose: transform Excel syntax into a valid Javascript/C++ syntax.
	 - operator "<>" becomes "!="
	 - operator "=" becomes "==" 
	 - postfix operator "%" becomes a division by 100
	 - immediate array notation {1,2;21,22} becomes standard array [ [1,2], [21,22] ]
	 - ranges names becomes pure alphanumeric variable names
	 - little known range union and intersection operator are handled
*/

import * as parser from "./excel_formula_parse";

export function parse_and_transfrom_test() {
	var xl_formulas = [
		[ '=A1:B2', '=A1_B2' ],
		[ '="a"+1.2+TRUE+A1+xyz+#VALUE!', '="a"+1.2+true+A1+xyz+null' ], // all operands
		[ '="a"&"b"', '="a"+"b"' ],	// & operator
		[ '=50', '=50' ],
		[ '=50%', '=50/100' ], 	// % operator
		[ '=3.1E-24-2.1E-24', '=3.1E-24-2.1E-24' ],
		[ '=1+3+5', '=1+3+5' ],
		[ '=3 * 4 + \n\t 5', '=3*4+5' ],
		[ '=1-(3+5) ', '=1-(3+5)' ],
		[ '=1=5', '=1==5' ],	// = operator
		[ '=1<>5', '=1!=5'],	// <> operator
		[ '=2^8', ], 	// ^ operator
		[ '=TRUE', '=true' ],
		[ '=FALSE', '=false' ],
		[ '={1.2,"s";TRUE,FALSE}', '=[[1.2,"s"],[true,false]]' ],	// array notation
		[ '=$A1', ],
		[ '=$B$2', ],
		[ '=Sheet1!A1', '=Sheet1_A1' ],
		[ '=\'Another sheet\'!A1', '=Another_sheet_A1' ],
		[ '=[Book1]Sheet1!$A$1', '=_Book1_Sheet1_$A$1' ],
		[ '=[data.xls]sheet1!$A$1', '=_data_xls_sheet1_$A$1' ],
		[ '=\'[Book1]Another sheet\'!$A$1', '=_Book1_Another_sheet_$A$1' ],
		[ '=\'[Book No2.xlsx]Sheet1\'!$A$1', '=_Book_No2_xlsx_Sheet1_$A$1' ],
		[ '=\'[Book No2.xlsx]Another sheet\'!$A$1', '=_Book_No2_xlsx_Another_sheet_$A$1' ],
		//[ '=[\'my data.xls\']\'my sheet\'!$A$1', ],
		[ '=SUM(B5:B15)', '=SUM(B5_B15)' ],
		[ '=SUM(B5:B15,D5:D15)', '=SUM(B5_B15,D5_D15)' ],
		[ '=SUM(B5:B15 A7:D7)', '=SUM(B5_B15 inter A7_D7)' ],
		[ '=SUM(sheet1!$A$1:$B$2)', '=SUM(sheet1_$A$1_$B$2)' ],
		//[ '=SUM((A:A 1:1))', '=SUM((A_A inter 1_1))' ], // 1_1 NOT VALID !!!!!
		//[ '=SUM((A:A,1:1))', '=SUM((A_A union 1_1))' ],
		[ '=SUM((A:A A1:B1))', '=SUM((A_A inter A1_B1))' ],
		[ '=SUM(D9:D11,E9:E11,F9:F11)', '=SUM(D9_D11,E9_E11,F9_F11)' ],
		[ '=SUM((D9:D11,(E9:E11,F9:F11)))', '=SUM((D9_D11 union (E9_E11 union F9_F11)))' ],
		//[ '=IF(P5=1.0,"NA",IF(P5=2.0,"A",IF(P5=3.0,"B",IF(P5=4.0,"C",IF(P5=5.0,"D",IF(P5=6.0,"E",IF(P5=7.0,"F",IF(P5=8.0,"G"))))))))', 
		//  "=(P5==1.0?\"NA\":(P5==2.0?\"A\":(P5==3.0?\"B\":(P5==4.0?\"C\":(P5==5.0?\"D\":(P5==6.0?\"E\":(P5==7.0?\"F\":(P5==8.0?\"G\":undefined))))))))" ],
		[ '=SUM(B2:D2*B3:D3)', '=SUM(B2_D2*B3_D3)' ],  // array formula
		//[ '=SUM( 123 + SUM(456) + IF(DATE(2002,1,6), 0, IF(ISERROR(R[41]C[2]),0, IF(R13C3>=R[41]C[2],0, '
		//	+'  IF(AND(R[23]C[11]>=55,R[24]C[11]>=20), R53C3, 0))))', 
		//	'=123+SUM(456)+(DATE(2002,1,6)?0:(ISERROR(R_41_C_2_)?0:(R13C3>=R_41_C_2_?0:(AND(R_23_C_11_>=55,R_24_C_11_>=20)?R53C3:0))))' ],
		//[ '=IF("a"={"a","b";"c",#N/A;-1,TRUE}, "yes", "no") &   "  more ""test"" text"', 
		//  '=(\"a\"==[[\"a\",\"b\"],[\"c\",null],[-1,true]]?\"yes\":\"no\")+\"  more \\\"test\\\" text\"' ],
		[ '=+ AName- (-+-+-2^6) = {"A","B"} + @SUM(R1C1) + (@ERROR.TYPE(#VALUE!) = 2)', 
		  '=AName-(---2^6)==[[\"A\",\"B\"]]+SUM(R1C1)+(ERROR.TYPE(null)==2)' ],
		//[ '=IF(R13C3>DATE(2002,1,6),0,IF(ISERROR(R[41]C[2]),0,IF(R13C3>=R[41]C[2],0, IF(AND(R[23]C[11]>=55,R[24]C[11]>=20),R53C3,0))))', 
		//  '=(R13C3>DATE(2002,1,6)?0:(ISERROR(R_41_C_2_)?0:(R13C3>=R_41_C_2_?0:(AND(R_23_C_11_>=55,R_24_C_11_>=20)?R53C3:0))))' ],
		//[ '=IF(R[39]C[11]>65,R[25]C[42],ROUND((R[11]C[11]*IF(OR(AND(R[39]C[11]>=55, R[40]C[11]>=20),AND(R[40]C[11]>=20,R11C3="YES")),R[44]C[11],R[43]C[11]))+(R[14]C[11] *IF(OR(AND(R[39]C[11]>=55,R[40]C[11]>=20),AND(R[40]C[11]>=20,R11C3="YES")), R[45]C[11],R[43]C[11])),0))', ],
		];

	var js_formulas = [
		'="a string \\n\\t with a few escape chars"',
		'="a string \\" with a tricky escape char"',
		'=true && false',
		"=A18==0 && F4==0"
		];
	
	let nb_errors = 0;

	function test(formulas, prms) {
		var formula_xl_js = [];
		var fcts={}, vars={};
		for (var f=0; f<formulas.length; f++) {
			var xlformula = Array.isArray(formulas[f]) ? formulas[f][0] : formulas[f];
			var jsexpected = Array.isArray(formulas[f]) ? formulas[f][1] || xlformula : xlformula;
			
			var tokens, ast;
			var jsformula = "="+parse_and_transfrom(xlformula,vars,fcts,prms);
			formula_xl_js.push([ xlformula, jsformula ]);
			
			if (jsformula!=jsexpected)
				nb_errors++;
		
		}
		/*
		info_msg("formula_xl_js:\n"+JSON.stringify(formula_xl_js,null,4));
		info_msg("vars:\n"+JSON.stringify(vars,null,4));
		info_msg("fcts:\n"+JSON.stringify(fcts,null,4));
		*/
	}
	test(xl_formulas, { xl: true });
	//test(js_formulas, { xjs: true });
    if (nb_errors) throw new Error("parse_and_transfrom_test: "+nb_errors+" errors.");
}

export function parse_and_transfrom(xlformula,vars,fcts,opt_prms) {
	var prms = opt_prms || {};
	
	// ExtendedJS grammar (with range and vector operations)
	// --or-- Excel grammar to plain JS grammar conversion 

	var tokens = parser.getTokens(xlformula);
	//info_msg("TOKENS:\n"+JSON.stringify(tokens,null,4));

	try {
		var ast = build_token_tree(tokens);
	} catch (e) {
		throw new Error("Error while parsing "+xlformula+"\n"+(e.stack||e));
	}
	//info_msg("TREE:\n"+JSON.stringify(ast,null,4));

	var jsformula = excel_to_js_formula(ast,vars,fcts,prms);
	return jsformula;
}

function build_token_tree(tokens) {
	var stack = [];
	var stack_top = { args : [[]] };
	
	while (tokens.moveNext()) {
		var token = tokens.current();

		if (token.subtype==parser.TOK_SUBTYPE_START) {
			stack.push(stack_top);
			stack_top = token;
			stack_top.args = [[]]; 
		} else if (token.type==parser.TOK_TYPE_ARGUMENT) {
			stack_top.args.push([]);
		} else if (token.subtype==parser.TOK_SUBTYPE_STOP) {
			var tmp = stack_top
			stack_top = stack.pop();
			
			var nb_args = stack_top.args.length;
			stack_top.args[nb_args-1].push(tmp);
		} else {
			var nb_args = stack_top.args.length;
			stack_top.args[nb_args-1].push(token);
		}
	}
	if (stack_top.args.length!=1)
		throw new Error("root formula contains multiple expressions.\n"+JSON.stringify(stack_top,null,4));
	return stack_top.args[0];
} 

function excel_to_js_formula(tokens, opt_vars?, opt_fcts?, opt_prms?) {
	var prms = opt_prms || {};
	var vars = opt_vars || {};
	var fcts = opt_fcts || {};
	var vars_count = Object.keys(vars).length;

	function excel_to_js_operand(token) {
		switch (token.subtype) {
			case parser.TOK_SUBTYPE_TEXT:
				return JSON.stringify(token.value);
			case parser.TOK_SUBTYPE_NUMBER:
				return token.value;
			case parser.TOK_SUBTYPE_LOGICAL:
				return token.value.toLowerCase();
			case parser.TOK_SUBTYPE_ERROR:
				return "null";
			case parser.TOK_SUBTYPE_RANGE:
				if (!vars[token.value]) {
					// let valid_id = "x"+vars_count;
					let valid_id = token.value.replace(/[^A-Za-z0-9$]/g, "_");
					for (let k in vars)
						if (valid_id==vars[k])
							throw new Error("argument name collision between "+JSON.stringify(k)+" and "+JSON.stringify(token.value));
					vars_count++;
					vars[token.value] = valid_id; 
				}
				return vars[token.value];
			default:
				throw new Error("unhandled subtype "+token.subtype+"\n"+JSON.stringify(token,null,4));
		}
	}
	
	function excel_to_js_array_operand(token) {
		var arr = new Array(token.args.length);
		for (var i=0; i<arr.length; i++) {
			if (token.args[i].length!=1)
				throw new Error("ARRAY is expected to be a list of ARRAYROW but found an expressions");

			arr[i] = new Array(token.args[i][0].args.length);
			for (var j=0; j<arr[i].length; j++) {
				var tmp = excel_to_js_formula(token.args[i][0].args[j]);
				arr[i][j] = JSON.parse(tmp);
			}
		}
		return JSON.stringify(arr);
	}

	function excel_to_js_function_call_formula(token) {
		fcts[token.value] = (fcts[token.value]||0) + 1;
		var res = token.value+"(";
		var nb_args = token.args.length;
		for (var i=0; i<nb_args; i++) {
			var expr = token.args[i];
			res += excel_to_js_formula(expr,vars,fcts);
			res += (i+1<nb_args ? "," : "");
		}
		return res+")";
	}

	function excel_to_js_if_then_else_trigram(token) {
		var nb_args = token.args.length;
		if (nb_args<2 || nb_args>3) 
		throw new Error("expect 2 or 3 arguments");

		var res = 
			"(" + excel_to_js_formula(token.args[0],vars,fcts) +
			"?" + excel_to_js_formula(token.args[1],vars,fcts) +
			":" + (nb_args==3?excel_to_js_formula(token.args[2],vars,fcts):"undefined") + 
			")";
		return res;
	}

	function excel_to_js_operator(token) {
		switch (token.subtype) {
			case parser.TOK_SUBTYPE_MATH:
				return token.value;
			case parser.TOK_SUBTYPE_LOGICAL:
				if (token.value=="=")
					return "==";
				else if (token.value=="<>")
					return "!=";
				else 
					return token.value;
			case parser.TOK_SUBTYPE_CONCAT:
				return "+"; // instead of &
			case parser.TOK_SUBTYPE_INTERSECT:
				return " inter "; // not a valid Javascript operator, will fail later but properly capture the intent.
				// return " && "; // a valid Javascript operator so that JS parser can produce an AST, BUT SEMANTICS CHANGED
			case parser.TOK_SUBTYPE_UNION:
				return " union "; // not a valid Javascript operator, will fail later but properly capture the intent.
				// return " || "; // a valid Javascript operator so that JS parser can produce an AST, BUT SEMANTICS CHANGED
			default:
				throw new Error("unhandled subtype "+token.subtype+"\n"+JSON.stringify(token));
		}
	}

	var res = "";
	for (var t=0; t<tokens.length; t++) {
		var token = tokens[t];
		
		switch (token.type) {

			// operators
			case "operator-infix":
				if (!prms.xjs)
					res += excel_to_js_operator(token);
				else 
					res += token.value;
				break;
			case "operator-prefix":
				res += token.value;
				//throw new Error("unhandled token "+token.type);
				break;
			case "operator-postfix":
				if (token.value!="%")
					throw new Error("unhandled token "+token.type+" "+token.value);
				res += "/100";
				break;
			
			// operands 
			case "operand":
				res += excel_to_js_operand(token);
				break;
		
			case "function":
			case "subexpression":
				if (!prms.xjs && token.value=="ARRAY")
					res += excel_to_js_array_operand(token);
				//else if (!prms.xjs && token.value=="IF")
				//	res += excel_to_js_if_then_else_trigram(token);
				else 
					res += excel_to_js_function_call_formula(token);
			
				break;
			case "white-space":
				res += token.value;
				break;
			default:
				throw new Error("unhandled "+token.type);
		}
	}
	return res;
}

function info_msg(msg) {
	console.log(msg);
}

parse_and_transfrom_test();
