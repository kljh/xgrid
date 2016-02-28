
function parse_and_transfrom_test() {
	var formulas = [
		'="a"+1.2+TRUE+A1+xyz+#VALUE!', // all operands
    '="a"&"b"',	// & operator
		'=50',
		'=50%', 	// % operator
		'=3.1E-24-2.1E-24',
		'=1+3+5',
		'=3 * 4 + 5',
		'=1-(3+5) ',
		'=1=5', 	// = operator
		'=1<>5',	// <> operator
		'=2^8', 	// ^ operator
		'=TRUE',
		'=FALSE',
		'={1.2,"s";TRUE,FALSE}',	// array notation
		'=$A1',
		'=$B$2',
		'=[data.xls]sheet1!$A$1',
		//'=[\'my data.xls\']\'my sheet\'!$A$1',
		'=SUM(B5:B15)',
		'=SUM(B5:B15,D5:D15)',
		'=SUM(B5:B15 A7:D7)',
		'=SUM(sheet1!$A$1:$B$2)',
		'=SUM((A:A 1:1))',
		'=SUM((A:A,1:1))',
		'=SUM((A:A A1:B1))',
		'=SUM(D9:D11,E9:E11,F9:F11)',
		'=SUM((D9:D11,(E9:E11,F9:F11)))',
		'=IF(P5=1.0,"NA",IF(P5=2.0,"A",IF(P5=3.0,"B",IF(P5=4.0,"C",IF(P5=5.0,"D",IF(P5=6.0,"E",IF(P5=7.0,"F",IF(P5=8.0,"G"))))))))',
		'=SUM(B2:D2*B3:D3)',  // array formula
		'=SUM( 123 + SUM(456) + IF(DATE(2002,1,6), 0, IF(ISERROR(R[41]C[2]),0, IF(R13C3>=R[41]C[2],0, '
      +'  IF(AND(R[23]C[11]>=55,R[24]C[11]>=20), R53C3, 0))))',
		'=IF("a"={"a","b";"c",#N/A;-1,TRUE}, "yes", "no") &   "  more ""test"" text"',
		'=+ AName- (-+-+-2^6) = {"A","B"} + @SUM(R1C1) + (@ERROR.TYPE(#VALUE!) = 2)',
		'=IF(R13C3>DATE(2002,1,6),0,IF(ISERROR(R[41]C[2]),0,IF(R13C3>=R[41]C[2],0, IF(AND(R[23]C[11]>=55,R[24]C[11]>=20),R53C3,0))))',
		'=IF(R[39]C[11]>65,R[25]C[42],ROUND((R[11]C[11]*IF(OR(AND(R[39]C[11]>=55, R[40]C[11]>=20),AND(R[40]C[11]>=20,R11C3="YES")),R[44]C[11],R[43]C[11]))+(R[14]C[11] *IF(OR(AND(R[39]C[11]>=55,R[40]C[11]>=20),AND(R[40]C[11]>=20,R11C3="YES")), R[45]C[11],R[43]C[11])),0))',
		];

	for (var f=0; f<formulas.length; f++) {
		var xlformula = formulas[f];
    print(f);
    print(xlformula);
    var tokens, ast;
    try {
      var tokens = getTokens(xlformula);
      //print("TOKENS:\n"+JSON.stringify(tokens,null,4));
      var ast = build_token_tree(tokens);
  		//print("TREE:\n"+JSON.stringify(ast,null,4));
      
      var jsformula = "="+excel_to_js_formula(ast);
  		print(jsformula);
      
		} catch (e) {
      print("ERROR: "+e+"\n"+e.stack);
    }
    print("\n------\n");
    //break;
	}
}

function build_token_tree(tokens) {
  var stack = [];
  var stack_top = { args : [[]] };
  
  while (tokens.moveNext()) {
    var token = tokens.current();

    if (token.subtype==TOK_SUBTYPE_START) {
      stack.push(stack_top);
      stack_top = token;
      stack_top.args = [[]]; 
    } else if (token.type==TOK_TYPE_ARGUMENT) {
      stack_top.args.push([]);
    } else if (token.subtype==TOK_SUBTYPE_STOP) {
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
    throw new Error(arguments.callee.name+": root formula contains multiple expressions.\n"+JSON.stringify(stack_top,null,4));
  return stack_top.args[0];
} 

function excel_to_js_formula(tokens) {

  function excel_to_js_operand(token) {
    switch (token.subtype) {
      case TOK_SUBTYPE_TEXT:
        return JSON.stringify(token.value);
      case TOK_SUBTYPE_NUMBER:
        return token.value;
      case TOK_SUBTYPE_LOGICAL:
        return token.value.toLowerCase();
      case TOK_SUBTYPE_ERROR:
        return "null";
      case TOK_SUBTYPE_RANGE:
        return token.value//+".value()";
      default:
        throw new Error(arguments.callee.name+": unhandled subtype "+token.subtype+"\n"+JSON.stringify(token,null,4));
    }
  }
  
  function excel_to_js_array_operand(token) {
    var arr = new Array(token.args.length);
    for (var i=0; i<arr.length; i++) {
      if (token.args[i].length!=1)
        throw new Error(arguments.callee.name+": ARRAY is expected to be a list of ARRAYROW but found an expressions");

      arr[i] = new Array(token.args[i][0].args.length);
      for (var j=0; j<arr[i].length; j++) {
        var tmp = excel_to_js_formula(token.args[i][0].args[j]);
        arr[i][j] = JSON.parse(tmp);
      }
    }
    return JSON.stringify(arr);
  }

  function excel_to_js_function_call_formula(token) {
    var res = token.value+"(";
    var nb_args = token.args.length;
    for (var i=0; i<nb_args; i++) {
      var expr = token.args[i];
      res += excel_to_js_formula(expr);
      res += (i+1<nb_args ? "," : "");
    }
    return res+")";
  }

  function excel_to_js_operator(token) {
    switch (token.subtype) {
      case TOK_SUBTYPE_MATH:
        return token.value;
        break;
      case TOK_SUBTYPE_LOGICAL:
        if (token.value=="=")
          return "==";
        else if (token.value=="<>")
          return "!=";
        else 
          return token.value;
        break;
      case TOK_SUBTYPE_CONCAT:
        return "+"; // instead of &
        break;
      case TOK_SUBTYPE_INTERSECT:
      case TOK_SUBTYPE_UNION:
      default:
        throw new Error(arguments.callee.name+": unhandled subtype "+token.subtype+"\n"+JSON.stringify(token));
    }
  }

  var res = "";
  for (var t=0; t<tokens.length; t++) {
    var token = tokens[t];
    
    switch (token.type) {

      // operators
      case "operator-infix":
        res += excel_to_js_operator(token);
        break;
      case "operator-prefix":
        res += token.value;
        //throw new Error(arguments.callee.name+": unhandled "+token.type);
        break;
      case "operator-postfix":
        if (token.value!="%")
          throw new Error(arguments.callee.name+": unhandled "+token.type+" "+token.value);
        res += "/100";
        break;
      
      // operands 
      case "operand":
        res += excel_to_js_operand(token);
        break;
    
      case "function":
      case "subexpression":
        if (token.value=="ARRAY")
          res += excel_to_js_array_operand(token);
        else 
          res += excel_to_js_function_call_formula(token);
      
        break;
      default:
        throw new Error(arguments.callee.name+": unhandled "+token.type);
    }
  }
  return res;
}

load("excel_formula_parse.js")
parse_and_transfrom_test();