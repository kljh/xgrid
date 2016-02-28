
function parse_and_transfrom_test() {
	var formulas = [
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
		'=[\'my data.xls\']\'my sheet\'!$A$1',
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
		'={SUM(B2:D2*B3:D3)}',
		'=SUM(123 + SUM(456) + (45DATE(2002,1,6),0,IF(ISERROR(R[41]C[2]),0,IF(R13C3>=R[41]C[2],0, IF(AND(R[23]C[11]>=55,R[24]C[11]>=20),R53C3,0))))',
		'=IF("a"={"a","b";"c",#N/A;-1,TRUE}, "yes", "no") &   "  more ""test"" text"',
		'=+ AName- (-+-+-2^6) = {"A","B"} + @SUM(R1C1) + (@ERROR.TYPE(#VALUE!) = 2)',
		'=IF(R13C3>DATE(2002,1,6),0,IF(ISERROR(R[41]C[2]),0,IF(R13C3>=R[41]C[2],0, IF(AND(R[23]C[11]>=55,R[24]C[11]>=20),R53C3,0))))',
		'=IF(R[39]C[11]>65,R[25]C[42],ROUND((R[11]C[11]*IF(OR(AND(R[39]C[11]>=55, R[40]C[11]>=20),AND(R[40]C[11]>=20,R11C3="YES")),R[44]C[11],R[43]C[11]))+(R[14]C[11] *IF(OR(AND(R[39]C[11]>=55,R[40]C[11]>=20),AND(R[40]C[11]>=20,R11C3="YES")), R[45]C[11],R[43]C[11])),0))',
		];

	for (var f=12; f<formulas.length; f++) {
		var xlformula = formulas[f];
		var tokens = getTokens(xlformula);
		var jsformula = transform_tokens(tokens);
		print(f);
    print(xlformula);
    print(jsformula);
    //print(JSON.stringify(tokens,null,4));
		print("\n------\n");
		
	}
}

function transform_tokens(tokens) {
  function transform_operand(token) {
      if (token.subtype=="text")
        return JSON.stringify(token.value);
      else if (token.value=="TRUE")
        return true;
      else if (token.value=="FALSE")
        return false;
      else 
        return token.value;
  }
  var res = "=";
  while (tokens.moveNext()) {
    var token = tokens.current();
    
    if (token.type=="operator-infix") {
      if (token.subtype=="concatenate")
        res += "+";
      else if (token.value=="=")
        res += "==";
      else if (token.value=="<>")
        res += "!=";
      else 
        res += token.value;
    }
    if (token.type=="operand") 
      res += transform_operand(token);

    if (token.subtype==TOK_SUBTYPE_START) 
      res += token.value + "(";
    if (token.subtype==TOK_SUBTYPE_STOP) 
      res += ")";
    
  }
  return res;
}

load("excel_formula_parse.js")
parse_and_transfrom_test();