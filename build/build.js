define("excel_range_parse", ["require", "exports"], function (require, exports) {
    "use strict";
    Object.defineProperty(exports, "__esModule", { value: true });
    var ReferenceStyle;
    (function (ReferenceStyle) {
        ReferenceStyle[ReferenceStyle["Auto"] = 0] = "Auto";
        ReferenceStyle[ReferenceStyle["A1"] = 1] = "A1";
        ReferenceStyle[ReferenceStyle["R1C1"] = 2] = "R1C1";
    })(ReferenceStyle = exports.ReferenceStyle || (exports.ReferenceStyle = {}));
    ;
    function parse_range_bijection_test() {
        var ranges_a1 = [
            'A1',
            'A$1',
            '$A1',
            '$A$1',
            'Z2',
            'AA3',
            'AZ4',
            'BA5',
            'ZZ6',
            'AAA7',
            'B5:B15',
            'sheet1!$A$1:$B$2',
            'Sheet1!Z1',
            '\'Another sheet\'!A1',
            '[Book1]Sheet1!$A$1',
            '[data.xls]sheet1!$A$1',
            '\'[Book1]Another sheet\'!$A$1',
            '\'[Book No2.xlsx]Sheet1\'!$A$1',
            '\'[Book No2.xlsx]Another sheet\'!$A$1',
            '[data.xls]sheet1!$A$1',
            '\'Let\'\'s test\'!A1',
            'A:A',
            'A:C',
            '1:1',
            '1:5',
            'named_cell',
            'alpha0',
            'R1',
            'C1',
        ];
        var ranges_r1c1 = [
            // R1C1 mode 
            'RC',
            'R[1]C',
            'RC[2]',
            'R[1]C[2]',
            'R13C3',
            'R[4]C7',
            'R23C[11]',
            'RC13',
            'R1C2:R[1]C[2]',
            // Not supported :
            //'R2', 	// ambiguous : full R1C1 row can be interpreted as A1 cell
            //'R[-2]',
            //'C5', 	// ambiguous : full R1C1 column can be interpreted as A1 cell
            //'C[3]',
            //'R3:R5', 	// ambiguous : full R1C1 rows can be interpreted as A1 range
            'named_cell',
            'radio_control_c1',
        ];
        var nb_errors = 0;
        function test(ranges, mode, rng_ref) {
            for (var r = 0; r < ranges.length; r++) {
                var rng_in = ranges[r];
                var tmp = parse_range(rng_in, mode, rng_ref);
                var rng_out = stringify_range(tmp, mode, rng_ref);
                //info_msg(JSON.stringify(rng_in));
                //info_msg(JSON.stringify(rng_out));
                //info_msg(JSON.stringify(tmp,null,4));
                // make sure everything is parsed
                var bOk = typeof tmp == "object" || rng_in == "named_cell" || rng_in == "alpha0" || rng_in == "radio_control_c1" || (rng_in == "R1C1" && mode == ReferenceStyle.A1);
                if (!bOk)
                    throw new Error("parse_range_bijection_test");
                // and bijective
                // assert.equal(rng_in,rng_out);
                if (rng_in != rng_out) {
                    console.log(rng_in, " <> ", rng_out);
                    nb_errors++;
                }
            }
        }
        var rng_ref = parse_range("A1", ReferenceStyle.A1);
        test(ranges_a1, ReferenceStyle.A1, rng_ref);
        test(ranges_r1c1, ReferenceStyle.R1C1, rng_ref);
        if (nb_errors)
            throw new Error("parse_range_bijection_test");
        function assert_equal(a, b) { if (a !== b)
            throw new Error("parse_range_bijection_test"); }
        var assert = { equal: assert_equal };
        assert.equal(JSON.stringify(parse_range('R1', ReferenceStyle.A1, rng_ref)), '{"col":18,"abs_col":false,"row":1,"abs_row":false}');
        assert.equal(JSON.stringify(parse_range('R[1]C[1]', ReferenceStyle.R1C1, rng_ref)), '{"row":2,"abs_row":false,"col":2,"abs_col":false}');
        //assert.equal(JSON.stringify(parse_range('R[1]', ReferenceStyle.R1C1, rng_ref)),
        //	'{"row":1,"abs_row":false,"abs_col":false,"end":{"row":1,"abs_row":false,"abs_col":false}');
        assert.equal(parse_column_name_to_index("A"), 1);
        assert.equal(parse_column_name_to_index("Z"), 26);
        assert.equal(parse_column_name_to_index("AA"), 27);
        assert.equal(parse_column_name_to_index("AZ"), 52);
        assert.equal(parse_column_name_to_index("BA"), 53);
        assert.equal(parse_column_name_to_index("AAA"), parse_column_name_to_index("ZZ") + 1);
        assert.equal(parse_column_name_to_index("ABA"), parse_column_name_to_index("AAZ") + 1);
        assert.equal(parse_column_name_to_index("BAA"), parse_column_name_to_index("AZZ") + 1);
        assert.equal(parse_column_index_to_name(parse_column_name_to_index("A")), "A");
        assert.equal(parse_column_index_to_name(parse_column_name_to_index("Z")), "Z");
        assert.equal(parse_column_index_to_name(parse_column_name_to_index("AA")), "AA");
        assert.equal(parse_column_index_to_name(parse_column_name_to_index("AZ")), "AZ");
        assert.equal(parse_column_index_to_name(parse_column_name_to_index("AAA")), "AAA");
        assert.equal(parse_column_index_to_name(parse_column_name_to_index("AZZ")), "AZZ");
    }
    exports.parse_range_bijection_test = parse_range_bijection_test;
    function parse_range(arg, mode, rng_ref) {
        var rex = /((.*!)?)([^!]*)/.exec(arg);
        var book_and_sheet = rex[2];
        var range = rex[3];
        if (book_and_sheet) {
            var quoted = (book_and_sheet[0] == '\'');
            var len = book_and_sheet.length;
            book_and_sheet = quoted ? book_and_sheet.substr(1, len - 3) : book_and_sheet.substr(0, len - 1);
            var rex = /(\[(.*)\])?(.*)/.exec(book_and_sheet);
            var book = rex[2];
            var sheet = rex[3];
            book = book ? book.replace(/''/g, '\'') : book;
            sheet = sheet ? sheet.replace(/''/g, '\'') : sheet;
        }
        var res = parse_cell_range(range, mode, rng_ref);
        res.book = book;
        res.sheet = sheet;
        return res;
    }
    exports.parse_range = parse_range;
    function parse_cell_range(range, mode, rng_ref) {
        var cells = range.split(":");
        if (cells.length > 2)
            throw new Error("range address invalid, contains more than on colons. " + range);
        var cell0 = parse_cell(cells[0], mode, rng_ref);
        var cell1 = parse_cell(cells[1], mode, rng_ref);
        if (cell0 && !isNaN(cell0.row) && !isNaN(cell0.col)
            || cell0 && cell1 && !isNaN(cell0.row) && !isNaN(cell1.row)
            || cell0 && cell1 && !isNaN(cell0.col) && !isNaN(cell1.col)) {
            // properly formated range
            var res = cell0;
            res.end = cell1;
            return res;
        }
        else {
            // named range
            return { reference_to_named_range: range };
        }
    }
    function parse_cell(cell, mode, rng_ref) {
        if (!cell)
            return;
        switch (mode) {
            case ReferenceStyle.Auto:
                var isA1 = A1.test(cell), isR1C1 = R1C1.test(cell);
                if (isA1 && !isR1C1)
                    return parse_cell_A1(cell, rng_ref);
                if (!isA1 && isR1C1)
                    return parse_cell_R1C1(cell, rng_ref);
                return { reference_to_named_range: cell };
            case ReferenceStyle.A1:
                return parse_cell_A1(cell, rng_ref);
            case ReferenceStyle.R1C1:
                return parse_cell_R1C1(cell, rng_ref);
            default:
                throw new Error("parse_cell: unknown reference style " + cell + " " + mode);
        }
    }
    var R1C1 = /^(R(\[?(-?[0-9]*)\]?))?(C(\[?(-?[0-9]*)\]?))?$/;
    function parse_cell_R1C1(cell, rng_ref) {
        var r1c1 = R1C1.exec(cell);
        if (r1c1) {
            if (r1c1[3] || r1c1[6] || cell == "RC") {
                var res = {
                    row: parseFloat(r1c1[3]),
                    abs_row: r1c1[2] == r1c1[3] && !!r1c1[2],
                    col: parseFloat(r1c1[6]),
                    abs_col: r1c1[5] == r1c1[6] && !!r1c1[5]
                };
                if (!res.abs_row)
                    res.row = rng_ref.row + (isNaN(res.row) ? 0 : res.row);
                if (!res.abs_col)
                    res.col = rng_ref.col + (isNaN(res.col) ? 0 : res.col);
                return res;
            }
        }
    }
    var A1 = /^(\$?([A-Z]{0,3}))?(\$?([0-9]*))?$/;
    function parse_cell_A1(cell, rng_ref) {
        if (!cell)
            return;
        var a1 = A1.exec(cell);
        if (a1) {
            if (a1[2] || a1[4]) {
                return {
                    col: parse_column_name_to_index(a1[2]),
                    abs_col: a1[1] != a1[2],
                    row: parseFloat(a1[4]),
                    abs_row: a1[3] != a1[4],
                };
            }
        }
    }
    function parse_column_name_to_index(col) {
        if (!col)
            return parseFloat("NaN");
        // we are NOT counting in base 25/26 : A is both 0 (when not the leading number) and 1 (when the leading number)
        // we are using symbols [A-Z] rather than [0-9A-P]
        //return parseInt(col, 26);
        var pos = 0;
        var mult = 1;
        var A = "A".charCodeAt(0);
        for (var k = 0; k < col.length; k++) {
            var v = col.charCodeAt(col.length - (k + 1)) - A;
            if (k == 0)
                pos += v + 1; // units
            if (k == 1)
                pos += v * 26 + 26; // tens
            if (k == 2)
                pos += v * 26 * 26 + 26 * 26; // hundreds
            if (k > 2)
                throw new Error("column on more than 3 letters");
        }
        return pos;
    }
    function stringify_range(rng, mode, rng_ref) {
        if (rng.reference_to_named_range)
            return rng.reference_to_named_range; // named range
        var scope = "";
        var bNeedQuotes = false;
        var reNeedsQuotes = /^[A-Za-z0-9.]*$/;
        var reSubstQuotes = /'/g;
        if (rng.book) {
            bNeedQuotes = bNeedQuotes || !reNeedsQuotes.test(rng.book);
            scope += "[" + rng.book.replace(reSubstQuotes, "''") + "]";
        }
        if (rng.sheet) {
            bNeedQuotes = bNeedQuotes || !reNeedsQuotes.test(rng.sheet);
            scope += rng.sheet.replace(reSubstQuotes, "''");
            if (bNeedQuotes)
                scope = "'" + scope + "'";
            scope += "!";
        }
        if (mode === ReferenceStyle.R1C1)
            return scope
                + (rng.row ? "R" + (rng.abs_row ? rng.row : (rng_ref.row == rng.row ? "" : "[" + (rng.row - rng_ref.row) + "]")) : "")
                + (rng.col ? "C" + (rng.abs_col ? rng.col : (rng_ref.col == rng.col ? "" : "[" + (rng.col - rng_ref.col) + "]")) : "")
                + (rng.end ? ":" + stringify_range(rng.end, mode, rng_ref) : "");
        else
            return scope
                + (rng.abs_col ? "$" : "") + parse_column_index_to_name(rng.col)
                + (rng.abs_row ? "$" : "") + (rng.row || "")
                + (rng.end ? ":" + stringify_range(rng.end, mode, rng_ref) : "");
    }
    exports.stringify_range = stringify_range;
    function parse_column_index_to_name(pos) {
        if (pos === undefined || pos === null || isNaN(pos))
            return ""; // for instance for 1:1 
        var col = "";
        var A = "A".charCodeAt(0);
        var k = 0; // number of characters used to encode
        do {
            if (k == 0)
                pos -= 1;
            if (k == 1)
                pos -= 1; // effectively -26 because we have divieded by 26 once
            if (k == 2)
                pos -= 1; // effectively -26*26 because we have divieded by 26 twice
            if (k > 2)
                throw new Error("column on more than 3 letters. " + pos);
            k++;
            var u = pos % 26;
            col = String.fromCharCode(A + u) + col;
            var pos = (pos - u) / 26;
        } while (pos != 0);
        return col;
    }
    // Nodejs stuff
    if (typeof module != "undefined") {
        var parser = require("./excel_formula_parse");
        // run tests if this file is called directly
        if (require.main === module) {
            parse_range_bijection_test();
        }
        module.exports.parse_range = parse_range;
        module.exports.stringify_range = stringify_range;
    }
});
/*

Excel Formula Parsing
http://ewbi.blogs.com/develops/2004/12/excel_formula_p.html

A list of Excel’s operators
http://chacocanyon.com/smm/readings/referenceoperators.shtml

Excel has 16 infix operators, 5 matchfix operators, 3 prefix operators and a postfix operator.

Infix operators
+ 	addition
- 	subtraction
* 	multiplication
/ 	division
^ 	exponentiation
> 	is greater than
< 	is less than
= 	is equal to
>= 	is greater than or equal to
<= 	is less than or equal to
<> 	is not equal to
& 	concatenation of strings
Space 	reference intersection
Comma (,) 	reference union
Colon (:) 	range reference defined by two cell references
! 	separate worksheet name from reference

Matchfix operators
Left Side 	Right Side 	Operation
" 	" 	string constant
{ 	} 	array constant
( 	) 	arithmetic grouping or function arguments or reference grouping
' 	' 	grouping worksheet name
[ 	] 	grouping workbook name, or relative reference in R1C1 style

Postfix operators
% 	percentage

Prefix operators
- 	negative
+ 	plus
$ 	Next component of an A1 reference is absolute

*/
define("excel_formula_parse", ["require", "exports"], function (require, exports) {
    "use strict";
    Object.defineProperty(exports, "__esModule", { value: true });
    exports.TOK_TYPE_NOOP = "noop";
    exports.TOK_TYPE_OPERAND = "operand";
    exports.TOK_TYPE_FUNCTION = "function";
    exports.TOK_TYPE_SUBEXPR = "subexpression";
    exports.TOK_TYPE_ARGUMENT = "argument";
    exports.TOK_TYPE_OP_PRE = "operator-prefix";
    exports.TOK_TYPE_OP_IN = "operator-infix";
    exports.TOK_TYPE_OP_POST = "operator-postfix";
    exports.TOK_TYPE_WSPACE = "white-space";
    exports.TOK_TYPE_UNKNOWN = "unknown";
    exports.TOK_SUBTYPE_START = "start";
    exports.TOK_SUBTYPE_STOP = "stop";
    // TOK_TYPE_OPERAND
    exports.TOK_SUBTYPE_TEXT = "text";
    exports.TOK_SUBTYPE_NUMBER = "number";
    exports.TOK_SUBTYPE_LOGICAL = "logical";
    exports.TOK_SUBTYPE_ERROR = "error";
    exports.TOK_SUBTYPE_RANGE = "range";
    // TOK_TYPE_OP_IN
    exports.TOK_SUBTYPE_MATH = "math"; // also TOK_SUBTYPE_LOGICAL
    exports.TOK_SUBTYPE_CONCAT = "concatenate";
    exports.TOK_SUBTYPE_INTERSECT = "intersect";
    exports.TOK_SUBTYPE_UNION = "union";
    function f_token(value, type, subtype) {
        this.value = value;
        this.type = type;
        this.subtype = subtype;
    }
    function f_tokens() {
        this.items = new Array();
        this.add = function (value, type, subtype) { if (!subtype)
            subtype = ""; var token = new f_token(value, type, subtype); this.addRef(token); return token; };
        this.addRef = function (token) { this.items.push(token); };
        this.index = -1;
        this.reset = function () { this.index = -1; };
        this.BOF = function () { return (this.index <= 0); };
        this.EOF = function () { return (this.index >= (this.items.length - 1)); };
        this.moveNext = function () { if (this.EOF())
            return false; this.index++; return true; };
        this.current = function () { if (this.index == -1)
            return null; return (this.items[this.index]); };
        this.next = function () { if (this.EOF())
            return null; return (this.items[this.index + 1]); };
        this.previous = function () { if (this.index < 1)
            return null; return (this.items[this.index - 1]); };
    }
    function f_tokenStack() {
        this.items = new Array();
        this.push = function (token) { this.items.push(token); };
        this.pop = function () {
            var token = this.items.pop();
            return (new f_token("", token.type, exports.TOK_SUBTYPE_STOP));
        };
        this.token = function () { return ((this.items.length > 0) ? this.items[this.items.length - 1] : null); };
        this.value = function () { return ((this.token()) ? this.token().value : ""); };
        this.type = function () { return ((this.token()) ? this.token().type : ""); };
        this.subtype = function () { return ((this.token()) ? this.token().subtype : ""); };
    }
    function getTokens(formula) {
        var tokens = new f_tokens();
        var tokenStack = new f_tokenStack();
        var offset = 0;
        var currentChar = function () { return formula.substr(offset, 1); };
        var doubleChar = function () { return formula.substr(offset, 2); };
        var nextChar = function () { return formula.substr(offset + 1, 1); };
        var EOF = function () { return (offset >= formula.length); };
        var token = "";
        var inString = false;
        var inPath = false;
        var inRange = false;
        var inError = false;
        while (formula.length > 0) {
            if (formula.substr(0, 1) == " ")
                formula = formula.substr(1);
            else {
                if (formula.substr(0, 1) == "=")
                    formula = formula.substr(1);
                break;
            }
        }
        var regexSN = /^[1-9]{1}(\.[0-9]+)?E{1}$/;
        while (!EOF()) {
            // state-dependent character evaluation (order is important)
            // double-quoted strings
            // embeds are doubled
            // end marks token
            if (inString) {
                if (currentChar() == "\"") {
                    if (nextChar() == "\"") {
                        token += "\"";
                        offset += 1;
                    }
                    else {
                        inString = false;
                        tokens.add(token, exports.TOK_TYPE_OPERAND, exports.TOK_SUBTYPE_TEXT);
                        token = "";
                    }
                }
                else {
                    token += currentChar();
                }
                offset += 1;
                continue;
            }
            // single-quoted strings (links)
            // embeds are double
            // end does not mark a token
            if (inPath) {
                if (currentChar() == "'") {
                    if (nextChar() == "'") {
                        token += "'";
                        offset += 1;
                    }
                    else {
                        inPath = false;
                    }
                }
                else {
                    token += currentChar();
                }
                offset += 1;
                continue;
            }
            // bracked strings (range offset or linked workbook name)
            // no embeds (changed to "()" by Excel)
            // end does not mark a token
            if (inRange) {
                if (currentChar() == "]") {
                    inRange = false;
                }
                token += currentChar();
                offset += 1;
                continue;
            }
            // error values
            // end marks a token, determined from absolute list of values
            if (inError) {
                token += currentChar();
                offset += 1;
                if ((",#NULL!,#DIV/0!,#VALUE!,#REF!,#NAME?,#NUM!,#N/A,").indexOf("," + token + ",") != -1) {
                    inError = false;
                    tokens.add(token, exports.TOK_TYPE_OPERAND, exports.TOK_SUBTYPE_ERROR);
                    token = "";
                }
                continue;
            }
            // scientific notation check
            if (("+-").indexOf(currentChar()) != -1) {
                if (token.length > 1) {
                    if (token.match(regexSN)) {
                        token += currentChar();
                        offset += 1;
                        continue;
                    }
                }
            }
            // independent character evaulation (order not important)
            // establish state-dependent character evaluations
            if (currentChar() == "\"") {
                if (token.length > 0) {
                    // not expected
                    tokens.add(token, exports.TOK_TYPE_UNKNOWN);
                    token = "";
                }
                inString = true;
                offset += 1;
                continue;
            }
            if (currentChar() == "'") {
                if (token.length > 0) {
                    // not expected
                    tokens.add(token, exports.TOK_TYPE_UNKNOWN);
                    token = "";
                }
                inPath = true;
                offset += 1;
                continue;
            }
            if (currentChar() == "[") {
                inRange = true;
                token += currentChar();
                offset += 1;
                continue;
            }
            if (currentChar() == "#") {
                if (token.length > 0) {
                    // not expected
                    tokens.add(token, exports.TOK_TYPE_UNKNOWN);
                    token = "";
                }
                inError = true;
                token += currentChar();
                offset += 1;
                continue;
            }
            // mark start and end of arrays and array rows
            if (currentChar() == "{") {
                if (token.length > 0) {
                    // not expected
                    tokens.add(token, exports.TOK_TYPE_UNKNOWN);
                    token = "";
                }
                tokenStack.push(tokens.add("ARRAY", exports.TOK_TYPE_FUNCTION, exports.TOK_SUBTYPE_START));
                tokenStack.push(tokens.add("ARRAYROW", exports.TOK_TYPE_FUNCTION, exports.TOK_SUBTYPE_START));
                offset += 1;
                continue;
            }
            if (currentChar() == ";") {
                if (token.length > 0) {
                    tokens.add(token, exports.TOK_TYPE_OPERAND);
                    token = "";
                }
                tokens.addRef(tokenStack.pop());
                tokens.add(",", exports.TOK_TYPE_ARGUMENT);
                tokenStack.push(tokens.add("ARRAYROW", exports.TOK_TYPE_FUNCTION, exports.TOK_SUBTYPE_START));
                offset += 1;
                continue;
            }
            if (currentChar() == "}") {
                if (token.length > 0) {
                    tokens.add(token, exports.TOK_TYPE_OPERAND);
                    token = "";
                }
                tokens.addRef(tokenStack.pop());
                tokens.addRef(tokenStack.pop());
                offset += 1;
                continue;
            }
            // trim white-space
            if (currentChar() == " " || currentChar() == "\t" || currentChar() == "\n") {
                if (token.length > 0) {
                    tokens.add(token, exports.TOK_TYPE_OPERAND);
                }
                token = "";
                while ((currentChar() == " " || currentChar() == "\t" || currentChar() == "\n") && (!EOF())) {
                    token += currentChar();
                    offset += 1;
                }
                tokens.add(token, exports.TOK_TYPE_WSPACE);
                token = "";
                continue;
            }
            // multi-character comparators
            if ((",>=,<=,<>,").indexOf("," + doubleChar() + ",") != -1) {
                if (token.length > 0) {
                    tokens.add(token, exports.TOK_TYPE_OPERAND);
                    token = "";
                }
                tokens.add(doubleChar(), exports.TOK_TYPE_OP_IN, exports.TOK_SUBTYPE_LOGICAL);
                offset += 2;
                continue;
            }
            // standard infix operators
            if (("+-*/^&=><").indexOf(currentChar()) != -1) {
                if (token.length > 0) {
                    tokens.add(token, exports.TOK_TYPE_OPERAND);
                    token = "";
                }
                tokens.add(currentChar(), exports.TOK_TYPE_OP_IN);
                offset += 1;
                continue;
            }
            // standard postfix operators
            if (("%").indexOf(currentChar()) != -1) {
                if (token.length > 0) {
                    tokens.add(token, exports.TOK_TYPE_OPERAND);
                    token = "";
                }
                tokens.add(currentChar(), exports.TOK_TYPE_OP_POST);
                offset += 1;
                continue;
            }
            // start subexpression or function
            if (currentChar() == "(") {
                if (token.length > 0) {
                    tokenStack.push(tokens.add(token, exports.TOK_TYPE_FUNCTION, exports.TOK_SUBTYPE_START));
                    token = "";
                }
                else {
                    tokenStack.push(tokens.add("", exports.TOK_TYPE_SUBEXPR, exports.TOK_SUBTYPE_START));
                }
                offset += 1;
                continue;
            }
            // function, subexpression, array parameters
            if (currentChar() == ",") {
                if (token.length > 0) {
                    tokens.add(token, exports.TOK_TYPE_OPERAND);
                    token = "";
                }
                if (!(tokenStack.type() == exports.TOK_TYPE_FUNCTION)) {
                    tokens.add(currentChar(), exports.TOK_TYPE_OP_IN, exports.TOK_SUBTYPE_UNION);
                }
                else {
                    tokens.add(currentChar(), exports.TOK_TYPE_ARGUMENT);
                }
                offset += 1;
                continue;
            }
            // stop subexpression
            if (currentChar() == ")") {
                if (token.length > 0) {
                    tokens.add(token, exports.TOK_TYPE_OPERAND);
                    token = "";
                }
                tokens.addRef(tokenStack.pop());
                offset += 1;
                continue;
            }
            // token accumulation
            token += currentChar();
            offset += 1;
        }
        // dump remaining accumulation
        if (token.length > 0)
            tokens.add(token, exports.TOK_TYPE_OPERAND);
        var tokens2 = exclude_ws(tokens);
        var tokens3 = infix_to_prefix_operator(tokens2);
        var tokens4 = exclude_noops(tokens3);
        return tokens4;
    }
    exports.getTokens = getTokens;
    function exclude_ws(tokens) {
        // move all tokens to a new collection, excluding all unnecessary white-space tokens
        var bKeepWs = false;
        if (bKeepWs)
            return tokens;
        var tokens2 = new f_tokens();
        var token, token_prev, token_next;
        while (tokens.moveNext()) {
            token = tokens.current();
            token_prev = tokens.previous();
            token_next = tokens.next();
            if (token.type == exports.TOK_TYPE_WSPACE) {
                if ((tokens.BOF()) || (tokens.EOF())) { }
                else if (!(((token_prev.type == exports.TOK_TYPE_FUNCTION) && (token_prev.subtype == exports.TOK_SUBTYPE_STOP)) ||
                    ((token_prev.type == exports.TOK_TYPE_SUBEXPR) && (token_prev.subtype == exports.TOK_SUBTYPE_STOP)) ||
                    (token_prev.type == exports.TOK_TYPE_OPERAND))) { }
                else if (!(((token_next.type == exports.TOK_TYPE_FUNCTION) && (token_next.subtype == exports.TOK_SUBTYPE_START)) ||
                    ((token_next.type == exports.TOK_TYPE_SUBEXPR) && (token_next.subtype == exports.TOK_SUBTYPE_START)) ||
                    (token_next.type == exports.TOK_TYPE_OPERAND))) { }
                else
                    tokens2.add(token.value, exports.TOK_TYPE_OP_IN, exports.TOK_SUBTYPE_INTERSECT);
                continue;
            }
            tokens2.addRef(token);
        }
        return tokens2;
    }
    function infix_to_prefix_operator(tokens2) {
        // switch infix "-" operator to prefix when appropriate, switch infix "+" operator to noop when appropriate, identify operand 
        // and infix-operator subtypes, pull "@" from in front of function names
        var token;
        while (tokens2.moveNext()) {
            token = tokens2.current();
            if ((token.type == exports.TOK_TYPE_OP_IN) && (token.value == "-")) {
                if (tokens2.BOF())
                    token.type = exports.TOK_TYPE_OP_PRE;
                else if (((tokens2.previous().type == exports.TOK_TYPE_FUNCTION) && (tokens2.previous().subtype == exports.TOK_SUBTYPE_STOP)) ||
                    ((tokens2.previous().type == exports.TOK_TYPE_SUBEXPR) && (tokens2.previous().subtype == exports.TOK_SUBTYPE_STOP)) ||
                    (tokens2.previous().type == exports.TOK_TYPE_OP_POST) ||
                    (tokens2.previous().type == exports.TOK_TYPE_OPERAND))
                    token.subtype = exports.TOK_SUBTYPE_MATH;
                else
                    token.type = exports.TOK_TYPE_OP_PRE;
                continue;
            }
            if ((token.type == exports.TOK_TYPE_OP_IN) && (token.value == "+")) {
                if (tokens2.BOF())
                    token.type = exports.TOK_TYPE_NOOP;
                else if (((tokens2.previous().type == exports.TOK_TYPE_FUNCTION) && (tokens2.previous().subtype == exports.TOK_SUBTYPE_STOP)) ||
                    ((tokens2.previous().type == exports.TOK_TYPE_SUBEXPR) && (tokens2.previous().subtype == exports.TOK_SUBTYPE_STOP)) ||
                    (tokens2.previous().type == exports.TOK_TYPE_OP_POST) ||
                    (tokens2.previous().type == exports.TOK_TYPE_OPERAND))
                    token.subtype = exports.TOK_SUBTYPE_MATH;
                else
                    token.type = exports.TOK_TYPE_NOOP;
                continue;
            }
            if ((token.type == exports.TOK_TYPE_OP_IN) && (token.subtype.length == 0)) {
                if (("<>=").indexOf(token.value.substr(0, 1)) != -1)
                    token.subtype = exports.TOK_SUBTYPE_LOGICAL;
                else if (token.value == "&")
                    token.subtype = exports.TOK_SUBTYPE_CONCAT;
                else
                    token.subtype = exports.TOK_SUBTYPE_MATH;
                continue;
            }
            if ((token.type == exports.TOK_TYPE_OPERAND) && (token.subtype.length == 0)) {
                if (isNaN(parseFloat(token.value)))
                    if ((token.value == 'TRUE') || (token.value == 'FALSE'))
                        token.subtype = exports.TOK_SUBTYPE_LOGICAL;
                    else
                        token.subtype = exports.TOK_SUBTYPE_RANGE;
                else
                    token.subtype = exports.TOK_SUBTYPE_NUMBER;
                continue;
            }
            if (token.type == exports.TOK_TYPE_FUNCTION) {
                if (token.value.substr(0, 1) == "@")
                    token.value = token.value.substr(1);
                continue;
            }
        }
        tokens2.reset();
        return tokens2;
    }
    function exclude_noops(tokens2) {
        // move all tokens to a new collection, excluding all noops
        var tokens = new f_tokens();
        while (tokens2.moveNext()) {
            if (tokens2.current().type != exports.TOK_TYPE_NOOP)
                tokens.addRef(tokens2.current());
        }
        tokens.reset();
        return tokens;
    }
});
/*
    Purpose: transform Excel syntax into a valid Javascript/C++ syntax.
     - operator "<>" becomes "!="
     - operator "=" becomes "=="
     - postfix operator "%" becomes a division by 100
     - immediate array notation {1,2;21,22} becomes standard array [ [1,2], [21,22] ]
     - ranges names becomes pure alphanumeric variable names
     - little known range union and intersection operator are handled
*/
define("excel_formula_transform", ["require", "exports", "excel_formula_parse"], function (require, exports, parser) {
    "use strict";
    Object.defineProperty(exports, "__esModule", { value: true });
    function parse_and_transfrom_test() {
        var xl_formulas = [
            ['=A1:B2', '=A1_B2'],
            ['{=A1+C2}', '=A1+C2'],
            ['="a"+1.2+TRUE+A1+xyz+#VALUE!', '="a"+1.2+true+A1+xyz+null'],
            ['="a"&"b"', '="a"+"b"'],
            ['=50', '=50'],
            ['=50%', '=50/100'],
            ['=3.1E-24-2.1E-24', '=3.1E-24-2.1E-24'],
            ['=1+3+5', '=1+3+5'],
            ['=3 * 4 + \n\t 5', '=3*4+5'],
            ['=1-(3+5) ', '=1-(3+5)'],
            ['=1=5', '=1==5'],
            ['=1<>5', '=1!=5'],
            ['=2^8',],
            ['=TRUE', '=true'],
            ['=FALSE', '=false'],
            ['={1.2,"s";TRUE,FALSE}', '=[[1.2,"s"],[true,false]]'],
            ['=$A1',],
            ['=$B$2',],
            ['=Sheet1!A1', '=Sheet1_A1'],
            ['=\'Another sheet\'!A1', '=Another_sheet_A1'],
            ['=[Book1]Sheet1!$A$1', '=_Book1_Sheet1_$A$1'],
            ['=[data.xls]sheet1!$A$1', '=_data_xls_sheet1_$A$1'],
            ['=\'[Book1]Another sheet\'!$A$1', '=_Book1_Another_sheet_$A$1'],
            ['=\'[Book No2.xlsx]Sheet1\'!$A$1', '=_Book_No2_xlsx_Sheet1_$A$1'],
            ['=\'[Book No2.xlsx]Another sheet\'!$A$1', '=_Book_No2_xlsx_Another_sheet_$A$1'],
            //[ '=[\'my data.xls\']\'my sheet\'!$A$1', ],
            ['=SUM(B5:B15)', '=SUM(B5_B15)'],
            ['=SUM(B5:B15,D5:D15)', '=SUM(B5_B15,D5_D15)'],
            ['=SUM(B5:B15 A7:D7)', '=SUM(B5_B15 inter A7_D7)'],
            ['=SUM(sheet1!$A$1:$B$2)', '=SUM(sheet1_$A$1_$B$2)'],
            //[ '=SUM((A:A 1:1))', '=SUM((A_A inter 1_1))' ], // 1_1 NOT VALID !!!!!
            //[ '=SUM((A:A,1:1))', '=SUM((A_A union 1_1))' ],
            ['=SUM((A:A A1:B1))', '=SUM((A_A inter A1_B1))'],
            ['=SUM(D9:D11,E9:E11,F9:F11)', '=SUM(D9_D11,E9_E11,F9_F11)'],
            ['=SUM((D9:D11,(E9:E11,F9:F11)))', '=SUM((D9_D11 union (E9_E11 union F9_F11)))'],
            //[ '=IF(P5=1.0,"NA",IF(P5=2.0,"A",IF(P5=3.0,"B",IF(P5=4.0,"C",IF(P5=5.0,"D",IF(P5=6.0,"E",IF(P5=7.0,"F",IF(P5=8.0,"G"))))))))', 
            //  "=(P5==1.0?\"NA\":(P5==2.0?\"A\":(P5==3.0?\"B\":(P5==4.0?\"C\":(P5==5.0?\"D\":(P5==6.0?\"E\":(P5==7.0?\"F\":(P5==8.0?\"G\":undefined))))))))" ],
            ['=SUM(B2:D2*B3:D3)', '=SUM(B2_D2*B3_D3)'],
            //[ '=SUM( 123 + SUM(456) + IF(DATE(2002,1,6), 0, IF(ISERROR(R[41]C[2]),0, IF(R13C3>=R[41]C[2],0, '
            //	+'  IF(AND(R[23]C[11]>=55,R[24]C[11]>=20), R53C3, 0))))', 
            //	'=123+SUM(456)+(DATE(2002,1,6)?0:(ISERROR(R_41_C_2_)?0:(R13C3>=R_41_C_2_?0:(AND(R_23_C_11_>=55,R_24_C_11_>=20)?R53C3:0))))' ],
            //[ '=IF("a"={"a","b";"c",#N/A;-1,TRUE}, "yes", "no") &   "  more ""test"" text"', 
            //  '=(\"a\"==[[\"a\",\"b\"],[\"c\",null],[-1,true]]?\"yes\":\"no\")+\"  more \\\"test\\\" text\"' ],
            ['=+ AName- (-+-+-2^6) = {"A","B"} + @SUM(R1C1) + (@ERROR.TYPE(#VALUE!) = 2)',
                '=AName-(---2^6)==[[\"A\",\"B\"]]+SUM(R1C1)+(ERROR.TYPE(null)==2)'],
        ];
        var js_formulas = [
            '="a string \\n\\t with a few escape chars"',
            '="a string \\" with a tricky escape char"',
            '=true && false',
            "=A18==0 && F4==0"
        ];
        var nb_errors = 0;
        function test(formulas, prms) {
            var formula_xl_js = [];
            var fcts = {}, vars = {};
            for (var f = 0; f < formulas.length; f++) {
                var xlformula = Array.isArray(formulas[f]) ? formulas[f][0] : formulas[f];
                var jsexpected = Array.isArray(formulas[f]) ? formulas[f][1] || xlformula : xlformula;
                var tokens, ast;
                var jsformula = "=" + parse_and_transfrom(xlformula, vars, fcts, prms);
                formula_xl_js.push([xlformula, jsformula]);
                if (jsformula != jsexpected)
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
        if (nb_errors)
            throw new Error("parse_and_transfrom_test: " + nb_errors + " errors.");
    }
    exports.parse_and_transfrom_test = parse_and_transfrom_test;
    function parse_and_transfrom(xlformula, vars, fcts, opt_prms) {
        var prms = opt_prms || {};
        var n = xlformula.length;
        if (n > 3 && xlformula[0] == "{" && xlformula[n - 1] == "}")
            xlformula = xlformula.substr(1, n - 2);
        // ExtendedJS grammar (with range and vector operations)
        // --or-- Excel grammar to plain JS grammar conversion 
        var tokens = parser.getTokens(xlformula);
        //info_msg("TOKENS:\n"+JSON.stringify(tokens,null,4));
        try {
            var ast = build_token_tree(tokens);
        }
        catch (e) {
            throw new Error("Error while parsing " + xlformula + "\n" + (e.stack || e));
        }
        //info_msg("TREE:\n"+JSON.stringify(ast,null,4));
        var jsformula = excel_to_js_formula(ast, vars, fcts, prms);
        return jsformula;
    }
    exports.parse_and_transfrom = parse_and_transfrom;
    function build_token_tree(tokens) {
        var stack = [];
        var stack_top = { args: [[]] };
        while (tokens.moveNext()) {
            var token = tokens.current();
            if (token.subtype == parser.TOK_SUBTYPE_START) {
                stack.push(stack_top);
                stack_top = token;
                stack_top.args = [[]];
            }
            else if (token.type == parser.TOK_TYPE_ARGUMENT) {
                stack_top.args.push([]);
            }
            else if (token.subtype == parser.TOK_SUBTYPE_STOP) {
                var tmp = stack_top;
                stack_top = stack.pop();
                var nb_args = stack_top.args.length;
                stack_top.args[nb_args - 1].push(tmp);
            }
            else {
                var nb_args = stack_top.args.length;
                stack_top.args[nb_args - 1].push(token);
            }
        }
        if (stack_top.args.length != 1)
            throw new Error("root formula contains multiple expressions.\n" + JSON.stringify(stack_top, null, 4));
        return stack_top.args[0];
    }
    function excel_to_js_formula(tokens, opt_vars, opt_fcts, opt_prms) {
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
                        var valid_id = token.value.replace(/[^A-Za-z0-9$]/g, "_");
                        for (var k in vars)
                            if (valid_id == vars[k])
                                throw new Error("argument name collision between " + JSON.stringify(k) + " and " + JSON.stringify(token.value));
                        vars_count++;
                        vars[token.value] = valid_id;
                    }
                    return vars[token.value];
                default:
                    throw new Error("unhandled subtype " + token.subtype + "\n" + JSON.stringify(token, null, 4));
            }
        }
        function excel_to_js_array_operand(token) {
            var arr = new Array(token.args.length);
            for (var i = 0; i < arr.length; i++) {
                if (token.args[i].length != 1)
                    throw new Error("ARRAY is expected to be a list of ARRAYROW but found an expressions");
                arr[i] = new Array(token.args[i][0].args.length);
                for (var j = 0; j < arr[i].length; j++) {
                    var tmp = excel_to_js_formula(token.args[i][0].args[j]);
                    arr[i][j] = JSON.parse(tmp);
                }
            }
            return JSON.stringify(arr);
        }
        function excel_to_js_function_call_formula(token) {
            fcts[token.value] = (fcts[token.value] || 0) + 1;
            var res = token.value + "(";
            var nb_args = token.args.length;
            for (var i = 0; i < nb_args; i++) {
                var expr = token.args[i];
                res += excel_to_js_formula(expr, vars, fcts);
                res += (i + 1 < nb_args ? "," : "");
            }
            return res + ")";
        }
        function excel_to_js_if_then_else_trigram(token) {
            var nb_args = token.args.length;
            if (nb_args < 2 || nb_args > 3)
                throw new Error("expect 2 or 3 arguments");
            var res = "(" + excel_to_js_formula(token.args[0], vars, fcts) +
                "?" + excel_to_js_formula(token.args[1], vars, fcts) +
                ":" + (nb_args == 3 ? excel_to_js_formula(token.args[2], vars, fcts) : "undefined") +
                ")";
            return res;
        }
        function excel_to_js_operator(token) {
            switch (token.subtype) {
                case parser.TOK_SUBTYPE_MATH:
                    return token.value;
                case parser.TOK_SUBTYPE_LOGICAL:
                    if (token.value == "=")
                        return "==";
                    else if (token.value == "<>")
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
                    throw new Error("unhandled subtype " + token.subtype + "\n" + JSON.stringify(token));
            }
        }
        var res = "";
        for (var t = 0; t < tokens.length; t++) {
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
                    if (token.value != "%")
                        throw new Error("unhandled token " + token.type + " " + token.value);
                    res += "/100";
                    break;
                // operands 
                case "operand":
                    res += excel_to_js_operand(token);
                    break;
                case "function":
                case "subexpression":
                    if (!prms.xjs && token.value == "ARRAY")
                        res += excel_to_js_array_operand(token);
                    else
                        res += excel_to_js_function_call_formula(token);
                    break;
                case "white-space":
                    res += token.value;
                    break;
                default:
                    throw new Error("unhandled " + token.type);
            }
        }
        return res;
    }
    function info_msg(msg) {
        console.log(msg);
    }
    parse_and_transfrom_test();
});
/**
 * Created by Charles on 13/03/2016.
 */
define("deps_graph", ["require", "exports", "excel_range_parse", "excel_formula_parse", "excel_range_parse"], function (require, exports, rp, parser, rng) {
    "use strict";
    Object.defineProperty(exports, "__esModule", { value: true });
    function info_msg(msg) {
        console.log(msg);
    }
    function dbg(name, variable) {
        info_msg(name + ":\n" + JSON.stringify(variable, null, 4));
    }
    function deps_graph_test() {
        var assert = require('assert');
        var no_edges = [
            ['A2', '="a"&"b"'],
            ['A3', '=50'],
            ['A4', '=50%']
        ];
        var adjacency_list = formulas_list_to_adjacency_list(no_edges);
        adjacency_list_to_dot(adjacency_list, 'C:/temp/graph1.svg');
        assert.equal(JSON.stringify(adjacency_list), JSON.stringify({
            '{"cell0":{"col":1,"row":2}}': [],
            '{"cell0":{"col":1,"row":3}}': [],
            '{"cell0":{"col":1,"row":4}}': []
        }));
        var simple_directed_cells_only = [
            ['W1', '=$A1'],
            ['U1', '=$B$2'],
            ['A1', '=$B$5']
        ];
        var adjacency_list = formulas_list_to_adjacency_list(simple_directed_cells_only);
        adjacency_list_to_dot(adjacency_list, 'C:/temp/graph2.svg');
        assert.equal(JSON.stringify(adjacency_list), JSON.stringify({
            '{"cell0":{"col":23,"row":1}}': ['{"cell0":{"col":1,"row":1}}'],
            '{"cell0":{"col":21,"row":1}}': ['{"cell0":{"col":2,"row":2}}'],
            '{"cell0":{"col":1,"row":1}}': ['{"cell0":{"col":2,"row":5}}']
        }));
        /*var all_formulas = [
            ['A1', '="a"+1.2+TRUE+A1+xyz+#VALUE!'], // all operands
            ['A2', '="a"&"b"'],	// & operator
            ['A3', '=50'],
            ['A4', '=50%'], 	// % operator
            ['A5', '=3.1E-24-2.1E-24'],
            ['A6', '=1+3+5'],
            ['A7', '=3 * 4 + \n\t 5'],
            ['A8', '=1-(3+5) '],
            ['A9', '=1=5'], 	// = operator
            ['B1', '=1<>5'],	// <> operator
            ['C1', '=2^8'], 	// ^ operator
            ['D1', '=TRUE'],
            ['E1', '=FALSE'],
            ['Q1', '={1.2,"s";TRUE,FALSE}'],	// array notation
            ['W1', '=$A1'],
            ['R1', '=$B$2'],
            ['T1', '=Sheet1!A1'],
            ['Y1', '=\'Another sheet\'!A1'],
            ['U1', '=[Book1]Sheet1!$A$1'],
            ['I1', '=[data.xls]sheet1!$A$1'],
            ['O1', '=\'[Book1]Another sheet\'!$A$1'],
            ['P1', '=\'[Book No2.xlsx]Sheet1\'!$A$1'],
            ['S1', '=\'[Book No2.xlsx]Another sheet\'!$A$1'],
            ['F1', '=[data.xls]sheet1!$A$1'],
            //['A1', '=[\'my data.xls\'],\'my sheet\'!$A$1'],
            ['G1', '=SUM(B5:B15)'],
            ['H1', '=SUM(B5:B15,D5:D15)'],
            ['J1', '=SUM(B5:B15 A7:D7)'],
            ['K1', '=SUM(sheet1!$A$1:$B$2)'],
            ['L1', '=SUM((A:A 1:1))'],
            ['Z1', '=SUM((A:A,1:1))'],
            ['X1', '=SUM((A:A A1:B1))'],
            ['V1', '=SUM(D9:D11,E9:E11,F9:F11)'],
            ['N1', '=SUM((D9:D11,(E9:E11,F9:F11)))'],
            ['M1', '=IF(P5=1.0,"NA",IF(P5=2.0,"A",IF(P5=3.0,"B",IF(P5=4.0,"C",IF(P5=5.0,"D",IF(P5=6.0,"E",IF(P5=7.0,"F",IF(P5=8.0,"G"))))))))'],
            ['M2', '=SUM(B2:D2*B3:D3)'],  // array formula
            ['M3', '=SUM( 123 + SUM(456) + IF(DATE(2002,1,6), 0, IF(ISERROR(R[41]C[2]),0, IF(R13C3>=R[41]C[2],0, '
                +'  IF(AND(R[23]C[11]>=55,R[24]C[11]>=20), R53C3, 0))))'],
            ['M4', '=IF("a"={"a","b";"c",#N/A;-1,TRUE}, "yes", "no") &   "  more ""test"" text"'],
            ['M5', '=+ AName- (-+-+-2^6) = {"A","B"} + @SUM(R1C1) + (@ERROR.TYPE(#VALUE!) = 2)'],
            ['M6', '=IF(R13C3>DATE(2002,1,6),0,IF(ISERROR(R[41]C[2]),0,IF(R13C3>=R[41]C[2],0, IF(AND(R[23]C[11]>=55,R[24]C[11]>=20),R53C3,0))))'],
            ['M7', '=IF(R[39]C[11]>65,R[25]C[42],ROUND((R[11]C[11]*IF(OR(AND(R[39]C[11]>=55, R[40]C[11]>=20),AND(R[40]C[11]>=20,R11C3="YES")),R[44]C[11],R[43]C[11]))+(R[14]C[11] *IF(OR(AND(R[39]C[11]>=55,R[40]C[11]>=20),AND(R[40]C[11]>=20,R11C3="YES")), R[45]C[11],R[43]C[11])),0))']
        ];
        
        var adjacency_list = formulas_list_to_adjacency_list(all_formulas);
    
        adjacency_list_to_dot(adjacency_list, 'C:/temp/graph1.dot');*/
    }
    exports.deps_graph_test = deps_graph_test;
    function formulas_list_to_adjacency_list(formulas /* [ [cell, formula_in_cell], ...]*/) {
        var adjacency_list = {};
        for (var f = 0; f < formulas.length; f++) {
            var cell = formulas[f][0];
            var xlformula = formulas[f][1];
            //try {
            var tokens = parser.getTokens(xlformula);
            // info_msg("TOKENS:\n"+JSON.stringify(tokens,null,4));
            var cell_deps = tokens_to_adjacent_list(tokens["items"]);
            var key = range_to_key(rng.parse_range(cell, rp.ReferenceStyle.A1));
            if (adjacency_list.hasOwnProperty(key))
                throw "Dependencies of cell " + key + " read more than once !";
            adjacency_list[key] = cell_deps;
            //} catch (e) {
            //	info_msg("ERROR: "+e+"\n"+e.stack);
            //}
        }
        //info_msg("adjacency_list:\n"+JSON.stringify(adjacency_list,null,4));
        return adjacency_list;
    }
    function tokens_to_adjacent_list(tokens) {
        var deps = [];
        for (var i = 0; i < tokens.length; ++i) {
            var token = tokens[i];
            if (token.subtype === "range") {
                var range_as_key = range_to_key(rng.parse_range(token.value, rp.ReferenceStyle.A1));
                if (deps.indexOf(range_as_key) != -1)
                    deps.push(range_as_key);
            }
        }
        return deps;
    }
    /*function eq_range_obj(left, right) {
        // no need for  (left.hasOwnProperty("workbook") && right.hasOwnProperty("workbook"))
        // undefined !== undefined --> false
        if (left["workbook"] !== right["workbook"])
            return false;
        if (left["sheet"] !== right["sheet"])
            return false;
        if (left["cell1"] !== right["cell1"])
            return false;
        if (left["cell0"] !== right["cell0"])
            return false;
        return true;
    }*/
    function range_to_key(range) {
        var dict_to_print = range;
        // remove key arg
        dict_to_print.arg = undefined;
        // remove keys col_abs and row_abs
        if (dict_to_print["cell0"] !== undefined) {
            dict_to_print["cell0"].row_abs = undefined;
            dict_to_print["cell0"].col_abs = undefined;
        }
        if (dict_to_print["cell1"] !== undefined) {
            dict_to_print["cell1"].row_abs = undefined;
            dict_to_print["cell1"].col_abs = undefined;
        }
        return JSON.stringify(dict_to_print);
    }
    function adjacency_list_to_dot(adj_list, output_path) {
        var fs = require('fs');
        var path = require('path');
        var child_process = require('child_process'); // to call dotty
        var output = 'digraph G {\n';
        for (var vertex in adj_list) {
            output += '\t"' + vertex.replace(/"/g, "'") + '"';
            var parents = adj_list[vertex];
            for (var i = 0; i < parents.length; ++i)
                output += ' -> "' + parents[i].replace(/"/g, "'") + '"';
            output += ';\n';
        }
        output += '}';
        var output_dir = path.dirname(output_path);
        var filename = path.basename(output_path);
        var extension = path.extname(output_path).replace(/\./g, "");
        var dot_file = path.join(output_dir, filename.replace(extension, 'dot'));
        fs.writeFile(dot_file, output, function (err) {
            if (err)
                throw err;
        });
        var path_to_dot = '"C:\\Program Files (x86)\\Graphviz2.38\\bin\\dot.exe"';
        var cmd = path_to_dot + " -T" + extension + ' ' + dot_file + ' -o ' + output_path;
        child_process.exec(cmd, function (error, stdout, stderr) {
            //dbg('stdout', stdout);
            //dbg('stderr', stderr);
            if (stderr) {
                throw stderr;
            }
            else {
                info_msg(output_path + " successfully created !");
            }
        });
    }
    exports.adjacency_list_to_dot = adjacency_list_to_dot;
    function contains(array, obj) {
        var i = this.length;
        while (i--) {
            if (this[i] === obj) {
                return true;
            }
        }
        return false;
    }
});
define("global_scope", ["require", "exports"], function (require, exports) {
    "use strict";
    Object.defineProperty(exports, "__esModule", { value: true });
    function sum(rng) {
        var total = 0;
        for (var i = 0; i < rng.length; i++)
            for (var j = 0; j < rng[i].length; j++)
                total += rng[i][j] || 0;
        return total;
    }
    exports.sum = sum;
    function trsp(rng) {
        var n = rng.length;
        var m = n == 0 ? 0 : rng[0].length;
        var res = new Array(m);
        for (var j = 0; j < m; j++) {
            res[j] = new Array(n);
            for (var i = 0; i < n; i++)
                res[j][i] = rng[i][j];
        }
        return res;
    }
    exports.trsp = trsp;
    function op_gen(f, a, b) {
        if (!Array.isArray(a) && !Array.isArray(b)) {
            // scalar-scalar operation
            return f(a, b);
        }
        else if (Array.isArray(a) && Array.isArray(b)) {
            // array-array operation
            throw new Error("op_gen: array-array operation not implemented");
        }
        else {
            throw new Error("op_gen: array-scalar operation not implemented");
        }
    }
});
/*
function op_add(a,b)  { return op_gen((a,b) => a+b, a, b); };     global["op_add"] = op_add;
function op_sub(a,b)  { return op_gen((a,b) => a-b, a, b); };     global["op_sub"] = op_sub;
function op_mult(a,b) { return op_gen((a,b) => a*b, a, b); };     global["op_mult"] = op_mult;
function op_div(a,b)  { return op_gen((a,b) => a/b, a, b); };     global["op_div"] = op_div;
function op_pow(a,b)  { return op_gen((a,b) => a^b, a, b); };     global["op_pow"] = op_pow;
function op_eq(a,b)   { return op_gen((a,b) => a===b, a, b); };   global["op_eq"] = op_eq;
function op_neq(a,b)  { return op_gen((a,b) => a!==b, a, b); };   global["op_neq"] = op_neq;
function op_gt(a,b)   { return op_gen((a,b) => a>b, a, b); };     global["op_gt"] = op_gt;
function op_lt(a,b)   { return op_gen((a,b) => a<b, a, b); };     global["op_lt"] = op_lt;
function op_gte(a,b)  { return op_gen((a,b) => a>=b, a, b); };    global["op_gte"] = op_gte;
function op_lte(a,b)  { return op_gen((a,b) => a<=b, a, b); };    global["op_lte"] = op_lte;
*/
/*
    Purpose: transform syntaxically valid Javascript formula expression into syntaxically correct expression.
     - operator overload for arrays: we want to support [1,2,3]+[5,6,7]
     - transform IF pseudo-function into trigram operator (to respect evaluation of the correct alternative only)
     - transform AND and OR pseudo-functions into && and || operators (to respect lazy evaluation of boolean expressions)
*/
define("js_formula_transform", ["require", "exports"], function (require, exports) {
    "use strict";
    Object.defineProperty(exports, "__esModule", { value: true });
    var acorn = require("acorn");
    var escodegen = require("escodegen");
    // if an object is a infix + operator
    function operator_override_ast(ast) {
        for (var k in ast) {
            if (Array.isArray(ast[k])) {
                for (var i = 0; i < ast[k].length; i++)
                    ast[k][i] = operator_override_ast(ast[k][i]);
            }
            else if (ast[k] !== null && typeof ast[k] == "object") {
                ast[k] = operator_override_ast(ast[k]);
            }
            else {
                // do nothing 
            }
        }
        if (ast.type == "BinaryExpression") {
            var op = void 0;
            switch (ast.operator) {
                // Excel operators
                case "+":
                    op = "add";
                    break;
                case "-":
                    op = "sub";
                    break;
                case "*":
                    op = "mult";
                    break;
                case "/":
                    op = "div";
                    break;
                case "^":
                    op = "pow";
                    break;
                case "=":
                    op = "eq";
                    break;
                case "<>":
                    op = "neq";
                    break;
                case ">":
                    op = "gt";
                    break;
                case "<":
                    op = "lt";
                    break;
                case ">=":
                    op = "gte";
                    break;
                case "<=":
                    op = "lte";
                    break;
                case "&":
                    op = "add";
                    break;
                // Excel operators transcompiles into JS operator so that JS parser could produce an AST
                case "==":
                    op = "eq";
                    break;
                case "!=":
                    op = "neq";
                    break;
                case "&&":
                    op = "inter";
                    break;
                case "||":
                    op = "union";
                    break;
            }
            op = "op_" + op;
            return {
                "type": "CallExpression",
                "callee": {
                    "type": "Identifier",
                    "name": op
                },
                "arguments": [
                    ast.left,
                    ast.right
                ]
            };
        }
        return ast;
    }
    function extract_variables_and_functions(ast, vars, fcts, excl, isProperty, isFct) {
        if (ast.type == "Identifier") {
            var id = ast.name;
            if (isProperty)
                return;
            if (excl[id])
                return;
            if (isFct) {
                fcts[id] = (fcts[id] || 0) + 1;
            }
            else {
                var valid_id = id;
                vars[id] = valid_id;
            }
        }
        else if (ast.type == "CallExpression") {
            // callee can be any expression
            extract_variables_and_functions(ast.callee, vars, fcts, excl, isProperty, true);
            // and arguments
            extract_variables_and_functions({ arguments: ast.arguments }, vars, fcts, excl, false, false);
        }
        else if (ast.type == "MemberExpression") {
            // we recurse on object part..
            extract_variables_and_functions(ast.object, vars, fcts, excl, isProperty, false);
            // and property part
            extract_variables_and_functions(ast.property, vars, fcts, excl, !ast.computed, false);
        }
        else if (ast.type == "FunctionExpression") {
            return; // no need to go deeper 
        }
        else if (ast.type == "VariableDeclarator") {
            extract_variables_and_functions(ast.init, vars, fcts, excl, isProperty, isFct);
        }
        else if (ast.type == "Property") {
            if (ast.computed)
                throw new Error("unhandled computed property"); // what does it mean a computed key-value pair 
            // ignore the key which can appear as an id rather as a string in the shorthand syntax { a: 123 }
            // just process the value
            extract_variables_and_functions(ast.value, vars, fcts, excl, isProperty, isFct);
        }
        else {
            for (var k in ast) {
                if (Array.isArray(ast[k])) {
                    for (var i = 0; i < ast[k].length; i++)
                        extract_variables_and_functions(ast[k][i], vars, fcts, excl, isProperty, isFct);
                }
                else if (ast[k] === null) {
                    // do nothing (null is an object)
                }
                else if (typeof ast[k] == "object") {
                    extract_variables_and_functions(ast[k], vars, fcts, excl, isProperty, isFct);
                }
                else {
                    // do nothing 
                }
            }
        }
    }
    function parse_and_transfrom(expr, opt_vars, opt_fcts, opt_prms) {
        var prms = opt_prms || {};
        var vars = opt_vars || {};
        var fcts = opt_fcts || {};
        // HACK TO ALLOW { a: 1, b: 2 }
        if (expr[0] == "{" || expr[0] == "[")
            expr = "(" + expr + ")";
        var ast = acorn.parse(expr, { ranges: false });
        //console.log(JSON.stringify(ast, null, 4));
        var excl = { "Math": 1, "setInterval": 1 };
        extract_variables_and_functions(ast, vars, fcts, excl, false, false);
        //console.log("expr: " + expr);
        //console.log("vars: " + JSON.stringify(vars, null, 4));
        //console.log("fcts: " + JSON.stringify(fcts, null, 4));
        var ast_op_over = operator_override_ast(ast);
        //console.log(JSON.stringify(ast_op_over, null, 4));
        // generate code
        var code = escodegen.generate(ast_op_over, { comment: false });
        // console.log("code: "+code);
        return code;
    }
    exports.parse_and_transfrom = parse_and_transfrom;
    function parse_and_transfrom_test() {
        var expressions = [
            ["o.a[i].f(x).c;", undefined,
                { o: "o", i: "i", x: "x" }],
            ["f(x).g(y).a[i];", undefined,
                { x: "x", y: "y", i: "i" }, { f: 1 }],
            ["a[i].b[j].c(x);", undefined,
                { a: "a", i: "i", j: "j", x: "x" }],
            ["f(x)[i].g[j](y);", undefined,
                { x: "x", i: "i", j: "j", y: "y" }, { f: 1 }],
            ["$F$3.replace('views.js', $F$6.views[$I4].file);",
                "$F$3.replace('views.js', $F$6.views[$I4].file);",
                { $F$3: "$F$3", $F$6: "$F$6", $I4: "$I4" }],
            ["var a = { ref: abc, src: 123 }",
                "var a = {\n    ref: abc,\n    src: 123\n};",
                { "abc": "abc" }],
            ["{ ref: abc, src: 123+456 }",
                "({\n    ref: abc,\n    src: op_add(123, 456)\n});",
                { "abc": "abc" }],
            ["[ 1.2, abc, true ]",
                "[\n    1.2,\n    abc,\n    true\n];",
                { "abc": "abc" }],
            ["setInterval(function () {\n    set_range_input_and_fire('C4', new Date());\n    model.evaluate_sheet(hot_render);\n}, 1000+B2);",
                "setInterval(function () {\n    set_range_input_and_fire('C4', new Date());\n    model.evaluate_sheet(hot_render);\n}, op_add(1000, B2));",
                { "B2": "B2" }],
            ["(function (x, y) { return x*y; })(A1,B1);",
                "(function (x, y) {\n    return op_mult(x, y);\n}(A1, B1));",
                { "A1": "A1", "B1": "B1" }],
            ['Math.abs(A1_B1.ref.addr)',
                'Math.abs(A1_B1.ref.addr);',
                { "A1_B1": "A1_B1" }],
            ['5*a.c + 4*b;',
                'op_add(op_mult(5, a.c), op_mult(4, b));',
                { "a": "a", "b": "b" }],
            ['42 + 3 + f( 5 + g( 77+99)); // answer',
                'op_add(op_add(42, 3), f(op_add(5, g(op_add(77, 99)))));',
                {}, { f: 1, g: 1 }],
            ['moment(C3).hours.ago(C5).fromNow();',
                'moment(C3).hours.ago(C5).fromNow();',
                { "C3": "C3", "C5": "C5" }, { moment: 1 }] // arguments order does not matter
        ];
        var nb_errors = 0;
        for (var e in expressions) {
            var expr_in = expressions[e][0];
            var expr_ref = expressions[e][1] || expr_in;
            var vars_ref = expressions[e][2] || undefined;
            var vars_out = {};
            var fcts_ref = expressions[e][3] || {};
            var fcts_out = {};
            var expr_out = parse_and_transfrom(expr_in, vars_out, fcts_out);
            if (expr_out != expr_ref)
                nb_errors++;
            if (vars_ref && JSON.stringify(vars_out) != JSON.stringify(vars_ref))
                nb_errors++;
            if (fcts_ref && JSON.stringify(fcts_out) != JSON.stringify(fcts_ref))
                nb_errors++;
        }
        if (nb_errors)
            throw new Error("parse_and_transfrom_test: " + nb_errors + " errors.");
    }
    exports.parse_and_transfrom_test = parse_and_transfrom_test;
});
define("lexer", ["require", "exports"], function (require, exports) {
    "use strict";
    Object.defineProperty(exports, "__esModule", { value: true });
    var dflt_token_re = {
        string: /^("([^"\\]|""|\\.)*"|'([^'\\]|''|\\.)*')/,
        number: /^[+-]?[0-9]+\.?[0-9]*(e[+-]?[0-9]+)?/,
        id: /^[@$a-zA-Z~][\w_:.]*/,
        op: /^(=|\+|-|\*|\/|\(|\)|\{|\}|\[|\]|<|>|,|;|:|!|%|&|.)/,
        ws: /^\s+/,
    };
    function lexer(txt_input, opt_token_re) {
        var txt = txt_input;
        var token_re = opt_token_re || dflt_token_re;
        var tokens = new Array();
        for (; txt.length > 0;) {
            //info_msg("\ntxt: "+txt);
            var token = undefined;
            for (var t in token_re) {
                var re = token_re[t];
                var res = re.exec(txt);
                //info_msg(t+": "+JSON.stringify(res));
                if (res != null) {
                    token = { type: t, txt: res[0] };
                    break;
                }
            }
            if (token) {
                //info_msg(JSON.stringify(token));
                tokens.push(token);
                txt = txt.substring(token.txt.length);
            }
            else {
                if (txt != "")
                    throw "unexpected token when hitting: " + txt + "\nfull expression: " + txt_input + "";
                // done !!
            }
        }
        return tokens;
    }
    exports.lexer = lexer;
    function handle_at_operator(tokens) {
        // use '@' prefix operator
        // to reference a graph node as opposed to a sheet range 
        // (it means we can manipulate objects or functions, not just scalars)
        for (var _i = 0, tokens_1 = tokens; _i < tokens_1.length; _i++) {
            var token = tokens_1[_i];
            if (token.txt[0] == '@')
                token.txt = /* "graph_nodes." + */ token.txt.substr(1);
        }
        return tokens;
    }
    exports.handle_at_operator = handle_at_operator;
    function rebuild_source_text(tokens) {
        var txt = "";
        for (var _i = 0, tokens_2 = tokens; _i < tokens_2.length; _i++) {
            var t = tokens_2[_i];
            txt += t.txt;
        }
        return txt;
    }
    exports.rebuild_source_text = rebuild_source_text;
    function lexer_test() {
        info_msg("lexer_test");
        var expressions = [
            '="B1 is not" & A1',
            '="this formula countains \\" and "" double quotes"',
            "='this formula countains \\' and '' single quotes'",
            "window.setViews=function(v) { window.views = v; }",
        ];
        var nb_errors = 0;
        for (var i = 0; i < expressions.length; i++) {
            var expr_in = expressions[i];
            var expr_tok = lexer(expr_in);
            var expr_out = rebuild_source_text(expr_tok);
            if (expr_out != expr_in) {
                nb_errors++;
                info_msg(JSON.stringify(expr_in));
                info_msg(JSON.stringify(expr_tok));
            }
        }
        if (nb_errors)
            throw new Error("lexer_test: " + nb_errors + " errors.");
    }
    exports.lexer_test = lexer_test;
    function info_msg(msg) {
        console.log(msg);
    }
    lexer_test();
});
define("sheet_exec", ["require", "exports", "excel_range_parse", "excel_range_parse", "excel_formula_transform", "js_formula_transform", "lexer"], function (require, exports, rp, range_parser, xlfx, jsfx, lexer) {
    "use strict";
    Object.defineProperty(exports, "__esModule", { value: true });
    /*
    function evaluate_formula(src, vars) {
        var var_names=[], var_values=[];
        for (var v in vars) {
            var_names.push(v);
            var_values.push(vars[v]);
        }
    
        var var_names_and_function_src = var_names.concat(["return "+src+";"]);
        var fct = Function.prototype.constructor.apply(undefined, var_names_and_function_src);
    
        var res = fct.apply(fct, var_values);
        return res;
    }
    */
    function spreadsheet_exec_test() {
        var fs = require('fs');
        var path = require('path');
        var folder_path = __dirname + '/../work/';
        var filepaths = fs.readdirSync(folder_path);
        for (var i = 0; i < filepaths.length; i++) {
            var filepath = path.resolve(folder_path + filepaths[i]);
            if (filepath.indexOf(".xjson") == -1)
                continue;
            info_msg("Executing " + filepath);
            var txt = fs.readFileSync(filepath, 'utf8');
            var data = JSON.parse(txt);
            var sheet = data.input;
            if (data.named_ranges)
                for (var k in data.named_ranges)
                    sheet[k] = data.named_ranges[k];
            //try {
            var node_values = sheet_exec(sheet, {
                "input_syntax": data.input_syntax || "excel",
                "xjs": true,
            });
            var tmp = filepath.split('/');
            tmp[tmp.length - 2] = "ref";
            var filepath_out = tmp.join('/');
            tmp.pop();
            var folder_out = tmp.join('/');
            //if (!fs.existsSync(folder_out)) fs.mkdirSync(folder_out);
            //fs.writeFileSync(filepath_out, JSON.stringify({input: node_values},null,4));
            data.ref_output = node_values;
            fs.writeFileSync(filepath, JSON.stringify(data, null, 2));
            //} catch (e) {
            //	info_msg("sheet "+JSON.stringify(sheet, null, 4));
            //	throw new Error("Error evaluating "+filepath+"\n"+(e.stack||e)+"\n");
            //}
        }
        info_msg("\nspreadsheet_exec_test DONE.");
    }
    exports.spreadsheet_exec_test = spreadsheet_exec_test;
    function sheet_exec(sheet, prms) {
        var indent = "";
        info_msg("sheet range and formular parsing...");
        var node_eval = {};
        var node_values = {};
        for (var r in sheet) {
            var val = sheet[r];
            if (val === "")
                continue;
            var target = range_parser.parse_range(r, rp.ReferenceStyle.A1);
            var isFormula = val.substr && val[0] == "=";
            var isFormulaArray = val.substr && val[0] == "{" && val[1] == "=" && val[val.length - 1] == "}";
            if (isFormula || isFormulaArray) {
                //try {
                var formula = isFormula ? val.substr(1) : val.substr(2, val.length - 3);
                var expr = void 0;
                var vars = {}, fcts = {};
                switch (prms.input_syntax) {
                    case "excel":
                        expr = xlfx.parse_and_transfrom(formula, vars, fcts, prms);
                        break;
                    case "javascript":
                        var formula2 = lexer.rebuild_source_text(lexer.handle_at_operator(lexer.lexer(formula)));
                        expr = jsfx.parse_and_transfrom(formula2, vars, fcts, prms);
                        break;
                    default:
                        throw new Error("unhandled input syntax.");
                }
                var ids = [];
                var args = [];
                for (var addr in vars) {
                    if (["Array", "Date", "Math", "Number", "Object", "String", "arguments", "caller", "this", "undefined", "JSON", "window"].indexOf(addr) != -1) {
                        warn_msg("WARN: reserved keyword '" + addr + "' used in " + r);
                        continue;
                    }
                    ids.push(vars[addr]);
                    args.push(addr);
                }
                //var func = new Function(ids, "try { var res="+expr+"; return res; } catch(e) { throw new Error('"+r+" '+(e.stack||e)+'\\n'); }");
                //var func = new Function(ids, "try { var res="+expr+"; return res; } catch(e) { throw new Error('"+r+" '+(e)+'\\n'); }");
                //var func = Function(ids, "try { var res="+expr+"; return res; } catch(e) { throw new Error('"+r+" '+(e)+'\\n'); }");
                var func = new Function(ids.join(","), "var res=" + expr + "; return res;");
                info_msg("expr: " + expr);
                if (target.end && !isFormulaArray) {
                    for (var i = target.row; i <= target.end.row; i++)
                        for (var j = target.col; j <= target.end.col; j++) {
                            var tmp_args = new Array(args.length);
                            for (var a = 0; a < args.length; a++)
                                tmp_args[a] = range_parser.stringify_range(move_range(range_parser.parse_range(args[a], rp.ReferenceStyle.A1), i - target.row, j - target.col), rp.ReferenceStyle.A1);
                            var tmp_target = { row: i, col: j };
                            var tmp_id = range_parser.stringify_range(tmp_target, rp.ReferenceStyle.A1);
                            node_eval[tmp_id] = {
                                target: tmp_target,
                                expr: expr,
                                fct: func,
                                args: tmp_args
                            };
                        }
                }
                else {
                    node_eval[r] = {
                        target: target,
                        expr: expr,
                        fct: func,
                        args: args
                    };
                }
                /*
                } catch (e) {
                    node_eval[r] = {
                        target: target,
                        expr: val,
                        fct: function(x) { return function() { throw new Error("ill formated expression: "+x); }; },
                        error: e.stack || (""+e)
                        };
                }
                */
            }
            else {
                node_eval[r] = {
                    target: target,
                    value: val,
                };
            }
        }
        info_msg("sheet static dependency tree ...");
        var formula_targets = Object.keys(node_eval);
        var formula_missing = sheet_static_dependencies(node_eval, true);
        info_msg("  formula_targets : #" + formula_targets.length + " " + formula_targets);
        info_msg("  formula_missing : #" + formula_missing.length + " " + formula_missing);
        info_msg("sheet evaluate...");
        /*
        info_msg("evaluation: "+JSON.stringify(node_eval, null, 4));
        */
        for (var rng in node_eval) {
            var res = evaluate_node(rng, node_eval, indent + " ");
            node_values[rng] = res;
            try {
                info_msg(rng + " value is " + JSON.stringify(res));
            }
            catch (e) {
                info_msg(rng + " value is " + res + " (not JSON stringifiable).");
            }
        }
        var adjacency_list = extract_adjacency_list(node_eval);
        var calculation_order = static_calculation_order(adjacency_list);
        return node_values;
    }
    function extract_adjacency_list(nodes) {
        var adjacency_list = {};
        for (var n in nodes) {
            if (!nodes[n])
                continue;
            info_msg("extract_adjacency_list node " + JSON.stringify(n));
            var args = nodes[n].args || [];
            var deps = nodes[n].deps ? Object.keys(nodes[n].deps) : [];
            if (JSON.stringify(args) != JSON.stringify(deps) && args.length && deps.length)
                throw new Error("args (for formula) and deps (for range aggregation) both present and inconsistent");
            adjacency_list[n] = deps.length ? deps : args;
        }
        return adjacency_list;
    }
    function static_calculation_order(adjacency_list) {
        var calculation_order = [];
        var adjacency_list_length = Object.keys(adjacency_list).length;
        for (var it = 0;; it++) {
            var nb_added_at_this_iteration = 0;
            for (var n in adjacency_list) {
                if (calculation_order.indexOf(n) != -1)
                    continue;
                var deps = adjacency_list[n];
                var depsOk = true;
                for (var d = 0; d < deps.length; d++) {
                    if (calculation_order.indexOf(deps[d]) == -1) {
                        depsOk = false;
                        break;
                    }
                }
                if (!depsOk)
                    continue;
                calculation_order.push(n);
                nb_added_at_this_iteration++;
            }
            if (calculation_order.length == adjacency_list_length)
                return calculation_order;
            if (nb_added_at_this_iteration === 0)
                throw new Error("circular reference detected");
        }
    }
    function sheet_static_dependencies(nodes, bCreateMissing) {
        var missing_deps = {};
        for (var n in nodes) {
            var args = nodes[n].args;
            if (args)
                for (var i = 0; i < args.length; i++) {
                    if (!(args[i] in nodes))
                        missing_deps[args[i]] = true;
                }
        }
        return Object.keys(missing_deps);
    }
    function create_on_the_fly_node(id, indent) {
        info_msg(indent + "create_on_the_fly_node: id=" + id);
        var rng = range_parser.parse_range(id, rp.ReferenceStyle.A1);
        //if (rng.end && (rng.end.row!=rng.row || rng.end.col!=rng.col || rng.end.row_abs || rng.col_abs)) {
        return { rng: rng };
        //} else {
        //	info_msg(indent+"create_on_the_fly_node: id="+id+" not needed");
        //	return; // throw new Error(id+" is not a multi-cell range");
        //}
    }
    function evaluate_node(id, nodes, indent) {
        info_msg(indent + "evaluate_node: id=" + id);
        if (id.split('.').shift() == "Math")
            return (Function("return " + id + ";"))();
        if (!(id in nodes))
            nodes[id] = create_on_the_fly_node(id, indent + " ");
        var node = nodes[id];
        var isNulNode = node === undefined || node == null;
        var isValNode = node && node.fct === undefined && node.rng === undefined;
        var isRngNode = node && node.fct === undefined && node.rng !== undefined;
        var isFctNode = node && node.fct !== undefined && node.rng === undefined;
        if (isNulNode)
            return null;
        else if (isValNode)
            return node.value;
        else if (isRngNode)
            return evaluate_range(id, nodes, indent + " ");
        else if (isFctNode)
            return evaluate_depth_first(id, nodes, indent + " ");
        else
            throw new Error(id + " unknown node type. " + JSON.stringify(node));
    }
    function get_range_dependencies(id, nodes) {
        var rng = nodes[id].rng;
        if (rng.reference_to_named_range)
            throw new Error("get_range_dependencies called on '" + id + "' named range (should only be called on A1:B2 style addresses)");
        // build list of dependencies
        var deps = nodes[id].deps;
        if (!deps) {
            deps = {};
            for (var n in nodes) {
                var node = nodes[n];
                if (!node.target)
                    continue;
                var inter = ranges_intersection(rng, node.target);
                if (inter)
                    deps[n] = inter;
            }
        }
        return deps;
    }
    function ranges_intersection(rng1, rng2) {
        if (rng1.reference_to_named_range || rng2.reference_to_named_range) {
            warn_msg("ranges_intersection called on named range (should only be called on A1:B2 style addresses)");
            return undefined;
        }
        var beg1 = rng1;
        var end1 = rng1.end || rng1;
        var beg2 = rng2;
        var end2 = rng2.end || rng2;
        var inter = {
            row: Math.max(beg1.row, beg2.row),
            col: Math.max(beg1.col, beg2.col)
        };
        var inter_end = {
            row: Math.min(end1.row, end2.row),
            col: Math.min(end1.col, end2.col)
        };
        if (inter.row > inter_end.row || inter.col > inter_end.col)
            // no interection 
            return undefined;
        if (inter.row == inter_end.row || inter.col == inter_end.col)
            // single cell intersection
            return inter;
        inter.end = inter_end;
        return inter;
    }
    function evaluate_range(id, nodes, indent) {
        info_msg(indent + "evaluate_range: id=" + id);
        // cached results
        if (nodes[id].value) {
            info_msg(indent + "evaluate_range: " + id + " value (cached) is " + JSON.stringify(res));
            return nodes[id].value;
        }
        var deps = get_range_dependencies(id, nodes);
        nodes[id].deps = deps;
        // evaluate dependencies
        for (var n in deps)
            evaluate_node(n, nodes, indent + " ");
        // assemble range 
        var rng = nodes[id].rng;
        var beg1 = rng;
        var end1 = rng.end || rng;
        var nb_row = end1.row - beg1.row + 1;
        var nb_col = end1.col - beg1.col + 1;
        // create an empty range
        var res = new Array(nb_row);
        for (var i = 0; i < nb_row; i++) {
            res[i] = new Array(nb_col);
            for (var j = 0; j < nb_col; j++)
                res[i][j] = null;
        }
        for (var n in nodes[id].deps) {
            // !! compatible with cells containing arrays ??
            var tmp = to_array_2d(evaluate_node(n, nodes, indent + " "));
            info_msg(indent + "evaluate_range: " + id + " dependency " + n + " value is " + JSON.stringify(tmp));
            var tmp_row = nodes[n].target.row;
            var tmp_col = nodes[n].target.col;
            var inter = nodes[id].deps[n];
            info_msg(indent + "evaluate_range: " + id + " intersection with " + n + " is " + JSON.stringify(inter));
            if (!inter.end)
                inter.end = { row: inter.row, col: inter.col };
            for (var i = inter.row; i <= inter.end.row; i++)
                for (var j = inter.col; j <= inter.end.col; j++)
                    res[i - rng.row][j - rng.col] = tmp[i - tmp_row][j - tmp_col];
        }
        // if asked for a scalar, then return a scalar
        if (!rng.end && res.length == 1 && res[0].length == 1)
            res = res[0][0];
        // cache results 
        nodes[id].value = res;
        info_msg(indent + "evaluate_range: " + id + " value (calculated) is " + JSON.stringify(res));
        return res;
    }
    function evaluate_depth_first(id, nodes, indent) {
        info_msg(indent + "evaluate_depth_first: id=" + id);
        var node = nodes[id];
        // cached results
        if (node.value) {
            info_msg(indent + "evaluate_depth_first: " + id + " value (cached) is " + JSON.stringify(node.value));
            return node.value;
        }
        // evaluate arguments first
        var args = [];
        var arg_map = {}; // for info only
        if (node.args) {
            args = new Array(node.args.length);
            for (var a = 0; a < node.args.length; a++) {
                args[a] = evaluate_node(node.args[a], nodes, indent + " ");
                arg_map[node.args[a]] = args[a];
            }
        }
        info_msg(indent + "evaluate_depth_first: " + id + " evaluating '" + node.expr + "' with args " + JSON.stringify(arg_map));
        node.t0 = new Date();
        try {
            var res = node.fct.apply(this, args);
            // cache reults
            node.value = res;
            node.error = undefined;
        }
        catch (e) {
            node.value = undefined;
            node.error = "" + e;
        }
        node.t1 = new Date();
        //try {
        info_msg(indent + "evaluate_depth_first: " + id + " value (calculated) is " + JSON.stringify(res));
        //} catch (e) {	info_msg(indent+"evaluate_depth_first: "+id+" value (calculated) is "+res); }
        return res;
    }
    function move_range(rng, i, j) {
        var tmp = JSON.parse(JSON.stringify(rng));
        if (tmp.row !== undefined && !tmp.abs_row)
            tmp.row += i;
        if (tmp.col !== undefined && !tmp.abs_col)
            tmp.col += j;
        if (tmp.end)
            tmp.end = move_range(tmp.end, i, j);
        return tmp;
    }
    function to_array_2d(v) {
        if (!Array.isArray(v))
            return to_array_2d([v]);
        if (v.length == 0 || !Array.isArray(v[0]))
            return [v];
        return v;
    }
    function info_msg(msg) {
        console.log(msg);
    }
    function warn_msg(msg) {
        console.warn(msg);
    }
});
// run tests if this file is called directly
// if (require.main === module)
//	spreadsheet_exec_test();
define("run_tests", ["require", "exports", "lexer", "excel_range_parse", "excel_formula_transform", "js_formula_transform", "sheet_exec"], function (require, exports, lex, rp, xlf, jsf, shx) {
    "use strict";
    Object.defineProperty(exports, "__esModule", { value: true });
    lex.lexer_test();
    rp.parse_range_bijection_test();
    xlf.parse_and_transfrom_test();
    jsf.parse_and_transfrom_test();
    shx.spreadsheet_exec_test();
});
//deps.deps_graph_test();
/*
function spreadsheet_scope() {
    this.A1 = 13
    this.f = function() { info_msg(this.A1); }
}
spreadsheet_scope.prototype.fun = function() { info_msg(this.A1); }
spreadsheet_scope.prototype.fun2 = function() { info_msg(A1); }

function scope_test() {
    var obj = {
        get A1 () {
            return 11;
        }
    }
    
    function f() { info_msg(this.A1); }
    f.call({ A1: 7 })
    f.call(obj)

    var ss = new spreadsheet_scope;
    ss.f();
    ss.fun();
    ss.fun2();

    ss.A2 = 17;
    ss.g = function() { info_msg(this.A2); }
    ss.g();

    ss.A3 = 23;
    ss.h = new Function("return this.A3;");
    info_msg(ss.h());
    
    var A5 = 5
    function f5() { info_msg(A5); }
    f5();

    var A6 = 6
    f6 = function() { info_msg(A6); }
    f6();
    
    var A7 = 7
    f7a = new Function("info_msg(A7);");
    f7 = f7a.bind(this);
    f7();
    
}

function info_msg(msg) {
    console.log(msg);
}
scope_test();

*/ 
//# sourceMappingURL=build.js.map