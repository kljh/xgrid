/*
	Purpose: transform syntaxically valid Javascript formula expression into syntaxically correct expression.
	 - operator overload for arrays: we want to support [1,2,3]+[5,6,7]
     - transform IF pseudo-function into trigram operator (to respect evaluation of the correct alternative only)
     - transform AND and OR pseudo-functions into && and || operators (to respect lazy evaluation of boolean expressions)
*/

var acorn = require("acorn");
var escodegen = require("escodegen");
var assert = require("assert");

// if an object is a infix + operator
function operator_override_ast(ast) {
    for (var k in ast) {
        if (Array.isArray(ast[k])) {
            for (var i=0; i<ast[k].length; i++)
                ast[k][i] = operator_override_ast(ast[k][i]);
        } else if (typeof ast[k]=="object") {
            ast[k] = operator_override_ast(ast[k]);
        } else {
            // do nothing 
        }
    }

    if (ast.type=="BinaryExpression") {
        let op;
        switch (ast.operator) {
            // Excel operators
            case "+": op = "add"; break;
            case "-": op = "sub"; break;
            case "*": op = "mult"; break;
            case "/": op = "div"; break;
            case "^": op = "pow"; break;
            case ">": op = "gt"; break;
            case "<": op = "lt"; break;
            case "=": op = "equal"; break;
            case ">=": op = "gte"; break;
            case "<=": op = "lte"; break;
            case "<>": op = "ne"; break;
            case "&": op = "concat"; break;
            // Excel operators transcompiles into JS operator so that JS parser could produce an AST
            case "==": op = "equal"; break;
            case "!=": op = "ne"; break;
            case "&&": op = "inter"; break;
            case "||": op = "union"; break;
        }
        op = "op_"+op;
            
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
        }
    }
    return ast;
} 


export function parse_and_transfrom(expr) {
    var ast = acorn.parse(expr, { ranges: false });
    //console.log(JSON.stringify(ast, null, 4));

    var ast_op_over = operator_override_ast(ast);
    //console.log(JSON.stringify(ast_op_over, null, 4));

    // generate code
    var code = escodegen.generate(ast_op_over, { comment: false });
    // console.log("code: "+code);

    return code;
}

export function parse_and_transfrom_test() {
    var expressions = [
        [ 'var x = 42 + 3 + f( 5 + g( 77+99)); // answer', 
          'var x = op_add(op_add(42, 3), f(op_add(5, g(op_add(77, 99)))));' ]
    ]

    var nb_errors = 0;
    for (let e in expressions) {
        var expr_in = expressions[e][0];
        var expr_ref = expressions[e][1];
        var expr_out = parse_and_transfrom(expr_in);
        if (expr_out!=expr_ref)
            nb_errors++;
    }

    assert(nb_errors==0);
}