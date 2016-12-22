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
        } else if (ast[k]!==null && typeof ast[k]=="object") {
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
            case "=": op = "eq"; break;
            case "<>": op = "neq"; break;
            case ">": op = "gt"; break;
            case "<": op = "lt"; break;
            case ">=": op = "gte"; break;
            case "<=": op = "lte"; break;
            case "&": op = "add"; break;
            // Excel operators transcompiles into JS operator so that JS parser could produce an AST
            case "==": op = "eq"; break;
            case "!=": op = "neq"; break;
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


function extract_variables_and_functions(ast, vars, fcts, excl, isProperty, isFct) : void {

    if (ast.type=="Identifier") {
        var id = ast.name;
        
        if (isProperty) 
            return; 

        if (excl[id])
            return;

        if (isFct) { 
            fcts[id] = (fcts[id]||0) + 1;
        } else {
            var valid_id = id;
            vars[id] = valid_id;
        }

    } else if (ast.type=="CallExpression") {
        // callee can be any expression
        extract_variables_and_functions(ast.callee, vars, fcts, excl, isProperty, true);
        // and arguments
        extract_variables_and_functions({ arguments: ast.arguments }, vars, fcts, excl, false, false);
    } else if (ast.type=="MemberExpression") {
        // we recurse on object part..
        extract_variables_and_functions(ast.object, vars, fcts, excl, isProperty, false);
        // and property part
        extract_variables_and_functions(ast.property, vars, fcts, excl, !ast.computed, false);  
    } else if(ast.type=="FunctionExpression") {
        return; // no need to go deeper 
    } else if (ast.type=="VariableDeclarator") {
        extract_variables_and_functions(ast.init, vars, fcts, excl, isProperty, isFct);       
    } else {
        for (var k in ast) {
            if (Array.isArray(ast[k])) {
                for (var i=0; i<ast[k].length; i++)
                    extract_variables_and_functions(ast[k][i], vars, fcts, excl, isProperty, isFct);
            } else if (ast[k]===null) {
                // do nothing (null is an object)
            } else if (typeof ast[k]=="object") {
                extract_variables_and_functions(ast[k], vars, fcts, excl, isProperty, isFct);
            } else {
                // do nothing 
            }
        }
    }
}

export function parse_and_transfrom(expr, opt_vars?, opt_fcts?, opt_prms?) {
    var prms = opt_prms || {};
	var vars = opt_vars || {};
	var fcts = opt_fcts || {};
	
    // HACK TO ALLOW { a: 1, b: 2 }
    if (expr[0]=="{" || expr[0]=="[")
        expr = "("+expr+")";

    var ast = acorn.parse(expr, { ranges: false });
    //console.log(JSON.stringify(ast, null, 4));

    var excl = { "Math":1, "setInterval":1 };
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

export function parse_and_transfrom_test() {
    var expressions = [
        [ "o.a[i].f(x).c;", undefined, 
          { o: "o", i: "i", x: "x" } ],
        [ "f(x).g(y).a[i];", undefined, 
          { x: "x", y: "y", i: "i" }, { f:1 } ],
        [ "a[i].b[j].c(x);", undefined, 
          { a: "a", i: "i", j: "j", x: "x" } ],
        [ "f(x)[i].g[j](y);", undefined, 
          { x: "x", i: "i", j: "j", y: "y" }, { f:1 } ],
        [ "$F$3.replace('views.js', $F$6.views[$I4].file);",
          "$F$3.replace('views.js', $F$6.views[$I4].file);",
          { $F$3: "$F$3", $F$6: "$F$6", $I4: "$I4" }],
        [ "var a = { ref: true, src: 123 }",
          "var a = {\n    ref: true,\n    src: 123\n};",
          { "ref": "ref", "src": "src" }],
        [ "{ ref: true, src: 123 }", // same as above without assignment, fails in Acorn if not surrounded by parentheses
          "({\n    ref: true,\n    src: 123\n});",
          { "ref": "ref", "src": "src" }],
        [ "[ 1.2, abc, true ]", // same as above without assignment, fails in Acorn if not surrounded by parentheses
          "[\n    1.2,\n    abc,\n    true\n];",
          { "abc": "abc" }],
        [ "setInterval(function () {\n    set_range_input_and_fire('C4', new Date());\n    model.evaluate_sheet(hot_render);\n}, 1000+B2);",
          "setInterval(function () {\n    set_range_input_and_fire('C4', new Date());\n    model.evaluate_sheet(hot_render);\n}, op_add(1000, B2));",
          { "B2": "B2" }],
        [ "(function (x, y) { return x*y; })(A1,B1);",
          "(function (x, y) {\n    return op_mult(x, y);\n}(A1, B1));",
           { "A1": "A1", "B1": "B1" }],
        [ 'Math.abs(A1_B1.ref.addr)',
          'Math.abs(A1_B1.ref.addr);', 
          { "A1_B1": "A1_B1" }],
        [ '5*a.c + 4*b;',
          'op_add(op_mult(5, a.c), op_mult(4, b));', 
          { "a": "a", "b": "b" }],
        [ '42 + 3 + f( 5 + g( 77+99)); // answer', 
          'op_add(op_add(42, 3), f(op_add(5, g(op_add(77, 99)))));',
          {}, { f:1, g:1 }],
        [ 'moment(C3).hours.ago(C5).fromNow();',
          'moment(C3).hours.ago(C5).fromNow();',
          { "C3": "C3", "C5": "C5" }, { moment:1 }] // arguments order does not matter
    ]

    var nb_errors = 0;
    for (let e in expressions) {
        var expr_in = expressions[e][0];
        var expr_ref = expressions[e][1] || expr_in;
        var vars_ref = expressions[e][2] || undefined;
        var vars_out = {};
        var fcts_ref = expressions[e][3] || {};
        var fcts_out = {};
        var expr_out = parse_and_transfrom(expr_in, vars_out, fcts_out);
        if (expr_out!=expr_ref) {
            nb_errors++;
        }
        if (vars_ref && JSON.stringify(vars_out)!=JSON.stringify(vars_ref))
            nb_errors++;
        if (fcts_ref && JSON.stringify(fcts_out)!=JSON.stringify(fcts_ref))
            nb_errors++;
    }

    assert(nb_errors==0);
}