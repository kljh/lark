"""

Dependencies:
pip install lark-parser

"""

import os, sys, traceback
from lark import Lark, Transformer, v_args


grammar = r"""
    ?start: statement*

    ?statement: ( option_stmt
        | function_definition_stmt
        | statement_expr ) ( ":" | NEWLINE )+

    ?statement_expr: ( on_error_stmt
        | dim_stmt
        | sub_call_stmt
        | assign_stmt
        | if_stmt
        | for_stmt
        | loop_stmt )

    !option_stmt: "Option"i ( "Explicit"i | "Base"i NUMBER )

    ?on_error_stmt: "On"i "Error"i ( "GoTo"i  (NAME|"0") | "Resume"i next )

    ?dim_stmt: ( "ReDim"i | "Dim"i | ( dim_stmt "," ) ) NAME [ "(" ranges ")" ] [ as NAME ]
    ?ranges: range | ranges "," range
    ?range: expr | ( expr to expr )

    ?function_definition_stmt.1: ( "Sub" | "Function" ) NAME "(" function_args ")" ( ":" | NEWLINE )  statement* "End" ( "Sub" | "Function" )
    ?function_args: function_args "," function_arg | function_arg |
    function_arg: [ "Optional" ] ["ByRef"] NAME [ as NAME ] [ "=" expr ]

    !assign_stmt: [ set ] lvalue "=" expr
    ?lvalue: ( name | func_call )
    // !! func_call is lvalue because we can have my2darray(1, 2) = "abc"

    ?expr: name | VB_STRING | NUMBER | func_call | expr_binary | expr_other
    ?expr_binary: expr ( "=" | "<>" | ">" | "<" | ">=" | "<=" | "+" | "-" | "*" | "/" | "&" | "And"i | "AndAlso"i | "Or"i | "OrElse"i | "Xor"i ) expr
    ?expr_other: ( "-" expr ) | ( "Not" expr ) | ( "(" expr ")" )

    ?name: ( VB_NAME  "." )*  VB_NAME

    !sub_call_stmt: name func_args
        | call name "(" func_args ")"

    !func_call: name "(" func_args ")"
    !func_args: func_args "," [ NAME ":=" ] func_arg | [ NAME ":=" ] func_arg |
    !func_arg: expr

    ?if_stmt: if expr then statement_expr [ else statement_expr]
        | if expr then NEWLINE+ statement+ ( elseif expr then NEWLINE+ statement+ )* [ else NEWLINE+ statement+] end if

    ?for_stmt: for each VB_NAME in expr NEWLINE+ statement+ next [VB_NAME]
        | for VB_NAME "=" expr to expr [ step expr ] NEWLINE+ statement+ next [VB_NAME]

    ?loop_stmt: "While"i expr NEWLINE+ statement+ "Wend"i
        | "Do"i  "While"i expr NEWLINE+ statement+ "Loop"i
        | "Do"i  "Until"i expr NEWLINE+ statement+ "Loop"i

    ?sum: product
        | sum "+" product   -> add
        | sum "-" product   -> sub

    ?product: atom
        | product "*" atom  -> mul
        | product "/" atom  -> div

    ?atom: NUMBER           -> number
         | "-" atom         -> neg
         | name             -> var
         | "(" sum ")"

    ?set : "Set"
    ?as : "AS"| "as" | "As"
    ?call: "CALL" | "call" | "Call"
    ?if : "If"
    ?elseif : "ElseIf"
    ?then : "Then"
    ?else : "Else"
    ?end : "End"

    ?for: "For"i
    ?to: "To"i
    ?step: "Step"i
    ?each : "Each"i
    ?in : "In"i
    ?next : "Next"i

    CONTINUE_LINE: /_\r*\n/
    %ignore CONTINUE_LINE

    COMMENT:  "'" /[^\n]*/ "\n"
    VB_STRING : "\"" /([^"]|"")*/ "\""
    %ignore COMMENT

    // %import common.CNAME -> NAME
    %import common.LETTER
    %import common.DIGIT
    NAME: LETTER ("_"|LETTER|DIGIT)*

    %import common.ESCAPED_STRING
    %import common.NUMBER
    %import common.WS_INLINE
    %import common.NEWLINE
    %ignore WS_INLINE

    // $ String, % Integer, F& Long, # Double, ! Single, @ Currency
    VB_NAME: NAME /[$%&#!@]?/

"""

vbparser = Lark(grammar, parser='lalr')  # 'lalr' or 'earley'

def generate_all(src_root):
    import os

    filepaths = []
    for dirpath, dirnames, filenames in os.walk(src_root):
        filepaths = filepaths + [ os.path.join(dirpath, f) for f in filenames if f.endswith(".bas") ]
    print("#files", len(filepaths))

    for f in filepaths:
        generate_one(f)

def generate_one(src_path):
    parse_one(src_path)

def parse_one(src_path):
    with open(src_path, "r") as f:
        lines = f.readlines()
        lines = lines[1:]
        txt = "".join(lines)

    print("src_path", src_path)

    try:
        ast = vbparser.parse(txt)
        print(ast.pretty())
    except Exception as e:
        traceback.print_exc()

        tmp = "".join([ f"{i+1}  {line}" for i, line in enumerate(lines) ])
        print(tmp)
        #raise e

def test():
    print(calc("a = 1+2"))
    print(calc("1+a*-3"))



if __name__ == '__main__':
    # test()

    src_root =r"..\..\xjs\vba"
    generate_all(src_root)

    generate_one(r"test.bas")
    #generate_one(src_root + r"\modWalkDir.bas")
    #generate_one(src_root + r"\modHttpRequest.bas")

