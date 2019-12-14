"""

Dependencies:
pip install lark-parser

"""

import os, sys, traceback
from lark import Lark, Transformer, v_args


grammar = r"""
    ?start: NEWLINE* statements

    !statements: ( ( ( statement | statements_inline ) smnt_new_line ) | ( COMMENT_EOL NEWLINE* ) | ( goto_anchor_stmt NEWLINE*  ) )*
    !statements_inline: statement ( ":" statement )+

    smnt_new_line: COMMENT_EOL NEWLINE* | NEWLINE+

    ?statement: ( option_stmt
        | preprocessor_if_stmt
        | function_declare_stmt
        | function_definition_stmt
        | on_error_stmt
        | goto_stmt
        | dim_stmt
        | sub_call_stmt
        | assign_stmt
        | if_stmt
        | for_stmt
        | loop_stmt
        | with_stmt
        | open_stmt
        | close_stmt
        | print_stmt )

    !option_stmt: "Option"i ( "Explicit"i | "Base"i NUMBER )

    ?on_error_stmt: "On"i "Error"i ( "GoTo"i  (NAME|"0") | "Resume"i next )
    !goto_stmt: "GoTo"i  VB_NAME
    !goto_anchor_stmt: VB_NAME GOTO_ANCHOR_EOL

    !dim_stmt: ( "Static" | private | "ReDim"i | "Dim"i | ( dim_stmt "," ) ) VB_NAME [ "(" ranges ")" ] [ as [ "New"i ] type_name ]
    ?ranges: range | ranges "," range
    ?range: expr | ( expr to expr )

    function_declare_stmt: ("Public"i|"Private"i)? "Declare"i "PtrSafe"i? ( "Sub"i | "Function"i ) NAME "Lib"i VB_STRING "(" function_args ")"  [ as NAME ]

    ?function_definition_stmt.2: ("Public"i|"Private"i)? ( "Sub"i | "Function"i ) NAME "(" function_args ")" [ as NAME ] smnt_new_line  statements end ( "Sub" | "Function" )
    ?function_args: function_args "," function_arg | function_arg |
    function_arg: [ "Optional" ] ["ByRef"|"ByVal"] NAME [ as NAME ] [ "=" expr ]

    // ("Const"i|"Public"i|"Private"i)?
    !assign_stmt: [ set ] lvalue "="  expr
    ?lvalue: "."? ( name | func_call )
    // !! func_call is lvalue because we can have my2darray(1, 2) = "abc"

    ?expr: name | VB_STRING | NUMBER | func_call | expr_binary | expr_other
    ?expr_binary: expr ( "=" | "<>" | ">" | "<" | ">=" | "<=" | "Is"i | "+" | "-" | "*" | "/" | "&" | "." | "And"i | "AndAlso"i | "Or"i | "OrElse"i | "Xor"i ) expr
    ?expr_other: ( "-" expr ) | ( "." expr ) | ( "Not" expr ) | ( "(" expr ")" ) | ( "New"i expr )

    !type_name: NAME
    ?name: (NAME|VB_NAME)  ( "." (NAME|VB_NAME) )*

    !sub_call_stmt: "."? name func_args
        | call name ( "(" func_args ")" ) ?

    !func_call: name "(" func_args ")"
    !func_args: func_args "," func_arg | func_args "," | "," func_args | func_arg |
    !func_arg: [ VB_NAME ":=" ] expr

    ?if_stmt: if expr then statement [ else statement ]
        | if expr then smnt_new_line statements ( elseif expr then smnt_new_line statements )* [ else smnt_new_line statements ] end if

    !preprocessor_if_stmt: "#If" expr then smnt_new_line statements ( "#" elseif expr then smnt_new_line statements )* [ "#Else" smnt_new_line statements ] "#End If"

    ?for_stmt: for each VB_NAME in expr smnt_new_line statements next [VB_NAME]
        | for VB_NAME "=" expr to expr [ step expr ] smnt_new_line statements next [VB_NAME]

    ?loop_stmt: "While"i expr smnt_new_line statements "Wend"i
        | "Do"i  "While"i expr smnt_new_line statements "Loop"i
        | "Do"i  "Until"i expr smnt_new_line statements "Loop"i
        | "Do"i  smnt_new_line statements "Loop"i "While"i expr

    ?with_stmt: with expr smnt_new_line statements end with
    !with_prefix: "."?

    open_stmt: "Open"i expr [open_mode_clause] [open_access_clause] [open_lock_clause] "As"i file_number [open_len_clause]
    open_mode_clause : for (  "Append" | "Binary" | "Input" | "Output" | "Random" )
    open_access_clause : "Access" ( "Read" | "Write" | ("Read" "Write") )
    open_lock_clause :  "Shared" |  ( "Lock" "Read" ) | ("Lock" "Write") |  ("Lock" "Read" "Write")
    open_len_clause : "Len" "=" expr
    close_stmt: "Close"i file_number
    print_stmt: "Print"i file_number "," func_args
    file_number : "#" expr

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

    ?private: "Private"i
    ?set : "Set"i
    ?as : "As"i
    ?call: "Call"i

    ?if : "If"i
    ?elseif : "ElseIf"i
    ?then : "Then"i
    ?else : "Else"i
    ?end : "End"i

    ?for: "For"i
    ?to: "To"i
    ?step: "Step"i
    ?each : "Each"i
    ?in : "In"i
    ?next : "Next"i

    ?with : "With"i

    GOTO_ANCHOR_EOL: /:\r*\n/


    CONTINUE_LINE: /_\r*\n/
    %ignore CONTINUE_LINE

    // PREPROCESSOR_LINE: /#[^\n]*(\r?\n\r?)+/

    COMMENT_EOL : /'[^\n]*\n/

    //COMMENT:  /'[^\n]*\n/
    VB_STRING : "\"" /([^"]|"")*/ "\""
    //%ignore COMMENT

    // %import common.CNAME -> NAME
    %import common.LETTER
    %import common.DIGIT
    NAME: LETTER ("_"|LETTER|DIGIT)*

    %import common.NUMBER
    %import common.NEWLINE
    %import common.WS_INLINE
    %ignore WS_INLINE

    // $ String, % Integer, & Long, # Double, ! Single, @ Currency
    VB_NAME: NAME /[$%&#!@]?/

"""

vbparser = Lark(grammar, parser='lalr')  # 'lalr' or 'earley'

def transpile_all(src_root):
    import os

    filepaths = []
    for dirpath, dirnames, filenames in os.walk(src_root):
        filepaths = filepaths + [ os.path.join(dirpath, f) for f in filenames if f.endswith(".bas") ]
    print("#files", len(filepaths), "\n")

    for i, f in enumerate(filepaths):
        print(f"{i+1}/{len(filepaths)}")
        try:
            transpile_one(f, verbose=False)
        except Exception as e:
            print(">!", str(e).split('\n')[0])
        print()

def transpile_one(src_path, verbose=False):
    import vba_py_generator as vba2py

    ast = parse_one(src_path, verbose)

    #src = vba2py.generate(ast)
    #with open(src_path+".py", "w") as f:
    #    f.write(src)

def parse_one(src_path, verbose=False):
    with open(src_path, "r") as f:
        lines = f.readlines()
        lines = lines[1:] # !!
        txt = "".join(lines)

    print("src_path", src_path, "\t\t", "#lines", len(lines))

    try:
        ast = vbparser.parse(txt)
        return ast
    except Exception as e:
        if verbose:
            traceback.print_exc()
            tmp = "".join([ f"{i+1}  {line}" for i, line in enumerate(lines) ])
            print(tmp)
        raise e

if __name__ == '__main__':

    src_root =r"..\..\xjs\vba"
    transpile_all(src_root)

    transpile_one(r"vba_demo_input.bas")
    #transpile_one(src_root + r"\modWalkDir.bas")
    #transpile_one(src_root + r"\modHttpRequest.bas")
    transpile_one(src_root + r"\modXtendedXl.bas")
