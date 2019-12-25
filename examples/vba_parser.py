"""

Dependencies:
pip install lark-parser

"""

import os, sys, traceback
from lark import Lark, Transformer, v_args

grammar = r"""
    ?start: statements

    !statements: ( ( ( statement | statements_inline ) smnt_new_line ) | comment_stmt | goto_anchor_stmt | empty_new_line )*
    !statements_inline: statement ( ":" statement )+

    smnt_new_line: comment_stmt | NEWLINE
    empty_new_line : NEWLINE
    comment_stmt : COMMENT_EOL

    ?statement: ( option_stmt
        | preprocessor_if_stmt
        | function_declare_stmt
        | function_definition_stmt
        | on_error_stmt
        | goto_stmt
        | exit_stmt
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

    !on_error_stmt: "On"i "Error"i ( "GoTo"i  (NAME|"0") | "Resume"i next )
    !goto_stmt: "GoTo"i  any_name
    !goto_anchor_stmt: any_name GOTO_ANCHOR_EOL

    !exit_stmt: "Exit"i  ( for | "Sub"i | "Function"i )

    !dim_stmt: ( "Static" | private | "ReDim"i | "Dim"i  ) dim_decl ( "," dim_decl ) *
    !dim_decl: TYPED_NAME [ "(" ranges ")" ] [ type_info ]
    ?ranges: range | ranges "," range
    !range: expr | ( expr to expr )

    function_declare_stmt: ("Public"i|"Private"i)? "Declare"i "PtrSafe"i? ( "Sub"i | "Function"i ) NAME "Lib"i VB_STRING "(" function_args ")"  [ as type_expr ]

    ?function_definition_stmt.2: ("Public"i|"Private"i)? function_or_sub NAME "(" function_args ")" type_info? smnt_new_line  statements end ( "Sub" | "Function" )
    !function_or_sub: "Sub"i | "Function"i
    function_args: [ function_arg ("," function_arg)* ]
    function_arg: [ "Optional" ] ["ByRef"|"ByVal"] NAME type_info? default_value?
    !type_info: as [ "New"i ]  type_expr
    !default_value: "=" expr

    // ("Public"i|"Private"i)?
    !assign_stmt: [ "Const"i | set | "Let"i ] lvalue "="  expr

    ?lvalue: with_access | any_name | member_access | func_call
    ?with_access: WITH_PREFIX any_name
    ?member_access: lvalue "." any_name
    // !! func_call is lvalue because we can have my2darray(1, 2) = "abc"

    // expr = value-expression / l-expression
    // value-expression = literal-expression / parenthesized-expression / typeof-is-expression /
    // new-expression / operator-expression
    // l-expression = simple-name-expression / instance-expression / member-access-expression /
    // index-expression / dictionary-access-expression / with-expression

    ?expr: lvalue | expr_literal | expr_parentheses | expr_new | expr_operator
    ?expr_literal: VB_STRING | NUMBER
    ?expr_operator: expr_binary | expr_unary
    !expr_binary: expr ( "^" | "=" | "<>" | ">" | "<" | ">=" | "<=" | "Is"i | "+" | "-" | "*" | "/" | "\\" | "&" | "." | "Mod"i | "And"i | "AndAlso"i | "Or"i | "OrElse"i | "Xor"i | "Like"i ) expr
    !expr_unary: ( "-" expr ) | ( "Not" expr )
    ?expr_parentheses: "(" expr_operator ")"
    ?expr_new: "New"i type_expr

    // expr_parentheses only apply on expr_operator to avoid consumung argument list

    // member-access-expression = l-expression NO-WS "." unrestricted-name
    // member-access-expression =/ l-expression line-continuation "." unrestricted-name

    // with-expression = with-member-access-expression / with-dictionary-access-expression
    // with-member-access-expression = "." unrestricted-name
    // with-dictionary-access-expression = "!" unrestricted-name
    // DICT ACCESS EXPRESION ??

    // Exponentiation ^
    // Unary negation -
    // Multiplicative *, /
    // Integer division \
    // Modulus Mod
    // Additive +, -
    // Concatenation &
    // Relational =, <>, <, >, <=, >=, Like, Is ??
    // Logical NOT Not
    // Logical AND And
    // Logical OR Or
    // Logical XOR Xor
    // Logical EQV Eqv ??
    // Logical IMP Imp ??


    !type_expr: NAME | ( NAME "." NAME )
    ?name: any_name  ( "." any_name )*
    any_name : NAME | TYPED_NAME | FOREIGN_NAME

    !sub_call_stmt: lvalue func_args
        | call lvalue

    !func_call: lvalue "(" func_args ")"
    !func_args: func_arg? ( "," func_arg? )*
    func_arg: [ TYPED_NAME ":=" ] expr

    ?if_stmt: if_stmt_inline | if_stmt_multilines
    if_stmt_inline : if expr then statement [ else statement ]
    if_stmt_multilines : if expr then smnt_new_line statements ( elseif expr then smnt_new_line statements )* [ else smnt_new_line statements ] end if

    !preprocessor_if_stmt: "#If" expr then smnt_new_line statements ( "#" elseif expr then smnt_new_line statements )* [ "#Else" smnt_new_line statements ] "#End If"

    ?for_stmt: for_each_stmt | for_in_stmt
    !for_each_stmt: for each TYPED_NAME in expr smnt_new_line statements next [TYPED_NAME]
    !for_in_stmt:     for TYPED_NAME "=" expr to expr  for_in_step? smnt_new_line statements next [TYPED_NAME]
    !for_in_step:      step expr

    ?loop_stmt: while_wend_stmt | while_do_loop_stmts | do_loop_while_stmts | until_do_loop_stmts | do_loop_until_stmts
    !while_wend_stmt : while_condition smnt_new_line statements "Wend"i
    !while_do_loop_stmts : "Do"i while_condition smnt_new_line statements "Loop"i
    !do_loop_while_stmts : "Do"i smnt_new_line statements "Loop"i while_condition
    !until_do_loop_stmts : "Do"i until_condition smnt_new_line statements "Loop"i
    !do_loop_until_stmts : "Do"i smnt_new_line statements "Loop"i until_condition
    !while_condition : "While"i expr
    !until_condition : "Until"i expr

    !with_stmt: with expr smnt_new_line statements end with
    !with_prefix: "."

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

    WITH_PREFIX: WS_INLINE "."

    %ignore WS_INLINE

    // $ String, % Integer, & Long, # Double, ! Single, @ Currency
    TYPED_NAME: NAME /[$%&#!@]?/

    FOREIGN_NAME : "[" NAME "]"

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

def transpile_one(src_path, verbose=False, generate=True):
    import vba_py_generator as vba2py

    ast = parse_one(src_path, verbose)
    if verbose:
        #print(ast.pretty())
        pass

    if generate:
        src = vba2py.generate(ast, verbose=verbose)
        with open(src_path+".py", "w") as f:
            f.write(src)
            print("Code written.")
            pass
        return src

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

    test_input = r"vba_demo_input.bas"
    test_input = src_root + r"\modXtendedXl.bas"
    transpile_one(test_input, verbose=False, generate=True)
