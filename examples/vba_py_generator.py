import traceback
import lark

TAB = "    "

TYPE_MAP = {
	"String": "str",
	"Long": "int",
	"Float": "float",
	"Double": "float",
	}
OP_MAP = {
	"=": "==",
	"<>": "!=",
	"Mod": "%" }

class GenerateInfo:
	def __init__(self):
		self.def_func_name = None # to handle return value
		self.with_stack = []

class CustomException(Exception):
    pass

class CustomRethrowException(Exception):
    def __init__(self, message, inner_error):
        super().__init__(message)
        #self.inner_error = inner_error

def generate(ast, verbose=False):
	info = GenerateInfo()

	src = []
	src.append("import os, sys\n")

	if verbose:
		print(ast.pretty())

	generate_stmts(ast, src, indent="")

	print("Code generated.")
	return '\n'.join(src)

def generate_stmts(ast, src, indent):
	assert get_type(ast) == "statements"

	stmts = ast.children
	if len(stmts)==0:
		src.append(f"{indent}pass")
	for s in stmts:
		generate_stmt(s, src, indent)

def generate_stmt(ast, src, indent):
	try:
		generate_stmt_throws(ast, src, indent)
	except Exception as e:
		traceback.print_exc()
		if is_node(ast):
			print("\nAST:\n", ast.pretty(), "\n")
		else:
			print("\nAST is TOKEN:\n", ast, "\n")

		src.append("")
		src.append(f"{indent}#!! Generation error here")
		src.append(f"{indent}raise Exception('Generation error here')")
		src.append("")
		#raise

def generate_stmt_throws(ast, src, indent):
	s = ast
	t = get_type(ast)
	n = len(ast.children) if is_node(ast) else None

	# print("generate_stmt", t )
	if t == "statements":
		generate_stmts(ast, src, indent)
	elif t == "statements_inline":
		for i in range(0, n, 2):
			generate_stmt(ast.children[i], src, indent)
	elif t == "option_stmt":
		pass
	elif t == "dim_stmt":
		pass
	elif t == "comment_stmt":
		txt = generate_fragment(ast)
		src.append(f"{indent}{txt}")
	elif t == "NEWLINE":
		src.append("")
	elif t == "function_definition_stmt":
		generate_function_definition_stmt(ast, src, indent)
	elif t == "sub_call_stmt":
		generate_sub_call_stmt(ast, src, indent)
	elif t == "assign_stmt":
		generate_assign_stmt(ast, src, indent)
	elif t == "if_stmt_multilines":
		generate_if_stmt(ast, src, indent)
	elif t == "if_stmt_inline":
		generate_if_stmt(ast, src, indent, inline = True)
	elif t == "for_in_stmt":
		generate_for_in_stmt(ast, src, indent)
	elif t == "smnt_new_line":
		if n==1:
			# only new line
			pass
		else:
			# comment and newline
			txt = generate_fragment(ast.children[0])
			src.append(f"{indent}{txt}")
	elif t == "empty_new_line":
		src.append("")
	elif t in [ "preprocessor_if_stmt", "on_error_stmt", "goto_stmt", "goto_anchor_stmt", "with_stmt", "loop_stmt", "exit_stmt", "if_stmt_inline", "for_each_stmt" ]:
		msg = f"{indent}raise Exception('Statement type '{t}' not supported !!')"
		print(msg)
		src.append(msg)
	elif t in [ "open_stmt", "print_stmt", "close_stmt" ]:
		msg = f"{indent}raise Exception('File operations (open, print, close) not supported !!')"
		print(msg)
		src.append(msg)
	else:
		raise Exception(f"Unknown statement type {t}")

def generate_function_definition_stmt(ast, src, indent):
	assert get_type(ast.children[0]) == "function_or_sub"
	assert get_type(ast.children[2]) == "function_args"

	type_anotation = ""
	vb_type = generate_ith_fragment(find_node_by_type(ast, "type_info"), 1)
	if vb_type in TYPE_MAP:
		type_anotation = " -> "+TYPE_MAP[vb_type]
	if ast.children[0].children[0].value == "Sub":
		type_anotation = " -> None "

	name = ast.children[1].value
	args_csv = generate_fragment(ast.children[2])
	stmts = find_node_by_type(ast, "statements")

	src.append(f"def {name}({args_csv}){type_anotation}:" )
	generate_stmts(stmts, src, indent+TAB)
	src.append(f"") # !!

def generate_sub_call_stmt(ast, src, indent):
	if get_type(ast.children[0]) == "call":

		if get_type(ast.children[1]) != "func_call":

			# no arguments, no parentheses, just "Call fct"
			name = generate_fragment(ast.children[1])
			src.append(f"{indent}{name}()")

		else:
			x = generate_func_call(ast.children[1])
			src.append(f"{indent}{x}");

	else:

		assert get_type(ast.children[0]) in [ "any_name", "member_access", "func_call" ]
		#assert get_type(ast.children[1]) == "func_args_sub"

		name = generate_fragment(ast.children[0])
		args_csv = generate_func_args(ast.children[1])

		src.append(f"{indent}{name}({args_csv})")

def generate_func_call(ast):
	n = len(ast.children)

	assert get_type(ast.children[0]) in [ "any_name", "member_access", "func_call" ]
	assert get_type(ast.children[1]) == "LPAR"
	assert get_type(ast.children[2]) == "func_args"
	assert get_type(ast.children[3]) == "RPAR"

	name = generate_fragment(ast.children[0])
	args_csv = generate_func_args(ast.children[2])

	return f"{name}({args_csv})"

def generate_func_args(ast):
	assert get_type(ast) == "func_args" or get_type(ast) == "func_args_sub"

	if len(ast.children)==0:
		return ""

	args = []
	prev_is_sep = True
	for arg in ast.children:
		if get_type(arg) == "COMMA":
			if prev_is_sep:
				args.append("None")
			else:
				prev_is_sep = True
		else:
			args.append(generate_fragment(arg))
			prev_is_sep = False
	if prev_is_sep:
		# ending with comma
		args.append("None")

	return ", ".join(args)

def generate_assign_stmt(ast, src, indent):
	if get_type(ast.children[1]) == "EQUAL":
		k=0
	elif get_type(ast.children[2]) == "EQUAL":
		k=1
	else:
		assert False

	var = generate_fragment(ast.children[k+0])
	val = generate_fragment(ast.children[k+2])

	src.append(f"{indent}{var} = {val}")

def generate_if_stmt(ast, src, indent, inline=False):
	n = len(ast.children)

	for k in range(0, n, 4 if inline else 5):
		t = get_type(ast.children[k+0])
		if t == "if" or t == "elseif":
			expr = generate_fragment(ast.children[k+1])
			stmts = ast.children[k+(3 if inline else 4)]

			ie = "if" if k==0 else "elif"
			src.append(f"{indent}{ie} {expr}:")
			generate_stmt(stmts, src, indent+TAB)
		elif t == "else" :
			stmts = ast.children[k+(1 if inline else 2)]

			src.append(f"{indent}else:")
			generate_stmt(stmts, src, indent+TAB)
		elif t == "end":
			pass
		else:
			assert False

def generate_for_in_stmt(ast, src, indent):
	n = len(ast.children)

	assert get_type(ast.children[0]) == "for"
	assert get_type(ast.children[2]) == "EQUAL"
	assert get_type(ast.children[4]) == "to"

	i  = generate_fragment(ast.children[1])
	i0 = generate_fragment(ast.children[3])
	i1 = generate_fragment(ast.children[5])
	di = generate_ith_fragment(find_node_by_type(ast, "for_in_step"), 1)
	stmts = find_node_by_type(ast, "statements")

	if di is None :
		src.append(f"{indent}for {i} in range({i0}, {i1}):")
	else:
		src.append(f"{indent}for {i} in range({i0}, {i1}, {di}):")
	generate_stmts(stmts, src, indent+TAB)

def generate_ith_fragment(ast, ith):
	if ast is None: return None
	return generate_fragment(ast.children[ith])

# generate expression fragments
def generate_fragment(ast):
	if ast is None: return None

	t = get_type(ast)
	n = len(ast.children) if is_node(ast) else None
	if t == "NUMBER":
		return ast.value
	elif t == "VB_STRING":
		return ast.value
	elif t == "NAME":
		return ast.value
	elif t == "TYPED_NAME":
		return ast.value
	elif t == "any_name":
		assert n==1
		return str(ast.children[0])
	elif t == "comment_stmt":
		assert len(ast.children)==1
		txt = ast.children[0].value
		txt = txt[1:].strip()
		return f"# {txt}"
	elif t == "function_args":
		return ", ".join([ generate_fragment(arg) for arg in ast.children ])
	elif t == "function_arg":
		name = generate_fragment(ast.children[0])

		type_anotation = ""
		vb_type = generate_ith_fragment(find_node_by_type(ast, "type_info"), 1)
		if vb_type in TYPE_MAP:
			type_anotation = ": "+TYPE_MAP[vb_type]

		default_value = generate_ith_fragment(find_node_by_type(ast, "default_value"), 1)
		if default_value is None:
			default_value = ""
		else:
			default_value = f" = {default_value}"

		return f"{name}{default_value}{type_anotation}"
	elif t == "func_call":
		return generate_func_call(ast)
	elif t == "func_arg":
		if n==1:
			return generate_fragment(ast.children[0])
		elif n==2:
			return generate_fragment(ast.children[0]) + " = " + generate_fragment(ast.children[1])
		else:
			assert False
	elif t == "member_access":
		assert len(ast.children)==2
		return generate_fragment(ast.children[0]) + "." + generate_fragment(ast.children[1])
	elif t == "expr_binary":
		assert n==3
		op = ast.children[1].value
		lhs = generate_fragment(ast.children[0])
		rhs = generate_fragment(ast.children[2])

		if op in OP_MAP:
			op = OP_MAP[op]
		return  f"{lhs} {op} {rhs}"
	elif t == "expr_unary":
		assert n==2
		op = ast.children[0].value
		arg = generate_fragment(ast.children[1])

		if op == "Not":
			op = "not"
		return  f"{op} {arg}"
	elif t == "type_expr":
		if n==1:
			return ast.children[0].value
		elif n==3:
			return ast.children[0].value + "." + ast.children[2].value
		else:
			assert False
	else:
		raise Exception(f"Unknown expression type {t}")

def get_type(ast):
	if is_node(ast):
		return ast.data
	else:
		return ast.type

def is_node(ast):
	return isinstance(ast, lark.tree.Tree)

def is_token(ast):
	return isinstance(ast, lark.lexer.Token)

def find_node_by_type(ast, data_type):
	for child in ast.children:
		if get_type(child)== data_type:
			return child

	# tmp = ast.find_data(data_type)    # recursive !
	# tmp = list(tmp)
	# if len(tmp)!=1:
	# 	print("#find_node_by_type", len(tmp))
	# return tmp[0]

if __name__ == '__main__':
	import vba_parser as p
	test_file = r"vba_demo_input.bas"
	#test_file = r"..\..\xjs\vba\modJsonStringify.bas"
	src = p.transpile_one(test_file, verbose=False, generate=True)
	print(src)