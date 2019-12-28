import traceback
import lark

TAB = "    "

TYPE_MAP = {
	"String": "str",
	"Boolean": "bool",
	"Byte": "int",
	"Integer": "int",
	"Long": "int",
	"LongLong": "int",
	"Single": "float",
	"Double": "float",
	# Currency, Date, Variant
	}

TYPE_SUFFIX_MAP = {
	"$": "str",
	"%": "int", 	# Integer
	"&": "int", 	# Long
	"#": "float", 	# Double
	"!": "float", 	# Single
	# "@":  "float", # Currency
	}

OP_MAP = {
	"=": "==",
	"<>": "!=",
	"And": "and", 	# approximate equivalence. Python does lazy evaluation
	"AndAlso": "and",
	"Or": "or", 	# approximate equivalence. Python does lazy evaluation
	"OrElse": "or",
	"&": "+",
	"Mod": "%",
	"Is": "is",
	}

class GenerateInfos:
	def __init__(self):
		self.indent = ""
		self.option_explicit = False
		self.option_base = 0
		self.def_func_name = None # to handle return value
		self.on_error_resume_next = False
		self.on_error_goto = None
		self.with_stack = []
		self.global_vars = []
		self.vars = self.global_vars

class CustomException(Exception):
    pass

class CustomRethrowException(Exception):
    def __init__(self, message, inner_error):
        super().__init__(message)
        #self.inner_error = inner_error

def generate(ast, verbose=False):
	infos = GenerateInfos()

	src = []
	src.append("import os, sys, numpy as np")
	src.append("")
	src.append("""#sys.path.append(r"..\..\lark\examples")""")
	src.append("from vba_py_app import *")
	src.append("")

	if verbose:
		print(ast.pretty())

	generate_stmts(ast, src, infos)

	print(f"Code generated. \t #lines{len(src)}")
	return '\n'.join(src)

has_unsupported_feature = False
def unsupported_feature(msg, src, infos):
	global has_unsupported_feature
	has_unsupported_feature = True

	print(">!", msg)
	msg = f"{infos.indent}raise Exception(\"{msg}\")";
	if src is not None:
		src.append(msg)
	else:
		return src
	#assert False

def generate_stmts(ast, src, infos):
	assert get_type(ast) == "statements"

	stmts = ast.children
	if len(stmts)==0:
		src.append(f"{infos.indent}pass")
	for s in stmts:
		generate_stmt(s, src, infos)

def generate_stmt(ast, src, infos):
	try:
		generate_stmt_throws(ast, src, infos)
	except Exception as e:
		traceback.print_exc()
		if is_node(ast):
			print("\nAST:\n", ast.pretty(), "\n")
		else:
			print("\nAST is TOKEN:\n", ast, "\n")

		src.append("")
		src.append(f"{infos.indent}#!! Generation error here")
		src.append(f"{infos.indent}raise Exception('Generation error here')")
		src.append("")
		#raise

def generate_stmt_throws(ast, src, infos):
	s = ast
	t = get_type(ast)
	n = len(ast.children) if is_node(ast) else None

	# print("generate_stmt", t )
	if t == "statements":
		generate_stmts(ast, src, infos)
	elif t == "statements_inline":
		for i in range(0, n, 2):
			generate_stmt(ast.children[i], src, infos)
	elif t == "option_stmt":
		generate_option_stmt(ast, src, infos)
	elif t == "dim_stmt":
		generate_dim_stmt(ast, src, infos)
	elif t == "comment_stmt":
		txt = generate_fragment(ast, infos)
		src.append(f"{infos.indent}{txt}")
	elif t == "NEWLINE":
		src.append("")
	elif t == "function_definition_stmt":
		generate_function_definition_stmt(ast, src, infos)
	elif t == "exit_stmt":
		generate_exit_stmt(ast, src, infos)
	elif t == "sub_call_stmt":
		generate_sub_call_stmt(ast, src, infos)
	elif t == "assign_stmt":
		generate_assign_stmt(ast, src, infos)
	elif t == "if_stmt_multilines":
		generate_if_stmt(ast, src, infos)
	elif t == "if_stmt_inline":
		generate_if_stmt(ast, src, infos, inline = True)
	elif t == "for_in_stmt":
		generate_for_in_stmt(ast, src, infos)
	elif t == "for_each_stmt":
		generate_for_each_stmt(ast, src, infos)
	elif t == "while_wend_stmt":
		generate_loop_stmt(ast, src, infos, pre_condition=True, while_condition=True)
	elif t == "while_do_loop_stmts":
		generate_loop_stmt(ast, src, infos, pre_condition=True, while_condition=True)
	elif t == "do_loop_while_stmts":
		generate_loop_stmt(ast, src, infos, pre_condition=False, while_condition=True)
	elif t == "until_do_loop_stmts":
		generate_loop_stmt(ast, src, infos, pre_condition=True, while_condition=False)
	elif t == "do_loop_until_stmts":
		generate_loop_stmt(ast, src, infos, pre_condition=False, while_condition=False)
	elif t == "with_stmt":
		generate_with_stmt(ast, src, infos)
	elif t == "on_error_stmt":
		generate_on_error_stmt(ast, src, infos)
	elif t == "goto_anchor_stmt":
		generate_goto_anchor_stmt(ast, src, infos)
	elif t == "smnt_new_line":
		if n==1:
			# only new line
			pass
		else:
			# comment and newline
			txt = generate_fragment(ast.children[0], infos)
			src.append(f"{infos.indent}{txt}")
	elif t == "empty_new_line":
		src.append("")
	elif t in [ "preprocessor_if_stmt", "goto_stmt", "exit_stmt", "if_stmt_inline", "for_each_stmt" ]:
		# "with_stmt",
		unsupported_feature(f"Unsupported statement type {t}.", src, infos)
	elif t in [ "open_stmt", "print_stmt", "close_stmt" ]:
		unsupported_feature(f"Unsupported file operations (open, print, close).", src, infos)
	else:
		raise Exception(f"Unknown statement type {t}")

def generate_option_stmt(ast, src, infos):
	opt = get_text(ast.children[1])
	if opt == "Explicit":
		infos.option_explicit = True
	elif opt == "Base":
		base = get_int(ast.children[2])
		assert base==0 or base==1 	# as per VBA spec
		infos.option_base = base
	else:
		unsupported_feature(f"Unsupported Option {opt}.", src, infos)

def generate_dim_stmt(ast, src, infos):
	decls = find_nodes_by_type(ast, "dim_decl")
	tmp = []
	for decl in decls:
		var = decl.children[0]
		typ = None
		has_new = False
		range_indices = None

		# implicit type
		if get_type(var) == "TYPED_NAME" and var.value[-1] in TYPE_SUFFIX_MAP:
			typ = TYPE_SUFFIX_MAP[var.value[-1]]
			var = generate_fragment(var, infos)
		# explicit type
		xtyp = find_node_by_type(decl, "type_info")
		if xtyp:
			vb_type = generate_fragment(xtyp.children[-1], infos)
			has_new = find_node_by_type(xtyp, "NEW") is not None
			if vb_type in TYPE_MAP:
				typ = TYPE_MAP[vb_type]
			if has_new:
				typ = vb_type 	# best we can do !!

		xrng = find_node_by_type(decl, "ranges")
		xrngs = find_nodes_by_type(decl, "range")
		if xrng is not None or len(xrngs)>0:
			range_indices = []
		for rng in xrngs:
			if len(rng.children) == 1:
				range_indices.append([ infos.option_base, generate_ith_fragment(rng, 0, infos) ])
			elif len(rng.children) == 3:
				range_indices.append([ generate_ith_fragment(rng, 0, infos), generate_ith_fragment(rng, 2, infos) ])
			else:
				raise False

		if range_indices is not None:
			shape = ", ".join([ f"{rng[1]}-{rng[0]}+1" for rng in range_indices])
			offset = ", ".join([ f"{rng[0]}" for rng in range_indices])
			if typ == "str" or typ == "int" or typ == "float":
				tmp.append(f"{var} = np.ndarray(shape=({shape},), offset_isnt_what_we_need=({offset},), dtype={typ})")
			else:
				tmp.append(f"{var} = np.ndarray(shape=({shape},), offset_isnt_what_we_need=({offset},))")
		elif typ == "ArrayList":
				tmp.append(f"{var} = []")
		else:
			if typ == "str":
				tmp.append(f"{var} = \"\"")
			elif typ == "int":
				tmp.append(f"{var} = 0")
			elif typ == "float":
				tmp.append(f"{var} = 0.0")
			elif has_new:
				tmp.append(f"{var} = {typ}()")
			elif typ is None:
				tmp.append(f"{var} = None")

		infos.vars.append(var)

	src.append(f"{infos.indent}{'; '.join(tmp)}")

def generate_function_definition_stmt(ast, src, infos):
	assert get_type(ast.children[0]) == "function_or_sub"
	assert get_type(ast.children[2]) == "function_args"

	type_anotation = ""
	vb_type = generate_ith_fragment(find_node_by_type(ast, "type_info"), 1, infos)
	if vb_type in TYPE_MAP:
		type_anotation = " -> "+TYPE_MAP[vb_type]
	if get_text(ast.children[0]) == "Sub":
		type_anotation = " -> None "

	name = ast.children[1].value
	args_csv = generate_fragment(ast.children[2], infos)
	stmts = find_node_by_type(ast, "statements")

	src.append(f"def {name}({args_csv}){type_anotation}:" )
	try:
		infos.def_func_name = name
		infos.indent +=TAB
		infos.vars = [] + infos.global_vars 	# clone global scope as local one
		generate_stmts(stmts, src, infos)

		if get_text(ast.children[0]) == "Function":
			src.append(f"{infos.indent}return {infos.def_func_name}")
	finally:
		infos.def_func_name = None
		infos.indent = infos.indent[:-len(TAB)]
		infos.vars = infos.global_vars 	# restore global scope

	src.append(f"") # !!

def generate_exit_stmt(ast, src, infos):
	assert get_type(ast.children[0]) == "EXIT"
	t = get_type(ast.children[1])

	if t == "FUNCTION":
		src.append(f"{infos.indent}return {infos.def_func_name}")
	elif t == "SUB":
		src.append(f"{infos.indent}return")
	elif t == "for":
		# check which for loopwe are exiting from !!
		src.append(f"{infos.indent}break")
	else:
		unsupported_feature("Unsupported Exit {t}", src, infos)

def generate_sub_call_stmt(ast, src, infos):
	if get_type(ast.children[0]) == "call":

		if get_type(ast.children[1]) != "func_call":

			# no arguments, no parentheses, just "Call fct"
			name = generate_fragment(ast.children[1], infos)
			src.append(f"{infos.indent}{name}()")

		else:
			x = generate_func_call(ast.children[1], infos)
			src.append(f"{infos.indent}{x}");

	else:

		assert get_type(ast.children[0]) in [ "any_name", "member_access", "func_call", "with_access" ]
		#assert get_type(ast.children[1]) == "func_args_sub"

		name = generate_fragment(ast.children[0], infos)
		args_csv = generate_func_args(ast.children[1], infos)

		src.append(f"{infos.indent}{name}({args_csv})")

def generate_func_call(ast, infos):
	n = len(ast.children)

	assert get_type(ast.children[0]) in [ "any_name", "member_access", "func_call", "with_access" ]
	assert get_type(ast.children[1]) == "LPAR"
	assert get_type(ast.children[2]) == "func_args"
	assert get_type(ast.children[3]) == "RPAR"

	name = generate_fragment(ast.children[0], infos)
	array_indices =  name in infos.vars or is_builtin_array(name)	# functions can't be assigned to variables, so that must be an array
	args_csv = generate_func_args(ast.children[2], infos, array_indices=array_indices)

	if array_indices:
		return f"{name}{args_csv}"

	return f"{name}({args_csv})"

def is_builtin_array(lhs):
	if lhs.endswith(")") or lhs.endswith("]"):
		return True
	return False

def generate_func_args(ast, infos, array_indices = False):
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
			args.append(generate_fragment(arg, infos))
			prev_is_sep = False
	if prev_is_sep:
		# ending with comma
		args.append("None")

	if array_indices:
		return "".join([ f"[{a}]" for a in args ])
	return ", ".join(args)

def generate_assign_stmt(ast, src, infos):
	if get_type(ast.children[1]) == "EQUAL":
		k=0
	elif get_type(ast.children[2]) == "EQUAL":
		k=1
	else:
		assert False

	var = generate_fragment(ast.children[k+0], infos)
	val = generate_fragment(ast.children[k+2], infos)

	if var == infos.def_func_name:
		# !! if we assign to the name of the function then we can't have recursive code
		# !! if we rename, we have to rename every assignment and evaluation
		pass
	else:
		infos.vars.append(var)

	src.append(f"{infos.indent}{var} = {val}")

def generate_if_stmt(ast, src, infos, inline=False):
	n = len(ast.children)

	for k in range(0, n, 4 if inline else 5):
		t = get_type(ast.children[k+0])
		if t == "if" or t == "elseif":
			expr = generate_fragment(ast.children[k+1], infos)
			stmts = ast.children[k+(3 if inline else 4)]

			ie = "if" if k==0 else "elif"
			src.append(f"{infos.indent}{ie} {expr}:")
			try:
				infos.indent +=TAB
				generate_stmt(stmts, src, infos)
			finally:
				infos.indent = infos.indent[:-len(TAB)]
		elif t == "else" :
			stmts = ast.children[k+(1 if inline else 2)]

			src.append(f"{infos.indent}else:")
			try:
				infos.indent +=TAB
				generate_stmt(stmts, src, infos)
			finally:
				infos.indent = infos.indent[:-len(TAB)]
		elif t == "end":
			pass
		else:
			assert False

def generate_for_in_stmt(ast, src, infos):
	n = len(ast.children)

	assert get_type(ast.children[0]) == "for"
	assert get_type(ast.children[2]) == "EQUAL"
	assert get_type(ast.children[4]) == "to"

	i  = generate_fragment(ast.children[1], infos)
	i0 = generate_fragment(ast.children[3], infos)
	i1 = generate_fragment(ast.children[5], infos)
	di = generate_ith_fragment(find_node_by_type(ast, "for_in_step"), 1, infos)
	stmts = find_node_by_type(ast, "statements")

	if di is None :
		src.append(f"{infos.indent}for {i} in range({i0}, {i1}):")
	else:
		src.append(f"{infos.indent}for {i} in range({i0}, {i1}, {di}):")
	try:
		infos.indent +=TAB
		generate_stmts(stmts, src, infos)
	finally:
		infos.indent = infos.indent[:-len(TAB)]

def generate_for_each_stmt(ast, src, infos):
	n = len(ast.children)

	assert get_type(ast.children[0]) == "for"
	assert get_type(ast.children[1]) == "each"
	assert get_type(ast.children[3]) == "in"

	k  = generate_fragment(ast.children[2], infos)
	v = generate_fragment(ast.children[4], infos)
	stmts = find_node_by_type(ast, "statements")

	src.append(f"{infos.indent}for {k} in {v}:")
	try:
		infos.indent +=TAB
		generate_stmts(stmts, src, infos)
	finally:
		infos.indent = infos.indent[:-len(TAB)]

def generate_loop_stmt(ast, src, infos, pre_condition, while_condition):
	cond = find_node_by_type(ast, "while_condition" if while_condition else "until_condition")
	stmts = find_node_by_type(ast, "statements")

	cond = generate_fragment(cond.children[1], infos)

	if pre_condition:
		if while_condition:
			src.append(f"{infos.indent}while {cond}:")
		else:
			src.append(f"{infos.indent}while not ({cond}):")
	else:
		src.append(f"{infos.indent}while True:")

	try:
		infos.indent +=TAB
		generate_stmts(stmts, src, infos)
		if not pre_condition:
			if while_condition:
				src.append(f"{infos.indent}if not({cond}):")
				src.append(f"{infos.indent}{TAB}break")
			else:
				src.append(f"{infos.indent}if {cond}:")
				src.append(f"{infos.indent}{TAB}break")
	finally:
		infos.indent = infos.indent[:-len(TAB)]

def generate_with_stmt(ast, src, infos):
	n = len(ast.children)

	assert get_type(ast.children[0]) == "with"

	var = f"__x{len(src)+1}"
	val  = generate_fragment(ast.children[1], infos)
	stmts = find_node_by_type(ast, "statements")

	src.append(f"{infos.indent}with {val} as {var}:")
	try:
		infos.with_stack.append(var)
		infos.indent +=TAB
		generate_stmts(stmts, src, infos)
	finally:
		infos.indent = infos.indent[:-len(TAB)]

def generate_with_access(ast, infos):
	n = len(ast.children)

	assert n == 2
	assert get_type(ast.children[0]) == "WITH_PREFIX"

	object = infos.with_stack[-1]
	member = generate_fragment(ast.children[1], infos)

	return f"{object}.{member}"

def generate_on_error_stmt(ast, src, infos):
	action = get_text(ast.children[2])
	tag = get_text(ast.children[3])
	if action == "Resume":
		assert tag == "next"
		infos.on_error_resume_next = True
		infos.on_error_goto = None
	elif action == "GoTo":
		infos.on_error_resume_next = False
		if tag == "0" or tag == 0:
			infos.on_error_goto = None
		else:
			src.append(f"{infos.indent}try:")
			infos.indent +=TAB
			infos.on_error_goto = ast.children[3].value
	else:
		assert False

def generate_goto_anchor_stmt(ast, src, infos):
	tag = generate_fragment(ast.children[0], infos)
	if infos.on_error_goto is not None:
		if infos.on_error_goto == tag:
			# matching On Error Goto Xyz
			infos.indent = infos.indent[:-len(TAB)]
			infos.on_error_goto = None
			src.append(f"{infos.indent}finally:")
			src.append(f"{infos.indent}{TAB}pass")
		else:
			print(infos.on_error_goto, "<>", tag)
			assert False
	else:
		src.append(f"{infos.indent}# Goto tag :{tag}")

def generate_ith_fragment(ast, ith, infos):
	if ast is None: return None
	return generate_fragment(ast.children[ith], infos)

# generate expression fragments
def generate_fragment(ast, infos):
	if ast is None: return None

	t = get_type(ast)
	n = len(ast.children) if is_node(ast) else None
	if t == "NUMBER":
		return ast.value
	elif t == "VB_STRING":
		txt = ast.value[1:-1]
		txt= txt.replace('""', '"')
		txt= txt.replace('\\', '\\\\').replace('"', '\\"')
		return f'"{txt}"'
	elif t == "NAME":
		return ast.value
	elif t == "TYPED_NAME":
		var = ast.value
		if var[-1] in TYPE_SUFFIX_MAP:
			var = var[:-1] 	# ignoring the type suffix
		return var
	elif t == "any_name":
		assert n==1
		return generate_fragment(ast.children[0], infos)
	elif t == "with_access":
		return generate_with_access(ast, infos)
	elif t == "comment_stmt":
		assert len(ast.children)==1
		txt = get_text(ast)
		txt = txt[1:].strip()
		return f"# {txt}"
	elif t == "function_args":
		return ", ".join([ generate_fragment(arg, infos) for arg in ast.children ])
	elif t == "function_arg":
		name = generate_fragment(ast.children[0], infos)

		type_anotation = ""
		vb_type = generate_ith_fragment(find_node_by_type(ast, "type_info"), 1, infos)
		if vb_type in TYPE_MAP:
			type_anotation = " : "+TYPE_MAP[vb_type]

		default_value = generate_ith_fragment(find_node_by_type(ast, "default_value"), 1, infos)
		if default_value is None:
			default_value = ""
		else:
			default_value = f" = {default_value}"

		return f"{name}{type_anotation}{default_value}"
	elif t == "func_call":
		return generate_func_call(ast, infos)
	elif t == "func_arg":
		if n==1:
			return generate_fragment(ast.children[0], infos)
		elif n==2:
			return generate_fragment(ast.children[0], infos) + " = " + generate_fragment(ast.children[1], infos)
		else:
			assert False
	elif t == "member_access":
		assert len(ast.children)==2
		return generate_fragment(ast.children[0], infos) + "." + generate_fragment(ast.children[1], infos)
	elif t == "expr_binary":
		assert n==3
		op = ast.children[1].value
		lhs = generate_fragment(ast.children[0], infos)
		rhs = generate_fragment(ast.children[2], infos)

		if op in OP_MAP:
			op = OP_MAP[op]
		if op in [ "Like", "Eqv", "Imp" ]:
			unsupported_feature(f"Unsupported {op} operator.", None, infos)
		else:
			return  f"{lhs} {op} {rhs}"
	elif t == "expr_unary":
		assert n==2
		op = get_text(ast.children[0])
		arg = generate_fragment(ast.children[1], infos)

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

def get_text(ast):
	if is_node(ast):
		n = len(ast.children)
		if n == 0: return ast.data
		if n == 1: return ast.children[0].value
		assert False
	else:
		return ast.value

def get_int(ast):
	return int(get_text(ast))

def is_node(ast):
	return isinstance(ast, lark.tree.Tree)

def is_token(ast):
	return isinstance(ast, lark.lexer.Token)

def find_node_by_type(ast, data_type):
	found = None
	for child in ast.children:
		if get_type(child)== data_type:
			assert found is None
			found = child
	return found

def find_nodes_by_type(ast, data_type):
	# recursive !!
	tmp = ast.find_data(data_type)
	tmp = list(tmp)
	return tmp

if __name__ == '__main__':
	import vba_parser as p
	test_file = r"vba_demo_input.bas"
	#test_file = r"..\..\xjs\vba\modJsonStringify.bas"
	src = p.transpile_one(test_file, verbose=False, generate=True)
	print()
	print(src)