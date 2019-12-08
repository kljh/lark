import lark
TAB = "    "

def generate(ast):
	src = []
	src.append("import os, sys\n")

	#print(ast.pretty())
	for s in ast.children:
		generate_stmt(s, src, indent="")

		#src.append(str(s))

	return '\n'.join(src)

def generate_stmt(ast, src, indent):
	s = ast.children[0]  	# concrete statement
	t = s.data  	# statement type
	if t == "option_stmt":
		pass
	elif t == "dim_stmt":
		pass
	elif t == "function_definition_stmt":
		generate_function_definition_stmt(s, src, indent)
	elif t == "sub_call_stmt":
		generate_sub_call_stmt(s, src, indent)
	elif t == "assign_stmt":
		generate_assign_stmt(s, src, indent)
	else:
		raise Exception(f"Unknown statement type {t}")

def generate_function_definition_stmt(ast, src, indent):
	#print(ast.pretty())

	name = ast.children[0].value
	args = ast.children[1]
	assert isinstance(ast.children[2], lark.lexer.Token) and  ast.children[2].type == "NEWLINE"
	stmts = ast.children[3:]  # children 2 is new line

	#print("stmts", len(stmts), stmts)

	src.append(f"def {name}():" )
	if len(stmts)==0:
		src.append(f"{indent+TAB}pass")
	for s in stmts:
		#print(s)
		generate_stmt(s, src, indent+TAB)

def generate_sub_call_stmt(ast, src, indent):
	print(ast.pretty())

	name = ast.children[0].value
	args_csv = "..."

	src.append(f"{indent}{name}({args_csv})" )

def generate_assign_stmt(ast, src, indent):
	#print(ast.pretty())
	var = ast.children[0].value
	val = ast.children[2].value

	src.append(f"{indent}{var} = {val}")

if __name__ == '__main__':
    import vba_to_script as p
    p.transpile_one(r"test.bas")