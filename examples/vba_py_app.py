import re
import unittest

from collections import namedtuple
from typing import List

# Application object

# Method 1: pywin32
import win32com.client
CreateObject = win32com.client.Dispatch

# Method 2
#import comtypes.client
#from comtypes.gen.Excel import xlRangeValueDefault
#CreateObject = comtypes.client.CreateObject

Application = CreateObject("Excel.Application")
#Application.Visible = True

# Debug object

def DebugPrint(*args):
	print(*args)

def DebugAssert(b):
	assert b

DebugObject = namedtuple('DebugObject', [ 'Print', 'Assert' ])
Debug = DebugObject(Print = DebugPrint, Assert = DebugAssert)

# Standard Public functions

Len = len
Asc = ord
Chr = chr

# 6.1.1.5 VbCompareMethod

vbBinaryCompare = 0
vbTextCompare = 1

# 6.1.2.2 Constants Module

vbBack = Chr(8)
vbCr = Chr(13)
vbCrLf = Chr(13) + Chr(10)
vbFormFeed = Chr(12)
vbLf = Chr(10)
vbNewLine = "\n"
vbNullChar = Chr(0)
vbTab = Chr(9)
vbVerticalTab = Chr(11)
# vbNullString # null string pointer
vbObjectError = -2147221504

# Standard Public functions (continued)

def Array(*args):
	return [ *args ]

def Join(arr, sep : str):
	return sep.join(arr)

def Mid(s : str, p : int, n : int):
	return s[p-1: p-1+n]

def Left(s : str, n : int):
	return s[:n]

def Right(s : str, n : int):
	return s[len(s)-n:]

def Replace(Expression : str, Find : str, Replace : str, Start : int = 1, Count : int = -1, Compare = vbBinaryCompare) -> str :
	# !! replace only one occurence
	return Expression.replace(Find, Replace)

def InStrRev(StringCheck: str, StringMatch: str, Start : int = -1, Compare : int = vbBinaryCompare):
	res = StringCheck.rfind(StringMatch)+1
	return res

class TestPublicFunctions(unittest.TestCase):
	def test_substring(self):
		self.assertEqual(Mid("abcdef", 2, 3), "bcd")
		self.assertEqual(Left("abcdef", 2), "ab")
		self.assertEqual(Right("abcdef", 2), "ef")
	def test_find(self):
		#self.assertEqual(InStr("abcdefbc", "bc"), 2)
		#self.assertEqual(InStr("abcdefbc", "ba"), 0)
		#self.assertEqual(InStr(1, "abcdefbc", "bc"), 2)
		#self.assertEqual(InStr(2, "abcdefbc", "bc"), 2)
		#self.assertEqual(InStr(3, "abcdefbc", "bc"), 7)
		self.assertEqual(InStrRev("abcdefbc", "bc"), 7)
		self.assertEqual(InStrRev("abcdefbc", "ba"), 0)

def Dir(path):
	import glob
	if path is None:
		assert False # re-entrant version not implemented
	res = glob.glob(path)
	if len(res)==0:
		return ""
	else:
		return res[0]

# ADO object

ADODBObject = namedtuple('ADODBObject', [ 'Connection', 'Recordset' ])
ADODB = ADODBObject(
	Connection = lambda: CreateObject("ADODB.Connection"),
	Recordset = lambda: CreateObject("ADODB.Recordset"))

adUseNone = 1
adUseServer = 2
adUseClient = 3

# ADOX object

ADOXObject = namedtuple('ADOXObject', [ 'Catalog' ])
ADOX = ADOXObject(
	Catalog = lambda: CreateObject("ADOX.Catalog"))


if __name__ == '__main__':
	Debug.Print("test Debug.Print", 1, True, 2.3)
	Debug.Assert(True)
	#Debug.Assert(False)

	print("test xlapp", Application)
	Application.Visible = True
	print("test xlapp Visible", Application.Visible)
	print("test xlapp", Application.Workbooks.Count)
	Application.Quit()

	fso = CreateObject("Scripting.FileSystemObject")
	print("test fso", fso)
	print("test fso", fso.GetAbsolutePathName("."))

	xhr = CreateObject("MSXML2.ServerXMLHTTP.6.0")
	print("test xhr", xhr)
	print("test xhr", xhr.Open("GET", "http://localhost:8080", False))

	adodb = CreateObject("ADODB.Connection")
	print("test adodb", adodb)

	list = CreateObject("System.Collections.ArrayList")
	print("test list", list)
	#list.Add("abc")
	#list.Add(1)
	#list.Add(2.3)

	dict = CreateObject("Scripting.Dictionary")
	print("test dict", dict)
	dict.Add("key", 12.3)
	print("test dict", "#keys", dict.Count)

	unittest.main()
