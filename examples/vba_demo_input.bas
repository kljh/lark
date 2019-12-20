' Extended Excel context menu
' Author : Claude Cochet
' Created : 2010-10-06

Const a = 3

Option Explicit
Option Base 5.3

' Demo VBA file for Python transpiler

' VBA function with types in the signature becomes a Python function with type annotations (PEP 483, 484)
Function tryDivide(a as Double, b as Double) as Double
	On Error GoTo nodivide
	a = a / b
	tryDivide = b / a

	Exit Function
nodivide:
	tryDivide = a * b
End Function

' VBA with Optional, ByRef/ByVal
Function f2(Optional ByRef s0 As String, Optional ByVal s1 = "v1", Optional s2 As String = "") As Long
	f2 = 5*7
End Function

Sub fct()
	' Static Variable
    Static bRefreshCalcTimeDisplay As Boolean

	Goto anchor
anchor:
	Dim i%, j%, nl%, nh%, ml%, mh%
	Dim i&, j&, nl&, nh&, ml&, mh& ' flt#
	ReDim header$(ml% To mh%)

	Dim a
	Dim b As Long
	Dim shp As New Shape

	Dim ctlg
	Set ctlg = New ADOX.Catalog
	Dim ctlg2 : Set ctlg2 = New ADOX.Catalog
	Dim cnx As New ADODB.Connection

	a = 3
    b = 3

	' Reminder : SHIFT is +, CTRL is ^ (caret), ALT is %, ALPHA ENTER is ~ or {RETURN}, NUMERIC KEYPAD ENTER is {ENTER}
	' Other keys: {UP}, {DOWN}, {LEFT}, {RIGHT}, {PGUP}, {PGDN}, {DOWN}, {ESC},  {DEL}, {CLEAR}, {BS}, {BREAK}, {HOME}, {INS}, {TAB}, {F1},.. {F15}
	Application.OnKey "{F1}", "SendKeysF2"          ' F1 is useless, time wasting
	' a comment after an inline comment
	Application.OnKey "^!", "ToggleTwoDigits"       ' extend standard shortcut

	If a() Then b() Else c = -1   ' trouble with b() !!

	If funcs = True Then

		' no parentheses below if no arguments
		Call tmp.Sort
		Call fct()
		Call fct( 123, 456 )
		Call fct( a := 123, b := 456 )
		Call fct( 123, New Dictionary )
		Call fct( 123, 456, )
		Call fct( 123,, 789 )
		Call fct( , 456, 789 )
		Call fct( 123, New Dictionary )
		names_dict.Add nm_sheet, New Dictionary

		Err.Raise vbObjectError, , "Oops"

	ElseIf subs = True Then
		sht.cells(1, 1).Activate
		sht.cells(1).Activate
		sht.cells(1, 1).Activate 123
		sht.cells(1).Activate 123
		sht.cells(1, 1).Activate abc
		sht.cells(1, 1).Activate abc, def
		sht.cells(1, 1).Activate abc, def.xyz
		sht.cells(1, 1).Activate abc, def.xyz  ' comment
	End If

	tmpChart.ChartArea.Width = shp.Width
	rs.fields(data(nl&)(j&)).Value = data(i&)(j&)
	Application.Union(rng, topLeftCell.CurrentArray).Select
	'names_dict.Item(nm_sheet).Add
	'names_dict.Item(nm_sheet).Add 123
	'names_dict.Item(nm_sheet).Add nm_bare
	'names_dict.Item(nm_sheet).Add nm_bare, nm
	'names_dict.Item(nm_sheet).Add nm_bare, nm.RefersTo
	'names_dict.Item(nm_sheet).Add nm_bare, nm.RefersTo ' nm.RefersTo or nm.RefersToR1C1

	For i = 1 To NbFiles() Step 7 Mod 3
		Dim FileNumber: FileNumber = FreeFile
		Open export_path For Output Access Write As #FileNumber
		Print #FileNumber, json
		Close #FileNumber
	Next i
End Sub

Sub loops()
    Debug.Print "While . . Wend"
    While False
        Debug.Print "zero+"
    Wend

    Debug.Print "Do While . . Loop"
    Do While False
        Debug.Print "zero+"
    Loop

    Debug.Print "Do . Loop While . "
    Do
        Debug.Print "one+"
    Loop While False

    Debug.Print "Do Until  . . Loop"
    Do Until True
        Debug.Print "zero+"
    Loop

    Debug.Print "Do . Loop Until . "
    Do
        Debug.Print "one+"
    Loop Until True
End Sub
