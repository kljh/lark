' Extended Excel context menu
' Author : Claude Cochet
' Created : 2010-10-06
' !! Error if you add space after comment
Const a = 3

Option Explicit
Option Base 5.3

' Demo VBA file for Python transpiler

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
	'!! Dim cnx As New ADODB.Connection

	a = 3
    b = 3

	' Reminder : SHIFT is +, CTRL is ^ (caret), ALT is %, ALPHA ENTER is ~ or {RETURN}, NUMERIC KEYPAD ENTER is {ENTER}
	' Other keys: {UP}, {DOWN}, {LEFT}, {RIGHT}, {PGUP}, {PGDN}, {DOWN}, {ESC},  {DEL}, {CLEAR}, {BS}, {BREAK}, {HOME}, {INS}, {TAB}, {F1},.. {F15}
	Application.OnKey "{F1}", "SendKeysF2"          ' F1 is useless, time wasting
	' a comment after an inline comment
	Application.OnKey "^!", "ToggleTwoDigits"       ' extend standard shortcut

	'!! Error if no parentheses below
	'!! Call tmp.Sort
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

	'!! Error below
	'sht.cells(1, 1).Activate
	'names_dict.Item(nm_sheet).Add nm_bare, nm.RefersTo ' nm.RefersTo or nm.RefersToR1C1

	Dim FileNumber: FileNumber = FreeFile
    Open export_path For Output Access Write As #FileNumber
    Print #FileNumber, json
    Close #FileNumber
End Sub
