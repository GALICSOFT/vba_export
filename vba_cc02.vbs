
Const RANGE_INFO = "D3"	'헤더범위 정보있는 영역
Const RANGE_DATA = "E3"	'데이터범위 정보있는 영역


Type Tschema
	name As String
	type As String
End Type

Dim g_sheetCount As Long   '전체 시트 갯수

Dim g_infoBgnRow As Long
Dim g_infoEndRow As Long
Dim g_infoBgnCol As Long
Dim g_infoEndCol As Long
Dim g_dataBgnRow As Long
Dim g_dataEndRow As Long
Dim g_dataBgnCol As Long
Dim g_dataEndCol As Long

Dim g_infoRow As Long
Dim g_infoCol As Long
Dim g_dataRow As Long
Dim g_dataCol As Long

Dim g_sheetName	As String	'sheet name

Dim	g_schema() As Tschema


Sub Export()
	Dim saveFile	As String	'sheet name
	Dim rows		As Long		'--변수 길이: Len()
	Dim MyPath		As String
	Dim curSheet	As String	' current sheet

	g_sheetCount = Sheets.Count	' sheet count

	'MyPath = CurDir
	'MsgBox ActiveWorkbook.path

	' save the current sheet
	curSheet = ActiveSheet.name

	For idx = 1 To g_sheetCount

	   'move the target sheet
		g_sheetName = Sheets(idx).name
		Worksheets(g_sheetName).Activate

		'get the range
		Set rngInfo = ActiveSheet.Range(ActiveSheet.Range(RANGE_INFO).Value)
		Set rngData = ActiveSheet.Range(ActiveSheet.Range(RANGE_DATA).Value)

		'setup the range
		g_infoRow = rngInfo.Rows.Count - 2
		g_infoCol = rngInfo.Columns.Count
		g_dataRow = rngData.Rows.Count - 2
		g_dataCol = rngData.Columns.Count

		g_infoBgnRow = rngInfo.Row
		g_infoEndRow = g_infoBgnRow + g_infoRow

		g_infoBgnCol = rngInfo.Column
		g_infoEndCol = g_infoBgnCol + g_infoCol

		g_dataBgnRow = rngData.Row
		g_dataEndRow = g_dataBgnRow + g_dataRow
		g_dataBgnCol = rngData.Column
		g_dataEndCol = g_dataBgnCol + g_dataCol


		'Redefine the schema
		ReDim g_schema(g_infoCol)

		'setup the schema
		For i = 1 To g_infoCol
			g_schema(i).name = Trim(Cells(g_infoBgnRow + 1, g_infoBgnCol - 1 + i))
			g_schema(i).type = Trim(Cells(g_infoBgnRow + 0, g_infoBgnCol - 1 + i))
		Next



		saveFile = ActiveWorkbook.path & "\" & g_sheetName & ".lua"
		Open saveFile For Output Shared As #1

		saveFile = ActiveWorkbook.path & "\" & g_sheetName & ".bin"
		Open saveFile For Binary Access Write As #2

			WriteTxt (idx)
			WriteBin (idx)

		Close #1
		Close #2
	Next

	Worksheets(curSheet).Activate	 '해당 시트로 이동하기

	MsgBox "ExportData completed!!!"
End Sub



Sub WriteTxt(sheetIdx As Long)
	Dim schema_name  As String
	Dim schema_type  As String
	Dim table_rec   As String
	Dim i As Long	 ' for col
	Dim j As Long	 ' for row


	'write the schema
	schema_name = "'" & g_schema(1).name & "'"
	schema_type = "'" & g_schema(1).type & "'"

	For i = 2 To g_infoCol
		schema_name = schema_name & "," & "'" & g_schema(i).name & "'"
		schema_type = schema_type & "," & "'" & g_schema(i).type & "'"
	Next

	Print #1, g_sheetName & "_schema_name = {" & schema_name & "}"
	Print #1, g_sheetName & "_schema_type = {" & schema_type & "}"

	Print #1, ""
	Print #1, "-- data ----------------------------------------"
	Print #1, ""

	'write record
	table_rec = g_sheetName & "_rec_data" & " = {}"
	Print #1, table_rec

	For i = 1 To g_infoRow

		schema_type = g_schema(1).type
		table_rec = GetRec(g_infoBgnRow + 1 + i, g_infoBgnCol + 0, schema_type)

		For j = 2 To g_infoCol
			schema_type = g_schema(j).type
			table_rec = table_rec & ", " & GetRec(g_infoBgnRow + 1 + i, g_infoBgnCol - 1 + j, schema_type)
		Next

		table_rec = g_sheetName & "_rec_data[" & i & "] = {" & table_rec & " }"
		Print #1, table_rec
	Next

	Call WriteLuaFunc
End Sub


Sub WriteBin(sheetIdx As Long)
	Dim schema_name(32)  As Byte
	Dim schema_type(16)  As Byte
	Dim table_rec   As String
	Dim i As Long	 ' for col
	Dim j As Long	 ' for row


	'write the schema
	For i = 1 To g_infoCol
		schema_name = schema_name & "," & "'" & g_schema(i).name & "'"
		schema_type = schema_type & "," & "'" & g_schema(i).type & "'"
	Next

	Print #2, g_sheetName & "_schema_name = {" & schema_name & "}"
	Print #2, g_sheetName & "_schema_type = {" & schema_type & "}"

	Print #2, ""
	Print #2, "-- data ----------------------------------------"
	Print #2, ""

	'write record
	table_rec = g_sheetName & "_rec_data" & " = {}"
	Print #2, table_rec

	For i = 1 To g_infoRow

		schema_type = g_schema(1).type
		table_rec = GetRec(g_infoBgnRow + 1 + i, g_infoBgnCol + 0, schema_type)

		For j = 2 To g_infoCol
			schema_type = g_schema(j).type
			table_rec = table_rec & ", " & GetRec(g_infoBgnRow + 1 + i, g_infoBgnCol - 1 + j, schema_type)
		Next

		table_rec = g_sheetName & "_rec_data[" & i & "] = {" & table_rec & " }"
		Print #2, table_rec
	Next
End Sub



Function GetRec(row As Long, col As Long, sType As String)
	Dim rec As String
	rec = Trim(Cells(row, col))

	If sType = "function" Or sType = "string" Then
		rec = "'" & rec & "'"
	End If

	GetRec = rec
End Function




Sub WriteLuaFunc()
	Print #1, ""
	Print #1, "-- table call ----------------------------------"
	Print #1, ""

	'Write table call function
	Print #1, "function GetRec_" & g_sheetName & "(row, sch_name)"
	Print #1, "  local idx = -1"
	Print #1, ""
	Print #1, "  if 0>row or row > #" & g_sheetName & "_rec_data then"
	Print #1, "	  return nil"
	Print #1, "  end"
	Print #1, ""
	Print #1, "  if 'number' == type(sch_name) then"
	Print #1, "	  idx = sch_name"
	Print #1, ""
	Print #1, "	  if sch_name> #" & g_sheetName & "_schema_name then"
	Print #1, "		  return nil"
	Print #1, "	  end"
	Print #1, ""
	Print #1, "	  return " & g_sheetName & "_rec_data[row][idx]"
	Print #1, "  end"
	Print #1, ""
	Print #1, "  for i=1, #" & g_sheetName & "_schema_name do"
	Print #1, "	  if sch_name == " & g_sheetName & "_schema_name[i] then"
	Print #1, "		  idx = i"
	Print #1, "		  break"
	Print #1, "	  end"
	Print #1, "  end"
	Print #1, ""
	Print #1, "  if -1 == idx then"
	Print #1, "	  return nil"
	Print #1, "  end"
	Print #1, ""
	Print #1, "  return " & g_sheetName & "_rec_data[row][idx]"
	Print #1, "end"
End Sub




Public Sub StringToByteArray( ByVal dst() As Byte, ByVal src As String)
	Dim strLen  As Integer
	strLen = Len(src)

    For i = 1 To strLen
		dst(i) = CByte(Asc(Mid(src, i, 1)))
	Next i
End Sub
