
Const RANGE_INFO = "J2"		'info range
Const RANGE_DATA = "J3"		'data range


Const DF_DATA_NONE		= 0
Const DF_DATA_INT		= 1
Const DF_DATA_DOUBLE	= 2
Const DF_DATA_STRING	= 3
Const DF_DATA_FUNCTION	= 4

Const DF_VERSION_SIZE	= 16
Const DF_OFFSET_INFO	= 32

' schema
Type Tschm
	sName As String
	nType As Long
	sType As String
	sData As String
End Type


Dim m_sheetCount As Long	'전체 시트 갯수
Dim m_sheetName  As String	'sheet sName

Dim m_infoBgnRow As Long
Dim m_infoBgnCol As Long

Dim m_infoRow	As Integer
Dim m_infoCol	As Integer
Dim m_infoSch()	As Tschm

Dim m_dataBgnRow As Long
Dim m_dataBgnCol As Long
Dim m_dataEndCol As Long

Dim m_dataRow	As Integer
Dim m_dataCol	As Integer
Dim m_dataSch()	As Tschm
Dim m_dataCel()	As String


Sub Export()
	Dim saveFile As String	'sheet sName
	Dim curSheet As String	'current sheet
	Dim i As Long			'for row
	Dim j As Long			'for col
	Dim n As Long			'for cell index
	Dim idx As Integer		'sheet index
	Dim sTmp As String		'Temp String


	m_sheetCount = Sheets.Count ' sheet count

	' save the current sheet
	curSheet = ActiveSheet.name

	For idx = 1 To m_sheetCount

	   'move the target sheet
		m_sheetName = Sheets(idx).name
		Worksheets(m_sheetName).Activate

		'get the range
		Set rngInfo = ActiveSheet.Range(ActiveSheet.Range(RANGE_INFO).Value)
		Set rngData = ActiveSheet.Range(ActiveSheet.Range(RANGE_DATA).Value)

		'setup the begin row and column
		m_infoBgnRow = rngInfo.row
		m_infoBgnCol = rngInfo.Column
		m_dataBgnRow = rngData.row
		m_dataBgnCol = rngData.Column


		'setup the row and col number
		m_infoRow = 0
		m_infoCol = 0
		m_dataRow = 0
		m_dataCol = 0

		For i = m_infoBgnRow To 32000
			sTmp = Cells(i, m_infoBgnCol)
			If 0 = Len(sTmp) Then
				Exit For
			End If

			m_infoRow = m_infoRow + 1
		Next

		For j = m_infoBgnCol To 32000
			sTmp = Cells(m_infoBgnRow, j)
			If 0 = Len(sTmp) Then
				Exit For
			End If

			m_infoCol = m_infoCol + 1
		Next

		For i = m_dataBgnRow + 2 To 32000
			sTmp = Cells(i, m_dataBgnCol)
			If 0 = Len(sTmp) Then
				Exit For
			End If

			m_dataRow = m_dataRow + 1
		Next

		For j = m_dataBgnCol To 32000
			sTmp = Cells(m_dataBgnRow, j)
			If 0 = Len(sTmp) Then
				Exit For
			End If

			m_dataCol = m_dataCol + 1
		Next
		

		If 0 < m_infoRow And 0 < m_infoRow And 0 < m_dataRow And 0 < m_dataRow Then

			'Redefine the schema
			ReDim m_infoSch(m_infoRow)
			ReDim m_dataSch(m_dataCol)
			ReDim m_dataCel(m_dataRow * m_dataCol)


			'setup the info schema and value
			Dim sType As String
			For i = 1 To m_infoRow
				m_infoSch(i).sName = Trim(Cells(m_infoBgnRow -1 + i, m_infoBgnCol + 0))
				m_infoSch(i).sType = Trim(Cells(m_infoBgnRow -1 + i, m_infoBgnCol + 1))
				m_infoSch(i).sData = Trim(Cells(m_infoBgnRow -1 + i, m_infoBgnCol + 2))

				sType = m_infoSch(i).sType
				If "int" = sType Then
					m_infoSch(i).nType = DF_DATA_INT

				ElseIf "double" = sType Then
					m_infoSch(i).nType = DF_DATA_DOUBLE

				ElseIf "string" = sType Then
					m_infoSch(i).nType = DF_DATA_STRING

				ElseIf "function" = sType Then
					m_infoSch(i).nType = DF_DATA_FUNCTION
				End If
			Next


			'setup the data schema
			For i = 1 To m_dataCol
				m_dataSch(i).sName = Trim(Cells(m_dataBgnRow + 1, m_dataBgnCol - 1 + i))
				m_dataSch(i).sType = Trim(Cells(m_dataBgnRow + 0, m_dataBgnCol - 1 + i))

				sType = m_dataSch(i).sType
				If "int" = sType Then
					m_dataSch(i).nType = DF_DATA_INT

				ElseIf "double" = sType Then
					m_dataSch(i).nType = DF_DATA_DOUBLE

				ElseIf "string" = sType Then
					m_dataSch(i).nType = DF_DATA_STRING

				ElseIf "function" = sType Then
					m_dataSch(i).nType = DF_DATA_FUNCTION
				End If
			Next


			'read the data
			For i = 1 To m_dataRow
				For j = 1 To m_dataCol
					n = (i -1) * m_dataCol + j
					m_dataCel(n) = Trim(Cells(m_dataBgnRow + 1 + i, m_dataBgnCol - 1 + j))
				Next
			Next



			saveFile = ActiveWorkbook.path & "\" & m_sheetName & ".lua"
			Open saveFile For Output Shared As #1
				WriteTxt (n)
			Close #1

			saveFile = ActiveWorkbook.path & "\" & m_sheetName & ".bin"
			Open saveFile For Binary Access Write As #2
				WriteBin (n)
			Close #2

		End If
	Next

	Worksheets(curSheet).Activate	'해당 시트로 이동하기

	'MsgBox "ExportData completed!!!"
End Sub


Sub WriteBin(sheetIdx As Long)
	Dim schema_n As String
	Dim schema_t As Long
	Dim schema_d As String
	Dim i As Long			'for row
	Dim j As Long			'for col
	Dim n As Long

	Dim tmpSpc() As Byte

	Dim db_version As String

	db_version = "DF_TBL 1.0"

	Put #2, , db_version

	Seek #2, DF_VERSION_SIZE + 1

	Put #2, , m_infoRow
	'Put #2, , m_infoCol
	Put #2, , m_dataRow
	Put #2, , m_dataCol

	Seek #2, DF_OFFSET_INFO + 1

	'write schema -----------------------------------------
	'write the info schema
	For i = 1 To m_infoRow
		schema_n = m_infoSch(i).sName
		schema_t = m_infoSch(i).nType

		'write the type
		Put #2, , schema_t

		'write the skill name
		Put #2, , schema_n

		'write the null space
		ReDim tmpSpc(27 - Len(schema_n))
		Put #2, , tmpSpc
	Next

	'write the data schema
	For i = 1 To m_dataCol
		schema_n = m_dataSch(i).sName
		schema_t = m_dataSch(i).nType

		'write the type
		Put #2, , schema_t

		'write the skill name
		Put #2, , schema_n

		'write the null space
		ReDim tmpSpc(27 - Len(schema_n))
		Put #2, , tmpSpc
	Next

	'write cells-------------------------------------------
	'write the info
	For i = 1 To m_infoRow
		schema_d = m_infoSch(i).sData
		schema_t = m_infoSch(i).nType

		PutBin schema_d, schema_t
	Next


	'write the data
	For i = 1 To m_dataRow
		For j = 1 To m_dataCol
			schema_t = m_dataSch(j).nType

			n = (i -1) * m_dataCol + j

			PutBin m_dataCel(n), schema_t
		Next
	Next

End Sub


Sub PutBin(s As String, t As Long)
	Dim sTmp As String
	Dim strLen As Integer
	Dim recInt As Long
	Dim recDbl As Double

	If DF_DATA_FUNCTION = t Or DF_DATA_STRING = t Then
	    sTmp = StrConv(s, vbFromUnicode)
		strLen = LenB(sTmp)
		Put #2, , strLen
		Put #2, , s

	ElseIf DF_DATA_INT = t Then
		recInt = CLng(s)
		Put #2, , recInt

	ElseIf DF_DATA_DOUBLE = t Then
		recDbl = CDbl(s)
		Put #2, , recDbl

	End If
End Sub


Sub WriteTxt(sheetIdx As Long)
	Dim schema_n As String
	Dim schema_t As String
	Dim schema_d As String
	Dim table_rc As String
	Dim i As Long			'for row
	Dim j As Long			'for col
	Dim n As Long

	'write the info schema
	schema_n = "'" & m_infoSch(1).sName & "'"
	schema_t = "'" & m_infoSch(1).sType & "'"
	schema_d = ToStrQut(m_infoSch(1).sData, m_infoSch(1).sType)

	For i = 2 To m_infoRow
		If 0 = i Mod 8 Then
			schema_n = schema_n & vbCrLf & Chr(9) & Chr(9)
			schema_t = schema_t & vbCrLf & Chr(9) & Chr(9)
			schema_d = schema_d & vbCrLf & Chr(9) & Chr(9)
		End If
		
		schema_n = schema_n & ", " & "'" & m_infoSch(i).sName & "'"
		schema_t = schema_t & ", " & "'" & m_infoSch(i).sType & "'"
		schema_d = schema_d & ", " & ToStrQut(m_infoSch(i).sData, m_infoSch(i).sType)
	Next

	Print #1, m_sheetName & "_info_schema_n = {"
	Print #1, Chr(9) & Chr(9) & schema_n
	Print #1, Chr(9) & Chr(9) & "}"

	Print #1, m_sheetName & "_info_schema_t = {"
	Print #1, Chr(9) & Chr(9) & schema_t
	Print #1, Chr(9) & Chr(9) & "}"

	Print #1, m_sheetName & "_info_data  = {" & schema_d & "}"
	Print #1, ""
	Print #1, "------------------------------------------------------------"
	Print #1, ""


	'write the data schema
	schema_n = "'" & m_dataSch(1).sName & "'"
	schema_t = "'" & m_dataSch(1).sType & "'"

	For i = 2 To m_dataCol
		schema_n = schema_n & ", " & "'" & m_dataSch(i).sName & "'"
		schema_t = schema_t & ", " & "'" & m_dataSch(i).sType & "'"
	Next

	Print #1, m_sheetName & "_schema_n = {" & schema_n & "}"
	Print #1, m_sheetName & "_schema_t = {" & schema_t & "}"

	Print #1, ""
	Print #1, ""

	'write data
	table_rc = m_sheetName & "_rec" & " = {}"
	Print #1, table_rc

	For i = 1 To m_dataRow
		schema_t = m_dataSch(1).sType

		n = (i -1) * m_dataCol + 1
		table_rc = ToStrQut(m_dataCel(n), schema_t)

		For j = 2 To m_dataCol
			schema_t = m_dataSch(j).sType

			n = (i -1) * m_dataCol + j
			table_rc = table_rc & ", " & ToStrQut(m_dataCel(n), schema_t)
		Next

		table_rc = m_sheetName & "_rec[" & i & "] = {" & table_rc & " }"
		Print #1, table_rc
	Next


	Print #1,""
	Print #1,"-- table call ----------------------------------"
	Print #1,""

	'Write table call function
	Print #1,"function GetRec_" & m_sheetName & "(row, sch_name)"
	Print #1,"  local col = -1"
	Print #1,""
	Print #1,"  if 0>row or row > #" & m_sheetName & "_rec then"
	Print #1,"	  return nil"
	Print #1,"  end"
	Print #1,""
	Print #1,"  if 'number' == type(sch_name) then"
	Print #1,"	  col = sch_name"
	Print #1,""
	Print #1,"	  if sch_name> #" & m_sheetName & "_schema_n then"
	Print #1,"		  return nil"
	Print #1,"	  end"
	Print #1,""
	Print #1,"	  return " & m_sheetName & "_rec[row][col]"
	Print #1,"  end"
	Print #1,""
	Print #1,"  for i=1, #" & m_sheetName & "_schema_n do"
	Print #1,"	  if sch_name == " & m_sheetName & "_schema_n[i] then"
	Print #1,"		  col = i"
	Print #1,"		  break"
	Print #1,"	  end"
	Print #1,"  end"
	Print #1,""
	Print #1,"  if -1 == col then"
	Print #1,"	  return nil"
	Print #1,"  end"
	Print #1,""
	Print #1,"  return " & m_sheetName & "_rec[row][col]"
	Print #1,"end"
	Print #1,""
	Print #1,""
End Sub


Function ToStrQut(in_str As String, sType As String)
	Dim rec As String
	rec = Trim(in_str)

	If sType = "function" Or sType = "string" Then
		rec = "'" & rec & "'"
	End If

	ToStrQut = rec
End Function


Function GetRec(row As Long, col As Long, sType As String)
	Dim rec As String
	rec = Trim(Cells(row, col))

	If sType = "function" Or sType = "string" Then
		rec = "'" & rec & "'"
	End If

	GetRec = rec
End Function

