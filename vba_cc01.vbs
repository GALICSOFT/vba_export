Const BeginRow = 5
Const BeginCol = 4

Type Tschema
    name As String
    type As String
End Type

Sub Export()

    Dim sheetCount As Long   '전체 시트 갯수
    Dim sheetName As String  'sheet name
    Dim saveFileName As String   'sheet name

    sheetCount = Sheets.Count


    Dim rows As Long    '--변수 길이: Len()

    '스키마를 구한다
    ' - 타입을 구한다
    ' - 이름을 구한다
    ' 이들을 파일에 저장한다

    Dim MyPath As String
    'MyPath = CurDir



    'MsgBox ActiveWorkbook.path

    Dim curSheet As String

    curSheet = ActiveSheet.name


    For i = 1 To sheetCount

        sheetName = Sheets(i).name

        Worksheets(sheetName).Activate   '해당 시트로 이동하기

        saveFileName = ActiveWorkbook.path & "\" & sheetName & ".lua"

        Open saveFileName For Output Shared As #1

            WriteText (i)

        Close #1
    Next

    Worksheets(curSheet).Activate     '해당 시트로 이동하기

    MsgBox "Export completed!!!"

End Sub




Sub WriteText(sheetIdx As Long)

    Dim Schema() As Tschema
    Dim row As Long
    Dim col As Long
    Dim i As Long     ' for col
    Dim j As Long     ' for row

    sheetName = Sheets(sheetIdx).name

    row = 0
    col = 0

    'setup the row and col number
    For i = BeginRow + 2 To 65000
        If 0 = Len(Cells(i, 5)) Then
            Exit For
        End If

        row = row + 1
    Next

    For i = BeginCol To 65000
        If 0 = Len(Cells(5, i)) Then
            Exit For
        End If

        col = col + 1
    Next

    'MsgBox row & " " & col


    'Redefine the schema
    ReDim Schema(col)

    'setup the schema
    For i = 1 To col
        Schema(i).name = Trim(Cells(BeginRow + 1, BeginCol - 1 + i))
        Schema(i).type = Trim(Cells(BeginRow + 0, BeginCol - 1 + i))
    Next


    'write the schema
    Dim schema_name  As String
    Dim schema_type  As String
    Dim table_rec   As String

    schema_name = "'" & Schema(1).name & "'"
    schema_type = "'" & Schema(1).type & "'"

    For i = 2 To col
        schema_name = schema_name & "," & "'" & Schema(i).name & "'"
        schema_type = schema_type & "," & "'" & Schema(i).type & "'"
    Next

    Print #1, sheetName & "_schema_name = {" & schema_name & "}"
    Print #1, sheetName & "_schema_type = {" & schema_type & "}"

    Print #1, ""
    Print #1, "-- data ----------------------------------------"
    Print #1, ""

    'write record
    table_rec = sheetName & "_rec_data" & " = {}"
    Print #1, table_rec

    For i = 1 To row

        schema_type = Schema(1).type
        table_rec = GetRec(BeginRow + 1 + i, BeginCol + 0, schema_type)

        For j = 2 To col
            schema_type = Schema(j).type
            table_rec = table_rec & ", " & GetRec(BeginRow + 1 + i, BeginCol - 1 + j, schema_type)
        Next

        table_rec = sheetName & "_rec_data[" & i & "] = {" & table_rec & " }"
        Print #1, table_rec
    Next

    Print #1, ""


    Print #1, ""
    Print #1, "-- table call ----------------------------------"
    Print #1, ""

    'Write table call function
    Print #1, "function GetRec_" & sheetName & "(row, sch_name)"
    Print #1, "  local idx = -1"
    Print #1, ""
    Print #1, "  if 0>row or row > #" & sheetName & "_rec_data then"
    Print #1, "      return nil"
    Print #1, "  end"
    Print #1, ""
    Print #1, "  if 'number' == type(sch_name) then"
    Print #1, "      idx = sch_name"
    Print #1, ""
    Print #1, "      if sch_name> #" & sheetName & "_schema_name then"
    Print #1, "          return nil"
    Print #1, "      end"
    Print #1, ""
    Print #1, "      return " & sheetName & "_rec_data[row][idx]"
    Print #1, "  end"
    Print #1, ""
    Print #1, "  for i=1, #" & sheetName & "_schema_name do"
    Print #1, "      if sch_name == " & sheetName & "_schema_name[i] then"
    Print #1, "          idx = i"
    Print #1, "          break"
    Print #1, "      end"
    Print #1, "  end"
    Print #1, ""
    Print #1, "  if -1 == idx then"
    Print #1, "      return nil"
    Print #1, "  end"
    Print #1, ""
    Print #1, "  return " & sheetName & "_rec_data[row][idx]"
    Print #1, "end"

End Sub


Function GetRec(row As Long, col As Long, sType As String)
    Dim rec As String
    rec = Trim(Cells(row, col))

    If sType = "function" Or sType = "string" Then
        rec = "'" & rec & "'"
    End If

    GetRec = rec

End Function






Sub 아이템전체데이터()

     '헤더범위 정보있는 영역
    Const HEADER_RANGE_INFO = "F2"
    '데이터범위 정보있는 영역
    Const DATA_RANGE_INFO = "F3"

    '워크시트 변수
    Dim thisSheet As Worksheet

    '헤더범위
    Dim rngHeaderRange As Range
    '데이터범위
    Dim rngDataRange As Range

    '헤더 실데이터
    Dim rngHeader As Range
    '데이터 실데이터
    Dim rngData As Range

    '데이터 컬럼수와 행수
    Dim colCount As Integer
    Dim rowCount As Integer

    Dim varStr As String

    '전체 시트 갯수
    Dim sheetCount As Integer
    '시트 이름 배열
    Dim sheetName() As String

    '데이터 사이즈
    Dim nSizeHeader As Long

    '전체 배열 사이즈
    Dim nSize As Long

    Dim i, j, k As Integer
    '오프셋
    Dim off As Long

    '데이터 변수
    Dim nData() As Byte

    '임시b
    Dim c As Range

    '전체 시트 갯수 가져오기
    sheetCount = Sheets.Count

    '시트 이름 배열 만들기
    ReDim sheetName(1 To 16) As String

    '전체 사이즈 초기화
    nSize = 0
    off = 0

    '시트 이름 가져와서 배열에 넣기
    '무기_리볼버(0) ~ 포션(15)까지
    For i = 1 To 16
        sheetName(i) = Sheets(i).name

         Worksheets(sheetName(i)).Activate       '해당 시트로 이동하기
        '현재 시트와 헤더범위와 데이터 범위를 설정
        Set thisSheet = ActiveSheet
        Set rngHeaderRange = thisSheet.Range(HEADER_RANGE_INFO)
        Set rngDataRange = thisSheet.Range(DATA_RANGE_INFO)

        '헤더와 데이터 실데이터 설정
        Set rngHeader = thisSheet.Range(rngHeaderRange.Value)
        Set rngData = thisSheet.Range(rngDataRange.Value)

        '컬럼수, 행수 설정
        'colCount = rngData.Columns.Count
        rowCount = rngData.rows.Count
        nSizeHeader = 0

        '헤더 돌면서 한 행 사이즈 체크
        For Each c In rngHeader

            Select Case Trim(c.Value)
            Case "byte"
                nSizeHeader = nSizeHeader + 1
                colCount = colCount + 1
            Case "short"
                nSizeHeader = nSizeHeader + 2
                colCount = colCount + 1
            Case "int"
                nSizeHeader = nSizeHeader + 4
                colCount = colCount + 1
            Case "long"
                nSizeHeader = nSizeHeader + 8
                colCount = colCount + 1
            Case "str8"
                nSizeHeader = nSizeHeader + 8
                colCount = colCount + 1
            Case "str16"
                nSizeHeader = nSizeHeader + 16
                colCount = colCount + 1
            Case "str24"
                nSizeHeader = nSizeHeader + 24
                colCount = colCount + 1
            Case "str32"
                nSizeHeader = nSizeHeader + 32
                colCount = colCount + 1
            Case "str40"
                nSizeHeader = nSizeHeader + 40
                colCount = colCount + 1
            Case "str64"
                nSizeHeader = nSizeHeader + 64
                colCount = colCount + 1
            Case "str128"
                nSizeHeader = nSizeHeader + 128
                colCount = colCount + 1
            End Select

        Next c

        nSize = nSize + 4 + nSizeHeader * rowCount

    Next i


    '데이터 배열잡기
    ReDim nData(0 To nSize - 1)

    '무기_리볼버(0) ~ 포션(15)까지의 데이터 저장
    For i = 1 To 16

         Worksheets(sheetName(i)).Activate       '해당 시트로 이동하기
        '현재 시트와 헤더범위와 데이터 범위를 설정
        Set thisSheet = ActiveSheet
        Set rngHeaderRange = thisSheet.Range(HEADER_RANGE_INFO)
        Set rngDataRange = thisSheet.Range(DATA_RANGE_INFO)


        '헤더와 데이터 실데이터 설정
        Set rngHeader = thisSheet.Range(rngHeaderRange.Value)
        Set rngData = thisSheet.Range(rngDataRange.Value)

        nSizeHeader = 0
         '헤더 돌면서 한 행 사이즈 체크
        For Each c In rngHeader

            Select Case Trim(c.Value)
            Case "byte"
                nSizeHeader = nSizeHeader + 1
                colCount = colCount + 1
            Case "short"
                nSizeHeader = nSizeHeader + 2
                colCount = colCount + 1
            Case "int"
                nSizeHeader = nSizeHeader + 4
                colCount = colCount + 1
            Case "long"
                nSizeHeader = nSizeHeader + 8
                colCount = colCount + 1
            Case "str8"
                nSizeHeader = nSizeHeader + 8
                colCount = colCount + 1
            Case "str16"
                nSizeHeader = nSizeHeader + 16
                colCount = colCount + 1
            Case "str24"
                nSizeHeader = nSizeHeader + 24
                colCount = colCount + 1
            Case "str32"
                nSizeHeader = nSizeHeader + 32
                colCount = colCount + 1
            Case "str40"
                nSizeHeader = nSizeHeader + 40
                colCount = colCount + 1
            Case "str64"
                nSizeHeader = nSizeHeader + 64
                colCount = colCount + 1
            Case "str128"
                nSizeHeader = nSizeHeader + 128
                colCount = colCount + 1
            End Select
        Next c


        '컬럼수, 행수 설정
        'colCount = rngData.Columns.Count
        rowCount = rngData.rows.Count

        '헤더 = 로우갯수(2) + 한줄(로우)사이즈(2)
        '컬럼갯수
        'Call writeShort(nData, off, colCount)
        'off = off + 2
        '로우갯수
        Call writeShort_little(nData, off, rowCount)
        off = off + 2
        '한로우사이즈
        Call writeShort_little(nData, off, nSizeHeader)
        off = off + 2

        For j = 1 To rngData.rows.Count

            For k = 1 To rngData.Columns.Count
                '데이터 타입에 따라서 저장
                Select Case Trim(rngHeader.Cells(1, k).Value)
                Case "byte"
                    nData(off) = (255 And Val(rngData.Cells(j, k).Value))
                    off = off + 1
                Case "short"
                    Call writeShort_little(nData, off, Val(rngData.Cells(j, k).Value))
                    off = off + 2
                Case "int"
                    Call writeInt_little(nData, off, Val(rngData.Cells(j, k).Value))
                    off = off + 4
                Case "long"
                    Call writeLong_little(nData, off, Val(rngData.Cells(j, k).Value))
                    off = off + 8
                Case "str8"
                    Call writeString(nData, off, rngData.Cells(j, k).Value, 8)
                    off = off + 8
                Case "str16"
                    Call writeString(nData, off, rngData.Cells(j, k).Value, 16)
                    off = off + 16
                Case "str24"
                    Call writeString(nData, off, rngData.Cells(j, k).Value, 24)
                    off = off + 24
                Case "str32"
                    Call writeString(nData, off, rngData.Cells(j, k).Value, 32)
                    off = off + 32
                Case "str40"
                    Call writeString(nData, off, rngData.Cells(j, k).Value, 40)
                    off = off + 40
                Case "str64"
                    Call writeString(nData, off, rngData.Cells(j, k).Value, 64)
                    off = off + 64
                Case "str128"
                    Call writeString(nData, off, rngData.Cells(j, k).Value, 128)
                    off = off + 128
                End Select

            Next k
        Next j
    Next i

    'varStr = CType(eSheetCount, String)
    Kill (ActiveWorkbook.path & "\..\..\..\BuildData\스프라이트데이터_피쳐폰\Data\itemData.bd")
    '데이터 저장하는 부분
    Dim saveFileName As String
    saveFileName = ActiveWorkbook.path & "\..\..\..\BuildData\스프라이트데이터_피쳐폰\Data\itemData.bd"

    Dim FILENUM As Integer
    FILENUM = FreeFile
    Open saveFileName For Binary Access Write As #FILENUM
        Put #FILENUM, , nData
    Close #FILENUM
    MsgBox "BuildData\스프라이트데이터_피쳐폰\Data 폴더에 itemData.bd 파일이 만들어졌습니다."
End Sub

Sub 장비세트설정()
        '헤더범위 정보있는 영역
    Const HEADER_RANGE_INFO = "F2"
    '데이터범위 정보있는 영역
    Const DATA_RANGE_INFO = "F3"

    '워크시트 변수
    Dim thisSheet As Worksheet

    '헤더범위
    Dim rngHeaderRange As Range
    '데이터범위
    Dim rngDataRange As Range

    '헤더 실데이터
    Dim rngHeader As Range
    '데이터 실데이터
    Dim rngData As Range

    '데이터 컬럼수와 행수
    Dim colCount As Integer
    Dim rowCount As Integer

    Dim varStr As String

    '현재 시트와 헤더범위와 데이터 범위를 설정
    Set thisSheet = ActiveSheet
    Set rngHeaderRange = thisSheet.Range(HEADER_RANGE_INFO)
    Set rngDataRange = thisSheet.Range(DATA_RANGE_INFO)

    '헤더와 데이터 실데이터 설정
    Set rngHeader = thisSheet.Range(rngHeaderRange.Value)
    Set rngData = thisSheet.Range(rngDataRange.Value)

    '컬럼수, 행수 설정
    'colCount = rngData.Columns.Count
    rowCount = rngData.rows.Count

    '데이터 사이즈
    Dim nSizeHeader As Long
    Dim nSize As Long

    '데이터 변수
    Dim nData() As Byte

    '임시b
    Dim c As Range

    '헤더 돌면서 한 행 사이즈 체크
    For Each c In rngHeader

        Select Case Trim(c.Value)
        Case "byte"
            nSizeHeader = nSizeHeader + 1
            colCount = colCount + 1
        Case "short"
            nSizeHeader = nSizeHeader + 2
            colCount = colCount + 1
        Case "int"
            nSizeHeader = nSizeHeader + 4
            colCount = colCount + 1
        Case "long"
            nSizeHeader = nSizeHeader + 8
            colCount = colCount + 1
        Case "str8"
            nSizeHeader = nSizeHeader + 8
            colCount = colCount + 1
        Case "str16"
            nSizeHeader = nSizeHeader + 16
            colCount = colCount + 1
        Case "str24"
            nSizeHeader = nSizeHeader + 24
            colCount = colCount + 1
        Case "str32"
            nSizeHeader = nSizeHeader + 32
            colCount = colCount + 1
        Case "str64"
            nSizeHeader = nSizeHeader + 64
            colCount = colCount + 1
        Case "str128"
            nSizeHeader = nSizeHeader + 128
            colCount = colCount + 1
        End Select

    Next c

    '추가 사이즈(4 = rowCount(2) + rowSize(2))
    nSize = 4 + nSizeHeader * rowCount

    '데이터 배열잡기
    ReDim nData(0 To nSize - 1)


    Dim i, j As Integer
    '오프셋
    Dim off As Long

    '데이터 돌면서 값가져오기
    off = 0

    '헤더 = 컬럼갯수(2) + 로우갯수(2) + 한줄(로우)사이즈(2)
    '컬럼갯수
    'Call writeShort(nData, off, colCount)
    'off = off + 2
    '로우갯수
    Call writeShort_little(nData, off, rowCount)
    off = off + 2
    '한로우사이즈
    Call writeShort_little(nData, off, nSizeHeader)
    off = off + 2


    For i = 1 To rngData.rows.Count

        For j = 1 To rngData.Columns.Count
            '데이터 타입에 따라서 저장
            Select Case Trim(rngHeader.Cells(1, j).Value)
            Case "byte"
                nData(off) = (255 And Val(rngData.Cells(i, j).Value))
                off = off + 1
            Case "short"
                Call writeShort_little(nData, off, Val(rngData.Cells(i, j).Value))
                off = off + 2
            Case "int"
                Call writeInt_little(nData, off, Val(rngData.Cells(i, j).Value))
                off = off + 4
            Case "long"
                Call writeLong_little(nData, off, Val(rngData.Cells(i, j).Value))
                off = off + 8
            Case "str8"
                Call writeString(nData, off, rngData.Cells(i, j).Value, 8)
                off = off + 8
            Case "str16"
                Call writeString(nData, off, rngData.Cells(i, j).Value, 16)
                off = off + 16
            Case "str24"
                Call writeString(nData, off, rngData.Cells(i, j).Value, 24)
                off = off + 24
            Case "str32"
                Call writeString(nData, off, rngData.Cells(i, j).Value, 32)
                off = off + 32
            Case "str64"
                Call writeString(nData, off, rngData.Cells(i, j).Value, 64)
                off = off + 64
            Case "str128"
                Call writeString(nData, off, rngData.Cells(i, j).Value, 128)
                off = off + 128
            End Select

        Next j

    Next i

    'varStr = CType(eSheetCount, String)

    '데이터 저장하는 부분
    Dim saveFileName As String
    'SaveFileName = ActiveWorkbook.Path & "\" & thisSheet.Name & "(" & Format(Now(), "yymmdd") & ").dat"
    saveFileName = ActiveWorkbook.path & "\..\..\..\BuildData\스프라이트데이터_피쳐폰\Data\setItem.bd"

    Dim FILENUM As Integer
    FILENUM = FreeFile
    Open saveFileName For Binary Access Write As #FILENUM
        Put #FILENUM, , nData
    Close #FILENUM
    MsgBox "BuildData\스프라이트데이터_피쳐폰\Data 폴더에 setItem.bd 파일이 만들어졌습니다"
End Sub

Sub 항아리테이블()
        '헤더범위 정보있는 영역
    Const HEADER_RANGE_INFO = "F2"
    '데이터범위 정보있는 영역
    Const DATA_RANGE_INFO = "F3"

    '워크시트 변수
    Dim thisSheet As Worksheet

    '헤더범위
    Dim rngHeaderRange As Range
    '데이터범위
    Dim rngDataRange As Range

    '헤더 실데이터
    Dim rngHeader As Range
    '데이터 실데이터
    Dim rngData As Range

    '데이터 컬럼수와 행수
    Dim colCount As Integer
    Dim rowCount As Integer

    Dim varStr As String

    '현재 시트와 헤더범위와 데이터 범위를 설정
    Set thisSheet = ActiveSheet
    Set rngHeaderRange = thisSheet.Range(HEADER_RANGE_INFO)
    Set rngDataRange = thisSheet.Range(DATA_RANGE_INFO)

    '헤더와 데이터 실데이터 설정
    Set rngHeader = thisSheet.Range(rngHeaderRange.Value)
    Set rngData = thisSheet.Range(rngDataRange.Value)

    '컬럼수, 행수 설정
    'colCount = rngData.Columns.Count
    rowCount = rngData.rows.Count

    '데이터 사이즈
    Dim nSizeHeader As Long
    Dim nSize As Long

    '데이터 변수
    Dim nData() As Byte

    '임시b
    Dim c As Range

    '헤더 돌면서 한 행 사이즈 체크
    For Each c In rngHeader

        Select Case Trim(c.Value)
        Case "byte"
            nSizeHeader = nSizeHeader + 1
            colCount = colCount + 1
        Case "short"
            nSizeHeader = nSizeHeader + 2
            colCount = colCount + 1
        Case "int"
            nSizeHeader = nSizeHeader + 4
            colCount = colCount + 1
        Case "long"
            nSizeHeader = nSizeHeader + 8
            colCount = colCount + 1
        Case "str8"
            nSizeHeader = nSizeHeader + 8
            colCount = colCount + 1
        Case "str16"
            nSizeHeader = nSizeHeader + 16
            colCount = colCount + 1
        Case "str24"
            nSizeHeader = nSizeHeader + 24
            colCount = colCount + 1
        Case "str32"
            nSizeHeader = nSizeHeader + 32
            colCount = colCount + 1
        Case "str64"
            nSizeHeader = nSizeHeader + 64
            colCount = colCount + 1
        Case "str128"
            nSizeHeader = nSizeHeader + 128
            colCount = colCount + 1
        End Select

    Next c

    '추가 사이즈(4 = rowCount(2) + rowSize(2))
    nSize = 4 + nSizeHeader * rowCount

    '데이터 배열잡기
    ReDim nData(0 To nSize - 1)


    Dim i, j As Integer
    '오프셋
    Dim off As Long

    '데이터 돌면서 값가져오기
    off = 0

    '헤더 = 컬럼갯수(2) + 로우갯수(2) + 한줄(로우)사이즈(2)
    '컬럼갯수
    'Call writeShort(nData, off, colCount)
    'off = off + 2
    '로우갯수
    Call writeShort_little(nData, off, rowCount)
    off = off + 2
    '한로우사이즈
    Call writeShort_little(nData, off, nSizeHeader)
    off = off + 2

     For i = 1 To rngData.rows.Count

        For j = 1 To rngData.Columns.Count
            '데이터 타입에 따라서 저장
            Select Case Trim(rngHeader.Cells(1, j).Value)
            Case "byte"
                nData(off) = (255 And Val(rngData.Cells(i, j).Value))
                off = off + 1
            Case "short"
                Call writeShort_little(nData, off, Val(rngData.Cells(i, j).Value))
                off = off + 2
            Case "int"
                Call writeInt_little(nData, off, Val(rngData.Cells(i, j).Value))
                off = off + 4
            Case "long"
                Call writeLong_little(nData, off, Val(rngData.Cells(i, j).Value))
                off = off + 8
            Case "str8"
                Call writeString(nData, off, rngData.Cells(i, j).Value, 8)
                off = off + 8
            Case "str16"
                Call writeString(nData, off, rngData.Cells(i, j).Value, 16)
                off = off + 16
            Case "str24"
                Call writeString(nData, off, rngData.Cells(i, j).Value, 24)
                off = off + 24
            Case "str32"
                Call writeString(nData, off, rngData.Cells(i, j).Value, 32)
                off = off + 32
            Case "str64"
                Call writeString(nData, off, rngData.Cells(i, j).Value, 64)
                off = off + 64
            Case "str128"
                Call writeString(nData, off, rngData.Cells(i, j).Value, 128)
                off = off + 128
            End Select

        Next j

    Next i

    'varStr = CType(eSheetCount, String)

    '데이터 저장하는 부분
    Dim saveFileName As String
    saveFileName = ActiveWorkbook.path & "\..\..\..\BuildData\스프라이트데이터_피쳐폰\Data\potTable.bd"

    Dim FILENUM As Integer
    FILENUM = FreeFile
    Open saveFileName For Binary Access Write As #FILENUM
        Put #FILENUM, , nData
    Close #FILENUM
    MsgBox "BuildData\스프라이트데이터_피쳐폰\Data 폴더에 potTable.bd 파일이 만들어졌습니다"
End Sub

Sub 세라아이템()
         '헤더범위 정보있는 영역
    Const HEADER_RANGE_INFO = "E2"
    '데이터범위 정보있는 영역
    Const DATA_RANGE_INFO = "E3"

    '워크시트 변수
    Dim thisSheet As Worksheet

    '헤더범위
    Dim rngHeaderRange As Range
    '데이터범위
    Dim rngDataRange As Range

    '헤더 실데이터
    Dim rngHeader As Range
    '데이터 실데이터
    Dim rngData As Range

    '데이터 컬럼수와 행수
    Dim colCount As Integer
    Dim rowCount As Integer

    Dim varStr As String

    '현재 시트와 헤더범위와 데이터 범위를 설정
    Set thisSheet = ActiveSheet
    Set rngHeaderRange = thisSheet.Range(HEADER_RANGE_INFO)
    Set rngDataRange = thisSheet.Range(DATA_RANGE_INFO)

    '헤더와 데이터 실데이터 설정
    Set rngHeader = thisSheet.Range(rngHeaderRange.Value)
    Set rngData = thisSheet.Range(rngDataRange.Value)

    '컬럼수, 행수 설정
    'colCount = rngData.Columns.Count
    rowCount = rngData.rows.Count

    '데이터 사이즈
    Dim nSizeHeader As Long
    Dim nSize As Long

    '데이터 변수
    Dim nData() As Byte

    '임시b
    Dim c As Range

    '헤더 돌면서 한 행 사이즈 체크
    For Each c In rngHeader

        Select Case Trim(c.Value)
        Case "byte"
            nSizeHeader = nSizeHeader + 1
            colCount = colCount + 1
        Case "short"
            nSizeHeader = nSizeHeader + 2
            colCount = colCount + 1
        Case "int"
            nSizeHeader = nSizeHeader + 4
            colCount = colCount + 1
        Case "long"
            nSizeHeader = nSizeHeader + 8
            colCount = colCount + 1
        Case "str8"
            nSizeHeader = nSizeHeader + 8
            colCount = colCount + 1
        Case "str16"
            nSizeHeader = nSizeHeader + 16
            colCount = colCount + 1
        Case "str24"
            nSizeHeader = nSizeHeader + 24
            colCount = colCount + 1
        Case "str32"
            nSizeHeader = nSizeHeader + 32
            colCount = colCount + 1
        Case "str64"
            nSizeHeader = nSizeHeader + 64
            colCount = colCount + 1
        Case "str128"
            nSizeHeader = nSizeHeader + 128
            colCount = colCount + 1
        End Select

    Next c

    '추가 사이즈(4 = rowCount(2) + rowSize(2))
    nSize = 4 + nSizeHeader * rowCount

    '데이터 배열잡기
    ReDim nData(0 To nSize - 1)


    Dim i, j As Integer
    '오프셋
    Dim off As Long

    '데이터 돌면서 값가져오기
    off = 0

    '헤더 = 컬럼갯수(2) + 로우갯수(2) + 한줄(로우)사이즈(2)
    '컬럼갯수
    'Call writeShort(nData, off, colCount)
    'off = off + 2
    '로우갯수
    Call writeShort_little(nData, off, rowCount)
    off = off + 2
    '한로우사이즈
    Call writeShort_little(nData, off, nSizeHeader)
    off = off + 2


    For i = 1 To rngData.rows.Count

        For j = 1 To rngData.Columns.Count
            '데이터 타입에 따라서 저장
            Select Case Trim(rngHeader.Cells(1, j).Value)
            Case "byte"
                nData(off) = (255 And Val(rngData.Cells(i, j).Value))
                off = off + 1
            Case "short"
                Call writeShort_little(nData, off, Val(rngData.Cells(i, j).Value))
                off = off + 2
            Case "int"
                Call writeInt_little(nData, off, Val(rngData.Cells(i, j).Value))
                off = off + 4
            Case "long"
                Call writeLong_little(nData, off, Val(rngData.Cells(i, j).Value))
                off = off + 8
            Case "str8"
                Call writeString(nData, off, rngData.Cells(i, j).Value, 8)
                off = off + 8
            Case "str16"
                Call writeString(nData, off, rngData.Cells(i, j).Value, 16)
                off = off + 16
            Case "str24"
                Call writeString(nData, off, rngData.Cells(i, j).Value, 24)
                off = off + 24
            Case "str32"
                Call writeString(nData, off, rngData.Cells(i, j).Value, 32)
                off = off + 32
            Case "str64"
                Call writeString(nData, off, rngData.Cells(i, j).Value, 64)
                off = off + 64
            Case "str128"
                Call writeString(nData, off, rngData.Cells(i, j).Value, 128)
                off = off + 128
            End Select

        Next j

    Next i

    'varStr = CType(eSheetCount, String)

    '데이터 저장하는 부분
    Dim saveFileName As String
    'SaveFileName = ActiveWorkbook.Path & "\" & thisSheet.Name & "(" & Format(Now(), "yymmdd") & ").dat"
    saveFileName = ActiveWorkbook.path & "\..\..\..\BuildData\스프라이트데이터_피쳐폰\Data\ceraItem.bd"

    Dim FILENUM As Integer
    FILENUM = FreeFile
    Open saveFileName For Binary Access Write As #FILENUM
        Put #FILENUM, , nData
    Close #FILENUM
    MsgBox "BuildData\스프라이트데이터_피쳐폰\Data 폴더에 ceraItem.bd파일이 만들어졌습니다"

End Sub






'***참고사항: vb에서 int는 일반적으로 2byte(short)이며 long는 4byte(int)이다.
' vb에서 시프트 연산자가 안되므로 왼쪽시프트는 * 2 ^ 시프트값 오른쪽은 / 2 ^ 시프트값
' 0xff같은 16진수 안되므로 그냥 십진수로 and 또는 or함. 255

'byte[2]->short로 변환 (받는건 리틀로 받는다. vb의 엔디안방식)

'public short readShort( byte[] data, int index ) {
'   return (short)(((data[index] & 0xff) << 8) | (data[index + 1] & 0xff));
'}

 Public Function readShort(ByRef data() As Byte, ByVal index As Long) As Integer
    readShort = ((data(index + 1) And 255) * 2 ^ 8) Or (data(index) And 255)
 End Function


'short->byte[2]로 변환 (보내는건 빅으로 현재프로젝트가 자바임)

'public void writeShort( byte[] data, int index, short i ) {
'   data[index] = (byte)(0xff & (i >> 8));
'   data[index + 1] = (byte)(0xff & i);
'}

Public Sub writeShort(ByRef data() As Byte, ByVal index As Long, ByVal i As Integer)
   data(index) = (255 And Int((i / 2 ^ 8)))
   data(index + 1) = (255 And i)
End Sub

Public Sub writeShort_little(ByRef data() As Byte, ByVal index As Long, ByVal i As Integer)
   data(index + 1) = (255 And Int((i / 2 ^ 8)))
   data(index) = (255 And i)
End Sub

'byte[4]->int로 변환

'public int readInt( byte[] data, int index ) {
'   return ((data[index] & 0xff) << 24) | ((data[index + 1] & 0xff) << 16) | ((data[index + 2] & 0xff) << 8) | (data[index + 3] & 0xff);
'}

Public Function readInt(ByRef data() As Byte, ByVal index As Long) As Long
    readInt = ((data(index + 3) And 255) * 2 ^ 24) Or ((data(index + 2) And 255) * 2 ^ 16) Or ((data(index + 1) And 255) * 2 ^ 8) Or (data(index) And 255)
End Function

'int->byte[4]수로 변환

'public void writeInt( byte[] data, int index, int i ) {
'   data[index] = (byte)(0xff & (i >> 24));
'   data[index + 1] = (byte)(0xff & (i >> 16));
'   data[index + 2] = (byte)(0xff & (i >> 8));
'   data[index + 3] = (byte)(0xff & i);
'}

Public Sub writeInt(ByRef data() As Byte, ByVal index As Long, ByVal i As Long)
   data(index) = (255 And Int((i / 2 ^ 24)))
   data(index + 1) = (255 And Int((i / 2 ^ 16)))
   data(index + 2) = (255 And Int((i / 2 ^ 8)))
   data(index + 3) = (255 And i)
End Sub

Public Sub writeInt_little(ByRef data() As Byte, ByVal index As Long, ByVal i As Long)
   data(index + 3) = (255 And Int((i / 2 ^ 24)))
   data(index + 2) = (255 And Int((i / 2 ^ 16)))
   data(index + 1) = (255 And Int((i / 2 ^ 8)))
   data(index) = (255 And i)
End Sub

'byte[8]->long 변환

'public int readLong( byte[] data, int index ) {
'   return ((data[index] & 0xff) << 56) | ((data[index + 1] & 0xff) << 48) | ((data[index + 2] & 0xff) << 40) | ((data[index + 3] & 0xff) << 32) | ((data[index + 4] & 0xff) << 24) | ((data[index + 5] & 0xff) << 16) | ((data[index + 6] & 0xff) << 8) | (data[index + 7] & 0xff);
'}

Public Function readLong(ByRef data() As Byte, ByVal index As Long) As Currency
    readLong = ((data(index + 7) And 255) * 2 ^ 56) Or ((data(index + 6) And 255) * 2 ^ 48) Or ((data(index + 5) And 255) * 2 ^ 40) Or ((data(index + 4) And 255) * 2 ^ 32) Or ((data(index + 3) And 255) * 2 ^ 24) Or ((data(index + 2) And 255) * 2 ^ 16) Or ((data(index + 1) And 255) * 2 ^ 8) Or (data(index) And 255)
End Function

'long->byte[8]로 변환

'    public void writeLong( byte[] data, int index, long i ) {
'        data[index] = (byte)(0xff & (i >> 56));
'        data[index + 1] = (byte)(0xff & (i >> 48));
'        data[index + 2] = (byte)(0xff & (i >> 40));
'        data[index + 3] = (byte)(0xff & (i >> 32));
'        data[index + 4] = (byte)(0xff & (i >> 24));
'        data[index + 5] = (byte)(0xff & (i >> 16));
'        data[index + 6] = (byte)(0xff & (i >> 8));
'        data[index + 7] = (byte)(0xff & i);
'    }

Public Sub writeLong(ByRef data() As Byte, ByVal index As Long, ByVal i As Currency)
    data(index) = (255 And Int((i / 2 ^ 56)))
    data(index + 1) = (255 And Int((i / 2 ^ 48)))
    data(index + 2) = (255 And Int((i / 2 ^ 40)))
    data(index + 3) = (255 And Int((i / 2 ^ 32)))
    data(index + 4) = (255 And Int((i / 2 ^ 24)))
    data(index + 5) = (255 And Int((i / 2 ^ 16)))
    data(index + 6) = (255 And Int((i / 2 ^ 8)))
    data(index + 7) = (255 & i)

End Sub

Public Sub writeLong_little(ByRef data() As Byte, ByVal index As Long, ByVal i As Currency)
    data(index + 7) = (255 And Int((i / 2 ^ 56)))
    data(index + 6) = (255 And Int((i / 2 ^ 48)))
    data(index + 5) = (255 And Int((i / 2 ^ 40)))
    data(index + 4) = (255 And Int((i / 2 ^ 32)))
    data(index + 3) = (255 And Int((i / 2 ^ 24)))
    data(index + 2) = (255 And Int((i / 2 ^ 16)))
    data(index + 1) = (255 And Int((i / 2 ^ 8)))
    data(index) = (255 & i)

End Sub

'String->byte[]로 변환

'    public void writeString( byte[] data, int index, String str, int size ) {
'        byte byteString[] = str.getBytes();
'        for( int i = 0 ; i < size ; i++ ) {
'            if( i < byteString.length ) {
'                data[index + i] = byteString[i];
'            } else {
'                data[index + i] = 0;
'            }
'        }
'    }

Public Sub writeString(ByRef data() As Byte, ByVal index As Long, ByVal str As String, ByVal size As Long)
    Dim i As Long

    Dim sTmp As String

    sTmp = StrConv(str, vbFromUnicode)


    For i = 0 To size - 1

        If (i < LenB(sTmp)) Then
            data(index + i) = AscB(MidB(sTmp, i + 1, 1))
        Else
            data(index + i) = 0
        End If

    Next i

End Sub

'byte[]->String로 변환

'    public String readString( byte[] data, int index, int length ) {
'        byte byteTemp[] = new byte[length];
'
'        for( int i = 0 ; i < length ; i++ ) {
'            byteTemp[i] = data[index + i];
'        }
'
'        String strTemp = new String( byteTemp );
'        return strTemp.trim();
'    }

Public Function readString(ByRef data() As Byte, ByVal index As Long, ByVal length As Long) As String
    Dim sTmp As String

    For i = 0 To length - 1
        sTmp = sTmp & ChrB(data(index + i))
    Next i

    readString = Trim(sTmp)
End Function
