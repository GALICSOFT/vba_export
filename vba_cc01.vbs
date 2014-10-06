Const BeginRow = 5
Const BeginCol = 4

Type Tschema
    name As String
    type As String
End Type

Sub Export()

    Dim sheetCount As Long   '��ü ��Ʈ ����
    Dim sheetName As String  'sheet name
    Dim saveFileName As String   'sheet name

    sheetCount = Sheets.Count


    Dim rows As Long    '--���� ����: Len()

    '��Ű���� ���Ѵ�
    ' - Ÿ���� ���Ѵ�
    ' - �̸��� ���Ѵ�
    ' �̵��� ���Ͽ� �����Ѵ�

    Dim MyPath As String
    'MyPath = CurDir



    'MsgBox ActiveWorkbook.path

    Dim curSheet As String

    curSheet = ActiveSheet.name


    For i = 1 To sheetCount

        sheetName = Sheets(i).name

        Worksheets(sheetName).Activate   '�ش� ��Ʈ�� �̵��ϱ�

        saveFileName = ActiveWorkbook.path & "\" & sheetName & ".lua"

        Open saveFileName For Output Shared As #1

            WriteText (i)

        Close #1
    Next

    Worksheets(curSheet).Activate     '�ش� ��Ʈ�� �̵��ϱ�

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






Sub ��������ü������()

     '������� �����ִ� ����
    Const HEADER_RANGE_INFO = "F2"
    '�����͹��� �����ִ� ����
    Const DATA_RANGE_INFO = "F3"

    '��ũ��Ʈ ����
    Dim thisSheet As Worksheet

    '�������
    Dim rngHeaderRange As Range
    '�����͹���
    Dim rngDataRange As Range

    '��� �ǵ�����
    Dim rngHeader As Range
    '������ �ǵ�����
    Dim rngData As Range

    '������ �÷����� ���
    Dim colCount As Integer
    Dim rowCount As Integer

    Dim varStr As String

    '��ü ��Ʈ ����
    Dim sheetCount As Integer
    '��Ʈ �̸� �迭
    Dim sheetName() As String

    '������ ������
    Dim nSizeHeader As Long

    '��ü �迭 ������
    Dim nSize As Long

    Dim i, j, k As Integer
    '������
    Dim off As Long

    '������ ����
    Dim nData() As Byte

    '�ӽ�b
    Dim c As Range

    '��ü ��Ʈ ���� ��������
    sheetCount = Sheets.Count

    '��Ʈ �̸� �迭 �����
    ReDim sheetName(1 To 16) As String

    '��ü ������ �ʱ�ȭ
    nSize = 0
    off = 0

    '��Ʈ �̸� �����ͼ� �迭�� �ֱ�
    '����_������(0) ~ ����(15)����
    For i = 1 To 16
        sheetName(i) = Sheets(i).name

         Worksheets(sheetName(i)).Activate       '�ش� ��Ʈ�� �̵��ϱ�
        '���� ��Ʈ�� ��������� ������ ������ ����
        Set thisSheet = ActiveSheet
        Set rngHeaderRange = thisSheet.Range(HEADER_RANGE_INFO)
        Set rngDataRange = thisSheet.Range(DATA_RANGE_INFO)

        '����� ������ �ǵ����� ����
        Set rngHeader = thisSheet.Range(rngHeaderRange.Value)
        Set rngData = thisSheet.Range(rngDataRange.Value)

        '�÷���, ��� ����
        'colCount = rngData.Columns.Count
        rowCount = rngData.rows.Count
        nSizeHeader = 0

        '��� ���鼭 �� �� ������ üũ
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


    '������ �迭���
    ReDim nData(0 To nSize - 1)

    '����_������(0) ~ ����(15)������ ������ ����
    For i = 1 To 16

         Worksheets(sheetName(i)).Activate       '�ش� ��Ʈ�� �̵��ϱ�
        '���� ��Ʈ�� ��������� ������ ������ ����
        Set thisSheet = ActiveSheet
        Set rngHeaderRange = thisSheet.Range(HEADER_RANGE_INFO)
        Set rngDataRange = thisSheet.Range(DATA_RANGE_INFO)


        '����� ������ �ǵ����� ����
        Set rngHeader = thisSheet.Range(rngHeaderRange.Value)
        Set rngData = thisSheet.Range(rngDataRange.Value)

        nSizeHeader = 0
         '��� ���鼭 �� �� ������ üũ
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


        '�÷���, ��� ����
        'colCount = rngData.Columns.Count
        rowCount = rngData.rows.Count

        '��� = �ο찹��(2) + ����(�ο�)������(2)
        '�÷�����
        'Call writeShort(nData, off, colCount)
        'off = off + 2
        '�ο찹��
        Call writeShort_little(nData, off, rowCount)
        off = off + 2
        '�ѷο������
        Call writeShort_little(nData, off, nSizeHeader)
        off = off + 2

        For j = 1 To rngData.rows.Count

            For k = 1 To rngData.Columns.Count
                '������ Ÿ�Կ� ���� ����
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
    Kill (ActiveWorkbook.path & "\..\..\..\BuildData\��������Ʈ������_������\Data\itemData.bd")
    '������ �����ϴ� �κ�
    Dim saveFileName As String
    saveFileName = ActiveWorkbook.path & "\..\..\..\BuildData\��������Ʈ������_������\Data\itemData.bd"

    Dim FILENUM As Integer
    FILENUM = FreeFile
    Open saveFileName For Binary Access Write As #FILENUM
        Put #FILENUM, , nData
    Close #FILENUM
    MsgBox "BuildData\��������Ʈ������_������\Data ������ itemData.bd ������ ����������ϴ�."
End Sub

Sub ���Ʈ����()
        '������� �����ִ� ����
    Const HEADER_RANGE_INFO = "F2"
    '�����͹��� �����ִ� ����
    Const DATA_RANGE_INFO = "F3"

    '��ũ��Ʈ ����
    Dim thisSheet As Worksheet

    '�������
    Dim rngHeaderRange As Range
    '�����͹���
    Dim rngDataRange As Range

    '��� �ǵ�����
    Dim rngHeader As Range
    '������ �ǵ�����
    Dim rngData As Range

    '������ �÷����� ���
    Dim colCount As Integer
    Dim rowCount As Integer

    Dim varStr As String

    '���� ��Ʈ�� ��������� ������ ������ ����
    Set thisSheet = ActiveSheet
    Set rngHeaderRange = thisSheet.Range(HEADER_RANGE_INFO)
    Set rngDataRange = thisSheet.Range(DATA_RANGE_INFO)

    '����� ������ �ǵ����� ����
    Set rngHeader = thisSheet.Range(rngHeaderRange.Value)
    Set rngData = thisSheet.Range(rngDataRange.Value)

    '�÷���, ��� ����
    'colCount = rngData.Columns.Count
    rowCount = rngData.rows.Count

    '������ ������
    Dim nSizeHeader As Long
    Dim nSize As Long

    '������ ����
    Dim nData() As Byte

    '�ӽ�b
    Dim c As Range

    '��� ���鼭 �� �� ������ üũ
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

    '�߰� ������(4 = rowCount(2) + rowSize(2))
    nSize = 4 + nSizeHeader * rowCount

    '������ �迭���
    ReDim nData(0 To nSize - 1)


    Dim i, j As Integer
    '������
    Dim off As Long

    '������ ���鼭 ����������
    off = 0

    '��� = �÷�����(2) + �ο찹��(2) + ����(�ο�)������(2)
    '�÷�����
    'Call writeShort(nData, off, colCount)
    'off = off + 2
    '�ο찹��
    Call writeShort_little(nData, off, rowCount)
    off = off + 2
    '�ѷο������
    Call writeShort_little(nData, off, nSizeHeader)
    off = off + 2


    For i = 1 To rngData.rows.Count

        For j = 1 To rngData.Columns.Count
            '������ Ÿ�Կ� ���� ����
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

    '������ �����ϴ� �κ�
    Dim saveFileName As String
    'SaveFileName = ActiveWorkbook.Path & "\" & thisSheet.Name & "(" & Format(Now(), "yymmdd") & ").dat"
    saveFileName = ActiveWorkbook.path & "\..\..\..\BuildData\��������Ʈ������_������\Data\setItem.bd"

    Dim FILENUM As Integer
    FILENUM = FreeFile
    Open saveFileName For Binary Access Write As #FILENUM
        Put #FILENUM, , nData
    Close #FILENUM
    MsgBox "BuildData\��������Ʈ������_������\Data ������ setItem.bd ������ ����������ϴ�"
End Sub

Sub �׾Ƹ����̺�()
        '������� �����ִ� ����
    Const HEADER_RANGE_INFO = "F2"
    '�����͹��� �����ִ� ����
    Const DATA_RANGE_INFO = "F3"

    '��ũ��Ʈ ����
    Dim thisSheet As Worksheet

    '�������
    Dim rngHeaderRange As Range
    '�����͹���
    Dim rngDataRange As Range

    '��� �ǵ�����
    Dim rngHeader As Range
    '������ �ǵ�����
    Dim rngData As Range

    '������ �÷����� ���
    Dim colCount As Integer
    Dim rowCount As Integer

    Dim varStr As String

    '���� ��Ʈ�� ��������� ������ ������ ����
    Set thisSheet = ActiveSheet
    Set rngHeaderRange = thisSheet.Range(HEADER_RANGE_INFO)
    Set rngDataRange = thisSheet.Range(DATA_RANGE_INFO)

    '����� ������ �ǵ����� ����
    Set rngHeader = thisSheet.Range(rngHeaderRange.Value)
    Set rngData = thisSheet.Range(rngDataRange.Value)

    '�÷���, ��� ����
    'colCount = rngData.Columns.Count
    rowCount = rngData.rows.Count

    '������ ������
    Dim nSizeHeader As Long
    Dim nSize As Long

    '������ ����
    Dim nData() As Byte

    '�ӽ�b
    Dim c As Range

    '��� ���鼭 �� �� ������ üũ
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

    '�߰� ������(4 = rowCount(2) + rowSize(2))
    nSize = 4 + nSizeHeader * rowCount

    '������ �迭���
    ReDim nData(0 To nSize - 1)


    Dim i, j As Integer
    '������
    Dim off As Long

    '������ ���鼭 ����������
    off = 0

    '��� = �÷�����(2) + �ο찹��(2) + ����(�ο�)������(2)
    '�÷�����
    'Call writeShort(nData, off, colCount)
    'off = off + 2
    '�ο찹��
    Call writeShort_little(nData, off, rowCount)
    off = off + 2
    '�ѷο������
    Call writeShort_little(nData, off, nSizeHeader)
    off = off + 2

     For i = 1 To rngData.rows.Count

        For j = 1 To rngData.Columns.Count
            '������ Ÿ�Կ� ���� ����
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

    '������ �����ϴ� �κ�
    Dim saveFileName As String
    saveFileName = ActiveWorkbook.path & "\..\..\..\BuildData\��������Ʈ������_������\Data\potTable.bd"

    Dim FILENUM As Integer
    FILENUM = FreeFile
    Open saveFileName For Binary Access Write As #FILENUM
        Put #FILENUM, , nData
    Close #FILENUM
    MsgBox "BuildData\��������Ʈ������_������\Data ������ potTable.bd ������ ����������ϴ�"
End Sub

Sub ���������()
         '������� �����ִ� ����
    Const HEADER_RANGE_INFO = "E2"
    '�����͹��� �����ִ� ����
    Const DATA_RANGE_INFO = "E3"

    '��ũ��Ʈ ����
    Dim thisSheet As Worksheet

    '�������
    Dim rngHeaderRange As Range
    '�����͹���
    Dim rngDataRange As Range

    '��� �ǵ�����
    Dim rngHeader As Range
    '������ �ǵ�����
    Dim rngData As Range

    '������ �÷����� ���
    Dim colCount As Integer
    Dim rowCount As Integer

    Dim varStr As String

    '���� ��Ʈ�� ��������� ������ ������ ����
    Set thisSheet = ActiveSheet
    Set rngHeaderRange = thisSheet.Range(HEADER_RANGE_INFO)
    Set rngDataRange = thisSheet.Range(DATA_RANGE_INFO)

    '����� ������ �ǵ����� ����
    Set rngHeader = thisSheet.Range(rngHeaderRange.Value)
    Set rngData = thisSheet.Range(rngDataRange.Value)

    '�÷���, ��� ����
    'colCount = rngData.Columns.Count
    rowCount = rngData.rows.Count

    '������ ������
    Dim nSizeHeader As Long
    Dim nSize As Long

    '������ ����
    Dim nData() As Byte

    '�ӽ�b
    Dim c As Range

    '��� ���鼭 �� �� ������ üũ
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

    '�߰� ������(4 = rowCount(2) + rowSize(2))
    nSize = 4 + nSizeHeader * rowCount

    '������ �迭���
    ReDim nData(0 To nSize - 1)


    Dim i, j As Integer
    '������
    Dim off As Long

    '������ ���鼭 ����������
    off = 0

    '��� = �÷�����(2) + �ο찹��(2) + ����(�ο�)������(2)
    '�÷�����
    'Call writeShort(nData, off, colCount)
    'off = off + 2
    '�ο찹��
    Call writeShort_little(nData, off, rowCount)
    off = off + 2
    '�ѷο������
    Call writeShort_little(nData, off, nSizeHeader)
    off = off + 2


    For i = 1 To rngData.rows.Count

        For j = 1 To rngData.Columns.Count
            '������ Ÿ�Կ� ���� ����
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

    '������ �����ϴ� �κ�
    Dim saveFileName As String
    'SaveFileName = ActiveWorkbook.Path & "\" & thisSheet.Name & "(" & Format(Now(), "yymmdd") & ").dat"
    saveFileName = ActiveWorkbook.path & "\..\..\..\BuildData\��������Ʈ������_������\Data\ceraItem.bd"

    Dim FILENUM As Integer
    FILENUM = FreeFile
    Open saveFileName For Binary Access Write As #FILENUM
        Put #FILENUM, , nData
    Close #FILENUM
    MsgBox "BuildData\��������Ʈ������_������\Data ������ ceraItem.bd������ ����������ϴ�"

End Sub






'***�������: vb���� int�� �Ϲ������� 2byte(short)�̸� long�� 4byte(int)�̴�.
' vb���� ����Ʈ �����ڰ� �ȵǹǷ� ���ʽ���Ʈ�� * 2 ^ ����Ʈ�� �������� / 2 ^ ����Ʈ��
' 0xff���� 16���� �ȵǹǷ� �׳� �������� and �Ǵ� or��. 255

'byte[2]->short�� ��ȯ (�޴°� ��Ʋ�� �޴´�. vb�� ����ȹ��)

'public short readShort( byte[] data, int index ) {
'   return (short)(((data[index] & 0xff) << 8) | (data[index + 1] & 0xff));
'}

 Public Function readShort(ByRef data() As Byte, ByVal index As Long) As Integer
    readShort = ((data(index + 1) And 255) * 2 ^ 8) Or (data(index) And 255)
 End Function


'short->byte[2]�� ��ȯ (�����°� ������ ����������Ʈ�� �ڹ���)

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

'byte[4]->int�� ��ȯ

'public int readInt( byte[] data, int index ) {
'   return ((data[index] & 0xff) << 24) | ((data[index + 1] & 0xff) << 16) | ((data[index + 2] & 0xff) << 8) | (data[index + 3] & 0xff);
'}

Public Function readInt(ByRef data() As Byte, ByVal index As Long) As Long
    readInt = ((data(index + 3) And 255) * 2 ^ 24) Or ((data(index + 2) And 255) * 2 ^ 16) Or ((data(index + 1) And 255) * 2 ^ 8) Or (data(index) And 255)
End Function

'int->byte[4]���� ��ȯ

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

'byte[8]->long ��ȯ

'public int readLong( byte[] data, int index ) {
'   return ((data[index] & 0xff) << 56) | ((data[index + 1] & 0xff) << 48) | ((data[index + 2] & 0xff) << 40) | ((data[index + 3] & 0xff) << 32) | ((data[index + 4] & 0xff) << 24) | ((data[index + 5] & 0xff) << 16) | ((data[index + 6] & 0xff) << 8) | (data[index + 7] & 0xff);
'}

Public Function readLong(ByRef data() As Byte, ByVal index As Long) As Currency
    readLong = ((data(index + 7) And 255) * 2 ^ 56) Or ((data(index + 6) And 255) * 2 ^ 48) Or ((data(index + 5) And 255) * 2 ^ 40) Or ((data(index + 4) And 255) * 2 ^ 32) Or ((data(index + 3) And 255) * 2 ^ 24) Or ((data(index + 2) And 255) * 2 ^ 16) Or ((data(index + 1) And 255) * 2 ^ 8) Or (data(index) And 255)
End Function

'long->byte[8]�� ��ȯ

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

'String->byte[]�� ��ȯ

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

'byte[]->String�� ��ȯ

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
