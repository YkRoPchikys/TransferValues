Attribute VB_Name = "TransferValues"
Sub TransferValues()
    Dim CurrentNameBook As String
    Dim folderPath As String
    Dim fileDialog As fileDialog
    Dim fileName As String
    Dim fileList As Collection
    Dim fileItem As Variant
    Dim selectedFile As String
    Dim wsList As String
    Dim i As Integer
    Dim targetWorkbook As Workbook
    Dim targetWorksheet As Worksheet
    Dim cell As Range
    Dim searchString As String
    Dim valueToTransfer As Variant
    Dim wsName As String
    Dim searchTerms As Variant
    Dim outputCells As Variant
    
    ' ��������� �������� ������� �����
    CurrentNameBook = ThisWorkbook.Name
    
    ' ������������� ���� � �����
    folderPath = "C:\���\����\�\�����\" ' �������� �� ���� ����
    
    ' �������������� ��������� ��� �������� ���� ������
    Set fileList = New Collection
    
    ' �������� ������ ���� ������ Excel � �����
    fileName = Dir(folderPath & "*.xlsx")
    Do While fileName <> ""
        fileList.Add fileName
        fileName = Dir
    Loop
    
    ' ���� ������ ���, ������� ��������� � ������� �� ���������
    If fileList.Count = 0 Then
        MsgBox "� ��������� ����� ��� ������ Excel."
        Exit Sub
    End If
    
    ' ������� ��������� � �������� ������ ��� ������
    wsList = "�������� ����� ����� ��� ��������:" & vbCrLf
    For i = 1 To fileList.Count
        wsList = wsList & i & ". " & fileList(i) & vbCrLf
    Next i
    
    ' ����������� � ������������ ����� �����
    selectedFile = InputBox(wsList, "����� �����")
    
    ' ���������, ��� ������� ���������� ��������
    If IsNumeric(selectedFile) Then
        i = CInt(selectedFile)
        If i >= 1 And i <= fileList.Count Then
            Workbooks.Open folderPath & fileList(i)
            Set targetWorkbook = Workbooks(fileList(i))
        Else
            MsgBox "������������ �����. ���������� �����."
            Exit Sub
        End If
    Else
        MsgBox "������������ ����. ���������� �����."
        Exit Sub
    End If
    
    ' ���������� ������ ������ ��� ������
    wsList = "�������� ����� �����:" & vbCrLf
    For i = 1 To targetWorkbook.Sheets.Count
        wsList = wsList & i & ". " & targetWorkbook.Sheets(i).Name & vbCrLf
    Next i
    
    ' ����������� � ������������ ����� �����
    wsName = InputBox(wsList, "����� �����")
    
    ' ���������, ��� ������� ���������� ��������
    If IsNumeric(wsName) Then
        i = CInt(wsName)
        If i >= 1 And i <= targetWorkbook.Sheets.Count Then
            Set targetWorksheet = targetWorkbook.Sheets(i)
        Else
            MsgBox "������������ �����. ���������� �����."
            Exit Sub
        End If
    Else
        MsgBox "������������ ����. ���������� �����."
        Exit Sub
    End If
    
    ' ���������� ������� ��� ������ � ��������������� ����� ��� �������
    searchTerms = Array("One", "Two", "Three", "Four", "Five", "Six") '��� �� ���� � ������� A
    outputCells = Array("R10", "R15", "R17", "R20", "R35", "R36") '������ � ������� �����, � ������� ����� ��������� ��������� ������
    
    ' �������� �� ���� ��������� ��������
    For i = LBound(searchTerms) To UBound(searchTerms)
        searchString = searchTerms(i)
        valueToTransfer = ""
        
        ' ����� ������ � ������� A
        For Each cell In targetWorksheet.Range("A:A")
            If cell.Value = searchString Then
                ' ������� ����������, �������� �������� �� ������� AI ���� ������
                valueToTransfer = cell.Offset(0, 34).Value ' 34 ������� ������ �� A (�� ���� AI)
                Exit For
            End If
        Next cell
        
        ' ���������, ��� �������� �������
        If valueToTransfer = "" Then
            MsgBox "����� '" & searchString & "' �� ������ � ������� A �� ����� " & targetWorksheet.Name & "."
        Else
            ' ��������� �������� � �������� �����
            ThisWorkbook.Sheets(1).Range(outputCells(i)).Value = valueToTransfer
        End If
    Next i
    
    ' �������� � ���������� ��������
    MsgBox "�������� ������� ����������."
    
    ' ��������� ������� ���� ��� ����������
    targetWorkbook.Close SaveChanges:=False
End Sub
