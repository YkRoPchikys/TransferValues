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
    
    ' Сохраняем название текущей книги
    CurrentNameBook = ThisWorkbook.Name
    
    ' Устанавливаем путь к папке
    folderPath = "C:\Ваш\Путь\К\Папке\" ' ИЗМЕНИТЬ НА СВОЙ ПУТЬ
    
    ' Инициализируем коллекцию для хранения имен файлов
    Set fileList = New Collection
    
    ' Получаем список всех файлов Excel в папке
    fileName = Dir(folderPath & "*.xlsx")
    Do While fileName <> ""
        fileList.Add fileName
        fileName = Dir
    Loop
    
    ' Если файлов нет, выводим сообщение и выходим из субрутины
    If fileList.Count = 0 Then
        MsgBox "В указанной папке нет файлов Excel."
        Exit Sub
    End If
    
    ' Создаем сообщение с перечнем файлов для выбора
    wsList = "Выберите номер файла для открытия:" & vbCrLf
    For i = 1 To fileList.Count
        wsList = wsList & i & ". " & fileList(i) & vbCrLf
    Next i
    
    ' Запрашиваем у пользователя выбор файла
    selectedFile = InputBox(wsList, "Выбор файла")
    
    ' Проверяем, что введено допустимое значение
    If IsNumeric(selectedFile) Then
        i = CInt(selectedFile)
        If i >= 1 And i <= fileList.Count Then
            Workbooks.Open folderPath & fileList(i)
            Set targetWorkbook = Workbooks(fileList(i))
        Else
            MsgBox "Недопустимый выбор. Попробуйте снова."
            Exit Sub
        End If
    Else
        MsgBox "Недопустимый ввод. Попробуйте снова."
        Exit Sub
    End If
    
    ' Составляем список листов для выбора
    wsList = "Выберите номер листа:" & vbCrLf
    For i = 1 To targetWorkbook.Sheets.Count
        wsList = wsList & i & ". " & targetWorkbook.Sheets(i).Name & vbCrLf
    Next i
    
    ' Запрашиваем у пользователя выбор листа
    wsName = InputBox(wsList, "Выбор листа")
    
    ' Проверяем, что введено допустимое значение
    If IsNumeric(wsName) Then
        i = CInt(wsName)
        If i >= 1 And i <= targetWorkbook.Sheets.Count Then
            Set targetWorksheet = targetWorkbook.Sheets(i)
        Else
            MsgBox "Недопустимый выбор. Попробуйте снова."
            Exit Sub
        End If
    Else
        MsgBox "Недопустимый ввод. Попробуйте снова."
        Exit Sub
    End If
    
    ' Определяем массивы для поиска и соответствующих ячеек для вставки
    searchTerms = Array("One", "Two", "Three", "Four", "Five", "Six") 'Что мы ищем в столбце A
    outputCells = Array("R10", "R15", "R17", "R20", "R35", "R36") 'Ячейки в текущем файле, в которые будем вставлять найденные данные
    
    ' Проходим по всем поисковым терминам
    For i = LBound(searchTerms) To UBound(searchTerms)
        searchString = searchTerms(i)
        valueToTransfer = ""
        
        ' Поиск текста в столбце A
        For Each cell In targetWorksheet.Range("A:A")
            If cell.Value = searchString Then
                ' Найдено совпадение, получаем значение из столбца AI этой строки
                valueToTransfer = cell.Offset(0, 34).Value ' 34 столбца вправо от A (то есть AI)
                Exit For
            End If
        Next cell
        
        ' Проверяем, что значение найдено
        If valueToTransfer = "" Then
            MsgBox "Текст '" & searchString & "' не найден в столбце A на листе " & targetWorksheet.Name & "."
        Else
            ' Вставляем значение в исходную книгу
            ThisWorkbook.Sheets(1).Range(outputCells(i)).Value = valueToTransfer
        End If
    Next i
    
    ' Сообщаем о завершении процесса
    MsgBox "Значения успешно перенесены."
    
    ' Закрываем целевой файл без сохранения
    targetWorkbook.Close SaveChanges:=False
End Sub
