Function getMaxTwoDArrayValue(arr) As Double
    maxValue = arr(LBound(arr), 1)
    For i = LBound(arr, 1) To UBound(arr, 1)
        If arr(i, 1) > maxValue Then maxValue = arr(i, 1)
    Next i
    getMaxTwoDArrayValue = maxValue
End Function

Function getMinTwoDArrayValue(arr) As Double
    minValue = arr(LBound(arr), 1)
    For i = LBound(arr, 1) To UBound(arr, 1)
        If arr(i, 1) < minValue Then minValue = arr(i, 1)
    Next i
    getMinTwoDArrayValue = minValue
End Function

Function twoDimArrayToOneDim(oldArr)
    Dim newArr As Variant
    ReDim newArr(1 To UBound(oldArr, 1) * UBound(oldArr, 2))
    For i = LBound(oldArr, 1) To UBound(oldArr, 1)
        newArr(i) = oldArr(i, 1)
    Next i
    twoDimArrayToOneDim = newArr
End Function


Sub merge_files()

    Set macroWb = ThisWorkbook
    Set newWs = macroWb.Sheets.Add(After:=macroWb.Sheets(macroWb.Sheets.Count))
    currTime = Array(Hour(Now), Minute(Now), Second(Now))
    newWs.Name = "Вывоз " & Date & "_" & currTime(0) & "_" & currTime(1) & "_" & currTime(2)
    
    filesToOpen = Application.GetOpenFilename(FileFilter:="All files (*.*), *.*", MultiSelect:=True, Title:="Выберите файлы")
    If TypeName(filesToOpen) = "Boolean" Then Exit Sub
    
    With Application
        .Calculation = xlCalculationManual
        .AskToUpdateLinks = False
        .DisplayAlerts = False
    End With

    ts_titles = Array("ТС", "ТС ", "Автомобиль", "Госномер ТС", "ГОС НОМЕР", "Гос.номер а/м", "Номеравто")

    fileIndex = 1
    For Each file In filesToOpen
        Set objectWb = Application.Workbooks.Open(fileName:=filesToOpen(fileIndex))
        With Sheets("Ввоз")
            If .AutoFilterMode Then .AutoFilter.ShowAllData
            Set findDate = .Range(.Cells(1, 1), .Cells(1, 20)).Find("Дата")
            For Each e In ts_titles
                Set findTS = .Range(.Cells(1, 1), .Cells(1, 20)).Find(e)
                If Not findTS Is Nothing Then Exit For
            Next e

            ' debug.print objectWb.name, findDate, findTS

            lastRowObj = .Cells(Rows.Count, 1).End(xlUp).Row

            dates = .Range(.Cells(2, findDate.Column), .Cells(lastRowObj, findDate.Column))
            ts = .Range(.Cells(2, findTS.Column), .Cells(lastRowObj, findTS.Column))
            dates = twoDimArrayToOneDim(dates)
            ts = twoDimArrayToOneDim(ts)
            Dim fileName() As String
            ReDim fileName(1 To UBound(dates))
            For i = LBound(fileName) To UBound(fileName)
                fileName(i) = objectWb.Name
            Next i
            ' Debug.Print objectWb.Name, " // ", UBound(dates), UBound(ts)
        End With
        
        With newWs
            .Cells(1, 1) = "Дата"
            .Cells(1, 2) = "Госномер"
            .Cells(1, 3) = "Госномер"
            .Cells(1, 4) = "Плечо"
            .Cells(1, 5) = "Файл"
            lastRowNewWs = .Cells(Rows.Count, 1).End(xlUp).Row
            .Cells(lastRowNewWs + 1, 1).Resize(UBound(dates), 1).Value = Application.Transpose(dates)
            .Cells(lastRowNewWs + 1, 2).Resize(UBound(ts), 1).Value = Application.Transpose(ts)
            .Cells(lastRowNewWs + 1, 4).Resize(UBound(fileName), 1).Value = 1
            .Cells(lastRowNewWs + 1, 5).Resize(UBound(fileName), 1).Value = Application.Transpose(fileName)
        End With

        Erase dates
        Erase ts
        Erase fileName
        objectWb.Close SaveChanges:=False
        fileIndex = fileIndex + 1
    Next


    With Application
        .Calculation = xlCalculationAutomatic
        .AskToUpdateLinks = True
        .DisplayAlerts = True
    End With

End Sub
