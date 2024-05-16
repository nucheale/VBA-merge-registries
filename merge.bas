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


Sub merge_files_step_1()

    'Загрузить файлы с реестрами объектов и одним реестром всех полигонов. файл с полигоном должен содержать слово полигон в названии

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

    ts_titles = Array("ТС", "ТС ", "Автомобиль", "Госномер ТС", "ГОС НОМЕР", "Гос.номер а/м", "Номеравто", "Гос. номер", "Госномер")

    fileIndex = 1
    For Each file In filesToOpen
        Set objectWb = Application.Workbooks.Open(fileName:=filesToOpen(fileIndex))
        If objectWb.Sheets.Count = 1 Then objectWb.Sheets(objectWb.Sheets.Count).Name = "Ввоз"
        With Sheets("Ввоз")
            If .AutoFilterMode Then .AutoFilter.ShowAllData
            Set findDate = .Range(.Cells(1, 1), .Cells(1, 20)).Find("Дата")
            For Each e In ts_titles
                Set findTS = .Range(.Cells(1, 1), .Cells(1, 20)).Find(e)
                If Not findTS Is Nothing Then Exit For
            Next e



            lastRowObj = .Cells(Rows.Count, 1).End(xlUp).Row

            dates = .Range(.Cells(2, findDate.Column), .Cells(lastRowObj, findDate.Column))
            ts = .Range(.Cells(2, findTS.Column), .Cells(lastRowObj, findTS.Column))
            dates = twoDimArrayToOneDim(dates)
            ts = twoDimArrayToOneDim(ts)
            Dim fileName() As String
            ReDim fileName(1 To UBound(dates))

            Dim steps As Variant 'плечо
            If InStr(LCase(objectWb.Name), "полигон") Then
                Set findStep = .Range(.Cells(1, 1), .Cells(1, 20)).Find(what:="Плечо", LookIn:=xlValues, lookAt:=xlWhole)
                steps = .Range(.Cells(2, findStep.Column), .Cells(lastRowObj, findStep.Column))
                steps = twoDimArrayToOneDim(steps)
            Else
                ReDim steps(1 To UBound(dates))
                For i = LBound(steps) To UBound(steps)
                    steps(i) = 1
                Next i
            End If

            ' https://regex101.com/r/MWdHhN/1

            Set regEx = CreateObject("VBScript.RegExp")
            regEx.Pattern = "\([^)]*\)"
            For i = LBound(dates) To UBound(dates)
                If IsDate(dates(i)) Then
                    If VarType(dates(i)) = vbDate Then
                        dates(i) = CLng(dates(i))
                    ElseIf VarType(dates(i)) = vbString Then
                        dates(i) = CLng(CDate(dates(i)))
                    End If
                End If
                fileName(i) = objectWb.Name
                ts(i) = Replace(ts(i), " ", "")
                ts(i) = Replace(ts(i), ".", "")
                ts(i) = regEx.Replace(ts(i), "")
            Next i
        End With
        
        With newWs
            .Cells(1, 1) = "Дата"
            .Cells(1, 2) = "Госномер"
            .Cells(1, 3) = "Госномер"
            .Cells(1, 4) = "Плечо"
            .Cells(1, 5) = "Файл"
            .Cells(1, 6) = "Дубликаты"
            .Cells(1, 7) = "Перевозчик"
            lastRowNewWs = .Cells(Rows.Count, 1).End(xlUp).Row
            .Cells(lastRowNewWs + 1, 1).Resize(UBound(dates), 1).Value = Application.Transpose(dates)
            .Cells(lastRowNewWs + 1, 2).Resize(UBound(ts), 1).Value = Application.Transpose(ts)
            .Cells(lastRowNewWs + 1, 4).Resize(UBound(steps), 1).Value = Application.Transpose(steps)
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


Sub merge_files_step_2()

    With Application
        .Calculation = xlCalculationManual
        .AskToUpdateLinks = False
        .DisplayAlerts = False
    End With

    Set macroWb = ThisWorkbook

    Set dictWb = Application.Workbooks.Open(fileName:=macroWb.Path & "\Сводная по транспортным средствам.xlsx")

    With dictWb.Sheets(1)
        If InStr(.Cells(2, 1), "Сводная по транспортным средствам") > 0 Then
            For i = 3 To 1 Step -1
                .Rows(i).Delete
            Next i
        End If
        lastRowDict = .Cells(Rows.Count, 3).End(xlUp).Row
        tsDict = .Range(.Cells(2, 3), .Cells(lastRowDict, 3))
        carriersDict = .Range(.Cells(2, 2), .Cells(lastRowDict, 2))
        tsDict = twoDimArrayToOneDim(tsDict)
        carriersDict = twoDimArrayToOneDim(carriersDict)
        For i = LBound(tsDict) To UBound(tsDict)
            tsDict(i) = Replace(tsDict(i), " ", "")
        Next i
        .Cells(2, 3).Resize(UBound(tsDict), 1).Value = Application.Transpose(tsDict)
    End With

    With macroWb.Sheets(macroWb.Sheets.Count)
        lastRowNewWs = .Cells(Rows.Count, 1).End(xlUp).Row
        lastColumnNewWs = .Cells(1, Columns.Count).End(xlToLeft).Column
                            
        dates = .Range(.Cells(2, 1), .Cells(lastRowNewWs, 1))
        ts = .Range(.Cells(2, 3), .Cells(lastRowNewWs, 3))
        dates = twoDimArrayToOneDim(dates)
        ts = twoDimArrayToOneDim(ts)

        Dim forDublicates() As String
        ReDim forDublicates(1 To UBound(dates))
        
        
        For i = LBound(dates) To UBound(dates)
            forDublicates(i) = dates(i) & ts(i)
        Next i
        
        .Range(.Cells(2, 1), .Cells(lastRowNewWs, 1)).ClearContents
        .Cells(2, 1).Resize(UBound(dates), 1).Value = Application.Transpose(dates)
        .Cells(2, 6).Resize(UBound(forDublicates), 1).Value = Application.Transpose(forDublicates)

        .Range(.Cells(1, 1), .Cells(lastRowNewWs, lastColumnNewWs)).RemoveDuplicates Columns:=6, Header:=xlYes

        Erase dates
        Erase ts
        lastRowNewWs = .Cells(Rows.Count, 1).End(xlUp).Row
        lastColumnNewWs = .Cells(1, Columns.Count).End(xlToLeft).Column
        dates = .Range(.Cells(2, 1), .Cells(lastRowNewWs, 1))
        ts = .Range(.Cells(2, 3), .Cells(lastRowNewWs, 3))
        dates = twoDimArrayToOneDim(dates)
        ts = twoDimArrayToOneDim(ts)
        Dim carriers() As Variant
        ReDim carriers(1 To UBound(ts))
        
        For i = LBound(ts) To UBound(ts)
            For e = LBound(tsDict) To UBound(tsDict)
                If ts(i) = tsDict(e) Then
                    carriers(i) = carriersDict(e)
                    Exit For
                End If
            Next e
        Next i

        .Cells(2, 7).Resize(UBound(carriers), 1).Value = Application.Transpose(carriers)

    End With

    dictWb.Close SaveChanges:=True

    With Application
        .Calculation = xlCalculationAutomatic
        .AskToUpdateLinks = True
        .DisplayAlerts = True
    End With

End Sub


'дальше разделение на листы
