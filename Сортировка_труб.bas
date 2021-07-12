Attribute VB_Name = "Module1"
Sub Одна_Кнопка_И_готово()
    Application.ScreenUpdating = False
Выделение_дубликатов_START
Выделение_дубликатов_END
Сортировка_трубопроводов
Создание_листов_по_шаблону
    Sheets("PipsList").Activate
Копирование_строк_дубликатов
Удаляем_пустые_строки
Итоговые_данные

End Sub

Private Sub Выделение_дубликатов_START()
    Sheets("PipsList").Activate
Dim MyRange As Range
Dim MyCell As Range
Set MyRange = Range("C2:E1000")
    For Each MyCell In MyRange
        If WorksheetFunction.CountIf(MyRange, MyCell) > 1 Then
        MyCell.Interior.ColorIndex = 36
        End If
    Next MyCell
End Sub

Private Sub Выделение_дубликатов_END()
    Sheets("PipsList").Activate
Dim MyRange As Range
Dim MyCell As Range
Set MyRange = Sheets("PipsList").Range("F2:H1000")
    For Each MyCell In MyRange
        If WorksheetFunction.CountIf(MyRange, MyCell) > 1 Then
        MyCell.Interior.ColorIndex = 40
        End If
    Next MyCell
End Sub

Private Sub Сортировка_трубопроводов()
    Sheets("PipsList").Activate
    With ActiveSheet.Sort
        .SortFields.Clear
        .SortFields.Add Key:=Range("A2"), Order:=1, CustomOrder:= _
        "Трубопровод 1,Трубопровод 2,Трубопровод 3,Трубопровод 4,Трубопровод 5,Трубопровод 6,Трубопровод 7,Трубопровод 8,Трубопровод 9,Трубопровод 10"
        .SortFields.Add Key:=Range("B2")
        .SortFields.Add Key:=Range("C2")
        .SetRange Range("A2:H1000")
        .Apply
    End With
End Sub

Private Sub Создание_листов_по_шаблону()
    
    Dim MyRange As Range
    Dim MyCell As Range
    Set MyRange = Sheets("PipsList").Range("A2:A1000")
    For Each MyCell In MyRange
        If MyCell = Sheets(Sheets.Count).Name Then
            ElseIf MyCell = 0 Then Exit For
                Else: Sheets("шаблон").Copy After:=Sheets(Sheets.Count)
                ActiveSheet.Name = MyCell
        End If
    Next MyCell
End Sub

Private Sub Копирование_строк_дубликатов()

    Dim Tr1 As String
    Dim r As Integer
    Dim i As Integer
    For i = 1 To 20
    Tr1 = "Трубопровод " & i
        For r = 2 To 100
            If Cells(r, 1) = Tr1 And ((Cells(r, 3).Interior.ColorIndex = 36 And Cells(r, 4).Interior.ColorIndex = 36 And Cells(r, 5).Interior.ColorIndex = 36) Or (Cells(r, 6).Interior.ColorIndex = 40 And Cells(r, 7).Interior.ColorIndex = 40 And Cells(r, 8).Interior.ColorIndex = 40)) Then
            Rows(r).Copy Sheets(Tr1).Rows(r + 4)
            End If
        Next r
    Next i
End Sub

Private Sub Удаляем_пустые_строки()
    
Dim ws As Worksheet
Dim SheetExists As Boolean
Dim i As Integer
Dim r As Integer
Dim Tr1 As String
For i = 1 To 20
Tr1 = "Трубопровод " & i
        
        For Each ws In ActiveWorkbook.Worksheets
            If ws.Name = Tr1 Then
            SheetExists = True
                
                Sheets(Tr1).Activate
                lastrow = Sheets(Tr1).UsedRange.Row - 1 + Sheets(Tr1).UsedRange.Rows.Count
                For r = lastrow To 1 Step -1
                If Application.CountA(Rows(r)) = 0 Then Rows(r).Delete
                Next r
            
        Else: SheetExists = False
        End If
        Next ws
Next i
End Sub

Private Sub Итоговые_данные()

Dim i As Integer
Dim Tr1 As String
For i = 1 To 20
Tr1 = "Трубопровод " & i
Sheets(Tr1).Activate
    On Error Resume Next
    Sheets(Tr1).Range("B1") = Tr1
        ' считаем длины трубопроводов в м.
        Sheets(Tr1).Range("I6").Formula = "=SQRT((F6-C6)^2+(G6-D6)^2+(H6-E6)^2)/1000"
            Range("A1").Select
            Selection.End(xlDown).Select
            r = Selection.Row
        Range("I6").Select
        Selection.AutoFill Destination:=Range("I6", "I" & r)
        Sheets(Tr1).Range("B2") = "=SUM(I:I)"
        
        ' считаем среднй диаметр трубопроводо
        Sheets(Tr1).Range("B3").FormulaR1C1 = "=AVERAGE(R6C:R100C)"

Next i
End Sub

