Attribute VB_Name = "Module1"
Sub ����_������_�_������()
    Application.ScreenUpdating = False
���������_����������_START
���������_����������_END
����������_�������������
��������_������_��_�������
    Sheets("PipsList").Activate
�����������_�����_����������
�������_������_������
��������_������

End Sub

Private Sub ���������_����������_START()
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

Private Sub ���������_����������_END()
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

Private Sub ����������_�������������()
    Sheets("PipsList").Activate
    With ActiveSheet.Sort
        .SortFields.Clear
        .SortFields.Add Key:=Range("A2"), Order:=1, CustomOrder:= _
        "����������� 1,����������� 2,����������� 3,����������� 4,����������� 5,����������� 6,����������� 7,����������� 8,����������� 9,����������� 10"
        .SortFields.Add Key:=Range("B2")
        .SortFields.Add Key:=Range("C2")
        .SetRange Range("A2:H1000")
        .Apply
    End With
End Sub

Private Sub ��������_������_��_�������()
    
    Dim MyRange As Range
    Dim MyCell As Range
    Set MyRange = Sheets("PipsList").Range("A2:A1000")
    For Each MyCell In MyRange
        If MyCell = Sheets(Sheets.Count).Name Then
            ElseIf MyCell = 0 Then Exit For
                Else: Sheets("������").Copy After:=Sheets(Sheets.Count)
                ActiveSheet.Name = MyCell
        End If
    Next MyCell
End Sub

Private Sub �����������_�����_����������()

    Dim Tr1 As String
    Dim r As Integer
    Dim i As Integer
    For i = 1 To 20
    Tr1 = "����������� " & i
        For r = 2 To 100
            If Cells(r, 1) = Tr1 And ((Cells(r, 3).Interior.ColorIndex = 36 And Cells(r, 4).Interior.ColorIndex = 36 And Cells(r, 5).Interior.ColorIndex = 36) Or (Cells(r, 6).Interior.ColorIndex = 40 And Cells(r, 7).Interior.ColorIndex = 40 And Cells(r, 8).Interior.ColorIndex = 40)) Then
            Rows(r).Copy Sheets(Tr1).Rows(r + 4)
            End If
        Next r
    Next i
End Sub

Private Sub �������_������_������()
    
Dim ws As Worksheet
Dim SheetExists As Boolean
Dim i As Integer
Dim r As Integer
Dim Tr1 As String
For i = 1 To 20
Tr1 = "����������� " & i
        
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

Private Sub ��������_������()

Dim i As Integer
Dim Tr1 As String
For i = 1 To 20
Tr1 = "����������� " & i
Sheets(Tr1).Activate
    On Error Resume Next
    Sheets(Tr1).Range("B1") = Tr1
        ' ������� ����� ������������� � �.
        Sheets(Tr1).Range("I6").Formula = "=SQRT((F6-C6)^2+(G6-D6)^2+(H6-E6)^2)/1000"
            Range("A1").Select
            Selection.End(xlDown).Select
            r = Selection.Row
        Range("I6").Select
        Selection.AutoFill Destination:=Range("I6", "I" & r)
        Sheets(Tr1).Range("B2") = "=SUM(I:I)"
        
        ' ������� ������ ������� ������������
        Sheets(Tr1).Range("B3").FormulaR1C1 = "=AVERAGE(R6C:R100C)"

Next i
End Sub

