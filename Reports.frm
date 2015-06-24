VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Reports 
   Caption         =   "Склад"
   ClientHeight    =   8895
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   15150
   OleObjectBlob   =   "Reports.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Reports"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim JobNames() As String

Sub ScanJobs()
On Error GoTo ExceptionControl:
Sheets("Каталог").Select
With Reports
    .JobsTree.Nodes.Clear
    TotalCats = Cells(4, 19).Value
    For i = InfoOffset To CInt(InfoOffset - 1 + TotalCats)
        .JobsTree.Nodes.Add(, , , Cells(i, 19).Value).Sorted = True
        .JobsTree.Nodes(i - InfoOffset + 1).Tag = "Cat"
    Next

    TotalJobs = Cells(4, 2).Value
    ReDim JobNames(InfoOffset To CInt(InfoOffset - 1 + TotalJobs))
    For i = InfoOffset To CInt(InfoOffset - 1 + TotalJobs)
  
            If (Cells(i, 8) = 1) Then
                .JobsTree.Nodes.Add(CInt(Cells(i, 1).Value - InfoOffset + 1), 4, CStr(Cells(i, 3)) & "z", Cells(i, 2).Value).Sorted = True
                JobNames(i) = Cells(i, 2).Value
            End If
    Next

    p = 1
    Do While p < .JobsTree.Nodes.Count
        If .JobsTree.Nodes(p).Children = 0 And .JobsTree.Nodes(p).Tag = "Cat" Then
            .JobsTree.Nodes.Remove (p)
            TotalCats = TotalCats - 1
            p = p - 1
        End If
        p = p + 1
    Loop
    .JobsTree.Tag = TotalCats
End With

Exit Sub
ExceptionControl:
Exception.Error_Box.Value = "Reports/ScanJobs()"
Exception.Show
End Sub

Private Sub Production_Button_Click()
Dim HiddenCount, Index, Children, CountLines As Integer
Dim ProductionPerIDPerDate() As Single
Dim Alternates() As Integer
On Error GoTo ExceptionControl:
Windows(WorkersBase).Activate
ScanJobs
Application.Calculation = xlCalculationManual

TotalIDs = Cells(4, 2).Value
Dimention = InfoOffset - 1 + CInt(TotalIDs)
ReDim ProductionPerIDPerDate(InfoOffset To Dimention, CInt(CDay_Box.Value) To 31, 1 To 1)
ReDim Alternates(InfoOffset To Dimention, 1 To 1)

MarksIgnored = False
If Mark_Chooser.Value = "Выделенные" Then Mark = 1
If Mark_Chooser.Value = "Без выделения" Then Mark = ""
If Mark_Chooser.Value = "Все" Then MarksIgnored = True
    
For i = FirstWorkersSheet To ActiveWorkbook.Sheets.Count
    Sheets(i).Select
    For j = InfoOffset + (CInt(CDay_Box.Value) - 1) * Lines To Lines * 31 + InfoOffset - 1
        If (CInt(Cells(j, 3).Value) > 5 And Cells(j, 4).Value <> "") And (Cells(j, 15).Value = Mark Or MarksIgnored) Then
            DayNum = 1 + (j - InfoOffset) \ Lines
            ID = CInt(Cells(j, 3).Value)
            Quantity = Cells(j, 4).Value
             AltDiam = Cells(j, 14).Value
            If AltDiam = "" Then
                ProductionPerIDPerDate(ID, DayNum, 1) = ProductionPerIDPerDate(ID, DayNum, 1) + Quantity
            Else
                BaseDiam = Right(JobNames(ID), Len(AltDiam) + 1)
                If (BaseDiam = "x" & AltDiam) Or (BaseDiam = "х" & AltDiam) Then
                    Cells(j, 14).ClearContents
                    ProductionPerIDPerDate(ID, DayNum, 1) = ProductionPerIDPerDate(ID, DayNum, 1) + Quantity
                Else
                    For k = 1 To Alternates(ID, 1) + 1
                        If Alternates(ID, k) = AltDiam Then
                            ProductionPerIDPerDate(ID, DayNum, k) = ProductionPerIDPerDate(ID, DayNum, k) + Quantity
                            Exit For
                        Else
                            If k = Alternates(ID, 1) + 1 Then
                                If k = UBound(Alternates, 2) Then
                                    ReDim Preserve ProductionPerIDPerDate(InfoOffset To Dimention, CInt(CDay_Box.Value) To 31, 1 To k + 1)
                                    ReDim Preserve Alternates(InfoOffset To Dimention, 1 To k + 1)
                                End If
                                Alternates(ID, k + 1) = CInt(AltDiam)
                                Alternates(ID, 1) = Alternates(ID, 1) + 1
                                ProductionPerIDPerDate(ID, DayNum, k + 1) = ProductionPerIDPerDate(ID, DayNum, k + 1) + Quantity
                                Exit For
                            End If
                        End If
                    Next k
                End If
            End If
        End If
    Next j
Next i

Sheets("Производство").Select
Cells(2, 2) = DateTime.Date
Cells(3, 2) = DateTime.Time
Range(Rows(7), Rows(2000)).Clear
Range(Columns(3), Columns(33)).EntireColumn.Hidden = True
Cells(1, 2).Value = "Выпуск продукции за " & MName(CMonth)

MarkLine = False
JobsTree.SingleSel = True
JobsTree.HideSelection = False
JobsTree.Top = -500
JobsTree.Nodes(1).FirstSibling.Selected = True
CountLines = 1
For i = 1 To JobsTree.Nodes.Count
    ID = CutZ(JobsTree.SelectedItem.Key)
    If ID > 0 Then
       For k = 1 To Alternates(ID, 1) + 1
            EmptyLine = True
            Index = InfoOffset + CountLines
                For j = CInt(CDay_Box.Value) To 31
                    If ProductionPerIDPerDate(ID, j, k) <> 0 Then
                        If k = 1 Then
                            Cells(Index, 2).Value = JobsTree.SelectedItem.Text
                        Else
                            Cells(Index, 2).Value = ReplaceToAlternate(JobsTree.SelectedItem.Text, Alternates(ID, k))
                        End If
                        Cells(Index, j + 2).Value = Round(ProductionPerIDPerDate(ID, j, k))
                        Cells(Index, j + 2).EntireColumn.Hidden = False
                        EmptyLine = False
                    End If
                Next j
            If Not EmptyLine Then
                Cells(Index, 34).FormulaR1C1 = "=SUM(RC[-31]:RC[-1])"
                Range(Cells(Index, 2), Cells(Index, 34)).Select
                Selection.RowHeight = 13
                FillAndBorders "#,##0", MarkLine
                Cells(Index, 34).Font.Bold = True
                MarkLine = Not MarkLine
                CountLines = CountLines + 1
            End If
        Next k
    Else
        If EmptyCatIndex = CountLines Then CountLines = CountLines - 1: MarkLine = Not MarkLine
        Index = InfoOffset + CountLines
        Cells(Index, 2).Value = JobsTree.SelectedItem.Text
        Range(Cells(Index, 2), Cells(Index, 33)).Select
        With Selection
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlBottom
            .WrapText = False
            .Orientation = 0
            .RowHeight = 20
            .AddIndent = False
            .IndentLevel = 0
            .ShrinkToFit = False
            .ReadingOrder = xlContext
            .MergeCells = True
        End With
        Range(Cells(Index, 2), Cells(Index, 34)).Select
        FillAndBorders "#,##0", MarkLine
        MarkLine = Not MarkLine
        CountLines = CountLines + 1
        EmptyCatIndex = CountLines
    End If
    
    If JobsTree.SelectedItem.Tag = "Cat" Then
        Children = JobsTree.SelectedItem.Children
        JobsTree.SelectedItem.Child.FirstSibling.Selected = True
    Else
        If Children > 1 Then
            JobsTree.SelectedItem.Next.Selected = True
            Children = Children - 1
        Else
            If i < JobsTree.Nodes.Count Then JobsTree.SelectedItem.Parent.Next.Selected = True
        End If
    End If
Next i
Erase ProductionPerIDPerDate
Erase Alternates
Erase JobNames

JobsTree.Top = 30
Columns(3).Select
ActiveWindow.FreezePanes = True
Cells(1, 1).Select
Application.Calculation = xlCalculationAutomatic

If NoPrint_Chk.Value = True Then
    Sheets("Производство").PrintOut
Else
    Application.DisplayStatusBar = True
    Reports.Hide
    Main.Hide
    ReportExit = True
End If

Exit Sub
ExceptionControl:
Exception.Error_Box.Value = "Reports/Production_Button_Click()"
Exception.Show
End Sub

Private Sub FeeReport_Button_Click()
On Error GoTo ExceptionControl:
Windows(WorkersBase).Activate
Sheets("Отчёт").Select

Selection.Font.Bold = False
Range(Rows(7), Rows(2000)).Clear

Cells(1, 3).Value = "Отчёт по зарплате за  " & MName(CMonth)
Cells(6, 3).Value = "Остаток за " & MName(LMonth)
Cells(6, 5).Value = "Выдано за " & MName(CMonth)

MarkLine = True
HiddenCount = 0
Sheets("Сотрудники").Select
Range("A3:G100").Select
Selection.Sort Key1:=Range("B3"), Order1:=xlAscending, Header:=xlGuess, _
    OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, DataOption1:=xlSortNormal
        
WeHaveWorkers = Cells(1, 2).Value

For i = 3 To WeHaveWorkers + 2
    Sheets("Сотрудники").Select
    If Cells(i, 4).Value = 1 Then
        HiddenCount = HiddenCount + 1
    Else
        Sheets(Cells(i, 3).Value).Select
        If Cells(1, 1).Value <> "" Then LastDay = "(по " & Cells(1, 1).Value & "-e число)" Else LastDay = "#нет данных#"
        Leftt = Cells(2, 10).Value
        Income = Cells(3, 10).Value
        Outcome = Cells(3, 11).Value
        Balance = Cells(1, 10).Value
        Namess = Cells(1, 2).Value & " " & Cells(2, 2).Value
        Sheets("Отчёт").Select
        Cells(3, 4) = DateTime.Date
        Cells(3, 5) = DateTime.Time
        Index = 4 + i - HiddenCount
        Cells(Index, 2) = Namess
        Cells(Index, 3) = Leftt
        Cells(Index, 4) = Income
        Cells(Index, 5) = Outcome
        Cells(Index, 6) = Balance
        Cells(Index, 7) = LastDay
        Range(Cells(Index, 6), Cells(Index, 6)).Select
        If Balance < 0 Then Selection.Font.Bold = True
        Range(Cells(Index, 2), Cells(Index, 6)).Select
        FillAndBorders "#,##0.00", MarkLine
        MarkLine = Not MarkLine
    End If
Next

If NoPrint_Chk.Value = True Then
    Sheets("Отчёт").PrintOut
Else
    Application.DisplayStatusBar = True
    Reports.Hide
    Main.Hide
    ReportExit = True
End If

Exit Sub
ExceptionControl:
Exception.Error_Box.Value = "Reports/FeeReport_Button_Click()"
Exception.Show
End Sub

Private Sub AvReport_Button_Click()
Dim Day(1 To 31) As Integer, Av(1 To 31) As String
On Error GoTo ExceptionControl:
Windows(WorkersBase).Activate
Sheets("АвансовыйОтчёт").Select
Cells(2, 2) = DateTime.Date
Cells(3, 2) = DateTime.Time

Range(Rows(7), Rows(2000)).Clear
Range(Columns(3), Columns(33)).EntireColumn.Hidden = True

Cells(1, 2).Value = "Авансовый отчёт за " & MName(CMonth)

MarkLine = True
HiddenCount = 0
Sheets("Сотрудники").Select
Range("A3:G100").Select
Selection.Sort Key1:=Range("B3"), Order1:=xlAscending, Header:=xlGuess, _
    OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, DataOption1:=xlSortNormal

WeHaveWorkers = Cells(1, 2).Value
Application.Calculation = xlCalculationManual
For i = 3 To WeHaveWorkers + 2
    Sheets("Сотрудники").Select
    If Cells(i, 4).Value = 1 Then
        HiddenCount = HiddenCount + 1
    Else
        Sheets(Cells(i, 3).Value).Select
        Namess = Cells(1, 2).Value & " " & Cells(2, 2).Value
        p = 0
        For j = InfoOffset To InfoOffset + 31 * Lines - Lines Step Lines
            If Cells(j, 11).Value <> 0 Then
                p = p + 1
                Day(p) = Cells(j, 1).Value
                Av(p) = CStr(Cells(j, 11).Value)
            End If
        Next j
        Sheets("АвансовыйОтчёт").Select
        Index = 4 + i - HiddenCount
        Cells(Index, 2).Value = Namess
        Cells(Index, 34).FormulaR1C1 = "=SUM(RC[-31]:RC[-1])"
        For j = 1 To p
            Cells(Index, Day(j) + 2).Value = Av(j)
            Cells(Index, Day(j) + 2).EntireColumn.Hidden = False
        Next j
        Range(Cells(Index, 2), Cells(Index, 34)).Select
        FillAndBorders "#,##0.00", MarkLine
        MarkLine = Not MarkLine
    End If
Next i
Application.Calculation = xlCalculationAutomatic

Erase Day
Erase Av

If NoPrint_Chk.Value = True Then
    Sheets("АвансовыйОтчёт").PrintOut
Else
     Application.DisplayStatusBar = True
     Reports.Hide
     Main.Hide
     ReportExit = True
End If

Exit Sub
ExceptionControl:
Exception.Error_Box.Value = "Reports/AvReport_Button_Click()"
Exception.Show
End Sub

Private Sub FillAndBorders(ByVal Format As String, ByVal MarkLine As Boolean)
Selection.NumberFormat = Format
With Selection.Borders(xlEdgeLeft)
    .LineStyle = xlDot
    .Weight = xlThin
    .ColorIndex = xlAutomatic
End With
With Selection.Borders(xlEdgeTop)
    .LineStyle = xlDot
    .Weight = xlThin
    .ColorIndex = xlAutomatic
End With
With Selection.Borders(xlEdgeBottom)
    .LineStyle = xlDot
    .Weight = xlThin
    .ColorIndex = xlAutomatic
End With
With Selection.Borders(xlEdgeRight)
    .LineStyle = xlDot
    .Weight = xlThin
    .ColorIndex = xlAutomatic
End With
With Selection.Borders(xlInsideVertical)
    .LineStyle = xlDot
    .Weight = xlThin
    .ColorIndex = xlAutomatic
End With
If MarkLine Then
    With Selection.Interior
        .ColorIndex = 15
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
    End With
End If
End Sub
