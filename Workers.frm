VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Workers 
   ClientHeight    =   10890
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   17085
   OleObjectBlob   =   "Workers.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Workers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim inRead, CommentTrigger, DayTrigger As Boolean
Dim ChosenMate, ConfirmObject As Integer

Sub ObjectsRecall()
WorkersTreeHolder.Visible = False
JobsTree.Visible = True
DayList.Visible = True
DateAndWorker_Frame.Visible = True
ConfirmObject = 0
ConfirmChoice_Button.Top = 420
ConfirmChoice_Button.Left = 120
End Sub

Sub ScanWorkers()
Dim i, p, TotalCats As Integer
On Error GoTo ExceptionControl:
With Workers
    .WorkersTreeHolder.Top = -500
    .WorkersTree.Visible = True
    .WorkersTreeHolder.Visible = True
    .WorkersTree.Nodes.Clear

    Sheets("Каталог").Select
    TotalCats = Cells(4, 23).Value
    For i = InfoOffset To CInt(InfoOffset - 1 + TotalCats)
        .WorkersTree.Nodes.Add(, , CStr(Cells(i, 24)) & "z", Cells(i, 23).Value).Sorted = True
        .WorkersTree.Nodes(i - InfoOffset + 1).Tag = "Cat"
    Next
  
    Sheets("Сотрудники").Select
    WeHaveWorkers = Cells(1, 2).Value
    For i = 3 To WeHaveWorkers + 2
        For p = 1 To TotalCats
            If (Cells(i, 4).Value <> 1) And (.WorkersTree.Nodes(p).Key = CStr(Cells(i, 6).Value) & "z") Then
                If (SelectUpdatesOnly And Cells(i, 1).Value = 1) Or Not SelectUpdatesOnly Then
                    .WorkersTree.Nodes.Add(p, 4, Cells(i, 3), Cells(i, 2) & " " & Cells(i, 5)).Sorted = True
                    If Not AdminMode Then .WorkersTree.Nodes(.WorkersTree.Nodes.Count).Tag = Cells(i, 7)
                    p = TotalCats
                End If
            End If
        Next p
    Next i
 
    p = 1
    Do While p < .WorkersTree.Nodes.Count
        If .WorkersTree.Nodes(p).Children = 0 And .WorkersTree.Nodes(p).Tag = "Cat" Then
            .WorkersTree.Nodes.Remove (p)
            p = p - 1
            TotalCats = TotalCats - 1
        End If
        p = p + 1
    Loop
    .WorkersTree.Tag = TotalCats
    .WorkersTreeHolder.Visible = False
End With

Exit Sub
ExceptionControl:
Exception.Error_Box.Value = "Workers/ScanWorkers()"
Exception.Show
End Sub

Sub ScanJobs()
Dim i, p, TotalCats As Integer
On Error GoTo ExceptionControl:
Sheets("Каталог").Select
With Workers
    .JobsTree.Visible = True

    .JobsTree.Nodes.Clear
    .BonusRate_Box.Value = Cells(4, 6).Value
    TotalCats = Cells(4, 19).Value
    For i = InfoOffset To CInt(InfoOffset - 1 + TotalCats)
        .JobsTree.Nodes.Add(, , , Cells(i, 19).Value).Sorted = True
        .JobsTree.Nodes(i - InfoOffset + 1).Tag = "Cat"
    Next

    TotalJobs = Cells(4, 2).Value
    ShowRates = Cells(4, 5).Value
    For i = InfoOffset To CInt(InfoOffset - 1 + TotalJobs)
        AddRate = ""
        If ShowRates = 1 Then
            If Cells(i, 5) = 0 Then AddRate = "  (" & CStr(Cells(i, 6)) & ")" Else AddRate = "  (" & CStr(Cells(i, 5)) & ")"
        End If
  
        If Not AdminMode Then
            If (Cells(i, 7) = 0) And (Cells(i, 9) = 0) Then _
                .JobsTree.Nodes.Add(CInt(Cells(i, 1).Value - InfoOffset + 1), 4, CStr(Cells(i, 3)) & "z", Cells(i, 2).Value & AddRate).Sorted = True
        Else
            If (Cells(i, 7) = 0) Then _
                .JobsTree.Nodes.Add(CInt(Cells(i, 1).Value - InfoOffset + 1), 4, CStr(Cells(i, 3)) & "z", Cells(i, 2).Value & AddRate).Sorted = True
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
Exception.Error_Box.Value = "Workers/ScanJobs()"
Exception.Show
End Sub

Sub FillControlList()
On Error GoTo ExceptionControl:
Sheets(NameChooser.Value).Activate
ControlList.ListItems.Clear
For i = InfoOffset To InfoOffset + 31 * Lines - Lines Step Lines
    If (Cells(i, 10).Value <> 0) Or (Cells(i, 11).Value <> 0) Or (Cells(i, 13).Value <> "") Then
        Dat = Cells(i, 1).Value
        Fee = Cells(i, 10).Value
        Pre = Cells(i, 11).Value
        Comment = ""
        If Cells(i, 13).Value <> "" Then
            For j = i To i + Lines - 1
                If Cells(j, 13).Value <> "" Then Comment = Cells(j, 13).Value Else Exit For
            Next j
        End If
        If Len(Dat) = 1 Then Dat = "0" & Dat
        With ControlList
            .ListItems.Add = Dat
            .ListItems.Item(.ListItems.Count).ListSubItems.Add = Fee
            .ListItems.Item(.ListItems.Count).ListSubItems.Add = Pre
            .ListItems.Item(.ListItems.Count).ListSubItems.Add = Comment
            If CInt(Dat) = CDay_Box.Value Then
                .ListItems.Item(.ListItems.Count).ForeColor = &HFF&
                For j = 1 To ControlList.ListItems.Item(.ListItems.Count).ListSubItems.Count
                    .ListItems.Item(.ListItems.Count).ListSubItems(j).ForeColor = &HFF&
                Next j
            End If
        End With
    End If
Next i
ControlList.Refresh

Exit Sub
ExceptionControl:
Exception.Error_Box.Value = "Workers/FillControlList()"
Exception.Show
End Sub

Sub ReadLockedInfo()
On Error GoTo ExceptionControl:
Sheets(NameChooser.Value).Select

Left_Box.Value = Cells(2, 10).Value
Income_Box.Value = Cells(4, 9).Value
Outcome_Box.Value = Cells(3, 11).Value
Tax_Box.Value = -Cells(4, 10).Value
Balance_Box.Value = Cells(1, 10).Value
Oklad_Box.Value = Cells(4, 2).Value

If Cells(3, 1).Value = "RO" Then
    MakeReadOnly_Chk.Value = True
    If Not AdminMode Then
        Apply_Button.Enabled = False
        Clear_Button.Enabled = False
        Delete_Button.Enabled = False
        ChooseMate_Button.Enabled = False
        Select_Button.Enabled = False
    End If
Else
    MakeReadOnly_Chk.Value = False
    If Not (AdminMode Or LMMode) Then
        Apply_Button.Enabled = True
        Clear_Button.Enabled = True
        Delete_Button.Enabled = True
        ChooseMate_Button.Enabled = True
        Select_Button.Enabled = True
    End If
End If

If Oklad_Box <> "" Then AboveOklad_Chk.Visible = True Else AboveOklad_Chk.Visible = False
If Balance_Box.Value >= 0 Then Balance_Label.ForeColor = &H8000& Else Balance_Label.ForeColor = &HFF&

Exit Sub
ExceptionControl:
Exception.Error_Box.Value = "Workers/ReadLockedInfo()"
Exception.Show
End Sub

Sub SetRandomMark()
On Error GoTo ExceptionControl:
Randomize
Cells(2, 1).Value = Round(10000000 * Rnd())

Exit Sub
ExceptionControl:
Exception.Error_Box.Value = "Workers/SetRandomMark()"
Exception.Show
End Sub

Sub RecordLine(ByVal Day, ByVal Job, Optional ByVal Mark As Boolean = False)
Dim Index, i, ADLen As Integer
On Error GoTo ExceptionControl:
If NameChooser.Value <> "" And CDay_Box.MatchFound Then Sheets(NameChooser.Value).Select Else Exit Sub
Index = Job + InfoOffset + Lines * (Day - 1) - 1

If Mark Then
    If Cells(Index, 2).Value <> "" Then
        If Cells(Index, 15).Value = 1 Then Cells(Index, 15).Value = "" Else Cells(Index, 15).Value = 1
        LogAction ("MarkLine " & CStr(Index))
    End If
    Exit Sub
End If

If (ID.Value <> "") And DayTrigger Then
    Application.Calculation = xlCalculationManual
    Cells(Index - Job + 1, 10).FormulaR1C1 = "=SUM(RC[-1]:R[8]C[-1])"
    Cells(Index, 9).FormulaR1C1 = "=(RC[-5]*(1-RC[-1])+RC[-3]*RC[-1])*RC[-2]"
    'OldAmount = Cells(index, 4)
    'OldId = Cells(index, 3)
    'OldTime = Cells(index, 6)
    ADLen = Len(AltDiam_Box.Value)
    If ADLen > 2 And ADLen < 5 Then
       Cells(Index, 2).Value = ReplaceToAlternate(JobName_Box.Value, CInt(AltDiam_Box.Value))
       Cells(Index, 14).Value = AltDiam_Box.Value
    Else
       Cells(Index, 2).Value = JobName_Box.Value
       Cells(Index, 14).Value = ""
    End If
    
    Cells(Index, 3).Value = ID.Value
    Cells(Index, 5).Value = Unit.Caption
    If CheckNumber(Amount_Box.Value) Then Cells(Index, 4).Value = Amount_Box.Value
    If CheckNumber(Time_Box.Value) Then Cells(Index, 6).Value = Time_Box.Value
    
    If Not AdminMode Then SetRandomMark
    
    If Oklad_Box = "" Or AboveOklad_Chk.Value = True Then
        If CheckNumber(Rate_Box.Value) Then Cells(Index, 7).Value = Rate_Box.Value
    Else
        Cells(Index, 7).ClearContents
    End If

    If Rate_Box.Tag = "Time" Then Cells(Index, 8).Value = 1 Else Cells(Index, 8).Value = 0

    Cells(Index, 2).Select
    If (Cells(Index, 2).Value = "") Then Selection.EntireRow.Hidden = True Else Selection.EntireRow.Hidden = False
    LogAction ("RecordLine " & CStr(1 + (Index - InfoOffset) \ Lines) & "(" & CStr(Index) & ");" & _
        ID.Value & ";" & Amount_Box.Value & ";" & Time_Box.Value & ";" & AltDiam_Box.Value)
    'AddAmount = Amount_Box.Value
    'AddTime = Time_Box.Value
    'If AddAmount = "" Then AddAmount = 0
    'If AddTime = "" Then AddTime = 0
    'If OldId = "" Then OldId = 0
    'Sheets("Каталог").Select
    'If (OldId <> 0) And (OldId <> 5) Then Cells(OldId, 11) = Cells(OldId, 11) - OldAmount
    'Cells(JobPosition(OldId) + 287, 6) = Cells(JobPosition(OldId) + 287, 6) - OldTime
    'End If
    'If ID.Value <> 5 Then Cells(ID.Value, 11) = Cells(ID.Value, 11) + AddAmount
    'Cells(JobPosition(ID.Value) + 287, 6) = Cells(JobPosition(ID.Value) + 287, 6) + AddTime
    'Sheets(NameChooser.Value).Select
    Application.Calculation = xlCalculationAutomatic
End If

If CommentTrigger And Comment_Box.Value <> "" Then
    For i = Index - Job + 1 To Index - Job + Lines
        If (Cells(i, 13).Value = "") Or (i = Index - Job + Lines) Then
            Cells(i, 13).Value = Comment_Box.Value
            LogAction ("NewComment " & CStr(Day) & ";" & Comment_Box.Value)
            Exit For
        End If
    Next i
    CommentTrigger = False
    If Not AdminMode Then SetRandomMark
End If

If AdminMode Then
    If CheckNumber(PrePay_Box.Value) And (CStr(Cells(Index - Job + 1, 11).Value) <> PrePay_Box.Value) Then
        Cells(Index - Job + 1, 11).Value = PrePay_Box.Value
        LogAction ("SetPrePay " & CStr(Day) & ";" & PrePay_Box.Value)
    End If
    If CheckNumber(Left_Box.Value) And (CStr(Cells(2, 10).Value) <> Left_Box.Value) Then
        Cells(2, 10).Value = Left_Box.Value
        LogAction ("SetLeft " & Left_Box.Value)
    End If
    If CheckNumber(Oklad_Box.Value) Then Cells(4, 2).Value = Oklad_Box.Value
    If MakeReadOnly_Chk.Value = True Then Cells(3, 1).Value = "RO" Else Cells(3, 1).Value = ""
End If
Cells(Index - Job + 1, 2).Select
If (Cells(Index - Job + 1, 2).Value = "") And _
   (Cells(Index - Job + 1, 11).Value = "") And _
   (Cells(Index - Job + 1, 13).Value = "") Then _
    Selection.EntireRow.Hidden = True _
    Else Selection.EntireRow.Hidden = False
If CInt(Day) > CInt(Cells(1, 1).Value) Then Cells(1, 1).Value = Day
ReadLockedInfo
FillDayList (CDay_Box.Value)
MakeShitLookGood
If LMMode Then TransferBalance NameChooser.Value, Balance_Box.Value
DayTrigger = False
FillControlList

Exit Sub
ExceptionControl:
Exception.Error_Box.Value = "Workers/RecordLine()"
Exception.Show
End Sub

Sub DeleteLine(ByVal Day, ByVal Job)
On Error GoTo ExceptionControl:
If NameChooser.Value <> "" And CDay_Box.MatchFound Then Sheets(NameChooser.Value).Select Else Exit Sub
Index = Job + InfoOffset + Lines * (Day - 1) - 1

'OldAmount = Cells(index, 4)
'OldId = Cells(index, 3)
'OldTime = Cells(index, 6)
'Sheets("Каталог").Select
'If OldId > 5 Then Cells(OldId, 11) = Cells(OldId, 11) - OldAmount
'Sheets(NameChooser.Value).Select
Application.Calculation = xlCalculationManual
Range(Cells(Index, 2), Cells(Index, 9)).ClearContents
Range(Cells(Index, 14), Cells(Index, 15)).ClearContents
If AdminMode Then Cells(Index, 3) = 4
Cells(Index, 2).Select
If (Cells(Index, 2).Value = "") And _
   (Cells(Index, 11).Value = "") And _
   (Cells(Index, 13).Value = "") Then Selection.EntireRow.Hidden = True
ReadLockedInfo
JobName_Box.Value = ""
ID = ""
Time_Box.Value = ""
Amount_Box.Value = ""
Rate_Box.Value = ""
Rate_Box.Tag = ""
Unit.Caption = ""
LogAction ("DeleteLine " & CStr(Index))
Application.Calculation = xlCalculationAutomatic
MakeShitLookGood
If LMMode Then TransferBalance NameChooser.Value, Balance_Box.Value
FillDayList (CDay_Box.Value)
FillControlList

Exit Sub
ExceptionControl:
Exception.Error_Box.Value = "Workers/DeleteLine()"
Exception.Show
End Sub

Sub ClearDay(ByVal Day)
Dim Index, i As Integer
On Error GoTo ExceptionControl:
If NameChooser.Value <> "" And CDay_Box.MatchFound Then Sheets(NameChooser.Value).Select Else Exit Sub
Index = InfoOffset + Lines * (Day - 1)
If CInt(Cells(1, 1)) = Day Then Cells(1, 1).ClearContents
For i = 0 To Lines - 1
    If (AdminMode) And (Cells(Index + i, 3) <> "") Then Cells(Index + i, 3) = 4
    'OldAmount = Cells(index + i, 4)
    'OldId = Cells(index + i, 3)
    'OldTime = Cells(index + i, 6)
    'If OldId > 5 Then
    'Sheets("Каталог").Select
    'Cells(OldId, 11) = Cells(OldId, 11) - OldAmount
    'Sheets(NameChooser.Value).Select
    'End If
Next
Application.Calculation = xlCalculationManual
If AdminMode Then
    Range(Cells(Index, 2), Cells(Index + Lines - 1, 2)).ClearContents
    Range(Cells(Index, 4), Cells(Index + Lines - 1, 9)).ClearContents
Else
    Range(Cells(Index, 2), Cells(Index + Lines - 1, 9)).ClearContents
End If
Range(Cells(Index, 13), Cells(Index + Lines - 1, 15)).ClearContents
Cells(Index, 11) = ""
Cells(Index, 10) = ""
Range(Cells(Index, 2), Cells(Index + Lines - 1, 2)).EntireRow.Hidden = True
LogAction ("ClearDay " & CStr(Day))
Application.Calculation = xlCalculationAutomatic
ReadLockedInfo
JobName_Box.Value = ""
ID = ""
Time_Box.Value = ""
Amount_Box.Value = ""
Rate_Box.Value = ""
Unit.Caption = ""
MakeShitLookGood
If LMMode Then TransferBalance NameChooser.Value, Balance_Box.Value
FillDayList (CDay_Box.Value)
FillControlList

Exit Sub
ExceptionControl:
Exception.Error_Box.Value = "Workers/ClearDay()"
Exception.Show
End Sub

Sub ReadLine(ByVal Day, ByVal Job)
On Error GoTo ExceptionControl:
If Day = "" Or NameChooser.Value = "" Or Not CDay_Box.MatchFound Then Exit Sub
    
inRead = True
Index = Job + InfoOffset + Lines * (Day - 1) - 1
Sheets(NameChooser.Value).Select
If Cells(Index, 3) > 4 Then ID.Value = Cells(Index, 3) Else ID.Value = ""
JobName_Box.Value = Cells(Index, 2).Value
Rate_Box.Value = Cells(Index, 7).Value
Rate_Box.Tag = ""
Time_Box.Value = Cells(Index, 6).Value
Unit.Caption = Cells(Index, 5).Value
Amount_Box.Value = Cells(Index, 4).Value
AltDiam_Box.Value = Cells(Index, 14).Value
If Unit.Caption = "" Then Amount_Box.Enabled = False Else Amount_Box.Enabled = True
If Oklad_Box.Value <> "" And Rate_Box.Value <> "" Then AboveOklad_Chk.Value = True
If Oklad_Box.Value <> "" And Rate_Box.Value = "" Then AboveOklad_Chk.Value = False
If Amount_Box.Enabled = False Then Time_Box.SetFocus Else Amount_Box.SetFocus
If JobName_Box.Value = "" Then JobsTree.SetFocus
inRead = False
DayTrigger = False

Exit Sub
ExceptionControl:
Exception.Error_Box.Value = "Workers/ReadLine()"
Exception.Show
End Sub

Function LastFilled(ByVal Day) As Integer
On Error GoTo ExceptionControl:
Index = InfoOffset + Lines * (Day - 1)
LastFilled = 0
For i = 1 To Lines
    If Cells(Index + i - 1, 2).Value <> "" Then LastFilled = i
Next

Exit Function
ExceptionControl:
Exception.Error_Box.Value = "Workers/LastFilled()"
Exception.Show
End Function

Private Sub AltDiam_Box_Change()
If AltDiam_Box.Value <> "" Then AltDiam_Box.Value = PointFilter(AltDiam_Box.Value, False, False, 4)
DayTrigger = True
End Sub

Private Sub Bonus_Button_Click()
On Error GoTo ExceptionControl:
ObjectsRecall
inRead = True
Sheets(NameChooser.Value).Select
Index = InfoOffset + Lines * (CDay_Box.Value - 1)
JobName_Box.Value = "Доплата " & CStr(BonusRate_Box.Value) & " %"
Amount_Box.Value = 1
Unit.Caption = " "
Rate_Box.Tag = "Amt"
Rate_Box.Value = Round(Cells(Index, 10).Value * BonusRate_Box.Value / 100, 2)
Time_Box.Value = ""
ID.Value = 5
If Apply_Button.Enabled = True Then Apply_Button.SetFocus
inRead = False

Exit Sub
ExceptionControl:
Exception.Error_Box.Value = "Workers/Bonus_Button_Click()"
Exception.Show
End Sub

Private Sub BonusRate_Box_Change()
If BonusRate_Box.Value <> "" Then BonusRate_Box.Value = PointFilter(BonusRate_Box.Value, False, False, 3)
End Sub

Private Sub CollapseJobs_Button_Click()
On Error GoTo ExceptionControl:
For i = 1 To JobsTree.Nodes.Count
    JobsTree.Nodes(i).Expanded = False
Next

Exit Sub
ExceptionControl:
Exception.Error_Box.Value = "Workers/CollapseJobs_Button_Click()"
Exception.Show
End Sub

Private Sub Comment_Box_Change()
If Not inRead Then CommentTrigger = True
End Sub

Private Sub Comment_Box_DropButtonClick()
ObjectsRecall
End Sub

Private Sub ConfirmChoice_Button_Click()
Select Case ConfirmObject
Case 1
       WorkersTree_DblClick
Case 2
       ControlList_DblClick
Case 3
       DayList_DblClick
Case 4
       JobsTree_DblClick
End Select
ObjectsRecall
ConfirmObject = 0
End Sub

Private Sub ControlList_Click()
ControlList.Refresh
ObjectsRecall
ConfirmObject = 2
ConfirmChoice_Button.Top = 120
ConfirmChoice_Button.Left = 570
End Sub

Private Sub DayList_Click()
DayList.Refresh
ObjectsRecall
ConfirmObject = 3
End Sub

Private Sub Delete_Button_Click()
ObjectsRecall
DeleteLine CDay_Box.Value, CJob_Box.Value
End Sub

Private Sub JobsTree_Click()
ObjectsRecall
ConfirmObject = 4
End Sub

Private Sub LastMonth_Label_Click()
ObjectsRecall
End Sub

Private Sub Div_Button_Click()
On Error GoTo ExceptionControl:
ObjectsRecall
If Amount_Box.Value <> "" And Amount_Box.Value <> 0 And CheckNumber(Amount_Box.Value) Then Amount_Box.Value = Round(Amount_Box.Value / 2, 2)
If Apply_Button.Enabled = True Then Apply_Button.SetFocus

Exit Sub
ExceptionControl:
Exception.Error_Box.Value = "Workers/Div_Button_Click()"
Exception.Show
End Sub

Private Sub Log_Button_Click()
Dim i, Stopp, Count As Integer
On Error GoTo ExceptionControl:
ObjectsRecall
If LogHolder.Visible = True Then
    LogHolder.Visible = False
Else
    LogHolder.Top = -500
    LogHolder.Left = 6
    LogHolder.Visible = True
    With Log
        Count = 1
        .ListItems.Clear
        .Sorted = False
        Stopp = CInt(Cells(InfoOffset - 1, 19).Value)
        If CInt(Cells(InfoOffset - 1, 21).Value) > Stopp Then Stopp = CInt(Cells(InfoOffset - 1, 21).Value)
        If Stopp = 0 Then Stopp = 600
        For i = InfoOffset To InfoOffset + Stopp
            If Cells(i, 19).Value = "" And Cells(i, 21).Value = "" Then
                Exit For
            End If
            If Cells(i, 19).Value <> "" Then
                .ListItems.Add = Cells(i, 19).Value
                .ListItems.Item(Count).ListSubItems.Add = Cells(i, 20).Value
                Count = Count + 1
            End If
            If Cells(i, 21).Value <> "" Then
                .ListItems.Add = Cells(i, 21).Value
                .ListItems.Item(Count).ListSubItems.Add = Cells(i, 22).Value
                Count = Count + 1
            End If
        Next i
        .Sorted = True
        LogHolder.Top = 6
    End With

Exit Sub
ExceptionControl:
Exception.Error_Box.Value = "Workers/Log_Button_Click()"
Exception.Show
End If
End Sub

Private Sub Select_Button_Click()
On Error GoTo ExceptionControl:
ObjectsRecall
If DayList.ListItems.Count > 0 Then
    If DayList.SelectedItem.Text = "" Or DayList.SelectedItem.Text = " " Then Exit Sub
    RecordLine CDay_Box.Value, CInt(DayList.SelectedItem.Text), True
    FillDayList CDay_Box.Value
End If

Exit Sub
ExceptionControl:
Exception.Error_Box.Value = "Workers/Select_Button_Click()"
Exception.Show
End Sub

Private Sub Triv_Button_Click()
On Error GoTo ExceptionControl:
ObjectsRecall
If Amount_Box.Value <> "" And Amount_Box.Value <> 0 And CheckNumber(Amount_Box.Value) Then Amount_Box.Value = Round(Amount_Box.Value / 3, 2)
If Apply_Button.Enabled = True Then Apply_Button.SetFocus

Exit Sub
ExceptionControl:
Exception.Error_Box.Value = "Workers/Triv_Button_Click()"
Exception.Show
End Sub

Private Sub SelectUpdatesOnly_Change()
ObjectsRecall
ScanWorkers
End Sub

Private Sub ControlList_DblClick()
On Error GoTo ExceptionControl:
If ControlList.ListItems.Count > 0 Then
    If ControlList.SelectedItem.Text <> "" Then CDay_Box.Value = CInt(ControlList.SelectedItem.Text)
End If

Exit Sub
ExceptionControl:
Exception.Error_Box.Value = "Workers/ControlList_DblClick()"
Exception.Show
End Sub

Private Sub DayList_DblClick()
On Error GoTo ExceptionControl:
Records = DayList.ListItems.Count
If Records > 0 Then
    If DayList.SelectedItem.Text = "" Then Exit Sub
    If (DayList.SelectedItem.Text = " ") Then
        If (Records - 2) < Lines Then CJob_Box.Value = Records - 1
    Else
        CJob_Box.Value = CInt(DayList.SelectedItem.Text)
    End If
End If

Exit Sub
ExceptionControl:
Exception.Error_Box.Value = "Workers/DayList_DblClick()"
Exception.Show
End Sub

Sub FillDayList(ByVal Day)
Dim Index, i, Records As Integer
On Error GoTo ExceptionControl:
If NameChooser.Value <> "" Then Sheets(NameChooser.Value).Select Else Exit Sub
If Day = "" Then Exit Sub
DayList.ListItems.Clear
Comment_Box.Clear
Comment_Box.Value = ""
CommentTrigger = False
Index = InfoOffset + Lines * (Day - 1)
Records = LastFilled(Day)
PrePay_Box.Value = Cells(Index, 11).Value
If (Records > 0) Or (Cells(Index, 13).Value <> "") Then
    TotalTime = 0
    For i = 1 To Lines
        Index = i + InfoOffset + Lines * (Day - 1) - 1
        If i <= Records Then
            JobName = Cells(Index, 2)
            Amount = Cells(Index, 4)
            UnitList = Cells(Index, 5)
            TimeList = Cells(Index, 6)
            TotalTime = TotalTime + TimeList
            RateList = Cells(Index, 7)
            Subtotal = Cells(Index, 9)
            With DayList
                .ListItems.Add = i
                .ListItems.Item(i).ListSubItems.Add = JobName
                .ListItems.Item(i).ListSubItems.Add = Amount
                .ListItems.Item(i).ListSubItems.Add = UnitList
                .ListItems.Item(i).ListSubItems.Add = TimeList
                .ListItems.Item(i).ListSubItems.Add = RateList
                .ListItems.Item(i).ListSubItems.Add = Subtotal
                If Cells(Index, 15) = "1" Then
                    .ListItems.Item(i).ForeColor = &HC000&
                    For j = 1 To .ListItems.Item(i).ListSubItems.Count
                        .ListItems.Item(i).ListSubItems(j).ForeColor = &HC000&
                    Next j
                End If
            End With
        End If
        If Cells(Index, 13).Value <> "" Then Comment_Box.AddItem (Cells(Index, 13).Value)
    Next i
    With DayList
        If .ListItems.Count > 0 Then
            Index = InfoOffset + Lines * (Day - 1)
            If .ListItems.Count < Lines Then
                .ListItems.Add = .ListItems.Count + 1
                If .ListItems.Count = CJob_Box.Value Then .ListItems.Item(CInt(CJob_Box.Value)).ForeColor = &HFF&
            Else
                .ListItems.Add = " "
            End If
                .ListItems.Add = " "
                .ListItems.Item(.ListItems.Count).ListSubItems.Add = ""
                .ListItems.Item(.ListItems.Count).ListSubItems.Add = ""
                .ListItems.Item(.ListItems.Count).ListSubItems.Add = "ВСЕГО"
                .ListItems.Item(.ListItems.Count).ListSubItems.Add = TotalTime
                .ListItems.Item(.ListItems.Count).ListSubItems.Add = "ИТОГО"
                .ListItems.Item(.ListItems.Count).ListSubItems.Add = Cells(Index, 10).Value
                .Refresh
        End If
    End With
    If Comment_Box.ListCount > 0 Then
        Comment_Box.Value = Comment_Box.List(Comment_Box.ListCount - 1)
        CommentTrigger = False
    End If
End If

If Records < Lines Then CJob_Box.Value = Records + 1
If Records > Lines - 1 Then CJob_Box.Value = Lines

Exit Sub
ExceptionControl:
Exception.Error_Box.Value = "Workers/FillDayList()"
Exception.Show
End Sub

Private Sub AboveOklad_Chk_Change()
DayTrigger = True
End Sub

Private Sub Amount_Box_Change()
If Amount_Box.Value <> "" Then Amount_Box.Value = PointFilter(Amount_Box.Value)
DayTrigger = True
End Sub

Private Sub ChooseWorker_Click()
On Error GoTo ExceptionControl:
ObjectsRecall
WorkersTreeHolder.Top = 6
WorkersTreeHolder.Left = 320
'For i = 1 To CInt(WorkersTree.Tag)
'  If WorkersTree.Nodes(i).Tag = "Cat" Then _
'        WorkersTree.Nodes(i).Expanded = False
' Next
'WorkersTree.Visible = True

DateAndWorker_Frame.Visible = False
WorkersTreeHolder.Visible = True
ConfirmObject = 1
ConfirmChoice_Button.Top = 46
ConfirmChoice_Button.Left = 270
If ChosenMate <> 0 Then
    WorkersTree.Nodes(ChosenMate).Selected = True
    ChosenMate = 0
End If
WorkersTree.SetFocus

Exit Sub
ExceptionControl:
Exception.Error_Box.Value = "Workers/ChooseWorker_Click()"
Exception.Show
End Sub

Private Sub ChooseMate_Button_Click()
On Error Resume Next
ObjectsRecall
WorkersTreeHolder.Top = 255
WorkersTreeHolder.Left = 290
ConfirmObject = 1
ConfirmChoice_Button.Top = 310
ConfirmChoice_Button.Left = 240
'For i = 1 To CInt(WorkersTree.Tag)
'  If WorkersTree.Nodes(i).Tag = "Cat" Then _
'        WorkersTree.Nodes(i).Expanded = False
' Next
'WorkersTree.Visible = True
WorkersTreeHolder.Visible = True
ChosenMate = WorkersTree.SelectedItem.Index
If ChosenMate = 0 Then ChosenMate = 1
DayList.Visible = False
WorkersTree.SetFocus
End Sub

Private Sub Frame2_Click()
ObjectsRecall
End Sub

Private Sub Frame3_Click()
ObjectsRecall
End Sub

Private Sub Frame4_Click()
ObjectsRecall
End Sub

Private Sub Frame6_Click()
ObjectsRecall
End Sub

Private Sub DateAndWorker_Frame_Click()
ObjectsRecall
End Sub

Private Sub JobName_Box_Change()
On Error GoTo ExceptionControl:
If (Left(Right(JobName_Box.Value, 4), 1) = "х") Or _
   (Left(Right(JobName_Box.Value, 4), 1) = "x") Or _
   (Left(Right(JobName_Box.Value, 5), 1) = "х") Or _
   (Left(Right(JobName_Box.Value, 5), 1) = "x") Then
    AltDiam_Box.Visible = True
    AltDiam_Label.Visible = True
Else
    AltDiam_Box.Visible = False
    AltDiam_Box.Value = ""
    AltDiam_Label.Visible = False
End If

Exit Sub
ExceptionControl:
Exception.Error_Box.Value = "Workers/JobName_Box_Change()"
Exception.Show
End Sub

Private Sub JobsTree_DblClick()
If JobsTree.SelectedItem.Key <> "" Then
    ID.Value = CutZ(JobsTree.SelectedItem.Key)
End If
End Sub

Private Sub Left_Box_Change()
If Left_Box.Value <> "" Then Left_Box.Value = PointFilter(Left_Box.Value)
End Sub

Private Sub MateChooser_Change()
If MateChooser.Value = "" Then CopyDay_Button.Enabled = False Else CopyDay_Button.Enabled = True
End Sub

Private Sub Oklad_Box_Change()
If Oklad_Box.Value <> "" Then Oklad_Box.Value = PointFilter(Oklad_Box.Value, False, False)
End Sub

Private Sub PrePay_Box_Change()
If PrePay_Box.Value <> "" Then PrePay_Box.Value = PointFilter(PrePay_Box.Value, False)
End Sub

Function isVisible(ByVal Day As Integer) As Boolean
On Error GoTo ExceptionControl:
isVisible = False
Index = InfoOffset + Lines * Day
If Cells(Index, 10) <> 0 Then
    isVisible = True
    Exit Function
End If
If Cells(Index, 11) > 0 Then
    isVisible = True
    Exit Function
End If
If Cells(Index, 13) <> "" Then
    isVisible = True
    Exit Function
End If

Exit Function
ExceptionControl:
Exception.Error_Box.Value = "Workers/isVisible()"
Exception.Show
End Function

Sub Mark(ByVal Day As Integer, ByVal PrevMarked As Boolean)
On Error GoTo ExceptionControl:
If PrevMarked Then Colorr = 2 Else Colorr = 15
Index = InfoOffset + Lines * Day
Range(Cells(Index, 1), Cells(Index + Lines - 1, 12)).Select
With Selection.Interior
     .ColorIndex = Colorr
     .Pattern = xlSolid
     .PatternColorIndex = xlAutomatic
End With

Exit Sub
ExceptionControl:
Exception.Error_Box.Value = "Workers/Mark()"
Exception.Show
End Sub

Sub MakeShitLookGood()
On Error GoTo ExceptionControl:
Last = Cells(1, 1).Value
PrevMarked = True
If Last = "" Then Last = MDays(CMonth)
For i = 0 To CInt(Last)
    If isVisible(i) = True Then
        PrevMarked = Not PrevMarked
        Mark i, PrevMarked
    End If
Next

Exit Sub
ExceptionControl:
Exception.Error_Box.Value = "Workers/MakeShitLookGood()"
Exception.Show
End Sub

Private Sub Print_Button_Click()
On Error Resume Next
ObjectsRecall
If NameChooser.Value <> "" Then
    Count = 0
    Sheets(NameChooser.Value).Select
    MakeShitLookGood
    If OnScreen_Chk.Value = True Then Sheets(NameChooser.Value).PrintOut
    If OnScreen_Chk.Value = False Then
        WorkersExit = True
        Application.ScreenUpdating = True
        Workers.Hide
        Main.Hide
        Cells(3, 1).Select
    End If
End If

Exit Sub
ExceptionControl:
Exception.Error_Box.Value = "Workers/Print_Button_Click()"
Exception.Show
End Sub

Private Sub Rate_Box_Change()
If Rate_Box.Value <> "" Then Rate_Box.Value = PointFilter(Rate_Box.Value)
DayTrigger = True
End Sub

Private Sub Time_Box_Change()
If Time_Box.Value <> "" Then Time_Box.Value = PointFilter(Time_Box.Value, False)
DayTrigger = True
End Sub

Private Sub Apply_Button_Click()
ObjectsRecall
RecordLine CDay_Box.Value, CJob_Box.Value
End Sub

Private Sub CDay_Box_Change()
On Error GoTo ExceptionControl:
If (NameChooser.Value <> "") And (Not ExtChange) And CDay_Box.MatchFound Then
    ObjectsRecall
    FillDayList (CDay_Box.Value)
    ReadLine Workers.CDay_Box.Value, Workers.CJob_Box.Value
    Label_FullDate.Caption = GetDayName(Workers.CDay_Box.Value) & ", " & Workers.CDay_Box.Value & " " & MName(CMonth, True)
    Workers.Caption = Workers.RealName_Box.Value & ": " & Workers.Label_FullDate.Caption
    MarkListLine ControlList, CDay_Box.Value
End If

Exit Sub
ExceptionControl:
Exception.Error_Box.Value = "Workers/CDay_Box_Change()"
Exception.Show
End Sub

Private Sub CJob_Box_Change()
ReadLine CDay_Box.Value, CJob_Box.Value
MarkListLine DayList, CJob_Box.Value
End Sub

Private Sub MarkListLine(ByRef List As ListView4, ByVal Line As Integer)
Dim i, j As Integer
On Error GoTo ExceptionControl:
With List
    For i = 1 To .ListItems.Count
        If .ListItems.Item(i).Text <> " " Then
            OldColor = .ListItems.Item(i).ForeColor
            If CInt(.ListItems.Item(i).Text) = Line Then
                If OldColor = &H0& Then NewColor = &HFF& Else NewColor = &H80&
                .ListItems.Item(i).ForeColor = NewColor
                For j = 1 To .ListItems.Item(i).ListSubItems.Count
                    .ListItems.Item(i).ListSubItems(j).ForeColor = NewColor
                Next j
            Else
                NewColor = OldColor
                If OldColor = &HFF& Then
                    NewColor = &H0&
                Else
                    If OldColor = &H80& Then NewColor = &HC000&
                End If
                If NewColor <> OldColor Then
                    .ListItems.Item(i).ForeColor = NewColor
                    For j = 1 To .ListItems.Item(i).ListSubItems.Count
                        .ListItems.Item(i).ListSubItems(j).ForeColor = NewColor
                    Next j
                End If
            End If
        End If
    Next i
    .Refresh
End With

Exit Sub
ExceptionControl:
Exception.Error_Box.Value = "Workers/MarkListLine()"
Exception.Show
End Sub

Private Sub Control_Button_Click()
FillControlList
End Sub

Private Sub CopyDay_Button_Click()
Dim Index, i As Integer
On Error GoTo ExceptionControl:
ObjectsRecall
If NameChooser.Value <> "" And MateChooser.Value <> "" And CDay_Box.MatchFound Then
    Sheets(MateChooser.Value).Select
    If Cells(3, 1).Value = "RO" And Not AdminMode Then
        Sheets(NameChooser.Value).Select
        b = MsgBox("Запись в лист сотрудника " & MateName_Box & " с рабочего места невозможна.", vbOKOnly, "Копирование отменено")
        Exit Sub
    End If
    For i = 1 To Lines
        Index = i + InfoOffset + Lines * (CDay_Box.Value - 1) - 1
        If Cells(Index, 2).Value <> "" Then
            If Not AdminMode Then
                Sheets(NameChooser.Value).Select
                b = MsgBox(MateName_Box & " уже записался на " & CDay_Box.Value & " " & MName(CMonth, True) & ".", vbOKOnly, "Копирование отменено")
                Exit Sub
            Else
                Query.Msg_label.Caption = "Все записи у " & MateName_Box & " за " & CDay_Box.Value & " " & MName(CMonth, True) & " будут перезаписаны. Выполнить копирование?"
                Query.NoButton.SetFocus
                Query.Show
                If Query.OK.Value = True Then
                    Exit For
                Else
                    Sheets(NameChooser.Value).Select
                    Exit Sub
                End If
            End If
        End If
    Next i
    Application.Calculation = xlCalculationManual
    For i = 1 To Lines
        Sheets(NameChooser.Value).Select
        Index = i + InfoOffset + Lines * (CDay_Box.Value - 1) - 1
 
        If Cells(Index, 2).Value <> "" Then
            'AddAmount = Cells(index, 4).Value
            'AddID = Cells(index, 3).Value
            Range(Cells(Index, 2), Cells(Index, 9)).Copy
            CopyAlternateDiam = Cells(Index, 14).Value
            MarkFlag = Cells(Index, 15).Value
            'Sheets("Каталог").Select
            'If AddID > 5 Then Cells(AddID, 11).Value = Cells(AddID, 11).Value + AddAmount
            'Sheets(MateChooser.Value).Select
            'OldAmount = Cells(index, 4).Value
            'OldId = Cells(index, 3).Value
            'If OldId <> "" And AddID > 5 Then
            'Sheets("Каталог").Select
            'Cells(OldId, 11).Value = Cells(OldId, 11).Value - OldAmount
            'End If
            Sheets(MateChooser.Value).Select
            Cells(Index, 2).PasteSpecial
            Cells(Index, 14).Value = CopyAlternateDiam
            Cells(Index, 15).Value = MarkFlag
            Cells(Index, 2).EntireRow.Hidden = False
        End If
    Next i
    Index = InfoOffset + Lines * (CDay_Box.Value - 1)
    Sheets(MateChooser.Value).Select
    Cells(Index, 10).FormulaR1C1 = "=SUM(RC[-1]:R[8]C[-1])"
    If CDay_Box.Value > Cells(1, 1).Value Then Cells(1, 1).Value = CDay_Box.Value
    If Not AdminMode Then
        SetRandomMark
        Cells(Index, 13).Value = "Копировал " & RealName_Box & " " & DateTime.Date & " " & DateTime.Time
    End If
    LogAction ("Copy day " & CDay_Box.Value & " from " & NameChooser.Value)
    Application.Calculation = xlCalculationAutomatic
    MakeShitLookGood
    If LMMode Then TransferBalance MateChooser.Value, Cells(1, 10).Value
    Sheets(NameChooser.Value).Select
End If

Exit Sub
ExceptionControl:
Exception.Error_Box.Value = "Workers/CopyDay_Button_Click()"
Exception.Show
End Sub

Private Sub Clear_Button_Click()
On Error GoTo ExceptionControl:
ObjectsRecall
If Not AdminMode Then
    Query.Msg_label.Caption = "Вы действительно хотите стереть все записи за " & CDay_Box.Value & " " & MName(CMonth, True) & "?"
    Query.NoButton.SetFocus
    Query.Show
    If Query.OK.Value = True Then ClearDay (CDay_Box.Value)
Else
    ClearDay (CDay_Box.Value)
End If

Exit Sub
ExceptionControl:
Exception.Error_Box.Value = "Workers/Clear_Button_Click()"
Exception.Show
End Sub

Private Sub ID_Change()
On Error GoTo ExceptionControl:
If Not inRead Then
    DayTrigger = True
    If ID.Value <> "" And ID.Value <> 0 Then
        ID = ID.Value
        Sheets("Каталог").Select
        JobName_Box.Value = Cells(ID, 2).Value
        AltDiam_Box.Value = ""
        Amount_Box.Enabled = True
        Unit.Caption = Cells(ID, 4)
        Amount_Box.SetFocus
        If Unit.Caption = "" Then
            Amount_Box.Enabled = False
            Amount_Box.Value = ""
            Time_Box.SetFocus
        End If
        Rate_Box.Value = Cells(ID, 5).Value
        Rate_Box.Tag = "Amt"
        If Rate_Box.Value = 0 Then
            Rate_Box.Value = Cells(ID, 6).Value
            Rate_Box.Tag = "Time"
            Time_Box.SetFocus
        End If
    End If
End If

Exit Sub
ExceptionControl:
Exception.Error_Box.Value = "Workers/ID_Change()"
Exception.Show
End Sub

Private Sub NameChooser_Change()
On Error GoTo ExceptionControl:
If NameChooser.Value <> "" Then
    Sheets(NameChooser.Value).Select
    FillControlList
    ReadLockedInfo
    Workers.Caption = RealName_Box.Value & ": " & Label_FullDate.Caption
    FillDayList (CDay_Box.Value)
    If NameChooser.Value <> "Образец" And Not AdminMode Then
        Add = ""
        If BlockIt.RcvHash = PinAdmin Then Add = " (Admin)"
        If BlockIt.RcvHash = PinSuperV Then Add = " (Supervisor)"
        LogAction ("Login with " & BlockIt.RcvHash & Add)
    End If
    If NameChooser.Value = MateChooser.Value Or Not AdminMode Then
        MateChooser.Value = ""
        MateName_Box.Value = ""
    End If
End If

Exit Sub
ExceptionControl:
Exception.Error_Box.Value = "Workers/NameChooser_Change()"
Exception.Show
End Sub

Private Sub Day_Spin_SpinDown()
On Error GoTo ExceptionControl:
ObjectsRecall
If CDay_Box.Value < MDays(CMonth) Then CDay_Box.Value = CDay_Box.Value + 1

Exit Sub
ExceptionControl:
Exception.Error_Box.Value = "Workers/Day_Spin_SpinDown()"
Exception.Show
End Sub

Private Sub Day_Spin_SpinUp()
On Error GoTo ExceptionControl:
ObjectsRecall
If CDay_Box.Value > 1 Then CDay_Box.Value = CDay_Box.Value - 1

Exit Sub
ExceptionControl:
Exception.Error_Box.Value = "Workers/Day_Spin_SpinUp()"
Exception.Show
End Sub

Private Sub Workers_Spin_SpinDown()
On Error Resume Next
ObjectsRecall
WorkersTreeHolder.Top = -500
WorkersTreeHolder.Visible = True
Total = WorkersTree.Nodes.Count
TotalCat = CInt(WorkersTree.Tag)
If ChosenMate <> 0 Then
    Index = ChosenMate
    ChosenMate = 0
Else
    Index = WorkersTree.SelectedItem.Index
End If
If Index < TotalCat Then Index = TotalCat
If Index <= Total - 1 Then
    If WorkersTree.Nodes(Index + 1).Tag <> "Cat" Then
        RealName_Box.Value = WorkersTree.Nodes(Index + 1).Text
        NameChooser.Value = WorkersTree.Nodes(Index + 1).Key
        WorkersTree.Nodes(Index + 1).Selected = True
    End If
End If
WorkersTreeHolder.Visible = False
End Sub

Private Sub Workers_Spin_SpinUp()
On Error GoTo over:
ObjectsRecall
WorkersTreeHolder.Top = -500
WorkersTreeHolder.Visible = True
If ChosenMate <> 0 Then
    Index = ChosenMate
    ChosenMate = 0
Else
    Index = WorkersTree.SelectedItem.Index
End If
If WorkersTree.Nodes(Index - 1).Tag <> "Cat" Then
    RealName_Box.Value = WorkersTree.Nodes(Index - 1).Text
    NameChooser.Value = WorkersTree.Nodes(Index - 1).Key
    WorkersTree.Nodes(Index - 1).Selected = True
Else
    WorkersTree.Nodes(Index).Selected = True
End If
over:
WorkersTreeHolder.Visible = False
End Sub

Private Sub UserForm_Click()
ObjectsRecall
End Sub

Private Sub WorkersClose_Button_Click()
On Error GoTo ExceptionControl:
ObjectsRecall
If CDay_Box.Value <> "" Then LastWorkersDay = CInt(CDay_Box.Value)
''LastPerson = NameChooser.Value
If Not AdminMode Then
    If Not LMMode Then
        RealName_Box.Value = ""
        NameChooser.Value = "Образец"
        NameChooser.Value = ""
        TS = TokenSum()
        Sheets("Каталог").Select
        Cells(2, 6).Value = TS
        ProcessFile WorkersBase, "SaveClose"
        PullBase = "pull.xls"
        Destination = Path & PullBase
        Source = Path & WorkersBase
        FileCopy Source, Destination
        ArcName = Path & "pull.7z"
        ArcFiles = Path & PullBase
        RunCommand (Archiver & " a -sdel " & ExchangeKey & " " & ArcName & " " & ArcFiles)
        RunCommand "ftp -v -s:" & Path & "ftp_client_send " & FtpStorageName, False
    End If
    Workers.Hide
Else
Workers.Hide
End If

Exit Sub
ExceptionControl:
Exception.Error_Box.Value = "Workers/WorkersClose_Button_Click()"
Exception.Show
End Sub

Private Sub Logout_Button_Click()
On Error GoTo ExceptionControl:
ObjectsRecall
RealName_Box.Value = ""
PrePay_Box.Value = ""
NameChooser.Value = "Образец"
NameChooser.Value = ""

Exit Sub
ExceptionControl:
Exception.Error_Box.Value = "Workers/Logout_Button_Click()"
Exception.Show
End Sub

Private Sub WorkersTree_Click()
ConfirmObject = 1
End Sub

Private Sub WorkersTree_DblClick()
On Error GoTo ExceptionControl:
If WorkersTree.SelectedItem.Key <> "" And WorkersTree.SelectedItem.Tag <> "Cat" Then
     If ChosenMate = 0 Then
        If Not AdminMode Then
            BlockIt.Pass = WorkersTree.SelectedItem.Tag
            BlockIt.PassOK = False
            BlockIt.AdminOverrides = True
            BlockIt.SupervisorOverrides = True
            BlockIt.Password_Box.SetFocus
            BlockIt.Show
        End If
        If BlockIt.PassOK Or AdminMode Then
            If WorkersTree.SelectedItem.Key = MateChooser.Value Then
                MateChooser.Value = ""
                MateName_Box.Value = ""
            End If
            RealName_Box.Value = WorkersTree.SelectedItem.Text
            NameChooser.Value = WorkersTree.SelectedItem.Key
            Amount_Box.Value = ""
            WorkersTreeHolder.Visible = False
            DateAndWorker_Frame.Visible = True
            JobsTree.Visible = True
            JobsTree.SetFocus
        Else
            ObjectsRecall
        End If
     Else
        If WorkersTree.SelectedItem.Key <> NameChooser.Value Then
            MateName_Box.Value = WorkersTree.SelectedItem.Text
            MateChooser.Value = WorkersTree.SelectedItem.Key
            CopyDay_Button.SetFocus
        End If
        WorkersTreeHolder.Visible = False
        DayList.Visible = True
     End If
End If

Exit Sub
ExceptionControl:
Exception.Error_Box.Value = "Workers/WorkersTree_DblClick()"
Exception.Show
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
BlockIt.Pass = PinAdmin
BlockIt.PassOK = False
BlockIt.AdminOverrides = False
BlockIt.SupervisorOverrides = False
BlockIt.Password_Box.SetFocus
BlockIt.Show
If BlockIt.PassOK = False And CloseMode = 0 Then Cancel = 1
End Sub

