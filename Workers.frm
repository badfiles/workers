VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Workers 
   ClientHeight    =   11025
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
Dim ChosenMate As Integer
Dim TestNode As Node
Const InfoOffset = 6
Const Lines = 9
Sub ObjectsRecall()
WorkersTreeHolder.Visible = False
JobsTree.Visible = True
DayList.Visible = True
DateAndWorker_Frame.Visible = True
End Sub
Sub ScanWorkers()
On Error GoTo ExceptionControl:
Workers.WorkersTreeHolder.Top = -500
Workers.WorkersTree.Visible = True
Workers.WorkersTreeHolder.Visible = True
Workers.WorkersTree.Nodes.Clear

Sheets("Каталог").Select
TotalCats = Cells(4, 23).Value
For i = InfoOffset To CInt(InfoOffset - 1 + TotalCats)
    Workers.WorkersTree.Nodes.Add(, , CStr(Cells(i, 24)) & "z", Cells(i, 23).Value).Sorted = True
    Workers.WorkersTree.Nodes(i - InfoOffset + 1).Tag = "Cat"
Next
  
Sheets("Сотрудники").Select
WeHaveWorkers = Cells(1, 2).Value
For i = 3 To WeHaveWorkers + 2
    For p = 1 To TotalCats
        If (Cells(i, 4).Value <> 1) And (Workers.WorkersTree.Nodes(p).Key = CStr(Cells(i, 6).Value) & "z") Then
            If (SelectUpdatesOnly And Cells(i, 1).Value = 1) Or Not SelectUpdatesOnly Then
                Workers.WorkersTree.Nodes.Add(p, 4, Cells(i, 3), Cells(i, 2) & " " & Cells(i, 5)).Sorted = True
                If Not AdminMode Then Workers.WorkersTree.Nodes(Workers.WorkersTree.Nodes.Count).Tag = Cells(i, 7)
                p = TotalCats
            End If
        End If
    Next p
Next i
 
p = 1
Do While p < Workers.WorkersTree.Nodes.Count
    If Workers.WorkersTree.Nodes(p).Children = 0 And Workers.WorkersTree.Nodes(p).Tag = "Cat" Then
        Workers.WorkersTree.Nodes.Remove (p)
        p = p - 1
        TotalCats = TotalCats - 1
    End If
    p = p + 1
Loop
Workers.WorkersTree.Tag = TotalCats
Workers.WorkersTreeHolder.Visible = False

Exit Sub
ExceptionControl:
ErrorForm.Error_Box.Value = "Workers/ScanWorkers()"
ErrorForm.Show
End Sub
Sub ScanJobs()
On Error GoTo ExceptionControl:
Sheets("Каталог").Select
Workers.JobsTree.Visible = True

Workers.JobsTree.Nodes.Clear
Workers.BonusRate_Box.Value = Cells(4, 6).Value
TotalCats = Cells(4, 19).Value
For i = InfoOffset To CInt(InfoOffset - 1 + TotalCats)
    Workers.JobsTree.Nodes.Add(, , , Cells(i, 19).Value).Sorted = True
    Workers.JobsTree.Nodes(i - InfoOffset + 1).Tag = "Cat"
Next

TotalJobs = Cells(4, 2).Value
ShowRates = Cells(4, 5).Value
For i = InfoOffset To CInt(InfoOffset - 1 + TotalJobs)
    AddRate = ""
    If ShowRates = 1 Then
        If Cells(i, 5) = 0 Then AddRate = "  (" & CStr(Cells(i, 6)) & ")" _
            Else AddRate = "  (" & CStr(Cells(i, 5)) & ")"
    End If
  
    If Not AdminMode Then
        If (Cells(i, 7) = 0) And (Cells(i, 9) = 0) Then _
            Workers.JobsTree.Nodes.Add(CInt(Cells(i, 1).Value - InfoOffset + 1), 4, _
            CStr(Cells(i, 3)) & "z", Cells(i, 2).Value & AddRate).Sorted = True
    Else
        If (Cells(i, 7) = 0) Then _
            Workers.JobsTree.Nodes.Add(CInt(Cells(i, 1).Value - InfoOffset + 1), 4, _
            CStr(Cells(i, 3)) & "z", Cells(i, 2).Value & AddRate).Sorted = True
    End If
Next

p = 1
Do While p < Workers.JobsTree.Nodes.Count
    If Workers.JobsTree.Nodes(p).Children = 0 And Workers.JobsTree.Nodes(p).Tag = "Cat" Then
        Workers.JobsTree.Nodes.Remove (p)
        TotalCats = TotalCats - 1
        p = p - 1
    End If
    p = p + 1
Loop
Workers.JobsTree.Tag = TotalCats


Exit Sub
ExceptionControl:
ErrorForm.Error_Box.Value = "Workers/ScanJobs()"
ErrorForm.Show
End Sub

Sub FillControlBox()
On Error GoTo ExceptionControl:
Sheets(NameChooser.Value).Activate
ControlList.ListItems.Clear
   
For i = InfoOffset To 276 Step Lines
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
        ControlList.ListItems.Add = Dat
        ControlList.ListItems.Item(ControlList.ListItems.Count).ListSubItems.Add = Fee
        ControlList.ListItems.Item(ControlList.ListItems.Count).ListSubItems.Add = Pre
        ControlList.ListItems.Item(ControlList.ListItems.Count).ListSubItems.Add = Comment
    End If
Next i

Exit Sub
ExceptionControl:
ErrorForm.Error_Box.Value = "Workers/FillControlBox()"
ErrorForm.Show
End Sub

Sub ReadLockedInfo()
On Error GoTo ExceptionControl:
Sheets(NameChooser.Value).Select

Left_Box.Value = Cells(2, 10).Value
Income_Box.Value = Cells(3, 10).Value
Outcome_Box.Value = Cells(3, 11).Value
Balance_Box.Value = Cells(1, 10).Value
Oklad_Box.Value = Cells(4, 2).Value

If Cells(3, 1).Value = "RO" Then
    MakeReadOnly_Chk.Value = True
    If Not AdminMode Then
        Apply_Button.Enabled = False
        Clear_Button.Enabled = False
        Delete_Button.Enabled = False
        ChooseMate_Button.Enabled = False
    End If
Else
    MakeReadOnly_Chk.Value = False
    If Not AdminMode And LastMonth_Label.Visible = False Then
        Apply_Button.Enabled = True
        Clear_Button.Enabled = True
        Delete_Button.Enabled = True
        ChooseMate_Button.Enabled = True
    End If
End If

If Oklad_Box <> "" Then AboveOklad_Chk.Visible = True Else AboveOklad_Chk.Visible = False
If Balance_Box.Value >= 0 Then Balance_Label.ForeColor = &H8000& Else Balance_Label.ForeColor = &HFF&

Exit Sub
ExceptionControl:
ErrorForm.Error_Box.Value = "Workers/ReadLockedInfo()"
ErrorForm.Show
End Sub

Sub SetRandomMark()
On Error GoTo ExceptionControl:
Randomize
Cells(2, 1).Value = Round(100000000 * Rnd())

Exit Sub
ExceptionControl:
ErrorForm.Error_Box.Value = "Workers/SetRandomMark()"
ErrorForm.Show
End Sub


Sub RecordInfo(ByVal Day, ByVal Job)
On Error GoTo ExceptionControl:
If NameChooser.Value = "" Then Exit Sub
Sheets(NameChooser.Value).Select
index = Job + InfoOffset + Lines * (Day - 1) - 1
If (ID.Value <> "") And DayTrigger Then
    Cells(index - Job + 1, 10).FormulaR1C1 = "=SUM(RC[-1]:R[8]C[-1])"
    Cells(index, 9).FormulaR1C1 = "=(RC[-5]*(1-RC[-1])+RC[-3]*RC[-1])*RC[-2]"
    'OldAmount = Cells(index, 4)
    'OldId = Cells(index, 3)
    'OldTime = Cells(index, 6)

    If (Len(AltDiam_Box.Value) > 2) And (Len(AltDiam_Box.Value) < 5) Then
       JobNameWithoutExDiam = Mid(JobName_Box.Value, 1, Len(JobName_Box.Value) - 3)
    
       If (Right(JobNameWithoutExDiam, 1) = "x") Or (Right(JobNameWithoutExDiam, 1) = "х") Then _
           Else JobNameWithoutExDiam = Mid(JobName_Box.Value, 1, Len(JobName_Box.Value) - 4)
        
       Cells(index, 2).Value = JobNameWithoutExDiam & AltDiam_Box.Value
       Cells(index, 14).Value = AltDiam_Box.Value
        
    Else
        
       Cells(index, 2).Value = JobName_Box.Value
       Cells(index, 14).Value = ""
        
    End If
    
    Cells(index, 3).Value = ID.Value
    If Amount_Box.Value <> Application.DecimalSeparator And Amount_Box.Value <> "-" And Amount_Box.Value <> "-" & Application.DecimalSeparator Then Cells(index, 4).Value = Amount_Box.Value
    Cells(index, 5).Value = Unit.Caption
    If Time_Box.Value <> Application.DecimalSeparator And Time_Box.Value <> "-" And Time_Box.Value <> "-" & Application.DecimalSeparator Then Cells(index, 6).Value = Time_Box.Value
    
    If Not AdminMode Then SetRandomMark
    
    If Oklad_Box = "" Or AboveOklad_Chk.Value = True Then
        If Rate_Box.Value <> Application.DecimalSeparator And Rate_Box.Value <> "-" And Rate_Box.Value <> "-" & Application.DecimalSeparator Then Cells(index, 7).Value = Rate_Box.Value
    Else
        Cells(index, 7).ClearContents
    End If

    If Rate_Box.Tag = "Time" Then Cells(index, 8).Value = 1
    If Rate_Box.Tag = "Amt" Then Cells(index, 8).Value = 0

    Cells(index, 2).Select
    
    If (Cells(index, 2).Value = "") Then _
        Selection.EntireRow.Hidden = True _
        Else Selection.EntireRow.Hidden = False
       
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
End If
      
If CommentTrigger And Comment_Box.Value <> "" Then
    For i = index - Job + 1 To index - Job + Lines
        If (Cells(i, 13).Value = "") Or (i = index - Job + Lines) Then
            Cells(i, 13).Value = Comment_Box.Value
            Exit For
        End If
    Next i
    CommentTrigger = False
    If Not AdminMode Then SetRandomMark
End If

If PrePay_Box.Value <> Application.DecimalSeparator And PrePay_Box.Value <> "-" And PrePay_Box.Value <> "-" & Application.DecimalSeparator Then Cells(index - Job + 1, 11).Value = PrePay_Box.Value
If Left_Box.Value <> Application.DecimalSeparator And Left_Box.Value <> "-" And Left_Box.Value <> "-" & Application.DecimalSeparator Then Cells(2, 10).Value = Left_Box.Value
If Oklad_Box.Value <> Application.DecimalSeparator And Oklad_Box.Value <> "-" And Oklad_Box.Value <> "-" & Application.DecimalSeparator Then Cells(4, 2).Value = Oklad_Box.Value
If MakeReadOnly_Chk.Value = True Then Cells(3, 1).Value = "RO" Else Cells(3, 1).Value = ""

Cells(index - Job + 1, 2).Select
If (Cells(index - Job + 1, 2).Value = "") And _
   (Cells(index - Job + 1, 11).Value = "") And _
   (Cells(index - Job + 1, 13).Value = "") Then _
    Selection.EntireRow.Hidden = True _
    Else Selection.EntireRow.Hidden = False

If CInt(Day) > CInt(Cells(1, 1).Value) Then Cells(1, 1).Value = Day
ReadLockedInfo
FillDayList (CDay_Box.Value)
MakeShitLookGood
If LMMode Then TransferBalance NameChooser.Value, Balance_Box.Value
DayTrigger = False
FillControlBox

Exit Sub
ExceptionControl:
ErrorForm.Error_Box.Value = "Workers/RecordInfo()"
ErrorForm.Show
End Sub
Sub DeleteInfo(ByVal Day, ByVal Job)
On Error GoTo ExceptionControl:
If NameChooser.Value = "" Then Exit Sub
index = Job + InfoOffset + Lines * (Day - 1) - 1
Sheets(NameChooser.Value).Select

'OldAmount = Cells(index, 4)
'OldId = Cells(index, 3)
'OldTime = Cells(index, 6)
'Sheets("Каталог").Select
'If OldId > 5 Then Cells(OldId, 11) = Cells(OldId, 11) - OldAmount
'Sheets(NameChooser.Value).Select

Range(Cells(index, 2), Cells(index, 9)).ClearContents
Cells(index, 14).ClearContents
If Not AdminMode Then Cells(index, 3) = 4
Cells(index, 2).Select
If (Cells(index, 2).Value = "") And _
    (Cells(index, 11).Value = "") And _
    (Cells(index, 13).Value = "") Then Selection.EntireRow.Hidden = True
ReadLockedInfo
JobName_Box.Value = ""
ID = ""
Time_Box.Value = ""
Amount_Box.Value = ""
Rate_Box.Value = ""
Rate_Box.Tag = ""
Unit.Caption = ""
MakeShitLookGood
If LMMode Then TransferBalance NameChooser.Value, Balance_Box.Value
FillDayList (CDay_Box.Value)
FillControlBox

Exit Sub
ExceptionControl:
ErrorForm.Error_Box.Value = "Workers/DeleteInfo()"
ErrorForm.Show
End Sub

Sub ClearDay(ByVal Day)
On Error GoTo ExceptionControl:
If NameChooser.Value = "" Then Exit Sub
index = InfoOffset + Lines * (Day - 1)
Sheets(NameChooser.Value).Select
If CInt(Cells(1, 1)) = Day Then Cells(1, 1).ClearContents
For i = 0 To Lines - 1
    If (AdminMode) And (Cells(index + i, 3) <> "") Then Cells(index + i, 3) = 4
    'OldAmount = Cells(index + i, 4)
    'OldId = Cells(index + i, 3)
    'OldTime = Cells(index + i, 6)
    'If OldId > 5 Then
    'Sheets("Каталог").Select
    'Cells(OldId, 11) = Cells(OldId, 11) - OldAmount
    'Sheets(NameChooser.Value).Select
    'End If
Next
'Sheets(NameChooser.Value).Select
If (AdminMode) Then
    Range(Cells(index, 2), Cells(index + Lines - 1, 2)).ClearContents
    Range(Cells(index, 4), Cells(index + Lines - 1, 9)).ClearContents
Else
    Range(Cells(index, 2), Cells(index + Lines - 1, 9)).ClearContents
End If
Range(Cells(index, 13), Cells(index + Lines - 1, 14)).ClearContents
Cells(index, 11) = ""
Cells(index, 10) = ""

Range(Cells(index, 2), Cells(index + Lines - 1, 2)).Select
Selection.EntireRow.Hidden = True

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
FillControlBox

Exit Sub
ExceptionControl:
ErrorForm.Error_Box.Value = "Workers/ClearDay()"
ErrorForm.Show
End Sub
Sub ReadLine(ByVal Day, ByVal Job)
On Error GoTo ExceptionControl:
If Day <> "" Then
    inRead = True
    index = Job + InfoOffset + Lines * (Day - 1) - 1

    Sheets(NameChooser.Value).Select

    If Cells(index, 3) > 4 Then ID.Value = Cells(index, 3) Else ID.Value = ""
    JobName_Box.Value = Cells(index, 2).Value
    Rate_Box.Value = Cells(index, 7).Value
    Rate_Box.Tag = ""
    Time_Box.Value = Cells(index, 6).Value
    Unit.Caption = Cells(index, 5).Value
    Amount_Box.Value = Cells(index, 4).Value
    AltDiam_Box.Value = Cells(index, 14).Value

    If Unit.Caption = "" Then Amount_Box.Enabled = False Else Amount_Box.Enabled = True
    If Oklad_Box.Value <> "" And Rate_Box.Value <> "" Then _
        AboveOklad_Chk.Value = True
    If Oklad_Box.Value <> "" And Rate_Box.Value = "" Then _
        AboveOklad_Chk.Value = False

    If Amount_Box.Enabled = False Then Time_Box.SetFocus Else Amount_Box.SetFocus
    If JobName_Box.Value = "" Then JobsTree.SetFocus
    inRead = False
    DayTrigger = False
End If

Exit Sub
ExceptionControl:
ErrorForm.Error_Box.Value = "Workers/ReadLine()"
ErrorForm.Show
End Sub
Function LastFilled(ByVal Day)
On Error GoTo ExceptionControl:
If Day <> "" Then
    index = InfoOffset + Lines * (Day - 1)
    LastFilled = 0
    For i = 1 To Lines
        If Cells(index + i - 1, 2).Value <> "" Then LastFilled = i
    Next
End If

Exit Function
ExceptionControl:
ErrorForm.Error_Box.Value = "Workers/LastFilled()"
ErrorForm.Show
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
index = InfoOffset + Lines * (CDay_Box.Value - 1)
JobName_Box.Value = "Доплата " & CStr(BonusRate_Box.Value) & " %"
Amount_Box.Value = 1
Unit.Caption = " "
Rate_Box.Tag = "Amt"
Rate_Box.Value = Cells(index, 10).Value * BonusRate_Box.Value / 100
Time_Box.Value = ""
ID.Value = 5
If Apply_Button.Enabled = True Then Apply_Button.SetFocus
inRead = False

Exit Sub
ExceptionControl:
ErrorForm.Error_Box.Value = "Workers/Bonus_Button_Click()"
ErrorForm.Show
End Sub

Private Sub BonusRate_Box_Change()
If BonusRate_Box.Value <> "" Then BonusRate_Box.Value = PointFilter(BonusRate_Box.Value, False, False, 3)
End Sub



Private Sub CollapseJobs_Button_Click()
p = 1
Do While p < JobsTree.Nodes.Count
    JobsTree.Nodes(p).Expanded = False
    p = p + 1
Loop
End Sub

Private Sub Comment_Box_Change()
If Not inRead Then CommentTrigger = True
End Sub

Private Sub Comment_Box_Click()
ObjectsRecall
End Sub

Private Sub Comment_Box_DropButtonClick()
ObjectsRecall
End Sub

Private Sub Delete_Button_Click()
ObjectsRecall
DeleteInfo CDay_Box.Value, CJob_Box.Value
End Sub

Private Sub Div_Button_Click()
On Error GoTo ExceptionControl:
ObjectsRecall
If Amount_Box.Value <> "" And Amount_Box.Value <> 0 And Amount_Box.Value <> "-" And Amount_Box.Value <> Application.DecimalSeparator And Amount_Box.Value <> "-" & Application.DecimalSeparator Then
    Amount_Box.Value = Round(Amount_Box.Value / 2, 2)
End If
If Apply_Button.Enabled = True Then Apply_Button.SetFocus

Exit Sub
ExceptionControl:
ErrorForm.Error_Box.Value = "Workers/Div_Button_Click()"
ErrorForm.Show
End Sub

Private Sub LastMonth_Label_Click()
ObjectsRecall
End Sub



Private Sub Triv_Button_Click()
On Error GoTo ExceptionControl:
ObjectsRecall
If Amount_Box.Value <> "" And Amount_Box.Value <> 0 And Amount_Box.Value <> "-" And Amount_Box.Value <> Application.DecimalSeparator And Amount_Box.Value <> "-" & Application.DecimalSeparator Then
    Amount_Box.Value = Round(Amount_Box.Value / 3, 2)
End If
If Apply_Button.Enabled = True Then Apply_Button.SetFocus

Exit Sub
ExceptionControl:
ErrorForm.Error_Box.Value = "Workers/Triv_Button_Click()"
ErrorForm.Show
End Sub

Private Sub SelectUpdatesOnly_Change()
ObjectsRecall
ScanWorkers
End Sub


Private Sub ControlList_DblClick()
On Error GoTo ExceptionControl:
ObjectsRecall
If ControlList.ListItems.Count > 0 Then
    If ControlList.SelectedItem.Text <> "" Then CDay_Box.Value = CInt(ControlList.SelectedItem.Text)
End If

Exit Sub
ExceptionControl:
ErrorForm.Error_Box.Value = "Workers/ControlList_DblClick()"
ErrorForm.Show
End Sub

Private Sub DayList_DblClick()
On Error GoTo ExceptionControl:
ObjectsRecall
If DayList.ListItems.Count > 0 Then
 
    If DayList.SelectedItem.Text = "" Then Exit Sub

    If DayList.SelectedItem.Text = " " Then _
            CJob_Box.Value = CInt(DayList.ListItems.Count - 1) _
            Else CJob_Box.Value = CInt(DayList.SelectedItem.Text)
End If

Exit Sub
ExceptionControl:
ErrorForm.Error_Box.Value = "Workers/DayList_DblClick()"
ErrorForm.Show
End Sub

Sub FillDayList(ByVal Day)
On Error GoTo ExceptionControl:

If NameChooser.Value <> "" Then Sheets(NameChooser.Value).Select
Records = LastFilled(Day)
DayList.ListItems.Clear
Comment_Box.Clear
Comment_Box.Value = ""
CommentTrigger = False
If NameChooser.Value <> "" And Day <> "" And Records <> 0 Then
    TotalTime = 0
    
    For i = 1 To Lines
        index = i + InfoOffset + Lines * (Day - 1) - 1
        If i <= Records Then
            Position = i
            JobName = Cells(index, 2)
            Amount = Cells(index, 4)
            UnitList = Cells(index, 5)
            TimeList = Cells(index, 6)
            TotalTime = TotalTime + TimeList
            RateList = Cells(index, 7)
            Subtotal = Cells(index, 9)
            DayList.ListItems.Add = Position
            DayList.ListItems.Item(i).ListSubItems.Add = JobName
            DayList.ListItems.Item(i).ListSubItems.Add = Amount
            DayList.ListItems.Item(i).ListSubItems.Add = UnitList
            DayList.ListItems.Item(i).ListSubItems.Add = TimeList
            DayList.ListItems.Item(i).ListSubItems.Add = RateList
            DayList.ListItems.Item(i).ListSubItems.Add = Subtotal
        End If
        If Cells(index, 13).Value <> "" Then Comment_Box.AddItem (Cells(index, 13).Value)
    Next i
    
    index = InfoOffset + Lines * (Day - 1)
 
    If DayList.ListItems.Count > 0 Then
        CJob_Box.Value = DayList.ListItems.Count + 1
        DayList.ListItems.Add = " "
        DayList.ListItems.Add = " "
        DayList.ListItems.Item(DayList.ListItems.Count).ListSubItems.Add = ""
        DayList.ListItems.Item(DayList.ListItems.Count).ListSubItems.Add = ""
        DayList.ListItems.Item(DayList.ListItems.Count).ListSubItems.Add = "ВСЕГО"
        DayList.ListItems.Item(DayList.ListItems.Count).ListSubItems.Add = TotalTime
        DayList.ListItems.Item(DayList.ListItems.Count).ListSubItems.Add = "ИТОГО"
        DayList.ListItems.Item(DayList.ListItems.Count).ListSubItems.Add = Cells(index, 10).Value
    End If
End If

If Records < Lines Then CJob_Box.Value = Records + 1
If Records > Lines - 1 Then CJob_Box.Value = Lines

If Day <> "" Then
    index = InfoOffset + Lines * (Day - 1)
    If Comment_Box.ListCount > 0 Then
        Comment_Box.Value = Comment_Box.List(Comment_Box.ListCount - 1)
        CommentTrigger = False
    End If
    PrePay_Box.Value = Cells(index, 11).Value
End If

Exit Sub
ExceptionControl:
ErrorForm.Error_Box.Value = "Workers/FillDayList()"
ErrorForm.Show
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
If ChosenMate <> 0 Then
    WorkersTree.Nodes(ChosenMate).Selected = True
    ChosenMate = 0
End If
WorkersTree.SetFocus

Exit Sub
ExceptionControl:
ErrorForm.Error_Box.Value = "Workers/ChooseWorker_Click()"
ErrorForm.Show
End Sub
Private Sub ChooseMate_Button_Click()
On Error Resume Next
ObjectsRecall
WorkersTreeHolder.Top = 255
WorkersTreeHolder.Left = 250
'For i = 1 To CInt(WorkersTree.Tag)
'  If WorkersTree.Nodes(i).Tag = "Cat" Then _
'        WorkersTree.Nodes(i).Expanded = False
' Next
'WorkersTree.Visible = True
WorkersTreeHolder.Visible = True
ChosenMate = WorkersTree.SelectedItem.index
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
ErrorForm.Error_Box.Value = "Workers/JobName_Box_Change()"
ErrorForm.Show
End Sub

Private Sub JobsTree_DblClick()
ObjectsRecall
If JobsTree.SelectedItem.Key <> "" Then
    'JobsTreeHolder.Visible = False
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
index = InfoOffset + Lines * Day
If Cells(index, 10) <> 0 Then
    isVisible = True
    Exit Function
End If
If Cells(index, 11) > 0 Then
    isVisible = True
    Exit Function
End If
If Cells(index, 13) <> "" Then
    isVisible = True
    Exit Function
End If

Exit Function
ExceptionControl:
ErrorForm.Error_Box.Value = "Workers/isVisible()"
ErrorForm.Show
End Function

Sub Mark(ByVal Day As Integer, ByVal PrevMarked As Boolean)
On Error GoTo ExceptionControl:
If PrevMarked Then Colorr = 2 Else Colorr = 15
index = InfoOffset + Lines * Day
Range(Cells(index, 1), Cells(index + Lines - 1, 12)).Select
With Selection.Interior
     .ColorIndex = Colorr
     .Pattern = xlSolid
     .PatternColorIndex = xlAutomatic
End With

Exit Sub
ExceptionControl:
ErrorForm.Error_Box.Value = "Workers/Mark()"
ErrorForm.Show
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
ErrorForm.Error_Box.Value = "Workers/MakeShitLookGood()"
ErrorForm.Show
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
        Workers.Hide
        Form.Hide
        Cells(3, 1).Select
    End If
End If
Exit Sub
ExceptionControl:
ErrorForm.Error_Box.Value = "Workers/Print_Button_Click()"
ErrorForm.Show
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
RecordInfo CDay_Box.Value, CJob_Box.Value
End Sub

Private Sub CDay_Box_Change()
On Error GoTo ExceptionControl:
ObjectsRecall
If NameChooser.Value <> "" Then
    FillDayList (CDay_Box.Value)

    Label_FullDate.Caption = GetDayName(CDay_Box.Value) & ", " & _
                            CDay_Box.Value & " " & MNameRusFix(CMonth)
                            
    Workers.Caption = RealName_Box.Value & ": " & Label_FullDate.Caption

    ReadLine CDay_Box.Value, CJob_Box.Value
End If

Exit Sub
ExceptionControl:
ErrorForm.Error_Box.Value = "Workers/CDay_Box_Change()"
ErrorForm.Show
End Sub
Private Sub CJob_Box_Change()
ReadLine CDay_Box.Value, CJob_Box.Value
End Sub
Private Sub Control_Button_Click()
FillControlBox
End Sub

Private Sub CopyDay_Button_Click()
On Error GoTo ExceptionControl:
ObjectsRecall
If NameChooser.Value <> "" And MateChooser.Value <> "" Then
    Sheets(MateChooser.Value).Select
    If Cells(3, 1).Value = "RO" And Not AdminMode Then
        Sheets(NameChooser.Value).Select
        b = MsgBox("Запись в лист сотрудника " & MateName_Box & " с рабочего места невозможна.", vbOKOnly, "Копирование отменено")
        Exit Sub
    End If
    For i = 1 To Lines
        index = i + InfoOffset + Lines * (CDay_Box.Value - 1) - 1
        If Cells(index, 2).Value <> "" Then
            If Not AdminMode Then
                Sheets(NameChooser.Value).Select
                b = MsgBox(MateName_Box & " уже записался на " & CDay_Box.Value & " " & MNameRusFix(CMonth) & ".", vbOKOnly, "Копирование отменено")
                Exit Sub
            Else
                InsureForm.Msg_label.Caption = "Все записи у " & MateName_Box & " за " & CDay_Box.Value & " " & MNameRusFix(CMonth) & " будут перезаписаны. Выполнить копирование?"
                InsureForm.NoButton.SetFocus
                InsureForm.Show
                If InsureForm.OK.Value = True Then
                    i = Lines + 1
                Else
                    Sheets(NameChooser.Value).Select
                    Exit Sub
                End If
            End If
        End If
    Next i
    
    For i = 1 To Lines
        Sheets(NameChooser.Value).Select
        index = i + InfoOffset + Lines * (CDay_Box.Value - 1) - 1
 
        If Cells(index, 2).Value <> "" Then
            'AddAmount = Cells(index, 4).Value
            'AddID = Cells(index, 3).Value
            Range(Cells(index, 2), Cells(index, 9)).Copy
            CopyAlternateDiam = Cells(index, 14).Value
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
            Cells(index, 2).PasteSpecial
            Cells(index, 14).Value = CopyAlternateDiam
            Cells(index, 2).Select
            Selection.EntireRow.Hidden = False
        End If
    Next i
    
    Sheets(MateChooser.Value).Select
    Cells(InfoOffset + 9 * (CDay_Box.Value - 1), 10).FormulaR1C1 = "=SUM(RC[-1]:R[8]C[-1])"

    If CDay_Box.Value > Cells(1, 1).Value Then _
    Cells(1, 1).Value = CDay_Box.Value

    If Not AdminMode Then SetRandomMark
    MakeShitLookGood
    If LMMode Then TransferBalance MateChooser.Value, Cells(1, 10).Value
Sheets(NameChooser.Value).Select
End If

Exit Sub
ExceptionControl:
ErrorForm.Error_Box.Value = "Workers/CopyDay_Button_Click()"
ErrorForm.Show
End Sub

Private Sub Clear_Button_Click()
On Error GoTo ExceptionControl:
ObjectsRecall
If Not AdminMode Then
    InsureForm.Msg_label.Caption = "Вы действительно хотите стереть все записи за " & CDay_Box.Value & " " & MNameRusFix(CMonth) & "?"
    InsureForm.NoButton.SetFocus
    InsureForm.Show
    If InsureForm.OK.Value = True Then ClearDay (CDay_Box.Value)
Else
    ClearDay (CDay_Box.Value)
End If

Exit Sub
ExceptionControl:
ErrorForm.Error_Box.Value = "Workers/Clear_Button_Click()"
ErrorForm.Show
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
ErrorForm.Error_Box.Value = "Workers/ID_Change()"
ErrorForm.Show
End Sub
Private Sub NameChooser_Change()
On Error GoTo ExceptionControl:
If NameChooser.Value <> "" Then
    Sheets(NameChooser.Value).Select
    FillControlBox
    ReadLockedInfo

    Workers.Caption = RealName_Box.Value & ": " & Label_FullDate.Caption

    FillDayList (CDay_Box.Value)
    
    If NameChooser.Value = MateChooser.Value Or Not AdminMode Then
        MateChooser.Value = ""
        MateName_Box.Value = ""
    End If
End If

Exit Sub
ExceptionControl:
ErrorForm.Error_Box.Value = "Workers/NameChooser_Change()"
ErrorForm.Show
End Sub



Private Sub Day_Spin_SpinDown()
On Error GoTo ExceptionControl:
ObjectsRecall
If CDay_Box.Value < MDays(CMonth) Then CDay_Box.Value = CDay_Box.Value + 1

Exit Sub
ExceptionControl:
ErrorForm.Error_Box.Value = "Workers/Day_Spin_SpinDown()"
ErrorForm.Show
End Sub

Private Sub Day_Spin_SpinUp()
On Error GoTo ExceptionControl:
ObjectsRecall
If CDay_Box.Value > 1 Then CDay_Box.Value = CDay_Box.Value - 1

Exit Sub
ExceptionControl:
ErrorForm.Error_Box.Value = "Workers/Day_Spin_SpinUp()"
ErrorForm.Show
End Sub
Private Sub Workers_Spin_SpinDown()
On Error Resume Next
ObjectsRecall
WorkersTreeHolder.Top = -500
WorkersTreeHolder.Visible = True

Total = WorkersTree.Nodes.Count
TotalCat = CInt(WorkersTree.Tag)
If ChosenMate <> 0 Then
    index = ChosenMate
    ChosenMate = 0
Else
    index = WorkersTree.SelectedItem.index
End If
If index < TotalCat Then index = TotalCat
If index <= Total - 1 Then
    If WorkersTree.Nodes(index + 1).Tag <> "Cat" Then
        RealName_Box.Value = WorkersTree.Nodes(index + 1).Text
        NameChooser.Value = WorkersTree.Nodes(index + 1).Key
        WorkersTree.Nodes(index + 1).Selected = True
    End If
End If
Endd:
WorkersTreeHolder.Visible = False
End Sub

Private Sub Workers_Spin_SpinUp()
On Error GoTo Endd:
ObjectsRecall
WorkersTreeHolder.Top = -500
WorkersTreeHolder.Visible = True

If ChosenMate <> 0 Then
    index = ChosenMate
    ChosenMate = 0
Else
    index = WorkersTree.SelectedItem.index
End If
If WorkersTree.Nodes(index - 1).Tag <> "Cat" Then
    RealName_Box.Value = WorkersTree.Nodes(index - 1).Text
    NameChooser.Value = WorkersTree.Nodes(index - 1).Key
    WorkersTree.Nodes(index - 1).Selected = True
Else
    WorkersTree.Nodes(index).Selected = True
End If
Endd:
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
    
        SaveClose (WorkersBase)
        PullBase = "pull.xls"
        Destination = Path & PullBase
        Source = Path & WorkersBase
        FileCopy Source, Destination
    
        ArcName = Path & "pull.7z"
        ArcFiles = Path & PullBase
        RunCommand (Archiver & " a -sdel " & ExchangeKey & " " & ArcName & " " & ArcFiles)
        a = Shell("ftp -v -s:" & Path & "ftp_client_send " & FtpStorageName, vbMinimizedNoFocus)
    End If
    Workers.Hide
Else
Workers.Hide
End If

Exit Sub
ExceptionControl:
ErrorForm.Error_Box.Value = "Workers/WorkersClose_Button_Click()"
ErrorForm.Show
End Sub

Private Sub Logout_Button_Click()
On Error GoTo ExceptionControl:
ObjectsRecall
RealName_Box.Value = ""
NameChooser.Value = "Образец"
NameChooser.Value = ""

Exit Sub
ExceptionControl:
ErrorForm.Error_Box.Value = "Workers/Logout_Button_Click()"
ErrorForm.Show
End Sub


Private Sub WorkersTree_DblClick()
On Error GoTo ExceptionControl:
If WorkersTree.SelectedItem.Key <> "" And WorkersTree.SelectedItem.Tag <> "Cat" Then
     
     If ChosenMate = 0 Then
     
        If Not AdminMode Then
            BlockIt.Pass = WorkersTree.SelectedItem.Tag
            BlockIt.PassOK = False
            BlockIt.Password_Box.SetFocus
            BlockIt.Show
        End If
     
        If (BlockIt.PassOK) Or (AdminMode) Then
     
            RealName_Box.Value = WorkersTree.SelectedItem.Text
            NameChooser.Value = WorkersTree.SelectedItem.Key
            Amount_Box.Value = ""
                    
            If WorkersTree.SelectedItem.Key = MateChooser.Value Then
                MateChooser.Value = ""
                MateName_Box.Value = ""
            End If
            
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
ErrorForm.Error_Box.Value = "Workers/WorkersTree_DblClick()"
ErrorForm.Show
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
BlockIt.Pass = PinAdmin
BlockIt.PassOK = False
BlockIt.Password_Box.SetFocus
BlockIt.Show
If BlockIt.PassOK = False And CloseMode = 0 Then Cancel = 1
End Sub

