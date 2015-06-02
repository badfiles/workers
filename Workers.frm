VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Workers 
   ClientHeight    =   11025
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   17115
   OleObjectBlob   =   "Workers.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Workers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Read As Integer
Dim PrevMarked, DetectChanges, WeChooseMate As Integer
Dim OldWCat, JCat0, JCat1, JCat2, JCat3 As String
Dim TestNode As Node
Const InfoOffset = 6
Const Lines = 9
Function ObjectsRecall()
WorkersTreeHolder.Visible = False
JobsTree.Visible = True
DayList.Visible = True
Frame7.Visible = True
End Function
Function ScanWorkers(Fil)
'On Error GoTo Start
Windows(Fil).Activate
Workers.WorkersTree.Visible = True
Workers.WorkersTreeHolder.Visible = True
Workers.WorkersTree.Nodes.Clear

Sheets("Каталог").Select
Total = Cells(4, 23).Value
For i = InfoOffset To CInt(InfoOffset - 1 + Total)
    Workers.WorkersTree.Nodes.Add(, , CStr(Cells(i, 24)) & "z", Cells(i, 23).Value).Sorted = True
    Workers.WorkersTree.Nodes(i - InfoOffset + 1).Tag = "Cat"
    
  Next
  
Sheets("Сотрудники").Select
  WeHaveWorkers = Cells(1, 2).Value
  For i = 3 To WeHaveWorkers + 2
     For p = 1 To Total
     If Cells(i, 4).Value <> 1 And Workers.WorkersTree.Nodes(p).Key = CStr(Cells(i, 6).Value) & "z" Then
         Workers.WorkersTree.Nodes.Add(p, 4, Cells(i, 3), _
                       Cells(i, 2) & " " & Cells(i, 5)).Sorted = True
      End If
     Next p
 Next i
 
 p = 1
 Do While p < Workers.WorkersTree.Nodes.Count
 
  If Workers.WorkersTree.Nodes(p).Children = 0 And Workers.WorkersTree.Nodes(p).Tag = "Cat" Then
         Workers.WorkersTree.Nodes.Remove (p)
         p = p - 1
         Total = Total - 1
     End If
  p = p + 1
 Loop
   

Workers.WorkersTree.Tag = Total
'Setup.WorkersTree.Visible = False
Workers.WorkersTreeHolder.Visible = False


GoTo Endd:
Start:
ErrorForm.Show
Endd:
End Function
Function ScanJobs()
Sheets("Каталог").Select
Workers.JobsTree.Visible = True
'Workers.JobsTreeHolder.Visible = True

Workers.JobsTree.Nodes.Clear
Workers.BonusRate_Box.Value = Cells(4, 6).Value
TotalCat = Cells(4, 19).Value
For i = InfoOffset To CInt(InfoOffset - 1 + TotalCat)
    Workers.JobsTree.Nodes.Add(, , , Cells(i, 19).Value).Sorted = True
    Workers.JobsTree.Nodes(i - InfoOffset + 1).Tag = "Cat"
    Next

Total = Cells(4, 2).Value
ShowRates = Cells(4, 5).Value
For i = InfoOffset To CInt(InfoOffset - 1 + Total)
  
  AddRate = ""
  If ShowRates = 1 Then
   If Cells(i, 5) = 0 Then AddRate = "  (" & CStr(Cells(i, 6)) & ")" _
   Else _
   AddRate = "  (" & CStr(Cells(i, 5)) & ")"
  End If
  
  If Cells(i, 7) = 0 Then _
  Workers.JobsTree.Nodes.Add _
  (CInt(Cells(i, 1).Value - InfoOffset + 1), 4, _
  CStr(Cells(i, 3)) & "z", Cells(i, 2).Value & AddRate).Sorted = True
Next

 p = 1
 Do While p < Workers.JobsTree.Nodes.Count
 
  If Workers.JobsTree.Nodes(p).Children = 0 And Workers.JobsTree.Nodes(p).Tag = "Cat" Then
         Workers.JobsTree.Nodes.Remove (p)
         TotalCat = TotalCat - 1
         p = p - 1
     End If
  p = p + 1
 Loop
Workers.JobsTree.Tag = TotalCat
'Workers.JobsTreeHolder.Visible = False

    
End Function




Function FillControlBox()

Sheets(NameChooser.Value).Activate
ControlList.ListItems.Clear
   
  For i = InfoOffset To 276 Step Lines
   If (Cells(i, 10).Value <> 0) Or (Cells(i, 11).Value <> 0) Or (Cells(i, 13).Value <> "") Then
     
    Dat = Cells(i, 1).Value
    Fee = Cells(i, 10).Value
    Pre = Cells(i, 11).Value
    Comment = Cells(i, 13).Value
    If Len(Dat) = 1 Then Dat = "0" & Dat
    
   ControlList.ListItems.Add = Dat
   ControlList.ListItems.Item(ControlList.ListItems.Count).ListSubItems.Add = Fee
   ControlList.ListItems.Item(ControlList.ListItems.Count).ListSubItems.Add = Pre
   ControlList.ListItems.Item(ControlList.ListItems.Count).ListSubItems.Add = Comment
   
   End If
  Next i
End Function

Function ReadLockedInfo()
'On Error GoTo Start
Sheets(NameChooser.Value).Select

Left_Box.Value = Cells(2, 10).Value
Income_Box.Value = Cells(3, 10).Value
Outcome_Box.Value = Cells(3, 11).Value
Balance_Box.Value = Cells(1, 10).Value
Oklad_Box.Value = Cells(4, 2).Value
If Oklad_Box <> "" Then AboveOklad_Chk.Visible = True _
Else AboveOklad_Chk.Visible = False

If Balance_Box.Value >= 0 Then Balance_Label.ForeColor = &H8000&
If Balance_Box.Value < 0 Then Balance_Label.ForeColor = &HFF&

GoTo Endd:
Start:
ErrorForm.Error_Box.Value = "RecordLockedInfo()"
ErrorForm.Show
Endd:
End Function
Function RecordInfo(Day, Job)
''On Error GoTo Start
Index = Job + InfoOffset + Lines * (Day - 1) - 1
Sheets(NameChooser.Value).Select
If (ID.Value <> "") And (DetectChanges = 1) Then


Cells(Index - Job + 1, 10).FormulaR1C1 = "=SUM(RC[-1]:R[8]C[-1])"

Cells(Index, 9).FormulaR1C1 = "=(RC[-5]*(1-RC[-1])+RC[-3]*RC[-1])*RC[-2]"


OldAmount = Cells(Index, 4)
OldId = Cells(Index, 3)
OldTime = Cells(Index, 6)

Cells(Index, 2).Value = JobName_Box.Value
Cells(Index, 3).Value = ID.Value
Cells(Index, 4).Value = Amount_Box.Value
Cells(Index, 5).Value = Unit.Caption
Cells(Index, 6).Value = Time_Box.Value

If Oklad_Box = "" Or AboveOklad_Chk.Value = True Then _
Cells(Index, 7).Value = Rate_Box.Value _
Else Cells(Index, 7).ClearContents

If Rate_Box.Tag = "Time" Then Cells(Index, 8).Value = 1
If Rate_Box.Tag = "Amt" Then Cells(Index, 8).Value = 0

Cells(Index, 2).Select
 If (Cells(Index, 2).Value = "") Then _
     Selection.EntireRow.Hidden = True _
 Else Selection.EntireRow.Hidden = False
       
AddAmount = Amount_Box.Value
AddTime = Time_Box.Value
If AddAmount = "" Then AddAmount = 0
If AddTime = "" Then AddTime = 0
If OldId = "" Then OldId = 0
 
Sheets("Каталог").Select

If OldId <> 0 Then Cells(OldId, 11) = Cells(OldId, 11) - OldAmount
'Cells(JobPosition(OldId) + 287, 6) = Cells(JobPosition(OldId) + 287, 6) - OldTime
'End If

Cells(ID.Value, 11) = Cells(ID.Value, 11) + AddAmount
'Cells(JobPosition(ID.Value) + 287, 6) = Cells(JobPosition(ID.Value) + 287, 6) + AddTime
Sheets(NameChooser.Value).Select

End If
      
Sheets(NameChooser.Value).Select
      
        
Cells(Index - Job + 1, 13).Value = Comment_Box.Value
Cells(Index - Job + 1, 11).Value = PrePay_Box.Value
Cells(2, 10).Value = Left_Box.Value
Cells(4, 2).Value = Oklad_Box.Value

Cells(Index - Job + 1, 2).Select
 If (Cells(Index - Job + 1, 2).Value = "") And _
    (Cells(Index - Job + 1, 11).Value = "") And _
    (Cells(Index - Job + 1, 13).Value = "") Then _
        Selection.EntireRow.Hidden = True _
 Else Selection.EntireRow.Hidden = False
If CInt(Day) > CInt(Cells(1, 1)) Then Cells(1, 1) = Day
ReadLockedInfo
FillDayList (CDay_Box.Value)

a = TransferBalanceToNextMonth(NameChooser.Value, Balance_Box.Value)

DetectChanges = 0
FillControlBox
GoTo Endd:
Start:
ErrorForm.Error_Box.Value = "RecordInfo()"
ErrorForm.Show
Endd:
End Function
Function DeleteInfo(Day, Job)
'On Error GoTo Start

Index = Job + InfoOffset + Lines * (Day - 1) - 1
Sheets(NameChooser.Value).Select

OldAmount = Cells(Index, 4)
OldId = Cells(Index, 3)
OldTime = Cells(Index, 6)


Sheets("Каталог").Select

If OldId > 5 Then Cells(OldId, 11) = Cells(OldId, 11) - OldAmount

Sheets(NameChooser.Value).Select



Range(Cells(Index, 2), Cells(Index, 9)).ClearContents

Cells(Index, 2).Select
If (Cells(Index, 2).Value = "") And (Cells(Index, 11).Value = "") And (Cells(Index, 13).Value = "") Then _
                                                                        Selection.EntireRow.Hidden = True

ReadLockedInfo
JobName_Box.Value = ""
ID = ""
Time_Box.Value = ""
Amount_Box.Value = ""
Rate_Box.Value = ""
Rate_Box.Tag = ""
Unit.Caption = ""
a = TransferBalanceToNextMonth(NameChooser.Value, Balance_Box.Value)
FillDayList (CDay_Box.Value)
FillControlBox
GoTo Endd:
Start:
ErrorForm.Error_Box.Value = "DeleteInfo()"
ErrorForm.Show
Endd:
End Function

Function ClearDay(Day)
'On Error GoTo Start

Index = InfoOffset + Lines * (Day - 1)
Sheets(NameChooser.Value).Select
 If CInt(Cells(1, 1)) = Day Then Cells(1, 1).ClearContents

For i = 0 To Lines - 1
OldAmount = Cells(Index + i, 4)
OldId = Cells(Index + i, 3)
OldTime = Cells(Index + i, 6)


If OldId <> 0 Then
Sheets("Каталог").Select
Cells(OldId, 11) = Cells(OldId, 11) - OldAmount
Sheets(NameChooser.Value).Select
End If


Next

Sheets(NameChooser.Value).Select

 Range(Cells(Index, 2), Cells(Index + Lines - 1, 9)).ClearContents
 Cells(Index, 11) = ""
 Cells(Index, 10) = ""

Range(Cells(Index, 2), Cells(Index + Lines - 1, 2)).Select
Selection.EntireRow.Hidden = True



ReadLockedInfo
JobName_Box.Value = ""
ID = ""
Time_Box.Value = ""
Amount_Box.Value = ""
Rate_Box.Value = ""
Unit.Caption = ""
a = TransferBalanceToNextMonth(NameChooser.Value, Balance_Box.Value)

FillDayList (CDay_Box.Value)
FillControlBox
GoTo Endd:
Start:
ErrorForm.Error_Box.Value = "ClearDay()"
ErrorForm.Show
Endd:
End Function
Function ReadLine(Day, Job)
'On Error GoTo Start
If Day <> "" Then
Read = 1
Index = Job + InfoOffset + Lines * (Day - 1) - 1

Sheets(NameChooser.Value).Select

ID.Value = Cells(Index, 3)
JobName_Box.Value = Cells(Index, 2).Value
Rate_Box.Value = Cells(Index, 7).Value
Rate_Box.Tag = ""
Time_Box.Value = Cells(Index, 6).Value
Unit.Caption = Cells(Index, 5).Value
Amount_Box.Value = Cells(Index, 4).Value


If Unit.Caption = "" Then Amount_Box.Enabled = False
If Unit.Caption <> "" Then Amount_Box.Enabled = True
If Oklad_Box.Value <> "" And Rate_Box.Value <> "" Then _
AboveOklad_Chk.Value = True
If Oklad_Box.Value <> "" And Rate_Box.Value = "" Then _
AboveOklad_Chk.Value = False

If Amount_Box.Enabled = False Then Time_Box.SetFocus Else Amount_Box.SetFocus
If JobName_Box.Value = "" Then JobsTree.SetFocus

Read = 0
DetectChanges = 0
End If
GoTo Endd:
Start:
ErrorForm.Error_Box.Value = "ReadLine()"
ErrorForm.Show
Endd:
End Function
Function LastFilled(Day)
If Day <> "" Then
Index = InfoOffset + Lines * (Day - 1)
LastFilled = 0
 For i = 1 To Lines
  If Cells(Index + i - 1, 2).Value <> "" Then LastFilled = i
 Next
End If
End Function

Private Sub Bonus_Button_Click()
Read = 1
Sheets(NameChooser.Value).Select
Index = InfoOffset + Lines * (CDay_Box.Value - 1)

JobName_Box.Value = "Доплата " & CStr(BonusRate_Box.Value) & " %"
Amount_Box.Value = 1
Unit.Caption = " "
Rate_Box.Tag = "Amt"
Rate_Box.Value = Cells(Index, 10).Value * BonusRate_Box.Value / 100
Time_Box.Value = ""
ID.Value = 5
Apply_Button.SetFocus
'a = RecordInfo(CDay_Box.Value, CJob_Box.Value)
Read = 0
End Sub

Private Sub BonusRate_Box_Change()
BonusRate_Box.Value = PointFilter(BonusRate_Box.Value)
End Sub


Private Sub Delete_Button_Click()
a = DeleteInfo(CDay_Box.Value, CJob_Box.Value)
End Sub

Private Sub Div_Button_Click()
If Amount_Box.Value <> "" And Amount_Box.Value <> 0 Then
Amount_Box.Value = Amount_Box.Value / 2
End If
Apply_Button.SetFocus
End Sub

Private Sub ControlList_DblClick()
If ControlList.ListItems.Count > 0 Then
 
 If ControlList.SelectedItem.Text <> "" Then CDay_Box.Value = CInt(ControlList.SelectedItem.Text)

End If
End Sub

Private Sub DayList_DblClick()
If DayList.ListItems.Count > 0 Then
 
If DayList.SelectedItem.Text = "" Then GoTo Endd

If DayList.SelectedItem.Text = " " Then _
            CJob_Box.Value = CInt(DayList.ListItems.Count - 1) _
            Else CJob_Box.Value = CInt(DayList.SelectedItem.Text)
 End If
Endd:
End Sub

Function FillDayList(Day)

'On Error GoTo Start
If NameChooser.Value <> "" Then Sheets(NameChooser.Value).Select

DoWeFill = LastFilled(Day)
DayList.ListItems.Clear

If NameChooser.Value <> "" And Day <> "" And DoWeFill <> 0 Then
 TotalTime = 0
 
 For i = 1 To DoWeFill
    Index = i + InfoOffset + Lines * (Day - 1) - 1
    Position = i
       
    JobName = Cells(Index, 2)
    Amount = Cells(Index, 4)
    UnitList = Cells(Index, 5)
    TimeList = Cells(Index, 6)
    TotalTime = TotalTime + TimeList
    RateList = Cells(Index, 7)
    Subtotal = Cells(Index, 9)
    
 DayList.ListItems.Add = Position
   DayList.ListItems.Item(i).ListSubItems.Add = JobName
   DayList.ListItems.Item(i).ListSubItems.Add = Amount
   DayList.ListItems.Item(i).ListSubItems.Add = UnitList
   DayList.ListItems.Item(i).ListSubItems.Add = TimeList
   DayList.ListItems.Item(i).ListSubItems.Add = RateList
   DayList.ListItems.Item(i).ListSubItems.Add = Subtotal

 Next i

 Index = InfoOffset + Lines * (Day - 1)
 
 If DayList.ListItems.Count > 0 Then
 CJob_Box.Value = DayList.ListItems.Count + 1
 DayList.ListItems.Add = " "
 DayList.ListItems.Add = " "
 DayList.ListItems.Item(DayList.ListItems.Count).ListSubItems.Add = ""
 DayList.ListItems.Item(DayList.ListItems.Count).ListSubItems.Add = ""
 DayList.ListItems.Item(DayList.ListItems.Count).ListSubItems.Add = "ВСЕГО"
 DayList.ListItems.Item(DayList.ListItems.Count).ListSubItems.Add = TotalTime
 DayList.ListItems.Item(DayList.ListItems.Count).ListSubItems.Add = "ИТОГО"
 DayList.ListItems.Item(DayList.ListItems.Count).ListSubItems.Add = Cells(Index, 10).Value
 End If



End If

If DoWeFill < Lines Then CJob_Box.Value = DoWeFill + 1
If DoWeFill > Lines - 1 Then CJob_Box.Value = Lines

If Day <> "" Then
Index = InfoOffset + Lines * (Day - 1)
Comment_Box.Value = Cells(Index, 13).Value
PrePay_Box.Value = Cells(Index, 11).Value
End If

GoTo Endd
Start:
ErrorForm.Error_Box.Value = "FillDayList()"
ErrorForm.Show
Endd:
End Function
Function PointFilter(Val) As String
PointFilter = Val
If Len(Val) = 0 Then GoTo Endd
If Right(Val, 1) = "1" Or _
   Right(Val, 1) = "2" Or _
   Right(Val, 1) = "3" Or _
   Right(Val, 1) = "4" Or _
   Right(Val, 1) = "5" Or _
   Right(Val, 1) = "6" Or _
   Right(Val, 1) = "7" Or _
   Right(Val, 1) = "8" Or _
   Right(Val, 1) = "9" Or _
   Right(Val, 1) = "0" Or _
   Right(Val, 1) = "-" Or _
   Right(Val, 1) = Application.DecimalSeparator _
                                                        Then CorrectEnter = True _
                                Else PointFilter = Left(Val, Len(Val) - 1)
Endd:
End Function

Private Sub AboveOklad_Chk_Change()
DetectChanges = 1
End Sub

Private Sub Amount_Box_Change()
Amount_Box.Value = PointFilter(Amount_Box.Value)
DetectChanges = 1
End Sub
Function isVisible(Day)
isVisible = 0
 Index = InfoOffset + Lines * Day
 If Cells(Index, 10) > 0 Then isVisible = 1
 If Cells(Index, 11) > 0 Then isVisible = 1
 If Cells(Index, 13) <> "" Then isVisible = 1
End Function

Function Mark(Day)
If PrevMarked = 1 Then Colorr = 2
If PrevMarked = 0 Then Colorr = 15
Index = InfoOffset + Lines * Day
    Range(Cells(Index, 1), Cells(Index + Lines - 1, 12)).Select
    With Selection.Interior
        .ColorIndex = Colorr
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
    End With
End Function

Private Sub ChooseWorker_Click()
WeChooseMate = 0
WorkersTreeHolder.Top = 6
WorkersTreeHolder.Left = 320
'For i = 1 To CInt(WorkersTree.Tag)
'  If WorkersTree.Nodes(i).Tag = "Cat" Then _
'        WorkersTree.Nodes(i).Expanded = False
' Next
'WorkersTree.Visible = True

Frame7.Visible = False
WorkersTreeHolder.Visible = True

'JobsTree.Visible = False

WorkersTree.SetFocus

End Sub
Private Sub ChooseMate_Button_Click()
WeChooseMate = 1
WorkersTreeHolder.Top = 255
WorkersTreeHolder.Left = 320
'For i = 1 To CInt(WorkersTree.Tag)
'  If WorkersTree.Nodes(i).Tag = "Cat" Then _
'        WorkersTree.Nodes(i).Expanded = False
' Next
'WorkersTree.Visible = True
WorkersTreeHolder.Visible = True
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

Private Sub Frame7_Click()
ObjectsRecall

End Sub

Private Sub JobsTree_DblClick()
If JobsTree.SelectedItem.Key <> "" Then
    'JobsTreeHolder.Visible = False
    ID.Value = CutZ(JobsTree.SelectedItem.Key)
    End If
End Sub


Private Sub Label16_Click()

End Sub

Private Sub Left_Box_Change()
Left_Box.Value = PointFilter(Left_Box.Value)
End Sub

Private Sub MateChooser_Change()
If MateChooser.Value = "" Then CopyDay_Button.Enabled = False
If MateChooser.Value <> "" Then CopyDay_Button.Enabled = True
End Sub


Private Sub Oklad_Box_Change()
Oklad_Box.Value = PointFilter(Oklad_Box.Value)
End Sub

Private Sub PrePay_Box_Change()
PrePay_Box.Value = PointFilter(PrePay_Box.Value)
End Sub

Private Sub Print_Button_Click()
If NameChooser.Value <> "" Then
Count = 0
Sheets(NameChooser.Value).Select
   Last = Cells(1, 1).Value
   If Last = "" Then Last = MDays(CMonth)
   For i = 0 To Last
    If isVisible(i) = 1 Then
    PrevMarked = Count Mod 2
    Mark (i)
    Count = Count + 1
    End If
    Next
 If OnScreen_Chk.Value = True Then Sheets(NameChooser.Value).PrintOut
 If OnScreen_Chk.Value = False Then
  WorkersExit = True
  Workers.Hide
  Form.Hide
  Cells(3, 1).Select
 End If
End If
End Sub
Private Sub Rate_Box_Change()
Rate_Box.Value = PointFilter(Rate_Box.Value)
DetectChanges = 1
End Sub

Private Sub RealName_Box_Change()

End Sub

Private Sub Time_Box_Change()
Time_Box.Value = PointFilter(Time_Box.Value)
DetectChanges = 1
End Sub
Private Sub Apply_Button_Click()
a = RecordInfo(CDay_Box.Value, CJob_Box.Value)

End Sub

Private Sub CDay_Box_Change()

If NameChooser.Value <> "" Then
FillDayList (CDay_Box.Value)

Label_FullDate.Caption = GetDayName(CDay_Box.Value) & ", " & _
                            CDay_Box.Value & " " & MNameRusFix(CMonth)
                            
Workers.Caption = RealName_Box.Value & ": " & Label_FullDate.Caption

a = ReadLine(CDay_Box.Value, CJob_Box.Value)
End If
End Sub
Private Sub CJob_Box_Change()

a = ReadLine(CDay_Box.Value, CJob_Box.Value)

End Sub
Private Sub Control_Button_Click()
FillControlBox
End Sub

Private Sub CopyDay_Button_Click()
If NameChooser.Value <> "" And MateChooser.Value <> "" Then
For i = 1 To Lines
Sheets(NameChooser.Value).Select
Index = i + InfoOffset + Lines * (CDay_Box.Value - 1) - 1
 
 If Cells(Index, 2).Value <> "" Then
 AddAmount = Cells(Index, 4).Value
 AddID = Cells(Index, 3).Value
 Range(Cells(Index, 2), Cells(Index, 9)).Copy

 Sheets("Каталог").Select
 If AddID > 5 Then Cells(AddID, 11).Value = Cells(AddID, 11).Value + AddAmount
 
 Sheets(MateChooser.Value).Select
 OldAmount = Cells(Index, 4).Value
 OldId = Cells(Index, 3).Value
 
    If OldId <> "" And AddID > 5 Then
     Sheets("Каталог").Select
     Cells(OldId, 11).Value = Cells(OldId, 11).Value - OldAmount
    End If
    
 Sheets(MateChooser.Value).Select
 Cells(Index, 2).PasteSpecial
 Cells(Index, 2).Select
 Selection.EntireRow.Hidden = False
 End If
 
Next i
Sheets(MateChooser.Value).Select

Cells(InfoOffset + 9 * (CDay_Box.Value - 1), 10).FormulaR1C1 = "=SUM(RC[-1]:R[8]C[-1])"

If CDay_Box.Value > Cells(1, 1).Value Then _
Cells(1, 1).Value = CDay_Box.Value
 
 a = TransferBalanceToNextMonth(MateChooser.Value, Cells(2, 10).Value)
   '' Workers.NameChooser.Value = MateChooser.Value
Sheets(NameChooser.Value).Select
End If
End Sub

Private Sub Clear_Button_Click()
ClearDay (CDay_Box.Value)
End Sub

Function CutZ(Val)
CutZ = CInt(Left(Val, Len(Val) - 1))
End Function

Private Sub ID_Change()
If Read = 0 Then
DetectChanges = 1
 If ID.Value <> "" And ID.Value <> 0 Then
ID = ID.Value
Sheets("Каталог").Select
JobName_Box.Value = Cells(ID, 2).Value

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
End Sub
Private Sub NameChooser_Change()
If NameChooser.Value <> "" Then
Sheets(NameChooser.Value).Select
FillControlBox
ReadLockedInfo

Workers.Caption = RealName_Box.Value & ": " & Label_FullDate.Caption

FillDayList (CDay_Box.Value)
  If NameChooser.Value = MateChooser.Value Then
  MateChooser.Value = ""
  MateName_Box.Value = ""
  End If
End If
End Sub

Private Sub Setup_Button_Click()
If CDay_Box.Value <> "" Then LastWorkersDay = CInt(CDay_Box.Value)
LastPerson = NameChooser.Value
Workers.Hide


Setup.ScanWorkers (WorkersBase)
            
            
Setup.ScanWCats
Setup.ScanJobs
Setup.ScanOrgs

Setup.ScanJCats
Setup.ScanOCats

Setup.NameChooser.Value = _
            Setup.WorkersTree.Nodes(CInt(Setup.WorkersTree.Tag) + 1).Key
Setup.jID.Value = _
            Workers.CutZ(Setup.JobsTree.Nodes(CInt(Setup.JobsTree.Tag) + 1).Key)
            
Setup.oID.Value = _
            Workers.CutZ(Setup.OrgsTree.Nodes(CInt(Setup.OrgsTree.Tag) + 1).Key)



Setup.cCatChooser.Value = Setup.cCatChooser.List(1)
Setup.jCatChooser.Value = Setup.jCatChooser.List(1)
Setup.oCatChooser.Value = Setup.oCatChooser.List(1)


Setup.Show
End Sub

Private Sub Day_Spin_SpinDown()
If CDay_Box.Value < MDays(CMonth) Then _
           CDay_Box.Value = CDay_Box.Value + 1
End Sub

Private Sub Day_Spin_SpinUp()
If CDay_Box.Value > 1 Then _
           CDay_Box.Value = CDay_Box.Value - 1
End Sub
Private Sub Workers_Spin_SpinDown()
On Error GoTo Endd
WorkersTreeHolder.Top = 1600
WorkersTreeHolder.Visible = True

Total = WorkersTree.Nodes.Count
Index = WorkersTree.SelectedItem.Index
If Index <= Total - 1 Then
Total = WorkersTree.Nodes.Count
Index = WorkersTree.SelectedItem.Index
 If WorkersTree.Nodes(Index + 1).Tag <> "Cat" Then
 RealName_Box.Value = WorkersTree.Nodes(Index + 1).Text
 NameChooser.Value = WorkersTree.Nodes(Index + 1).Key
 WorkersTree.Nodes(Index + 1).Selected = True
 End If
End If
Endd:
WorkersTreeHolder.Visible = False
End Sub

Private Sub Workers_Spin_SpinUp()
On Error GoTo Endd
WorkersTreeHolder.Top = 1600
WorkersTreeHolder.Visible = True
Index = WorkersTree.SelectedItem.Index

If WorkersTree.Nodes(Index - 1).Tag <> "Cat" Then
RealName_Box.Value = WorkersTree.Nodes(Index - 1).Text
NameChooser.Value = WorkersTree.Nodes(Index - 1).Key
WorkersTree.Nodes(Index - 1).Selected = True
Else
WorkersTree.Nodes(Index).Selected = True
End If
Endd:
WorkersTreeHolder.Visible = False
End Sub

Private Sub UserForm_Click()
ObjectsRecall
End Sub

Private Sub WorkersClose_Button_Click()
If CDay_Box.Value <> "" Then LastWorkersDay = CInt(CDay_Box.Value)
LastPerson = NameChooser.Value
Workers.Hide
End Sub


Private Sub WorkersTree_DblClick()
If WorkersTree.SelectedItem.Key <> "" And WorkersTree.SelectedItem.Tag <> "Cat" Then
     
     If WeChooseMate = 0 Then
     RealName_Box.Value = WorkersTree.SelectedItem.Text
     NameChooser.Value = WorkersTree.SelectedItem.Key
      If WorkersTree.SelectedItem.Key = MateChooser.Value Then
      MateChooser.Value = ""
      MateName_Box.Value = ""
      End If
     WorkersTreeHolder.Visible = False
     Frame7.Visible = True
     
     JobsTree.Visible = True
     JobsTree.SetFocus
     Else
     
       If WorkersTree.SelectedItem.Key <> NameChooser.Value Then
       MateName_Box.Value = WorkersTree.SelectedItem.Text
       MateChooser.Value = WorkersTree.SelectedItem.Key
       WorkersTreeHolder.Visible = False
       DayList.Visible = True
       CopyDay_Button.SetFocus
       End If
     End If
  
End If
End Sub

