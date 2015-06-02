VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Form 
   Caption         =   "ООО ""ТехПолиМет"" Автоматизированная система учёта продукции [Главный модуль] v2.0"
   ClientHeight    =   11025
   ClientLeft      =   6045
   ClientTop       =   6330
   ClientWidth     =   15270
   OleObjectBlob   =   "Form.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Leftt, Income, outcome, Balance, LastName, Namess As String

Function SetCaption(Fil, Side)
''Protect
If Side = 1 Then

End If
''unprotect
If Side = -1 Then

End If
End Function

Function OpenFile(Fil)
On Error GoTo FIn
SetCaption(Fil, 1) = 0
Workbooks.Open FileName:=Fil
FIn:
End Function
Function SaveClose(Fil)
On Error GoTo FIn
Windows(Fil).Activate
ActiveWorkbook.Save
ActiveWorkbook.Close
FIn:
End Function

Function JustSave(Fil)
On Error GoTo FIn
Windows(Fil).Activate
ActiveWorkbook.Save
FIn:
End Function
Function GetAv(Line)
 For i = 1 To Len(Line) - 1
 If Mid(Line, i, 1) = "#" Then GetAv = Right(Line, Len(Line) - i)
 Next
End Function
Function GetDay(Line) As Integer
 For i = 1 To Len(Line) - 1
 If Mid(Line, i, 1) = "#" Then GetDay = CInt(Left(Line, i - 1))
 Next
End Function

Private Sub AvReport_Button_Click()
'On Error GoTo Start
Windows(WorkersBase).Activate
 Sheets("АвансовыйОтчёт").Select
Cells(2, 2) = DateTime.Date
Cells(3, 2) = DateTime.TIME


Range("B7:AH200").Clear
Range("C7:AG7").Select
Selection.EntireColumn.Hidden = True

Cells(1, 2).Value = "Авансовый отчёт за " & MName(CMonth)

 Start = 0
 HiddenCount = 0
 Sheets("Сотрудники").Select
    
    Range("B2:F100").Select
    Selection.Sort Key1:=Range("B3"), Order1:=xlAscending, Header:=xlGuess, _
        OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, _
        DataOption1:=xlSortNormal
        
WeHaveWorkers = Cells(1, 2).Value

For i = 3 To WeHaveWorkers + 2
Sheets("Сотрудники").Select
If Cells(i, 4).Value = 1 Then HiddenCount = HiddenCount + 1
If Cells(i, 4).Value = 0 Then
ii = i - HiddenCount

Sheets(Cells(i, 3).Value).Select
Namess = Cells(1, 2).Value & " " & Cells(2, 2).Value
AvRepColl.Clear
  For j = 6 To 276 Step 9
   If Cells(j, 11).Value <> 0 Then
     AvRepColl.AddItem (CStr(Cells(j, 1).Value) & "#" & CStr(Cells(j, 11).Value))
   End If
  Next j

 Sheets("АвансовыйОтчёт").Select
RepOffset = 4 + ii
Cells(RepOffset, 2).Value = Namess
Cells(RepOffset, 34).FormulaR1C1 = "=SUM(RC[-31]:RC[-1])"
 For j = 0 To AvRepColl.ListCount - 1
 Clmn = GetDay(AvRepColl.List(j)) + 2
 Av = GetAv(AvRepColl.List(j))
 
 Cells(RepOffset, Clmn).Value = Av
 Cells(RepOffset, Clmn).Select
 Selection.EntireColumn.Hidden = False

 Next j
    
Range(Cells(RepOffset, 2), Cells(RepOffset, 34)).Select
Selection.NumberFormat = "#,##0.00"
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

If Start Mod 2 = 0 Then
    With Selection.Interior
        .ColorIndex = 15
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
    End With
End If

Start = Start + 1

End If
Next
'  Columns("B:AH").EntireColumn.AutoFit

If NoPrintAvReport_Chk.Value = True Then Sheets("АвансовыйОтчёт").PrintOut
If NoPrintAvReport_Chk.Value = False Then
     Form.Hide
     ReportExit = True
End If


GoTo Endd:
Start:
ErrorForm.Error_Box.Value = "AvReport()"
ErrorForm.Show
Endd:

End Sub

Private Sub Block_Button_Click()
BlockIt.Show
End Sub

Private Sub Chamber_Button_Click()
FiltersReady = 0

Orders.CDay_Box.Clear
Orders.Day_Filter.Clear
Orders.Day_Filter.AddItem ("Все")

  For i = 1 To MDays(CMonth)
   Orders.CDay_Box.AddItem (i)
   Orders.Day_Filter.AddItem (i)
  Next
Orders.CDay_Box.Value = DateTime.Day(DateTime.Date)
Orders.Day_Filter.Value = "Все"

Orders.Dol_Chooser.Clear
Orders.Dol_Chooser.AddItem ("Все")
Orders.Dol_Chooser.AddItem ("Долги")
Orders.Dol_Chooser.Value = "Долги"

Orders.Opl_Chooser.Clear
Orders.Opl_Chooser.AddItem ("Все")
Orders.Opl_Chooser.AddItem ("нал")
Orders.Opl_Chooser.AddItem ("б/н")
Orders.Opl_Chooser.Value = "б/н"

Orders.Opl_Chooser_w.Clear
Orders.Opl_Chooser_w.AddItem ("нал")
Orders.Opl_Chooser_w.AddItem ("б/н")
Orders.Opl_Chooser_w.Value = "б/н"

Orders.RoundType.Clear
Orders.RoundType.AddItem ("в большую сторону")
Orders.RoundType.AddItem ("в меньшую сторону")
Orders.RoundType.Value = "в большую сторону"




Orders.Label_FullDate.Caption = GetDayName(Orders.CDay_Box.Value) & ", " & _
                            Orders.CDay_Box.Value & " " & MNameRusFix(CMonth)

Orders.ScanOrgs (WorkersBase)
Orders.ScanJobs
Orders.ScanOCats
FiltersReady = 1
Orders.Region_Filter.Value = "Все"

Orders.OrgName_Box.Value = _
            Orders.OrgsTree.Nodes(CInt(Orders.OrgsTree.Tag) + 1).Text
Orders.oID.Value = _
            Workers.CutZ(Orders.OrgsTree.Nodes(CInt(Orders.OrgsTree.Tag) + 1).Key)

Orders.OrgsTreeHolder.Visible = True
Orders.OrgsTree.Nodes(CInt(Orders.OrgsTree.Tag) + 1).Selected = True
Orders.oCat.Value = _
            Workers.CutZ(Orders.OrgsTree.SelectedItem.Parent.Key)
Orders.OrgsTreeHolder.Visible = False

'If LastPerson <> "" Then Workers.NameChooser.Value = LastPerson Else _
'                         Workers.NameChooser.Value = Workers.NameChooser.List(0)

Orders.Show

End Sub

Private Sub GenerateNextMonth_Click()

If DateTime.Month(DateTime.Date) = CMonth Then

  b = MsgBox(NextMonth & " ещё не наступил (или уже прошёл :-D)", vbOKOnly, "Внимание")
  GoTo Endd
  
  End If

InsureForm.CommandButton2.SetFocus
InsureForm.Show
If InsureForm.OK.Value = True Then
InsureForm.OK.Value = False

SaveClose (WorkersBase)

 ArcMonth = CMonth - 1
 ArcYear = CYear
 If ArcMonth = 0 Then
    ArcMonth = 12
    ArcYear = CYear - 1
    End If

''Name Path + "lWorkers.xls" As Path + "iworkers.xls"
 
 ArcName = Path + "Archive\Valid\" & MNameEng(ArcMonth) & "_" & ArcYear
 ArcFile = Path + "lWorkers.xls"
 a = Shell("C:\Program Files\WinRar\WinRar.EXE m -ep " & ArcName & " " & ArcFile, vbMinimizedNoFocus)
 

On Error GoTo Smart
Do
AppActivate (a)
Loop Until 1 > 2
Smart:

 
 Destination = Path & "lWorkers.xls"
 Source = Path & WorkersBase


FileCopy Source, Destination

Workbooks.Open FileName:=Path + WorkersBase

 Windows(WorkersBase).Activate
 
 Sheets("Каталог").Select
  CYear = Cells(1, 3).Value
  CMonth = Cells(2, 3).Value
  Cells(2, 3).Value = Cells(2, 3).Value + 1
  If CMonth = 12 Then
     Cells(2, 3).Value = 1
     Cells(1, 3).Value = Cells(1, 3).Value + 1
  End If
  
  CYear = Cells(1, 3).Value
  CMonth = Cells(2, 3).Value
    Cells(2, 2).Value = MName(CMonth)

 For i = 9 To ActiveWorkbook.Sheets.Count
  Sheets(i).Select
  
  Cells(2, 10).Value = Cells(1, 10).Value
  
  Cells(1, 1).ClearContents
  Range("b6:k284").ClearContents
  Range("m6:m284").ClearContents
  
  Rows("6:284").Select
  Selection.EntireRow.Hidden = True
 
 Next
ReportExit = True
FormShow
Endd:
End If
End Sub

Private Sub SaveAndClose_Click()
    SaveClose (WorkersBase)
    
   Path = Workbooks("Index.XLS").Path + "\"

 ArcName = Path + "Archive\LastState"
 ArcFiles = Path + "*Workers.xls"
 a = Shell("C:\Program Files\WinRar\WinRar.EXE a -ep -y " & ArcName & " " & ArcFiles, vbMinimizedNoFocus)

On Error GoTo Start
Do
AppActivate (a)
Loop Until 1 > 2
Start:
    Windows("Index.xls").Close (SaveChanges = xlDoNotSaveChanges)
    
    Form.Hide
    
Application.Quit

End Sub

Private Sub SaveState_Click()
On Error Resume Next

JustSave (WorkersBase)

End Sub

Private Sub SwitchToLastMonth_Click()

If Not IsOpened("lWorkers.xls") Then Workbooks.Open FileName:=Path & "lWorkers.xls" _
Else SaveClose ("lWorkers.xls")
ReportExit = True
FormShow
End Sub


Private Sub Workers_Button_Click()

Workers.CDay_Box.Clear
     
  For i = 1 To MDays(CMonth)
   Workers.CDay_Box.AddItem (i)
  Next


If LastWorkersDay <> 0 Then
Workers.CDay_Box.Value = LastWorkersDay
Else
Workers.CDay_Box.Value = DateTime.Day(DateTime.Date)
End If


Workers.Label_FullDate.Caption = GetDayName(Workers.CDay_Box.Value) & ", " & _
                            Workers.CDay_Box.Value & " " & MNameRusFix(CMonth)
Workers.IncomeLabel.Caption = "Приход за " & MName(CMonth)
Workers.OutComeLabel.Caption = "Расход за " & MName(CMonth)
Workers.LeftLabel.Caption = "Остаток за " & MName(LMonth)

Workers.ScanWorkers (WorkersBase)
Workers.ScanJobs


Workers.RealName_Box.Value = _
            Workers.WorkersTree.Nodes(CInt(Workers.WorkersTree.Tag) + 1).Text
Workers.NameChooser.Value = _
            Workers.WorkersTree.Nodes(CInt(Workers.WorkersTree.Tag) + 1).Key
            Workers.WorkersTree.Nodes(CInt(Workers.WorkersTree.Tag) + 1).Selected = True
'If LastPerson <> "" Then Workers.NameChooser.Value = LastPerson Else _
'                         Workers.NameChooser.Value = Workers.NameChooser.List(0)
Workers.Show
End Sub

Private Sub FeeReport_Button_Click()
'On Error GoTo Start
Windows(WorkersBase).Activate
 Sheets("Отчёт").Select

Selection.Font.Bold = False
Range("B7:G100").Clear

Cells(1, 3).Value = "Отчёт по зарплате за  " & MName(CMonth)
Cells(6, 3).Value = "Остаток за " & MName(LMonth)
Cells(6, 5).Value = "Выдано за " & MName(CMonth)

 Start = 0
 HiddenCount = 0
 Sheets("Сотрудники").Select
    
    Range("B2:F100").Select
    Selection.Sort Key1:=Range("B3"), Order1:=xlAscending, Header:=xlGuess, _
        OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, _
        DataOption1:=xlSortNormal
        
WeHaveWorkers = Cells(1, 2).Value

For i = 3 To WeHaveWorkers + 2
Sheets("Сотрудники").Select
If Cells(i, 4).Value = 1 Then HiddenCount = HiddenCount + 1
If Cells(i, 4).Value = 0 Then
ii = i - HiddenCount

Sheets(Cells(i, 3).Value).Select

If Cells(1, 1).Value <> "" Then LastDay = "(по " & Cells(1, 1).Value & "-e число)" _
Else LastDay = "#нет данных#"
Leftt = Cells(2, 10).Value
Income = Cells(3, 10).Value
outcome = Cells(3, 11).Value
Balance = Cells(1, 10).Value
Namess = Cells(1, 2).Value & " " & Cells(2, 2).Value

 Sheets("Отчёт").Select

Cells(3, 4) = DateTime.Date
Cells(3, 5) = DateTime.TIME
RepOffset = 4 + ii
Cells(RepOffset, 2) = Namess
Cells(RepOffset, 3) = Leftt
Cells(RepOffset, 4) = Income
Cells(RepOffset, 5) = outcome
Cells(RepOffset, 6) = Balance
Cells(RepOffset, 7) = LastDay
Range(Cells(RepOffset, 6), Cells(RepOffset, 6)).Select
If Balance < 0 Then Selection.Font.Bold = True
    
Range(Cells(RepOffset, 2), Cells(RepOffset, 6)).Select
Selection.NumberFormat = "#,##0.00"
   


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

If Start Mod 2 = 0 Then
    With Selection.Interior
        .ColorIndex = 15
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
    End With
End If

Start = Start + 1

End If



Next


If NoPrintFeeReport_Chk.Value = True Then Sheets("Отчёт").PrintOut
If NoPrintFeeReport_Chk.Value = False Then
     Form.Hide
     ReportExit = True
End If


GoTo Endd:
Start:
ErrorForm.Error_Box.Value = "FeeReport()"
ErrorForm.Show
Endd:
End Sub

Private Sub Setup_Button_Click()

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
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
''If CloseMode = 0 Then Cancel = 1
End Sub


