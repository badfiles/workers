﻿VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Form 
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
                            Orders.CDay_Box.Value & " " & MName(CMonth, True)

Orders.ScanOrgs (WorkersBase)
Orders.ScanJobs
Orders.ScanOCats
FiltersReady = 1
Orders.Region_Filter.Value = "Все"

Orders.OrgName_Box.Value = Orders.OrgsTree.Nodes(CInt(Orders.OrgsTree.Tag) + 1).Text
Orders.oID.Value = CutZ(Orders.OrgsTree.Nodes(CInt(Orders.OrgsTree.Tag) + 1).Key)

Orders.OrgsTreeHolder.Visible = True
Orders.OrgsTree.Nodes(CInt(Orders.OrgsTree.Tag) + 1).Selected = True
Orders.oCat.Value = CutZ(Orders.OrgsTree.SelectedItem.Parent.Key)
Orders.OrgsTreeHolder.Visible = False

'If LastPerson <> "" Then Workers.NameChooser.Value = LastPerson Else _
'                         Workers.NameChooser.Value = Workers.NameChooser.List(0)

Orders.Show

End Sub


Private Sub GenerateNextMonth_Click()
On Error GoTo ExceptionControl:
If DateTime.Month(DateTime.Date) <> NMonth Then
    b = MsgBox(NextMonth & " ещё не наступил (или уже прошёл :-D)", vbOKOnly, "Внимание")
    Exit Sub
End If

InsureForm.NoButton.SetFocus
InsureForm.Msg_label.Caption = "После перехода на новый месяц будет невозможно " & SwitchToLastMonth.Caption & ". Продолжаем?"
InsureForm.Show
If InsureForm.OK.Value = True Then
    SaveClose (WorkersBase)

    ArcMonth = CMonth - 1
    ArcYear = CYear
    If ArcMonth = 0 Then
        ArcMonth = 12
        ArcYear = CYear - 1
    End If

    ''Name Path + "lWorkers.xls" As Path + "iworkers.xls"
 
    ArcName = Path + "Archive\Valid\" & MNameEng(ArcMonth) & "_" & ArcYear
    ArcFiles = Path + "lWorkers.xls"
    
    RunCommand (Archiver & " a " & ArcKey & " " & ArcName & " " & ArcFiles)
     
    Destination = Path & "lWorkers.xls"
    Source = Path & WorkersBase
    FileCopy Source, Destination

    Workbooks.Open Filename:=Path + WorkersBase

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
    
    Cells(1, 6).Value = Cells(2, 6).Value
    Cells(2, 6).Value = ""
    
    For i = 9 To ActiveWorkbook.Sheets.Count
        Sheets(i).Select
        Cells(2, 10).Value = Cells(1, 10).Value
        Cells(1, 1).ClearContents
        Range("b6:k284").ClearContents
        Range("m6:n284").ClearContents
        Rows("6:284").Select
        Selection.EntireRow.Hidden = True
    Next
    ReportExit = True
    FormShow
End If

Exit Sub
ExceptionControl:
ErrorForm.Error_Box.Value = "Form/GenerateNextMonth_Click()"
ErrorForm.Show
End Sub

Sub DropSensitiveData()
On Error GoTo ExceptionControl:
    Sheets("АвансовыйОтчёт").Select
    Range("a7:bb684").ClearContents
    Sheets("Отчёт").Select
    Range("a7:bb684").ClearContents
    Sheets("Каталог").Select

Exit Sub
ExceptionControl:
ErrorForm.Error_Box.Value = "Form/DropSensitiveData()"
ErrorForm.Show
End Sub


Sub DoneForNow(ByVal CloseBase As Boolean)
On Error GoTo ExceptionControl:
If AdminMode Then
    Windows(WorkersBase).Activate
    DropSensitiveData
    SaveClose (WorkersBase)
    
    PushBase = "push.xls"
    Destination = Path & PushBase
    Source = Path & WorkersBase
    FileCopy Source, Destination

    ArcName = Path & "push.7z"
    Kill (ArcName)
    ArcFiles = Path & PushBase & " " & Path & "index-c.xls"
    RunCommand (Archiver & " a " & ExchangeKey & " " & ArcName & " " & ArcFiles)
    RunCommand (Archiver & " rn " & ExchangeKey & " " & ArcName & " index-c.xls index.xls")
    Kill (Path & PushBase)
    
    If CloseBase Then
        ArcName = Path + "Archive\LastState.7z"
        ArcFiles = Path + "*Workers.xls"
        RunCommand (Archiver & " a " & ArcKey & " " & ArcName & " " & ArcFiles)
        a = Shell("ftp -v -s:" & Path & "ftp_server_send_all " & FtpStorageName, vbMinimizedNoFocus)
    Else
        a = Shell("ftp -v -s:" & Path & "ftp_server_send " & FtpStorageName, vbMinimizedNoFocus)
    End If
End If

Exit Sub
ExceptionControl:
ErrorForm.Error_Box.Value = "Form/DoneForNow()"
ErrorForm.Show
End Sub

Private Sub RunTC_Button_Click()
BlockIt.Pass = PinAdmin
BlockIt.PassOK = False
BlockIt.Password_Box.SetFocus
BlockIt.Show
If BlockIt.PassOK Then a = Shell("c:\Program Files\WINCMD\totalcmd.exe", vbMaximizedFocus)
End Sub

Private Sub SaveAndClose_Click()
On Error GoTo ExceptionControl:
DoneForNow (True)
Windows("Index.xls").Close (SaveChanges = xlDoNotSaveChanges)
Form.Hide
Application.Quit

Exit Sub
ExceptionControl:
ErrorForm.Error_Box.Value = "Form/SaveAndClose_Click()"
ErrorForm.Show
End Sub

Private Sub SaveState_Click()
DoneForNow (False)
FormShow
End Sub

Private Sub SwitchToLastMonth_Click()
On Error Resume Next
If LMMode Then
    LastWorkersDay = 0
    If AdminMode Then
        Windows(WorkersBase).Activate
        DropSensitiveData
        SaveClose ("lWorkers.xls")
        ArcName = Path & "lm.7z"
        ArcFiles = Path & "lWorkers.xls"
        RunCommand (Archiver & " a " & ExchangeKey & " " & ArcName & " " & ArcFiles)
    Else
        Windows(WorkersBase).Activate
        ActiveWorkbook.Close SaveChanges:=False
        Kill (Path & WorkersBase)
    End If
Else
    LastWorkersDay = 31
    If Not AdminMode Then
        ArcFiles = "lWorkers.xls"
        ArcName = Path & "lm.7z"
        RunCommand ("ftp -v -s:" & Path & "ftp_client_get_lm " & FtpStorageName)
        RunCommand (Archiver & " e -y " & ExchangeKey & " " & ArcName & " -o" & Path & " " & ArcFiles)
        Kill (ArcName)
    End If
    Workbooks.Open Filename:=Path & "lWorkers.xls"
End If
ReportExit = True
FormShow
End Sub


Private Sub InitWorkers()
On Error GoTo ExceptionControl:
Form.Top = 0
Form.Left = 0
Windows(WorkersBase).Activate

ExtChange = True
Workers.CDay_Box.Clear
For i = 1 To MDays(CMonth)
    Workers.CDay_Box.AddItem (i)
Next
If LastWorkersDay <> 0 Then
    Workers.CDay_Box.Value = LastWorkersDay
Else
    Workers.CDay_Box.Value = DateTime.Day(DateTime.Date)
End If
If Workers.CDay_Box.Value > MDays(CMonth) Then Workers.CDay_Box.Value = MDays(CMonth)
ExtChange = False

Workers.Label_FullDate.Caption = GetDayName(Workers.CDay_Box.Value) & ", " & Workers.CDay_Box.Value & " " & MName(CMonth, True)

Workers.IncomeLabel.Caption = "Заработано за " & MName(CMonth)
Workers.OutComeLabel.Caption = "Выдано за " & MName(CMonth)
Workers.LeftLabel.Caption = "Остаток за " & MName(LMonth)

Workers.ScanWorkers
Workers.ScanJobs

If LMMode Then
    Workers.LastMonth_Label.Visible = True
    Workers.MakeReadOnly_Chk.Visible = False
Else
    Workers.LastMonth_Label.Visible = False
    Workers.MakeReadOnly_Chk.Visible = True
End If

Exit Sub
ExceptionControl:
ErrorForm.Error_Box.Value = "Form/InitWorkers()"
ErrorForm.Show
End Sub

Private Sub Workers_Button_Click()
On Error GoTo ExceptionControl:
MainReInit
If AdminMode Then
        
    ArcFiles = "pull.xls"
    RunCommand ("ftp -v -s:" & Path & "ftp_server_get " & FtpStorageName)
    ArcName = Path + "pull.7z"
    RunCommand (Archiver & " e -y " & ExchangeKey & " " & ArcName & " -o" & Path & " " & ArcFiles)
    
    PullOnServer

    InitWorkers
    Workers.WorkersTreeHolder.Visible = True
    Workers.RealName_Box.Value = ""
    Workers.NameChooser.Value = ""
    Workers.RealName_Box.Value = Workers.WorkersTree.Nodes(CInt(Workers.WorkersTree.Tag) + 1).Text
    Workers.NameChooser.Value = Workers.WorkersTree.Nodes(CInt(Workers.WorkersTree.Tag) + 1).Key
    Workers.WorkersTree.Nodes(CInt(Workers.WorkersTree.Tag) + 1).Selected = True
    Workers.WorkersTreeHolder.Visible = False
    Workers.Show
Else
    If Not LMMode Then
        Windows(WorkersBase).Activate
        Sheets("Каталог").Select
        ReferenceTokens = Cells(2, 6).Value

        ArcFiles = "push.xls index.xls"
        ArcName = Path & "push.7z"
        RunCommand ("ftp -v -s:" & Path & "ftp_client_get " & FtpStorageName)
        RunCommand (Archiver & " e -y " & ExchangeKey & " " & ArcName & " -o" & Path & " " & ArcFiles)
        
        PushBase = "push.xls"
        Destination = Path & "tWorkers.xls"
        Source = Path & PushBase

        If Not IsOpened("push.xls") Then Workbooks.Open Filename:=Path + PushBase
        Windows(PushBase).Activate
        Sheets("Каталог").Select

        If (Cells(1, 6).Value = ReferenceTokens) Or (Cells(2, 6).Value = ReferenceTokens) Then
            ActiveWorkbook.Close SaveChanges:=False
            Windows(WorkersBase).Activate
            ActiveWorkbook.Close SaveChanges:=False
            FileCopy Source, Destination
            MainReInit
        Else
            ActiveWorkbook.Close
            Windows(WorkersBase).Activate
        End If
    End If
    InitWorkers
    Workers.MakeReadOnly_Chk.Visible = False
    If LMMode Then
        Workers.Apply_Button.Enabled = False
        Workers.Clear_Button.Enabled = False
        Workers.Delete_Button.Enabled = False
        Workers.ChooseMate_Button.Enabled = False
    End If
    Workers.Show
End If
'If LastPerson <> "" Then Workers.NameChooser.Value = LastPerson Else _
'                         Workers.NameChooser.Value = Workers.NameChooser.List(0)
Exit Sub
ExceptionControl:
ErrorForm.Error_Box.Value = "Form/Workers_Button_Click()"
ErrorForm.Show
End Sub

Private Sub FeeReport_Button_Click()
On Error GoTo ExceptionControl:
Windows(WorkersBase).Activate
Sheets("Отчёт").Select

Selection.Font.Bold = False
Range("B7:G100").Clear

Cells(1, 3).Value = "Отчёт по зарплате за  " & MName(CMonth)
Cells(6, 3).Value = "Остаток за " & MName(LMonth)
Cells(6, 5).Value = "Выдано за " & MName(CMonth)

MarkLine = True
HiddenCount = 0
Sheets("Сотрудники").Select
Range("A3:G100").Select
Selection.Sort Key1:=Range("B3"), Order1:=xlAscending, Header:=xlGuess, _
    OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, _
    DataOption1:=xlSortNormal
        
WeHaveWorkers = Cells(1, 2).Value

For i = 3 To WeHaveWorkers + 2
    Sheets("Сотрудники").Select
    If Cells(i, 4).Value = 1 Then
        HiddenCount = HiddenCount + 1
    Else
        ii = i - HiddenCount
        Sheets(Cells(i, 3).Value).Select
        If Cells(1, 1).Value <> "" Then LastDay = "(по " & Cells(1, 1).Value & "-e число)" Else LastDay = "#нет данных#"
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
        FillAndBorders (MarkLine)
        MarkLine = Not MarkLine
    End If
Next

If NoPrintFeeReport_Chk.Value = True Then
    Sheets("Отчёт").PrintOut
Else
    Form.Hide
    ReportExit = True
End If

Exit Sub
ExceptionControl:
ErrorForm.Error_Box.Value = "Form/FeeReport_Button_Click()"
ErrorForm.Show
End Sub

Private Sub AvReport_Button_Click()
Dim Day(1 To 31) As Integer, Av(1 To 31) As String
On Error GoTo ExceptionControl:
Windows(WorkersBase).Activate
Sheets("АвансовыйОтчёт").Select
Cells(2, 2) = DateTime.Date
Cells(3, 2) = DateTime.TIME

Range("B7:AH200").Clear
Range("C7:AG7").Select
Selection.EntireColumn.Hidden = True

Cells(1, 2).Value = "Авансовый отчёт за " & MName(CMonth)

MarkLine = True
HiddenCount = 0
Sheets("Сотрудники").Select
Range("A3:G100").Select
Selection.Sort Key1:=Range("B3"), Order1:=xlAscending, Header:=xlGuess, _
    OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, _
    DataOption1:=xlSortNormal

WeHaveWorkers = Cells(1, 2).Value
For i = 3 To WeHaveWorkers + 2
    Sheets("Сотрудники").Select
    If Cells(i, 4).Value = 1 Then
        HiddenCount = HiddenCount + 1
    Else
        ii = i - HiddenCount
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
        RepOffset = 4 + ii
        Cells(RepOffset, 2).Value = Namess
        Cells(RepOffset, 34).FormulaR1C1 = "=SUM(RC[-31]:RC[-1])"
        For j = 1 To p
            Clmn = Day(j) + 2
            Cells(RepOffset, Clmn).Value = Av(j)
            Cells(RepOffset, Clmn).Select
            Selection.EntireColumn.Hidden = False
        Next j
    
        Range(Cells(RepOffset, 2), Cells(RepOffset, 34)).Select
        Selection.NumberFormat = "#,##0.00"
        FillAndBorders (MarkLine)
        MarkLine = Not MarkLine
    End If
Next

If NoPrintAvReport_Chk.Value = True Then
    Sheets("АвансовыйОтчёт").PrintOut
Else
     Form.Hide
     ReportExit = True
End If

Exit Sub
ExceptionControl:
ErrorForm.Error_Box.Value = "Form/AvReport_Button_Click()"
ErrorForm.Show
End Sub
Private Sub FillAndBorders(ByVal MarkLine As Boolean)
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

Private Sub Setup_Button_Click()
On Error GoTo ExceptionControl:
Setup.ScanWorkers (WorkersBase)
Setup.ScanWCats
Setup.ScanJobs
Setup.ScanOrgs
Setup.ScanJCats
Setup.ScanOCats
Setup.NameChooser.Value = Setup.WorkersTree.Nodes(CInt(Setup.WorkersTree.Tag) + 1).Key
Setup.jID.Value = CutZ(Setup.JobsTree.Nodes(CInt(Setup.JobsTree.Tag) + 1).Key)
Setup.oID.Value = CutZ(Setup.OrgsTree.Nodes(CInt(Setup.OrgsTree.Tag) + 1).Key)
Setup.cCatChooser.Value = Setup.cCatChooser.List(1)
Setup.jCatChooser.Value = Setup.jCatChooser.List(1)
Setup.oCatChooser.Value = Setup.oCatChooser.List(1)
Setup.Show

Exit Sub
ExceptionControl:
ErrorForm.Error_Box.Value = "Form/Setup_Button_Click()"
ErrorForm.Show
End Sub
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
BlockIt.Pass = PinAdmin
BlockIt.PassOK = False
BlockIt.Password_Box.SetFocus
BlockIt.Show
If BlockIt.PassOK = False And CloseMode = 0 Then Cancel = 1
End Sub


