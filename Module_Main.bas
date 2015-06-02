Attribute VB_Name = "Module_MAin"
Public CYear, CMonth, LMonth, NMonth As Integer
Public WorkersBase, Path, NextMonth As String
Public AtLast As Boolean
Public ReportExit, WorkersExit, LMMode As Boolean
Public LastWorkersDay As Integer
Public FiltersReady As Integer
Public LastPerson As String

Public Const FtpStorageName = "10.10.11.1"

Public Const PinAdmin = "17ED0255"
'Public Const PinAdmin = "free"
'Public Const ExchangeKey = "-mhe -pLLHNM8y56vmjthURFnvhnvw9o4y6"
'Public Const ArcKey = "-mhe -pIUBNMyjbo7ytno4RYDthvklfhg59b"
Public Const Archiver = "c:\Program Files\7-zip\7z.exe"
Public Const ExchangeKey = ""
Public Const ArcKey = ""
Public Const Version = "U-3.3.102"

Public Const AdminMode = True
'Public Const AdminMode = False

Declare Function GetSystemMetrics32 Lib "user32" Alias "GetSystemMetrics" (ByVal nIndex As Long) As Long
Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public Sub OpenFile(ByVal Fil As String)
On Error GoTo ExceptionControl:
'SetCaption(Fil, 1) = 0
Workbooks.Open Filename:=Fil

Exit Sub
ExceptionControl:
ErrorForm.Error_Box.Value = "Main/OpenFile()"
ErrorForm.Show
End Sub
Public Sub SaveClose(ByVal Fil As String)
On Error GoTo ExceptionControl:
Windows(Fil).Activate
ActiveWorkbook.Save
ActiveWorkbook.Close

Exit Sub
ExceptionControl:
ErrorForm.Error_Box.Value = "Main/SaveClose()"
ErrorForm.Show
End Sub

Public Sub JustSave(ByVal Fil As String)
On Error GoTo ExceptionControl:
Windows(Fil).Activate
ActiveWorkbook.Save

Exit Sub
ExceptionControl:
ErrorForm.Error_Box.Value = "Main/JustSave()"
ErrorForm.Show
End Sub
Public Sub RunCommand(ByVal Command As String)
On Error GoTo Endd:
pid = Shell(Command, vbMinimizedNoFocus)
Do
Sleep (500)
AppActivate (pid)
Loop Until 1 > 2
Endd:
End Sub

Public Function AfterRecord(ListName)

Sheets(ListName).Select
ActiveSheet.Protect Password = "trytoguess", DrawingObjects:=True, Contents:=True, Scenarios:=True

End Function

Public Function BeforeRecord(ListName)

Sheets(ListName).Select
ActiveSheet.Unprotect Password = "trytoguess"

End Function

Public Function TokenSum() As Long
On Error GoTo ExceptionControl:
TokenSum = 0
For i = 9 To ActiveWorkbook.Sheets.Count
    Sheets(i).Select
    TokenSum = Cells(2, 1).Value + TokenSum
Next

Exit Function
ExceptionControl:
ErrorForm.Error_Box.Value = "Main/TokenSum()"
ErrorForm.Show
End Function
Public Function PointFilter(Val, Optional AllowNeg As Boolean = True, Optional AllowPoint As Boolean = True, Optional MaxLength As Integer = 9) As String
On Error GoTo ExceptionControl:
PointFilter = Val
String_Len = Len(Val)
If String_Len = 0 Then Exit Function
LastChar = Right(Val, 1)
If LastChar = "1" Or _
   LastChar = "2" Or _
   LastChar = "3" Or _
   LastChar = "4" Or _
   LastChar = "5" Or _
   LastChar = "6" Or _
   LastChar = "7" Or _
   LastChar = "8" Or _
   LastChar = "9" Or _
   LastChar = "0" Or _
   LastChar = "-" Or _
   LastChar = Application.DecimalSeparator _
Then
    If (LastChar = "-") Then If (AllowNeg = False) Or (String_Len > 1) Then PointFilter = Left(Val, String_Len - 1)
    If (LastChar = Application.DecimalSeparator) Then _
     If (AllowPoint = False) Or (InStr(1, Val, Application.DecimalSeparator) <> String_Len) Then PointFilter = Left(Val, String_Len - 1)
    If String_Len > MaxLength Then PointFilter = Left(Val, String_Len - 1)
Else
    PointFilter = Left(Val, String_Len - 1)
End If

Exit Function
ExceptionControl:
ErrorForm.Error_Box.Value = "Main/PointFilter()"
ErrorForm.Show
End Function

Public Sub TransferBalanceToNextMonth(ByVal Name, ByVal Balance)
On Error GoTo ExceptionControl:
Windows("Workers.xls").Activate
If GetWorkerID(Name) <> 0 Then
    Sheets(Name).Select
    Cells(2, 10).Value = Balance
End If
Windows("lWorkers.xls").Activate

Exit Sub
ExceptionControl:
ErrorForm.Error_Box.Value = "Main/TransferBalanceToNextMonth()"
ErrorForm.Show
End Sub

Public Function GetDayName(Num) As String
On Error GoTo ExceptionControl:
DateString = "1/" & CMonth & "/" & CYear
stDay = DateTime.Weekday(DateTime.DateValue(DateString))
ShowDay = Abs(-1 + Num + stDay - 1) Mod 7 + 1
GetDayName = DName(ShowDay)

Exit Function
ExceptionControl:
ErrorForm.Error_Box.Value = "Main/TransferBalanceToNextMonth()"
ErrorForm.Show
End Function

Public Function DName(Num) As String
Select Case Num
Case 2
       DName = "Понедельник"
Case 3
       DName = "Вторник"
Case 4
       DName = "Среда"
Case 5
       DName = "Четверг"
Case 6
       DName = "Пятница"
Case 7
       DName = "Суббота"
Case 1
       DName = "Воскресенье"
Case 0
       DName = "Ноль"
End Select
End Function

Public Function MName(Num) As String
Select Case Num
Case 1
       MName = "Январь"
Case 2
       MName = "Февраль"
Case 3
       MName = "Март"
Case 4
       MName = "Апрель"
Case 5
       MName = "Май"
Case 6
       MName = "Июнь"
Case 7
       MName = "Июль"
Case 8
       MName = "Август"
Case 9
       MName = "Сентябрь"
Case 10
       MName = "Октябрь"
Case 11
       MName = "Ноябрь"
Case 12
       MName = "Декабрь"
Case Else
       MName = "#Месяц не определён#"
End Select
End Function
Public Function MNameEng(Num) As String
Select Case Num
Case 1
       MNameEng = "Jan"
Case 2
       MNameEng = "Feb"
Case 3
       MNameEng = "Mar"
Case 4
       MNameEng = "Apr"
Case 5
       MNameEng = "May"
Case 6
       MNameEng = "Jun"
Case 7
       MNameEng = "Jul"
Case 8
       MNameEng = "Aug"
Case 9
       MNameEng = "Sep"
Case 10
       MNameEng = "Oct"
Case 11
       MNameEng = "Nov"
Case 12
       MNameEng = "Dec"
Case Else
       MNameEng = "#Error#"
End Select
End Function

Public Function MNameRusFix(Num) As String
Select Case Num
Case 1
       MNameRusFix = "Января"
Case 2
       MNameRusFix = "Февраля"
Case 3
       MNameRusFix = "Марта"
Case 4
       MNameRusFix = "Апреля"
Case 5
       MNameRusFix = "Мая"
Case 6
       MNameRusFix = "Июня"
Case 7
       MNameRusFix = "Июля"
Case 8
       MNameRusFix = "Августа"
Case 9
       MNameRusFix = "Сентября"
Case 10
       MNameRusFix = "Октября"
Case 11
       MNameRusFix = "Ноября"
Case 12
       MNameRusFix = "Декабря"
Case Else
       MNameRusFix = "#Месяц не определён#"
End Select
End Function
Public Function MDays(Num) As Integer
Select Case Num
Case 1
       MDays = 31
Case 2
       MDays = 28
Case 3
       MDays = 31
Case 4
       MDays = 30
Case 5
       MDays = 31
Case 6
       MDays = 30
Case 7
       MDays = 31
Case 8
       MDays = 31
Case 9
       MDays = 30
Case 10
       MDays = 31
Case 11
       MDays = 30
Case 12
       MDays = 31
Case Else
       MDays = 31
End Select
If CYear Mod 4 = 0 And Num = 2 Then MDays = 29
End Function

Public Function IsOpened(Fil) As Boolean
On Error GoTo ExceptionControl:
IsOpened = False
For i = 1 To Workbooks.Count
  If Workbooks(i).Name = Fil Then IsOpened = True
Next

Exit Function
ExceptionControl:
ErrorForm.Error_Box.Value = "Main/IsOpened()"
ErrorForm.Show
End Function
Public Function GetWorkerID(ByVal WorkerKey As String) As Integer
On Error GoTo ExceptionControl:
GetWorkerID = 0
Sheets("Сотрудники").Select
WeHaveWorkers = Cells(1, 2).Value
For i = 3 To CInt(WeHaveWorkers) + 3
    If WorkerKey = Cells(i, 3).Value Then
        GetWorkerID = i
        Exit Function
    End If
Next

Exit Function
ExceptionControl:
ErrorForm.Error_Box.Value = "Main/GetWorkerID()"
ErrorForm.Show
End Function
Public Function CutZ(Val As String) As Integer
On Error GoTo ExceptionControl:
CutZ = CInt(Left(Val, Len(Val) - 1))

Exit Function
ExceptionControl:
ErrorForm.Error_Box.Value = "Main/CutZ()"
ErrorForm.Show
End Function
Public Sub PullOnServer()
On Error GoTo ExceptionControl:
Dim PushArray(1 To 284), PullArray(1 To 284), CommentArray(1 To 284) As Boolean

PullBase = "pull.xls"
Sheets("Каталог").Select
LastMonthTokens = Cells(1, 6).Value
ThisMonthTokens = Cells(2, 6).Value

If Not IsOpened(PullBase) Then Workbooks.Open Filename:=Path + PullBase

Windows(PullBase).Activate
Sheets("Каталог").Select
PullYear = Cells(1, 3).Value
PullMonth = Cells(2, 3).Value
PulledTokens = Cells(2, 6).Value

If (PullYear <> CYear) Or (PullMonth <> CMonth) Then
    'pull from another month
    ActiveWorkbook.Close
    Windows(WorkersBase).Activate
Else
    If ThisMonthTokens <> PulledTokens Then
        PullLists = ActiveWorkbook.Sheets.Count
            For i = 9 To PullLists
                Windows(PullBase).Activate
                Sheets(i).Select
                PullToken = Cells(2, 1).Value
                LastDay = Cells(1, 1).Value
                DesiredDestination = Sheets(i).Name
                Windows(WorkersBase).Activate
                DestinationID = GetWorkerID(DesiredDestination)
                If DestinationID <> 0 Then
                    Sheets("Сотрудники").Select
                    Cells(DestinationID, 1).Value = 0
                    Sheets(DesiredDestination).Select
                    If Cells(2, 1).Value <> PullToken Then
                        Cells(2, 1).Value = PullToken
                        Cells(1, 1).Value = LastDay
                    
                        For j = 6 To 284
                            PushArray(j) = False
                            If Cells(j, 3).Value = "" Then PushArray(j) = True
                        Next j
                           
                        Sheets("Сотрудники").Select
                        Cells(DestinationID, 1).Value = 1
                        Windows(PullBase).Activate
                        Sheets(i).Select
                           
                        For j = 6 To 284
                           PullArray(j) = False
                           CommentArray(j) = False
                           If Cells(j, 2).Value <> "" Then PullArray(j) = True
                           If Cells(j, 13).Value <> "" Then CommentArray(j) = True
                        Next j
                     
                        For j = 6 To 284
                           If (PushArray(j) And PullArray(j)) = True Then
                                Windows(PullBase).Activate
                                Sheets(i).Select
                                CopyAlternateDiam = Cells(j, 14).Value
                                CopyComment = Cells(j, 13).Value
                                Range(Cells(j, 2), Cells(j, 9)).Copy
                                Windows(WorkersBase).Activate
                                Sheets(DesiredDestination).Select
                                Cells(j, 2).PasteSpecial
                                Cells(j, 14).Value = CopyAlternateDiam
                                Cells(j, 13).Value = CopyComment
                                Cells(j, 2).Select
                                Selection.EntireRow.Hidden = False
                                If Cells(j, 10).FormulaR1C1 = "" Then Cells(j, 10).FormulaR1C1 = "=SUM(RC[-1]:R[8]C[-1])"
                            Else
                                If (CommentArray(j) = True) Then
                                    Windows(PullBase).Activate
                                    Sheets(i).Select
                                    CopyComment = Cells(j, 13).Value
                                    Windows(WorkersBase).Activate
                                    Sheets(DesiredDestination).Select
                                    Cells(j, 13).Value = CopyComment
                                End If
                            End If
                        Next j
                        If LMMode Then TransferBalanceToNextMonth DesiredDestination, Cells(1, 10).Value
                    End If
                End If
            Next i
    Else
    'pull already done
    End If
    Windows(PullBase).Activate
    ActiveWorkbook.Close
    Windows(WorkersBase).Activate
    Sheets("Каталог").Select
    Cells(2, 6).Value = PulledTokens
    If LMMode Then
        Windows("Workers.xls").Activate
        Sheets("Каталог").Select
        Cells(1, 6).Value = PulledTokens
        Windows(WorkersBase).Activate
    End If
End If

Exit Sub
ExceptionControl:
ErrorForm.Error_Box.Value = "Main/PullOnServer()"
ErrorForm.Show
End Sub

Public Sub MainReInit()
On Error GoTo ExceptionControl:
  
If Not AdminMode Then
    WorkersBase = "tWorkers.xls"
    Form.Caption = "ООО ""Диск"" Система расчёта сдельной оплаты [Рабочее место] " & Version
    Workers.Bonus_Button.Visible = False
    Workers.BonusRate_Box.Visible = False
    Workers.Bonus_Label.Visible = False
    Workers.Logout_Button.Visible = True
    Workers.AboveOklad_Chk.Visible = False
    Workers.SelectUpdatesOnly.Visible = False
    Form.GenerateNextMonth.Enabled = False
    Form.SaveAndClose.Enabled = False
    Form.SaveState.Enabled = False
    Form.Setup_Button.Enabled = False
    Form.FeeReport_Button.Enabled = False
    Form.AvReport_Button.Enabled = False
    Form.Chamber_Button.Enabled = False
Else
    WorkersBase = "Workers.xls"
    Form.Caption = "ООО ""Диск"" Система расчёта сдельной оплаты [Администратор] " & Version
    Workers.Left_Box.Locked = False
    Workers.Rate_Box.Enabled = True
    Workers.Workers_Spin.Enabled = True
    Workers.PrePay_Box.Enabled = True
    Workers.Bonus_Button.Visible = True
    Workers.BonusRate_Box.Visible = True
    Workers.Bonus_Label.Visible = True
    Workers.OnScreen_Chk.Enabled = True
    Workers.Oklad_Box.Enabled = True
    Workers.Logout_Button.Visible = False
End If

Path = Workbooks("Index.xls").Path + "\"

If Not IsOpened(WorkersBase) Then Workbooks.Open Filename:=Path + WorkersBase

If IsOpened("lWorkers.xls") Then
    WorkersBase = "lWorkers.xls"
    LMMode = True
Else
    LMMode = False
End If

Windows(WorkersBase).Activate
Sheets("Каталог").Select
CYear = Cells(1, 3).Value
CMonth = Cells(2, 3).Value
 
LMonth = CMonth - 1
NMonth = CMonth + 1
If LMonth = 0 Then LMonth = 12
If NMonth = 13 Then NMonth = 1
  
NextMonth = MName(NMonth)

Form.GenerateNextMonth.Caption = "Перейти на " & NextMonth
If LMMode Then Form.SwitchToLastMonth.Caption = "Закрыть " & MName(CMonth) Else _
                                 Form.SwitchToLastMonth.Caption = "Открыть " & MName(LMonth)

If AdminMode Then
    If LMMode Then
        Form.GenerateNextMonth.Enabled = False
        Form.SaveAndClose.Enabled = False
        Form.SaveState.Enabled = False
        Form.Setup_Button.Enabled = False
    Else
        Form.GenerateNextMonth.Enabled = True
        Form.SaveAndClose.Enabled = True
        Form.SaveState.Enabled = True
        Form.Setup_Button.Enabled = True
    End If
End If

Exit Sub
ExceptionControl:
ErrorForm.Error_Box.Value = "MainReInit()"
ErrorForm.Show
End Sub
  
Public Sub FormShow()
On Error Resume Next
Form.Top = 0
Form.Left = 0
Form.Width = Round(GetSystemMetrics32(0) * 72 / 96)
Form.Height = Round(GetSystemMetrics32(1) * 72 / 96)
MainReInit
Form.Show

''If ReportExit = False And WorkersExit = False Then BlockIt.Show

'ReportExit = False

'If WorkersExit = True Then
'    WorkersExit = False
 '   Workers.Show
'   End If

'If WorkersExit = False Then Form.Show

End Sub

Sub Choose()
Attribute Choose.VB_ProcData.VB_Invoke_Func = "q\n14"
FormShow
End Sub


Public Function Recovery()

BlockIt.Show

Archive.Show
End Function
