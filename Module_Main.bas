Attribute VB_Name = "Module_MAin"
Public CYear, CMonth, LMonth, NMonth As Integer
Public WorkersBase, Path, NextMonth As String
Public AtLast As Boolean
Public ReportExit, WorkersExit As Boolean
Public LastWorkersDay As Integer
Public FiltersReady As Integer
Public LastPerson As String


Public Function AfterRecord(ListName)

Sheets(ListName).Select
ActiveSheet.Protect Password = "trytoguess", DrawingObjects:=True, Contents:=True, Scenarios:=True

End Function

Public Function BeforeRecord(ListName)

Sheets(ListName).Select
ActiveSheet.Unprotect Password = "trytoguess"

End Function

Public Function TransferBalanceToNextMonth(Name, Value)
On Error GoTo Endd
If IsOpened("lWorkers.xls") Then
Windows("Workers.xls").Activate
 Sheets(Name).Select
 Cells(2, 10).Value = Value
End If
Endd:
If IsOpened("lWorkers.xls") Then Windows("lWorkers.xls").Activate
End Function
Public Function GetDayName(Num) As String
On Error GoTo Endd:
DateString = "1/" & CMonth & "/" & CYear
stDay = DateTime.Weekday(DateTime.DateValue(DateString))
ShowDay = _
Abs(-1 + Num + stDay - 1) Mod 7 + 1
GetDayName = DName(ShowDay)
Endd:
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
IsOpened = False
For i = 1 To Workbooks.Count
  If Workbooks(i).Name = Fil Then IsOpened = True
Next
End Function

Public Function FormShow()
  On Error Resume Next
  
  ''If ReportExit = False And WorkersExit = False Then BlockIt.Show
  BlockIt.Exit_b.Visible = False
  
   WorkersBase = "Workers.xls"
   
   If IsOpened("lWorkers.xls") Then WorkersBase = "lWorkers.xls"
   
   
   Path = Workbooks("Index.XLS").Path + "\"
   
   If Not IsOpened(WorkersBase) Then Workbooks.Open FileName:=Path + WorkersBase
       
       
 Windows(WorkersBase).Activate
 
 Sheets("Каталог").Select
  CYear = Cells(1, 3).Value
  CMonth = Cells(2, 3).Value
  
   
   Form.SwitchToLastMonth.Caption = "Закрыть базы данных за " & MName(CMonth)
   Form.GenerateNextMonth.Enabled = False
   Form.SaveAndClose.Enabled = False
   Form.SaveState.Enabled = False
   
  LMonth = CMonth - 1
  NMonth = CMonth + 1
  If LMonth = 0 Then LMonth = 12
  If NMonth = 13 Then NMonth = 1
  
  NextMonth = MName(NMonth)
  If Not IsOpened("lWorkers.xls") Then
  Form.SwitchToLastMonth.Caption = "Открыть базы данных за " & MName(LMonth)
  Form.GenerateNextMonth.Enabled = True
  Form.GenerateNextMonth.Caption = "Создать базы данных за " & NextMonth
  Form.SaveAndClose.Enabled = True
  Form.SaveState.Enabled = True
   
  End If

''If ReportExit = False And WorkersExit = False Then BlockIt.Show

ReportExit = False

If WorkersExit = True Then
    WorkersExit = False
    Workers.Show
    End If

If WorkersExit = False Then Form.Show

End Function

Sub Choose()
Attribute Choose.VB_ProcData.VB_Invoke_Func = "q\n14"
FormShow
End Sub


Public Function Recovery()

BlockIt.Show

Archive.Show
End Function
