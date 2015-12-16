VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Setup 
   Caption         =   "Настройки"
   ClientHeight    =   5895
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8400
   OleObjectBlob   =   "Setup.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Setup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Read As Integer
Dim Detonate As Integer


' ##############################################Start WorkersTab surr###############################

Function wRecordInfo(ID)
'On Error GoTo Endd
Read = 1
If NameChooser.Value <> BaseName_Box.Value Then wUpdateBaseName
NameChooser.Value = BaseName_Box.Value
WorkersProcess (ID)
ScanWorkers (WorkersBase)

Endd:
Read = 0

End Function

Function ScanWorkers(Fil)
'On Error GoTo Start
Windows(Fil).Activate
Setup.WorkersTree.Visible = True
Setup.WorkersTreeHolder.Visible = True
Setup.WorkersTree.Nodes.Clear

Sheets("Каталог").Select
Total = Cells(4, 23).Value
For i = InfoOffset To CInt(InfoOffset - 1 + Total)
    Setup.WorkersTree.Nodes.Add(, , CStr(Cells(i, 24)) & "z", Cells(i, 23).Value).Sorted = True
    Setup.WorkersTree.Nodes(i - InfoOffset + 1).Tag = "Cat"
    
  Next
  
Sheets("Сотрудники").Select
  WeHaveWorkers = Cells(1, 2).Value
  For i = 3 To WeHaveWorkers + 3
     For p = 1 To Total
     If Setup.WorkersTree.Nodes(p).Key = CStr(Cells(i, 6).Value) & "z" Then
         Setup.WorkersTree.Nodes.Add(p, 4, Cells(i, 3), _
                       Cells(i, 2) & " " & Cells(i, 5)).Sorted = True
      End If
     Next p
 Next i
 
 p = 1
 Do While p < Setup.WorkersTree.Nodes.Count
 
  If Setup.WorkersTree.Nodes(p).Children = 0 And Setup.WorkersTree.Nodes(p).Tag = "Cat" Then
         Setup.WorkersTree.Nodes.Remove (p)
         p = p - 1
         Total = Total - 1
     End If
  p = p + 1
 Loop
   

Setup.WorkersTree.Tag = Total
'Setup.WorkersTree.Visible = False
Setup.WorkersTreeHolder.Visible = False


GoTo Endd:
Start:
Exception.Show
Endd:
End Function




Function wReadLockedInfo()
'On Error GoTo Endd
ID = GetWorkerID(NameChooser.Value)

Read = 1
LastName_Box.Value = Cells(ID, 2).Value
wCatChooser.Value = wCatChooser.List(Cells(ID, 6).Value - InfoOffset)
Names_Box.Value = Cells(ID, 5).Value
NewPin_Box.Value = ""
BaseName_Box.Value = NameChooser.Value
isHidden_wmark.Value = Cells(ID, 4)
Read = 0
Endd:
End Function

Function WorkersProcess(ID)
Sheets("Сотрудники").Select
Cells(ID, 2).Value = LastName_Box.Value
Cells(ID, 3).Value = BaseName_Box.Value

Cells(ID, 6).Value = wCatChooser.ListIndex + InfoOffset
Cells(ID, 5).Value = Names_Box.Value
If NewPin_Box.Value <> "" Then Cells(ID, 7).Value = BlockIt.CalcStr(NewPin_Box.Value)
NewPin_Box.Value = ""

If isHidden_wmark.Value = True Then Cells(ID, 4).Value = 1 _
        Else Cells(ID, 4).Value = 0

Sheets(BaseName_Box.Value).Select


Cells(1, 2).Value = LastName_Box.Value
Cells(2, 2).Value = Names_Box.Value
'Cells(1, 4).Value = wCatChooser.ListIndex + InfoOffset
'If isHidden_wmark.Value = True Then Cells(2, 1).Value = 1 _
        Else Cells(2, 1).Value = 0


End Function

Function wUpdateBaseName()
Sheets("Сотрудники").Select
ID = GetWorkerID(NameChooser.Value)
Cells(ID, 3).Value = BaseName_Box.Value

ActiveWorkbook.Sheets(NameChooser.Value).Name = BaseName_Box.Value
End Function

Private Sub BaseName_Box_Change()
If (BaseName_Box.Value <> "") And (BaseName_Box.Value <> NameChooser.Value) Then _
wAdd_Button.Enabled = True
If BaseName_Box.Value = NameChooser.Value Then wAdd_Button.Enabled = False
End Sub

Private Sub CommandButton1_Click()
FineTuning_Main.Show
End Sub


Private Sub wAdd_Button_Click()
If CheckBaseName(0) = 0 Then
Read = 1
Sheets("Образец").Select
Sheets("Образец").Copy After:=Sheets(ActiveWorkbook.Sheets.Count)
ActiveWorkbook.Sheets(ActiveWorkbook.Sheets.Count).Name = BaseName_Box.Value
Sheets("Сотрудники").Select
ID = Cells(1, 2).Value + 3
Cells(1, 2).Value = Cells(1, 2).Value + 1
WorkersProcess (ID)
ScanWorkers (WorkersBase)
NameChooser.Value = BaseName_Box.Value
Read = 0
wAdd_Button.Enabled = False
End If
End Sub

Private Sub wChange_Button_Click()
ID = GetWorkerID(NameChooser.Value)
If CheckBaseName(ID) = 0 Then wRecordInfo (ID)
End Sub
Function CheckBaseName(ExcludePosition)
CheckBaseName = 0
Sheets("Сотрудники").Select
  WeHaveWorkers = Cells(1, 2).Value
For i = 3 To WeHaveWorkers + 3
If BaseName_Box.Value = Cells(i, 3).Value And i <> ExcludePosition Then
   CheckBaseName = 1
   i = WeHaveWorkers + 4
   End If
Next
End Function

Private Sub WorkersTree_DblClick()
If WorkersTree.SelectedItem.Key <> "" And WorkersTree.SelectedItem.Tag <> "Cat" Then
     NameChooser.Value = WorkersTree.SelectedItem.Key
     WorkersTreeHolder.Visible = False
End If
End Sub

Private Sub NameChooser_Change()
If Read = 0 Then wReadLockedInfo
End Sub
Private Sub ChooseWorker_Click()
'For i = 1 To CInt(WorkersTree.Tag)
'  If WorkersTree.Nodes(i).Tag = "Cat" Then _
 '       WorkersTree.Nodes(i).Expanded = False
' Next
'WorkersTree.Visible = True
WorkersTreeHolder.Visible = True

WorkersTree.SetFocus

End Sub

Private Sub Frame3_Click()
WorkersTreeHolder.Visible = False

End Sub

Private Sub LastName_Box_Change()
If Read = 0 Then BaseName_Box.Value = LastName_Box.Value
End Sub
''################################### End WorkersTab Surr _#####################################################



''################################### Start WCatTab Surr _#####################################################

Function GetWCatID(CatName)

Sheets("Каталог").Select
Total = Cells(4, 23).Value

For i = InfoOffset To CInt(InfoOffset - 1 + Total)
 If Cells(i, 23).Value = CatName Then
 GetWCatID = CInt(Cells(i, 24).Value)
 i = 1 + CInt(InfoOffset - 1 + Total)
 End If

Next
If CatName = "" Then GetWCatID = 1

End Function

Function ScanWCats()
Sheets("Каталог").Select
Setup.cCatChooser.Clear
Setup.wCatChooser.Clear
Total = Cells(4, 23).Value
For i = InfoOffset To CInt(InfoOffset - 1 + Total)
    Setup.cCatChooser.AddItem (Cells(i, 23).Value)
    Setup.wCatChooser.AddItem (Cells(i, 23).Value)
    Next
End Function
Function wcReadLockedInfo()
On Error GoTo Endd

ID = GetWCatID(cCatChooser.Value)

CatName_Box.Value = cCatChooser.Value


Endd:
End Function

Private Sub CatName_Box_Change()
If (CatName_Box.Value <> "") And (CatName_Box.Value <> cCatChooser.Value) Then cAdd_Button.Enabled = True
If CatName_Box.Value = cCatChooser.Value Then cAdd_Button.Enabled = False
End Sub

Private Sub CatSpin_SpinUp()
If cCatChooser.ListIndex > 0 Then _
cCatChooser.Value = cCatChooser.List(cCatChooser.ListIndex - 1)

End Sub

Private Sub CatSpin_SpinDown()
If cCatChooser.ListIndex < cCatChooser.ListCount - 1 Then _
cCatChooser.Value = cCatChooser.List(cCatChooser.ListIndex + 1)
End Sub


Private Sub cCatChooser_Change()
If Read = 0 Then wcReadLockedInfo
End Sub
Function ProcessWCat(ID)
Read = 1

Cells(ID, 23).Value = CatName_Box.Value

Cells(ID, 24).Value = ID

ScanWCats
ScanWorkers (WorkersBase)

Read = 0
cCatChooser.Value = CatName_Box.Value
End Function
Function UpdateWCatName()
Read = 1
ID = GetWCatID(cCatChooser.Value)
Cells(ID, 23).Value = CatName_Box.Value
Read = 0
End Function

Private Sub cChange_Button_Click()
If CatName_Box.Value <> cCatChooser.Value Then UpdateWCatName
ProcessWCat (GetWCatID(CatName_Box.Value))
End Sub


Private Sub cAdd_Button_Click()
Read = 1

ID = Cells(4, 23).Value + InfoOffset
Cells(4, 23).Value = Cells(4, 23).Value + 1
ProcessWCat (ID)
Read = 0

cAdd_Button.Enabled = False

End Sub

''################################### End WCatTab Surr _#####################################################


''################################### Start JobTab Surr _#####################################################

Function ScanJobs()
Sheets("Каталог").Select
Setup.JobsTree.Visible = True
Setup.JobsTreeHolder.Visible = True

Setup.JobsTree.Nodes.Clear
TotalCat = Cells(4, 19).Value
For i = InfoOffset To CInt(InfoOffset - 1 + TotalCat)
    Setup.JobsTree.Nodes.Add(, , , Cells(i, 19).Value).Sorted = True
    Setup.JobsTree.Nodes(i - InfoOffset + 1).Tag = "Cat"
    Next

Total = Cells(4, 2).Value
For i = InfoOffset To CInt(InfoOffset - 1 + Total)
'  If Cells(i, 7) = 0 Then
  Setup.JobsTree.Nodes.Add(CInt(Cells(i, 1).Value - InfoOffset + 1), 4, CStr(Cells(i, 3)) & "z", Cells(i, 2).Value).Sorted = True
Next

 p = 1
 Do While p < Setup.JobsTree.Nodes.Count
 
  If Setup.JobsTree.Nodes(p).Children = 0 And Setup.JobsTree.Nodes(p).Tag = "Cat" Then
         Setup.JobsTree.Nodes.Remove (p)
         TotalCat = TotalCat - 1
         p = p - 1
     End If
  p = p + 1
 Loop
Setup.JobsTree.Tag = TotalCat
Setup.JobsTreeHolder.Visible = False

    
End Function
Private Sub Frame1_Click()
Setup.JobsTreeHolder.Visible = False
End Sub
Function jReadLockedInfo()
'On Error GoTo Endd
Sheets("Каталог").Select

ID = jID.Value

JobName_Box.Value = Cells(ID, 2).Value
CatChooser.Value = CatChooser.List(Cells(ID, 1).Value - 6)
Unit_Box.Value = Cells(ID, 4).Value
UnitRate_Box.Value = Cells(ID, 5)
TimeRate_Box.Value = Cells(ID, 6)
isHidden_jMark.Value = Cells(ID, 7).Value
OnSale_Mark.Value = Cells(ID, 8).Value
OnReport_Mark.Value = Cells(ID, 9).Value
Price.Value = Cells(ID, 10).Value


Endd:
End Function
Function ProcessJobs(ID)
Cells(ID, 2).Value = JobName_Box.Value
Cells(ID, 4).Value = Unit_Box.Value
Cells(ID, 5) = UnitRate_Box.Value
Cells(ID, 6) = TimeRate_Box.Value
Cells(ID, 1) = InfoOffset + CatChooser.ListIndex
Cells(ID, 10).Value = Price.Value

If isHidden_jMark.Value = False Then Cells(ID, 7).Value = 0 _
                               Else Cells(ID, 7).Value = 1

If OnSale_Mark.Value = False Then Cells(ID, 8).Value = 0 _
                             Else Cells(ID, 8).Value = 1

If OnReport_Mark.Value = False Then Cells(ID, 9).Value = 0 _
                               Else Cells(ID, 9).Value = 1
End Function


Private Sub ChooseJob_Click()
'For i = 1 To CInt(JobsTree.Tag)
'  If JobsTree.Nodes(i).Tag = "Cat" Then _
'        JobsTree.Nodes(i).Expanded = False
 'Next
Frame4.Visible = False
JobsTreeHolder.Visible = True
JobsTree.SetFocus

End Sub

Private Sub JobName_Box_Change()
If (JobName_Box.Value <> "") Then jAdd_Button.Enabled = True
End Sub

Private Sub jID_Change()
If Read = 0 Then jReadLockedInfo
End Sub

Private Sub JobsTree_DblClick()
If JobsTree.SelectedItem.Key <> "" Then
    jID.Value = CutZ(JobsTree.SelectedItem.Key)
    JobsTreeHolder.Visible = False
    End If
jAdd_Button.Enabled = False
End Sub
Private Sub TimeRate_Box_Change()
If TimeRate_Box.Value <> "" Then TimeRate_Box.Value = PointFilter(TimeRate_Box.Value, True, True)
If Detonate = 0 Then
 Detonate = 1
 UnitRate_Box.Value = "0"
 Detonate = 0
 End If
End Sub

Private Sub UnitRate_Box_Change()
 If UnitRate_Box.Value <> "" Then UnitRate_Box.Value = PointFilter(UnitRate_Box.Value, True, True)
 If Detonate = 0 Then
 Detonate = 1
 TimeRate_Box.Value = "0"
 Detonate = 0
 End If
End Sub
Private Sub JobSpin_SpinUp()
If JobChooser.ListIndex > 0 Then _
JobChooser.Value = JobChooser.List(JobChooser.ListIndex - 1)

End Sub
Private Sub JobSpin_SpinDown()
If JobChooser.ListIndex < JobChooser.ListCount - 1 Then _
JobChooser.Value = JobChooser.List(JobChooser.ListIndex + 1)
End Sub
Private Sub OnSale_Mark_Change()
If OnSale_Mark.Value = True Then Frame4.Visible = True
If OnSale_Mark.Value = False Then Frame4.Visible = False
End Sub

Private Sub Price_Change()
If Price.Value <> "" Then Price.Value = PointFilter(Price.Value)
End Sub
Private Sub jChange_Button_Click()
ProcessJobs (jID.Value)
ScanJobs
End Sub
Private Sub jAdd_Button_Click()
Read = 1
Sheets("Каталог").Select
ID = Cells(4, 2).Value + InfoOffset
ProcessJobs (ID)
Cells(ID, 3).Value = ID
Cells(ID, 2).Value = JobName_Box.Value
Cells(4, 2).Value = Cells(4, 2).Value + 1

ScanJobs

Read = 0

jAdd_Button.Enabled = False
jID.Value = ID
Endd:
End Sub


''################################### Start JCatTab Surr _#####################################################
Function jGetCatID(CatName)

Sheets("Каталог").Select
Total = Cells(4, 19).Value

For i = 6 To CInt(5 + Total)
 If Cells(i, 19).Value = CatName Then
 jGetCatID = CInt(Cells(i, 20).Value)
 i = 1 + CInt(5 + Total)
 End If

Next
If CatName = "" Then jGetCatID = 1

End Function

Function ScanJCats()
On Error GoTo Start
Setup.CatChooser.Clear
Setup.jCatChooser.Clear
Sheets("Каталог").Select
Total = Cells(4, 19).Value

For i = 6 To CInt(5 + Total)
    Setup.CatChooser.AddItem (Cells(i, 19).Value)
    Setup.jCatChooser.AddItem (Cells(i, 19).Value)
    
    Next
GoTo Endd:
Start:
Exception.Show
Endd:

End Function

Private Sub jCatName_Box_Change()
If (jCatName_Box.Value <> "") And (jCatName_Box.Value <> jCatChooser.Value) Then jcAdd_Button.Enabled = True
If jCatName_Box.Value = jCatChooser.Value Then jcAdd_Button.Enabled = False
End Sub
Private Sub jCatChooser_Change()
If Read = 0 Then jCatName_Box.Value = jCatChooser.Value
End Sub

Function ReInit()


ScanJCats
ScanJobs
Setup.jCatChooser.Value = jCatName_Box.Value

End Function

Private Sub jcChange_Button_Click()
If jCatChooser.Value <> jCatName_Box.Value Then UpdateJCatName
Read = 1
ReInit
Read = 0
End Sub
Private Sub jCatSpin_SpinUp()
If jCatChooser.ListIndex > 0 Then _
jCatChooser.Value = jCatChooser.List(jCatChooser.ListIndex - 1)

End Sub

Private Sub jCatSpin_SpinDown()
If jCatChooser.ListIndex < jCatChooser.ListCount - 1 Then _
jCatChooser.Value = jCatChooser.List(jCatChooser.ListIndex + 1)
End Sub



Private Sub jcAdd_Button_Click()
Read = 1

ID = Cells(4, 19).Value + 5 + 1


Cells(ID, 19).Value = jCatName_Box.Value
Cells(ID, 20).Value = ID

Cells(4, 19).Value = Cells(4, 19).Value + 1

ReInit
jcAdd_Button.Enabled = False

Read = 0
End Sub
Function UpdateJCatName()
Read = 1
ID = jGetCatID(jCatChooser.Value)
Cells(ID, 19).Value = jCatName_Box.Value
ReInit
Read = 0
End Function



''################################### Stop  JCatTab Surr _#####################################################

''################################### Start  OCatTab Surr _#####################################################
Function oGetCatID(CatName)

Sheets("Каталог").Select
Total = Cells(4, 31).Value

For i = 6 To CInt(5 + Total)
 If Cells(i, 31).Value = CatName Then
 oGetCatID = CInt(Cells(i, 32).Value)
 i = 1 + CInt(5 + Total)
 End If

Next
If CatName = "" Then oGetCatID = 1

End Function

Function ScanOCats()
On Error GoTo Start
Setup.OrgChooser.Clear
Setup.oCatChooser.Clear
Sheets("Каталог").Select
Total = Cells(4, 31).Value

For i = 6 To CInt(5 + Total)
    Setup.OrgChooser.AddItem (Cells(i, 31).Value)
    Setup.oCatChooser.AddItem (Cells(i, 31).Value)
    
    Next
GoTo Endd:
Start:
Exception.Show
Endd:

End Function

Private Sub oCatName_Box_Change()
If (oCatName_Box.Value <> "") And (oCatName_Box.Value <> oCatChooser.Value) Then ocAdd_Button.Enabled = True
If oCatName_Box.Value = oCatChooser.Value Then ocAdd_Button.Enabled = False
End Sub
Private Sub oCatChooser_Change()
If Read = 0 Then oCatName_Box.Value = oCatChooser.Value
End Sub

Function oReInit()


ScanOCats
ScanOrgs
Setup.oCatChooser.Value = oCatName_Box.Value

End Function

Private Sub ocChange_Button_Click()
If oCatChooser.Value <> oCatName_Box.Value Then UpdateOCatName
Read = 1
oReInit
Read = 0
End Sub
Function UpdateOCatName()
Read = 1
ID = oGetCatID(oCatChooser.Value)
Cells(ID, 31).Value = oCatName_Box.Value
oReInit
Read = 0
End Function

Private Sub oCatSpin_SpinUp()
If oCatChooser.ListIndex > 0 Then _
oCatChooser.Value = oCatChooser.List(oCatChooser.ListIndex - 1)

End Sub

Private Sub oCatSpin_SpinDown()
If oCatChooser.ListIndex < oCatChooser.ListCount - 1 Then _
oCatChooser.Value = oCatChooser.List(oCatChooser.ListIndex + 1)
End Sub
Private Sub ocAdd_Button_Click()
Read = 1

ID = Cells(4, 31).Value + 5 + 1
Cells(ID, 31).Value = oCatName_Box.Value
Cells(ID, 32).Value = ID
Cells(4, 31).Value = Cells(4, 31).Value + 1
oReInit
ocAdd_Button.Enabled = False

Read = 0
End Sub

''################################### Stop  JCatTab Surr _#####################################################

''################################### Start OgrTab Surr _#####################################################

Function ScanOrgs()
Sheets("Каталог").Select
Setup.OrgsTree.Visible = True
Setup.OrgsTreeHolder.Visible = True

Setup.OrgsTree.Nodes.Clear
TotalCat = Cells(4, 31).Value
For i = InfoOffset To CInt(InfoOffset - 1 + TotalCat)
    Setup.OrgsTree.Nodes.Add(, , , Cells(i, 31).Value).Sorted = True
    Setup.OrgsTree.Nodes(i - InfoOffset + 1).Tag = "Cat"
    Next

Total = Cells(4, 27).Value
For i = InfoOffset To CInt(InfoOffset - 1 + Total)
'  If Cells(i, 7) = 0 Then
  Setup.OrgsTree.Nodes.Add(CInt(Cells(i, 28).Value - InfoOffset + 1), 4, CStr(Cells(i, 25)) & "z", Cells(i, 27).Value).Sorted = True
Next

 p = 1
 Do While p < Setup.OrgsTree.Nodes.Count
 
  If Setup.OrgsTree.Nodes(p).Children = 0 And Setup.OrgsTree.Nodes(p).Tag = "Cat" Then
         Setup.OrgsTree.Nodes.Remove (p)
         TotalCat = TotalCat - 1
         p = p - 1
     End If
  p = p + 1
 Loop
Setup.OrgsTree.Tag = TotalCat
Setup.OrgsTreeHolder.Visible = False

    
End Function
Private Sub Frame7_Click()
Setup.OrgsTreeHolder.Visible = False
End Sub
Function oReadLockedInfo()
'On Error GoTo Endd
Sheets("Каталог").Select

ID = oID.Value

OrgName_Box.Value = Cells(ID, 27).Value
OrgChooser.Value = OrgChooser.List(Cells(ID, 28).Value - 6)
isHidden_oMark.Value = Cells(ID, 26).Value


Endd:
End Function
Function ProcessOrgs(ID)
Cells(ID, 27).Value = OrgName_Box.Value
Cells(ID, 28) = InfoOffset + OrgChooser.ListIndex

If isHidden_oMark.Value = False Then Cells(ID, 26).Value = 0 _
                               Else Cells(ID, 26).Value = 1

End Function


Private Sub ChooseOrg_Click()
'For i = 1 To CInt(JobsTree.Tag)
'  If JobsTree.Nodes(i).Tag = "Cat" Then _
'        JobsTree.Nodes(i).Expanded = False
 'Next
OrgsTreeHolder.Visible = True
OrgsTree.SetFocus

End Sub

Private Sub OrgName_Box_Change()
If (OrgName_Box.Value <> "") Then oAdd_Button.Enabled = True
End Sub

Private Sub oID_Change()
If Read = 0 Then oReadLockedInfo
End Sub

Private Sub OrgsTree_DblClick()
If OrgsTree.SelectedItem.Key <> "" Then
    oID.Value = CutZ(OrgsTree.SelectedItem.Key)
    OrgsTreeHolder.Visible = False
    End If
oAdd_Button.Enabled = False
End Sub
Private Sub oChange_Button_Click()
ProcessOrgs (oID.Value)
ScanOrgs
End Sub
Private Sub oAdd_Button_Click()
Read = 1
Sheets("Каталог").Select
ID = Cells(4, 27).Value + InfoOffset
ProcessOrgs (ID)
Cells(ID, 25).Value = ID
Cells(ID, 27).Value = OrgName_Box.Value
Cells(4, 27).Value = Cells(4, 27).Value + 1

ScanOrgs

Read = 0

oAdd_Button.Enabled = False
oID.Value = ID
Endd:
End Sub




Private Sub UserForm_Click()
WorkersTree.Visible = False
JobsTree.Visible = False
OrgsTree.Visible = False
End Sub


