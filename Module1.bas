Attribute VB_Name = "Module1"

Sub ����2()
Attribute ����2.VB_ProcData.VB_Invoke_Func = " \n14"
Dim fId As Integer '�Y�O���q��ƽг]Long
Dim bCnt As Integer
bCnt = CInt(InputBox("�п�J�ثe����t�Ϥu�@ï�ƶq"))
For fId = 1 To bCnt

Workbooks.Open Filename:=ThisWorkbook.Path & "\����" & fId & "�t.xlsx"

ActiveWorkbook.Sheets(1).Activate '�Ĥ@�i��Ұ�
'MsgBox ("���t�ϸ�Ʀ@" & ActiveSheet.UsedRange.Rows.Count & "��")
  
  '�бN���s�n�������K�W�b����U��
Sub ����1()
'
' ����1 ����
'

'
End Sub
Sub ����2()
'
' ����2 ����
'

'
    SolverOk SetCell:="$F$6", MaxMinVal:=1, ValueOf:=0, ByChange:="$F$4:$F$6", _
        Engine:=2, EngineDesc:="Simplex LP"
    SolverOk SetCell:="$F$6", MaxMinVal:=1, ValueOf:=0, ByChange:="$F$4:$F$6", _
        Engine:=2, EngineDesc:="Simplex LP"
    SolverSolve
End Sub

ActiveWorkbook.Save
ActiveWorkbook.Close

Next
End Sub

