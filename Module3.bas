Attribute VB_Name = "Module3"
Sub Macro1()
Attribute Macro1.VB_ProcData.VB_Invoke_Func = "l\n14"
'
' Macro1 Macro
'

'
If Not (Range("A:A").EntireColumn.Hidden) Then

    Range( _
        "A:A,D:D,M:M,N:N,O:O,Q:Q,R:R,S:S,V:V,W:W,X:X,Y:Y,Z:Z,AA:AA,AB:AB,AC:AC,AD:AD,AE:AE,AF:AF,AL:AL,AM:AM,AN:AN,AO:AO,AP:AP,AQ:AQ,AS:AS" _
        ).EntireColumn.Hidden = True
Else
    Range("A:AS").EntireColumn.Hidden = 0
End If
End Sub
Function compareStr(add1, add2, criteria) '�ȶ������ַ����Ƕ������ض���
On Error GoTo err
compareStr = False
str1 = Right(Split(add1, criteria)(0), 2)
str2 = Right(Split(add2, criteria)(0), 2)
' If (add1 Like ("*" & str2 & "*")) Or (add2 Like ("*" & str1 & "*")) Then
'  compareStr = True
' End If
 If (add1 Like ("*" & str2 & "*")) And Not add2 = Split(add2, criteria)(0) Then
 compareStr = True
 End If
 If (add2 Like ("*" & str1 & "*")) And Not add1 = Split(add1, criteria)(0) Then
 compareStr = True
 End If
 
Exit Function
err:
compareStr = False
End Function
Function checkAdd(add1, add2)
On Error GoTo err
checkAdd = False
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 If compareStr(add1, add2, "��") And compareStr(add1, add2, "��") Then
   checkAdd = True
   Exit Function
 End If
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 If compareStr(add1, add2, "��") And compareStr(add1, add2, "��") Then
   checkAdd = True
   Exit Function
 End If
 ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 If compareStr(add1, add2, "��") And compareStr(add1, add2, "��") Then
   checkAdd = True
   Exit Function
 End If
 ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 If compareStr(add1, add2, "��") And compareStr(add1, add2, "Է") Then
   checkAdd = True
   Exit Function
 End If
 ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 If compareStr(add1, add2, "��") And compareStr(add1, add2, "��԰") Then
   checkAdd = True
   Exit Function
 End If
 ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 If compareStr(add1, add2, "��") And compareStr(add1, add2, "��ί��") Then
   checkAdd = True
   Exit Function
 End If
  '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 If compareStr(add1, add2, "��") And compareStr(add1, add2, "��Ԣ") Then
   checkAdd = True
   Exit Function
 End If
  '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 If compareStr(add1, add2, "��") And compareStr(add1, add2, "��ͥ") Then
   checkAdd = True
   Exit Function
 End If
   '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 If compareStr(add1, add2, "��") And compareStr(add1, add2, "����") Then
   checkAdd = True
   Exit Function
 End If
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 If compareStr(add1, add2, "��") And compareStr(add1, add2, "��ҵ԰") Then
   checkAdd = True
   Exit Function
 End If
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 If compareStr(add1, add2, "��") And compareStr(add1, add2, "��԰") Then
   checkAdd = True
   Exit Function
 End If
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 If compareStr(add1, add2, "��") And compareStr(add1, add2, "��") Then
   checkAdd = True
   Exit Function
 End If
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 If compareStr(add1, add2, "��") And compareStr(add1, add2, "��ҵ��") Then
   checkAdd = True
   Exit Function
 End If
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 If compareStr(add1, add2, "��") And compareStr(add1, add2, "�Ƽ�") Then
   checkAdd = True
   Exit Function
 End If
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 If compareStr(add1, add2, "��") And compareStr(add1, add2, "�Ƽ�") Then
   checkAdd = True
   Exit Function
 End If
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 If compareStr(add1, add2, "��") And compareStr(add1, add2, "ׯ") Then
   checkAdd = True
   Exit Function
 End If
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 If compareStr(add1, add2, "��") And compareStr(add1, add2, "ׯ") Then
   checkAdd = True
   Exit Function
 End If
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 If compareStr(add1, add2, "��") And compareStr(add1, add2, "�㳡") Then
   checkAdd = True
   Exit Function
 End If
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 If compareStr(add1, add2, "��") And compareStr(add1, add2, "��ҵ����") Then
   checkAdd = True
   Exit Function
 End If
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

err:
checkAdd = False
End Function
Sub TEST()
str1 = "ɽ��ʡ��������ɽ����ȸɽ·159�Ž�������"

str2 = "ɽ��������ɽ����ɽ���´�������·���һ·�����������B��201��"

MsgBox compareStr(str1, str2, "��")
End Sub

Sub docheckadd()
On Error GoTo err
totalrow = maxRow(5)
k = 0
Load UserForm1
UserForm1.Show 0
For i = 1 To totalrow
str1 = Cells(i, 7)
str2 = Cells(i, 8)
str3 = Cells(i, 9)
str4 = Cells(i, 11) & Cells(i, 10)
If Cells(i, 6) = "a.�־�ס��ַ" Then
result = checkAdd(str1, str4)
ElseIf Cells(i, 6) = "b.������ַ" Then
result = checkAdd(str2, str4)
ElseIf Cells(i, 6) = "e.������ַ" Then
result = checkAdd(str3, str4)
Else
result = False
End If
If result = True And Cells(i, 34) = "" Then
Cells(i, 34) = "A"
k = k + 1
End If
UserForm1.ProgressBar1.Value = i / totalrow * 100
Next i
UserForm1.Hide
MsgBox "�Ǽ���" & k & "����¼����������ʱ�����ж�Ϊ��ȷ�ľ������ˣ���Ҫ�����������к˶ԡ�"
Unload UserForm1
Exit Sub
err:
Application.StatusBar = ""
UserForm1.Hide
Unload UserForm1
MsgBox "û��ִ��������Ϊ����ܲ�����Ҫ���ı��"

End Sub
Sub dcheck()
On Error GoTo err
Start = Timer
Application.StatusBar = "����ʹ�ð������Ⱥ��֮��ƴ�����......"
Dim x, y
x = sols(2)
y = sols(27)
Load UserForm1
UserForm1.Show 0
For i = 0 To UBound(x)

 For j = 0 To UBound(y)
 Call doCheckDevice(i, j)
 Next j
 UserForm1.ProgressBar1.Value = i / UBound(x) * 100

Next i
UserForm1.Hide
Unload UserForm1

Total = Timer - Start
Application.StatusBar = ""
MsgBox "�ɹ���������񣡺�ʱ��" & Total & "�롣"
Exit Sub
err:
Application.StatusBar = ""
UserForm1.Hide
Unload UserForm1
MsgBox "û��ִ��������Ϊ����ܲ�����Ҫ���ı��"
End Sub
Function maxRow(col As Integer) 'get max row index
    r = ActiveWorkbook.ActiveSheet.Cells(65536, col).End(xlUp).Row
maxRow = r
End Function
Function sols(col As Integer) 'get sole values of the very column
    Dim Dic
    Dim i As Integer, r As Integer
    Dim str As String
    r = ActiveWorkbook.ActiveSheet.Cells(65536, col).End(xlUp).Row
    If r = 1 Then Exit Function '�����һ��û��������ô�˳�����
    Set Dic = CreateObject("scripting.dictionary")  '�����ֵ����
    For i = 1 To r              '����һ��������ӵ��ֵ��keyֵ��
        Dic(CStr(Cells(i, col))) = ""
    Next
    sols = Dic.keys              '�����ֵ�key������
    Set Dic = Nothing           '���ٶ���
End Function
Function sameRows(col As Integer, str As String) 'get row numbers of the same values
On Error Resume Next
    Dim Dic
    Dim i As Integer, r As Integer
    r = ActiveWorkbook.ActiveSheet.Cells(65536, col).End(xlUp).Row
    If r = 1 Then Exit Function '�����һ��û��������ô�˳�����
    'r = UBound(arr())
    Set Dic = CreateObject("scripting.dictionary")  '�����ֵ����
    For i = 1 To r              '����һ��������ӵ��ֵ��keyֵ��
      If Cells(i, col) = str Then
        Dic(CStr(i)) = ""
      End If
    Next
        sameRows = Dic.keys
    Set Dic = Nothing
End Function
Function d(lcs, theday) '��ȡĳרԱĳ��ǩ�������к�
Dim rownums(), Dic
rownums = sameRows(2, CStr(lcs))
r = UBound(rownums())
    Set Dic = CreateObject("scripting.dictionary")  '�����ֵ����
    For i = 0 To r              '����һ��������ӵ��ֵ��keyֵ��
      If CStr(Cells(rownums(i), 27)) = theday Then
        Dic(CStr(rownums(i))) = ""
      End If
    Next
d = Dic.keys
Set Dic = Nothing
End Function

Sub doCheckDevice(lcsIndex, theDayIndex) 'check the very lcs' device of the day
Dim Dic
rownums = d(sols(2)(lcsIndex), sols(27)(theDayIndex))
r = UBound(rownums)
    Set Dic = CreateObject("scripting.dictionary")  '�����ֵ����
    For i = 0 To r              '����һ��������ӵ��ֵ��keyֵ��
        Dic(CStr(Cells(rownums(i), 12))) = ""
    Next
devicecode = Dic.keys
If UBound(devicecode) = 0 Then
For i = 0 To r
If Cells(rownums(i), 35) = "" Then Cells(rownums(i), 35) = "Y"
Next i
End If
Set Dic = Nothing

End Sub






