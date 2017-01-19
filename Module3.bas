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
Function compareStr(add1, add2, criteria) '比对两个字符串是都含有特定字
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
 If compareStr(add1, add2, "区") And compareStr(add1, add2, "村") Then
   checkAdd = True
   Exit Function
 End If
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 If compareStr(add1, add2, "区") And compareStr(add1, add2, "社") Then
   checkAdd = True
   Exit Function
 End If
 ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 If compareStr(add1, add2, "区") And compareStr(add1, add2, "塘") Then
   checkAdd = True
   Exit Function
 End If
 ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 If compareStr(add1, add2, "区") And compareStr(add1, add2, "苑") Then
   checkAdd = True
   Exit Function
 End If
 ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 If compareStr(add1, add2, "区") And compareStr(add1, add2, "家园") Then
   checkAdd = True
   Exit Function
 End If
 ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 If compareStr(add1, add2, "区") And compareStr(add1, add2, "居委会") Then
   checkAdd = True
   Exit Function
 End If
  '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 If compareStr(add1, add2, "区") And compareStr(add1, add2, "公寓") Then
   checkAdd = True
   Exit Function
 End If
  '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 If compareStr(add1, add2, "区") And compareStr(add1, add2, "华庭") Then
   checkAdd = True
   Exit Function
 End If
   '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 If compareStr(add1, add2, "区") And compareStr(add1, add2, "大厦") Then
   checkAdd = True
   Exit Function
 End If
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 If compareStr(add1, add2, "区") And compareStr(add1, add2, "工业园") Then
   checkAdd = True
   Exit Function
 End If
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 If compareStr(add1, add2, "区") And compareStr(add1, add2, "花园") Then
   checkAdd = True
   Exit Function
 End If
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 If compareStr(add1, add2, "镇") And compareStr(add1, add2, "村") Then
   checkAdd = True
   Exit Function
 End If
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 If compareStr(add1, add2, "镇") And compareStr(add1, add2, "工业区") Then
   checkAdd = True
   Exit Function
 End If
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 If compareStr(add1, add2, "镇") And compareStr(add1, add2, "科技") Then
   checkAdd = True
   Exit Function
 End If
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 If compareStr(add1, add2, "区") And compareStr(add1, add2, "科技") Then
   checkAdd = True
   Exit Function
 End If
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 If compareStr(add1, add2, "镇") And compareStr(add1, add2, "庄") Then
   checkAdd = True
   Exit Function
 End If
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 If compareStr(add1, add2, "区") And compareStr(add1, add2, "庄") Then
   checkAdd = True
   Exit Function
 End If
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 If compareStr(add1, add2, "区") And compareStr(add1, add2, "广场") Then
   checkAdd = True
   Exit Function
 End If
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 If compareStr(add1, add2, "区") And compareStr(add1, add2, "商业中心") Then
   checkAdd = True
   Exit Function
 End If
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

err:
checkAdd = False
End Function
Sub TEST()
str1 = "山东省临沂市兰山区金雀山路159号金阳大厦"

str2 = "山东临沂兰山区兰山办事处临西三路与金一路交汇金阳大厦B座201室"

MsgBox compareStr(str1, str2, "村")
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
If Cells(i, 6) = "a.现居住地址" Then
result = checkAdd(str1, str4)
ElseIf Cells(i, 6) = "b.工作地址" Then
result = checkAdd(str2, str4)
ElseIf Cells(i, 6) = "e.户籍地址" Then
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
MsgBox "登记了" & k & "条记录，对于我暂时不能判断为正确的就留空了，需要你们人类自行核对。"
Unload UserForm1
Exit Sub
err:
Application.StatusBar = ""
UserForm1.Hide
Unload UserForm1
MsgBox "没有执行任务，因为这可能不是我要检查的表格。"

End Sub
Sub dcheck()
On Error GoTo err
Start = Timer
Application.StatusBar = "正在使用安吉莉娜洪荒之力拼命检查......"
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
MsgBox "成功完成了任务！耗时：" & Total & "秒。"
Exit Sub
err:
Application.StatusBar = ""
UserForm1.Hide
Unload UserForm1
MsgBox "没有执行任务，因为这可能不是我要检查的表格。"
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
    If r = 1 Then Exit Function '如果第一列没有数据那么退出程序
    Set Dic = CreateObject("scripting.dictionary")  '创建字典对象
    For i = 1 To r              '将第一列数据添加到字典的key值中
        Dic(CStr(Cells(i, col))) = ""
    Next
    sols = Dic.keys              '返回字典key的数组
    Set Dic = Nothing           '销毁对象
End Function
Function sameRows(col As Integer, str As String) 'get row numbers of the same values
On Error Resume Next
    Dim Dic
    Dim i As Integer, r As Integer
    r = ActiveWorkbook.ActiveSheet.Cells(65536, col).End(xlUp).Row
    If r = 1 Then Exit Function '如果第一列没有数据那么退出程序
    'r = UBound(arr())
    Set Dic = CreateObject("scripting.dictionary")  '创建字典对象
    For i = 1 To r              '将第一列数据添加到字典的key值中
      If Cells(i, col) = str Then
        Dic(CStr(i)) = ""
      End If
    Next
        sameRows = Dic.keys
    Set Dic = Nothing
End Function
Function d(lcs, theday) '获取某专员某天签到所有行号
Dim rownums(), Dic
rownums = sameRows(2, CStr(lcs))
r = UBound(rownums())
    Set Dic = CreateObject("scripting.dictionary")  '创建字典对象
    For i = 0 To r              '将第一列数据添加到字典的key值中
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
    Set Dic = CreateObject("scripting.dictionary")  '创建字典对象
    For i = 0 To r              '将第一列数据添加到字典的key值中
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






