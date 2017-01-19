VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Private Sub Workbook_Open()
On Error Resume Next
Windows("钉钉核查工具.xlsb").Visible = False
    Call MenuSetup(True)
End Sub
  
'-----------------------------------------------
'在Excel中添加菜单和菜单项按钮（Excel启动时候添加）
'-----------------------------------------------
Public Function MenuSetup(blSetUp As Boolean)
    Dim myMenu As CommandBarPopup
    Dim mycontrol As CommandBarControl
    Dim i As Integer
    Dim sMenuItemName As String     '菜单项的名称
    Dim sMenuItemFunc As String     '菜单项的调用的函数名称
    Dim strM As String              '菜单名称
    Dim strMenuItem() As String     '菜单项名称
   
    On Error Resume Next
      
    '初始化菜单项
    ReDim strMenuItem(3, 2)    'VBA数组下界从1开始
    '菜单项1
    strMenuItem(1, 1) = "打卡设备检查"
    strMenuItem(1, 2) = "dcheck"
    '菜单项2
    strMenuItem(2, 1) = "签到地址核查"
    strMenuItem(2, 2) = "docheckadd"
      
    Application.ScreenUpdating = False
      
    '---添加菜单1
    strM = "钉钉核查工具"
'    If Application.CommandBars.FindControl(strT) = True Then Application.CommandBars.FindControl(strT).Delete
    
    Set myMenu = Application.CommandBars(1).Controls(strM)       '判断我的菜单是否存在？
    If err Then
        err.Clear
        Set myMenu = Application.CommandBars(1).Controls.Add(Type:=msoControlPopup, temporary:=False)
        myMenu.Caption = strM
    End If
      
    If blSetUp Then
            '---添加菜单项目1
            For i = 1 To UBound(strMenuItem)      '数组第一维的大小
                sMenuItemName = strMenuItem(i, 1)
                sMenuItemFunc = strMenuItem(i, 2)
                  
                Set mycontrol = myMenu.Controls(sMenuItemName)   '判断子程序是否存在
                If err Then
                    err.Clear
                    Set mycontrol = myMenu.Controls.Add(Type:=msoControlButton, temporary:=False) '在菜栏最后位置增加一个按钮
                    With mycontrol
                        .Caption = sMenuItemName                    '菜单项显示名称
                        .OnAction = sMenuItemFunc                   '左键单击该菜单项按钮便运行的过程
                        .Style = msoButtonCaption                   '只显示文字
                    End With
                End If
            Next
    Else
        Application.CommandBars(1).Controls(strT).Delete
         
    End If
      
    Application.ScreenUpdating = True
    If err Then err.Clear
End Function
  
Function isSZnumber(number)
isSZnumber = False
For i = 1 To 4260
If Left(number, 7) = CStr(Sheets("sznumber").Cells(i, 1)) Then isSZnumber = True
Next i
End Function
Sub colorSzNum()
On Error Resume Next
For Each c In Selection
If Not c.Value = "" Then
If isSZnumber(c.Value) Then c.Font.Color = -16776961
End If
Next c
End Sub
