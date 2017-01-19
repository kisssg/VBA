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
Windows("�����˲鹤��.xlsb").Visible = False
    Call MenuSetup(True)
End Sub
  
'-----------------------------------------------
'��Excel�����Ӳ˵��Ͳ˵��ť��Excel����ʱ�����ӣ�
'-----------------------------------------------
Public Function MenuSetup(blSetUp As Boolean)
    Dim myMenu As CommandBarPopup
    Dim mycontrol As CommandBarControl
    Dim i As Integer
    Dim sMenuItemName As String     '�˵��������
    Dim sMenuItemFunc As String     '�˵���ĵ��õĺ�������
    Dim strM As String              '�˵�����
    Dim strMenuItem() As String     '�˵�������
   
    On Error Resume Next
      
    '��ʼ���˵���
    ReDim strMenuItem(3, 2)    'VBA�����½��1��ʼ
    '�˵���1
    strMenuItem(1, 1) = "���豸���"
    strMenuItem(1, 2) = "dcheck"
    '�˵���2
    strMenuItem(2, 1) = "ǩ����ַ�˲�"
    strMenuItem(2, 2) = "docheckadd"
      
    Application.ScreenUpdating = False
      
    '---���Ӳ˵�1
    strM = "�����˲鹤��"
'    If Application.CommandBars.FindControl(strT) = True Then Application.CommandBars.FindControl(strT).Delete
    
    Set myMenu = Application.CommandBars(1).Controls(strM)       '�ж��ҵĲ˵��Ƿ���ڣ�
    If err Then
        err.Clear
        Set myMenu = Application.CommandBars(1).Controls.Add(Type:=msoControlPopup, temporary:=False)
        myMenu.Caption = strM
    End If
      
    If blSetUp Then
            '---���Ӳ˵���Ŀ1
            For i = 1 To UBound(strMenuItem)      '�����һά�Ĵ�С
                sMenuItemName = strMenuItem(i, 1)
                sMenuItemFunc = strMenuItem(i, 2)
                  
                Set mycontrol = myMenu.Controls(sMenuItemName)   '�ж��ӳ����Ƿ����
                If err Then
                    err.Clear
                    Set mycontrol = myMenu.Controls.Add(Type:=msoControlButton, temporary:=False) '�ڲ������λ������һ����ť
                    With mycontrol
                        .Caption = sMenuItemName                    '�˵�����ʾ����
                        .OnAction = sMenuItemFunc                   '��������ò˵��ť�����еĹ���
                        .Style = msoButtonCaption                   'ֻ��ʾ����
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