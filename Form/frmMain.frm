VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.MDIForm frmMain 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFFFFF&
   Caption         =   "�칫��Ʒ����ϵͳ"
   ClientHeight    =   9210
   ClientLeft      =   285
   ClientTop       =   705
   ClientWidth     =   11400
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   2  '��Ļ����
   WindowState     =   2  'Maximized
   Begin MSComDlg.CommonDialog Dlg 
      Left            =   2265
      Top             =   3375
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox picSB 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   300
      Left            =   0
      ScaleHeight     =   20
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   760
      TabIndex        =   0
      Top             =   8910
      Width           =   11400
      Begin VB.Image Image2 
         Height          =   240
         Left            =   75
         Picture         =   "frmMain.frx":1982
         Top             =   45
         Width           =   240
      End
      Begin VB.Image Image1 
         Height          =   240
         Left            =   3150
         Picture         =   "frmMain.frx":1D0C
         Top             =   45
         Width           =   240
      End
      Begin VB.Label LBSB 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   180
         Index           =   2
         Left            =   3465
         TabIndex        =   2
         Top             =   75
         Width           =   90
      End
      Begin VB.Label LBSB 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��ӭʹ�ñ�ϵͳ"
         Height          =   180
         Index           =   1
         Left            =   375
         TabIndex        =   1
         Top             =   75
         Width           =   1260
      End
      Begin VB.Shape Shb2 
         BorderColor     =   &H00A6A6A6&
         Height          =   270
         Left            =   3090
         Top             =   30
         Width           =   6885
      End
      Begin VB.Image imgLB 
         Height          =   180
         Left            =   10080
         MousePointer    =   8  'Size NW SE
         Top             =   120
         Width           =   180
      End
      Begin VB.Shape Shb1 
         BorderColor     =   &H00A6A6A6&
         Height          =   270
         Left            =   30
         Top             =   30
         Width           =   3015
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "�ļ�(&F)"
      Begin VB.Menu mnuDBBackUp 
         Caption         =   "�������ݿ�(&B)"
         Shortcut        =   ^B
      End
      Begin VB.Menu mnuDBResume 
         Caption         =   "�ָ����ݿ�(&R)"
      End
      Begin VB.Menu mnuFileSP2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuLogin 
         Caption         =   "�����û�(&L)"
         Shortcut        =   ^L
      End
      Begin VB.Menu mnuFileSP3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "�˳�(&X)"
         Index           =   0
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "��ͼ(&V)"
      Enabled         =   0   'False
      Begin VB.Menu mnuGuide 
         Caption         =   "������(&W)"
         Checked         =   -1  'True
         Shortcut        =   ^W
      End
      Begin VB.Menu mnuTB 
         Caption         =   "������(&T)"
         Checked         =   -1  'True
         Shortcut        =   ^T
      End
      Begin VB.Menu mnuSB 
         Caption         =   "״̬��(&H)"
         Checked         =   -1  'True
         Shortcut        =   ^H
      End
   End
   Begin VB.Menu mnuData 
      Caption         =   "���ݹ���(&M)"
      Begin VB.Menu mnuSale 
         Caption         =   "���õǼ�(&S)"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuBalance 
         Caption         =   "���õǼ�(&O)"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuDataManageSpec9 
         Caption         =   "-"
      End
      Begin VB.Menu mnuReturn 
         Caption         =   "�黹�Ǽ�(&L)"
      End
      Begin VB.Menu mnuDataManageSpec20 
         Caption         =   "-"
      End
      Begin VB.Menu mnuStock 
         Caption         =   "���Ǽ�(&I)"
         Shortcut        =   ^I
      End
      Begin VB.Menu mnuDataManageSpec0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuInv 
         Caption         =   "�����ʾ(&V)"
         Shortcut        =   ^V
      End
   End
   Begin VB.Menu mnuSysSetup 
      Caption         =   "ϵͳ����(&M)"
      Begin VB.Menu mnuParaManage 
         Caption         =   "������������(&K)"
         Shortcut        =   ^K
      End
      Begin VB.Menu mnuSysSetupSpec 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSysUsermanage 
         Caption         =   "�û�����(&U)"
         Shortcut        =   ^U
      End
   End
   Begin VB.Menu mnuTool 
      Caption         =   "ʵ�ù���(&T)"
      Begin VB.Menu mnuCalcu 
         Caption         =   "������(&C)"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuNote 
         Caption         =   "���±�(&Q)"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "����(&H)"
      Begin VB.Menu mnuContent 
         Caption         =   "����(&C)"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuHelpSP1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "���ڱ����(&A)"
         Shortcut        =   ^A
      End
   End
   Begin VB.Menu mnutc 
      Caption         =   "�˳�"
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'�϶������API
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Dim CanResize As Boolean
Public LastFrm As Long

Private Sub cmdAbout_Click()
    mnuContent_Click
End Sub

Private Sub cmdClose_Click()
    picLeft.Visible = False
    mnuGuide.Checked = False
    SaveINI "Main", "Guide", "n"
End Sub

Public Sub cmdLeft_Click(Index As Integer)
End Sub

Private Sub imgLB_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Button = 1 Then
        Call ReleaseCapture
        Call SendMessage(hwnd, &HA1, 17, 0)
    End If
End Sub

Private Sub imgLogin_Click()

End Sub

Private Sub MDIForm_Load()
    Dim rs As ADODB.Recordset
    Dim strSQL As String
    Dim boolIsRun As Boolean    '������ϵͳ
    
   
   '��ȡ����λ��,��ͼ��Ϣ
    If GetINI("Main", "Left") = "" Then
        Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2
    Else
        Move GetLongINI("Main", "Left"), GetLongINI("Main", "Top"), GetLongINI("Main", "Width"), GetLongINI("Main", "Height")
        Dim j As Long
        j = GetLongINI("Main", "WindowState")
        If j = 2 Then Me.WindowState = 2
    End If
    CanResize = True
    If GetINI("Main", "Guide") = "n" Then
        picLeft.Visible = False
        mnuGuide.Checked = False
    End If
    If GetINI("Main", "ToolBar") = "n" Then
        picTB.Visible = False
        mnuTB.Checked = False
    End If
    If GetINI("Main", "StateBar") = "n" Then
        picSB.Visible = False
        mnuSB.Checked = False
    End If
    
    Set rs = New ADODB.Recordset
    DBConnect
    rs.Open "select  * from startSys where sIsStart=true", Conn, 1, 1
    
    boolIsRun = False
    
    If rs.RecordCount < 1 Then GoTo contin
        
    boolIsRun = rs("sIsStart")
    
contin:
    
    rs.Close
    Set rs = Nothing
    Conn.Close
    
    
End Sub

Private Sub MDIForm_Resize()
On Error Resume Next
    If CanResize = False Then Exit Sub
    If Me.Width < 9900 Then Me.Width = 9900
    If Me.Height < 8370 Then Me.Height = 8370
    SaveINI "Main", "WindowState", CStr(WindowState)
    If Me.WindowState = 0 Then
        SaveINI "Main", "Width", CStr(Width)
        SaveINI "Main", "Height", CStr(Height)
    End If
    picSB_Resize
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
On Error Resume Next
    DBConnect
    Conn.Execute "update options set list1Index=" & curList1Index & ",list2index=" & curList2Index & ",list3index=" & curList3Index & ",List5index=" & curList5Index
    Conn.Close
    Set frmMain = Nothing
End Sub

Private Sub mnuAbout_Click()
    MsgBox "�칫��Ʒ����ϵͳ V1.0" & Chr(13) & Chr(13) & _
          "������2011.10", vbInformation, "�칫��Ʒ����ϵͳ"
End Sub

Private Sub mnuBalance_Click()
    SetCmdState False
    blBorrow = False
    frmBorrow.Show
    
End Sub

Private Sub mnuBase_Click()
    SetCmdState False
    frmBaseInput.Show
End Sub

Private Sub mnuCalcu_Click()
    Dim RetVal As String
    RetVal = Shell("C:\WINDOWS\system32\calc.exe", 1)

End Sub

Private Sub mnuContent_Click()
    frmHelp.Show vbModal
End Sub

Private Sub mnuDataGoods_Click()

End Sub

Private Sub mnuDBBackUp_Click()
    On Error GoTo errmsg
    
    If Conn.state <> 0 Then
        Conn.Close
    End If
    
    If DirExists(GetApp & "bak") = 0 Then
        MkDir GetApp & "bak"
    End If
    
    Dlg.Filter = "�칫��Ʒ�����ļ�(*.gds)|*.gds"
    Dlg.FileName = "DATA" & Format(Now(), "yyyy-mm-dd hh.mm.ss") & ".gds"
    Dlg.DialogTitle = "���ݱ���"
    Dlg.InitDir = GetApp & "bak"
    Dlg.CancelError = True
    Dlg.ShowSave
    
    FileCopy GetApp & "data.gds", Dlg.FileName
    MsgBox "���ݱ��ݳɹ���", vbInformation, "���ݱ���"
    
    Exit Sub

errmsg:
    If Err.number = 32755 Then Exit Sub     '32755���û����ȡ����ť
    MsgBox Err.Description, vbInformation, "���ݱ���"
End Sub

Private Sub mnuDBResume_Click()
    On Error GoTo errmsg
    
    If Conn.state <> 0 Then
        Conn.Close
    End If
    If DirExists(GetApp & "bak") <> 0 Then
        Dlg.InitDir = GetApp & "bak"
    End If
    
    Dlg.Filter = "�칫��Ʒ�����ļ�(*.gds)|*.gds"
    Dlg.DialogTitle = "���ݻָ�"
    Dlg.CancelError = True
    Dlg.ShowOpen
    
    If MsgBox("���棺���ݻָ�����" & Dlg.FileName & "�����ݸ������������ݡ�", vbExclamation + vbYesNo, "���ݻָ�") = vbNo Then Exit Sub
    If MsgBox("ȷ�Ͻ������ݻָ���?", vbExclamation + vbYesNo, "���ݻָ�") = vbNo Then Exit Sub
    FileCopy Dlg.FileName, GetApp & "data.mik"
    MsgBox "���ݻָ��ɹ���", vbInformation, "���ݻָ�"
    
    
    Exit Sub

errmsg:
    If Err.number = 32755 Then Exit Sub     '32755���û����ȡ����ť
    MsgBox Err.Description, vbInformation, "���ݻָ�"

End Sub

Private Sub mnuExIncome_Click()
    'On Error GoTo errmsg
    Dim xlApp As excel.Application
    Dim xlBook As excel.Workbook
    Dim xlSheet As excel.Worksheet
    Dim xlRange As excel.Range
    Dim rs As ADODB.Recordset
    Dim rsIncome As ADODB.Recordset
    Dim strSQL As String
    Dim i, row, startRow, n As Integer
    Dim strFormat As String
    Dim strHTBH, strXMBH As String '��ͬ���,��Ŀ���
    Dim dblTotal As Double    '��֧���
    
    startRow = 3  '�ӵ�3�п�ʼ���
    
    Set rs = New ADODB.Recordset
    Set rsIncome = New ADODB.Recordset
    DBConnect
    
    If DirExists(GetApp & "Doc") = 0 Then
        MkDir GetApp & "Doc"
    End If
    
    strSQL = "select  id,htbh,htmc,htzj,jsj" & " " & _
             "from main" & " " & _
             "order by main.lrrq desc"
    

    rs.Open strSQL, Conn, 1, 1
    If rs.EOF Then
        MsgBox "δ�ҵ���ؼ�¼��������ֹ��", vbExclamation, "�����տ����һ����"
        rs.Close
        Conn.Close
        Exit Sub
    End If
    
    Dlg.Filter = "MS Excel�ļ�(*.xls)|*.xls"
    Dlg.FileName = "�տ����һ����(" & Format(Now(), "yyyy-mm-dd") & ")"
    Dlg.DialogTitle = "�����տ����һ����"
    Dlg.InitDir = GetApp & "Doc"
    Dlg.CancelError = True
    Dlg.ShowSave
    
    strFormat = ";;;##,##0.00;##,##0.00;yyyy��mm��dd��;##,##0.00;##,##0.00"
    arrayFormat = Split(strFormat, ";")
    
    Set xlApp = CreateObject("Excel.Application")
    Set xlBook = xlApp.Workbooks.Open(GetApp & "templets\�տ����һ����.xls")
    xlApp.Visible = False
    Set xlSheet = xlBook.Worksheets("Sheet1")
    
    strXMBH = ""    '��Ŀ���
    strHTBH = ""   '��ͬ���
    n = 0
    row = 0
    
    
    
    Do While Not rs.EOF
        n = n + 1
        
        xlSheet.Cells(startRow + row, 1) = Trim(CStr(n))  '��4�У�1��
        
        If IsNull(rs("jsj")) Then      'Ԥ���֧���
            dblTotal = 0
        Else
            dblTotal = CDbl(rs("jsj"))
        End If
            
        For i = 1 To 4 '1-��ͬ���,....
            If Not IsNull(rs.Fields(i).value) Then
                xlSheet.Cells(startRow + row, 1 + i) = IIf(arrayFormat(i) <> "", Format(CStr(rs.Fields(i).value), arrayFormat(i)), rs.Fields(i).value)
                    
            End If
        Next
        
        strSQL = "select skrq,skje from income where zhtid=" & rs("id") & " order by skrq"
        rsIncome.Open strSQL, Conn, 1, 1
            
    
        If rsIncome.RecordCount < 1 Then
            row = row + 1
        Else
        
            
            Do While Not rsIncome.EOF
            
                For i = 0 To 1    '�տ����
                    If Not IsNull(rsIncome.Fields(i).value) Then
                        xlSheet.Cells(startRow + row, 6 + i) = IIf(arrayFormat(5 + i) <> "", Format(CStr(rsIncome.Fields(i).value), arrayFormat(5 + i)), rsIncome.Fields(i).value)
                    End If
                
                Next
            
                If Not IsNull(rsIncome("skje")) Then    '�����տ����
                    dblTotal = dblTotal - CDbl(rsIncome("skje"))
                End If
                xlSheet.Cells(startRow + row, 8) = IIf(arrayFormat(7) <> "", Format(CStr(dblTotal), arrayFormat(7)), CStr(dblTotal))
                
                rsIncome.MoveNext
                row = row + 1
            Loop
            
            If rsIncome.RecordCount > 1 Then
                For i = 1 To 4
                    xlSheet.Range(xlSheet.Cells(startRow + row - 1, i), xlSheet.Cells(startRow + row - rsIncome.RecordCount, i)).Merge
                Next
                xlSheet.Range(xlSheet.Cells(startRow + row - 1, 9), xlSheet.Cells(startRow + row - rsIncome.RecordCount, 9)).Merge
            
            End If
        
        End If
        
        rsIncome.Close
        
        rs.MoveNext
    Loop
    
    Set xlRange = xlSheet.Range(xlSheet.Cells(startRow, 1), xlSheet.Cells(startRow + row - 1, 9))
    
    With xlRange.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With xlRange.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With xlRange.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With xlRange.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With xlRange.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With xlRange.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With

    
    
    rs.Close
    Conn.Close
    Set rs = Nothing
    Set Conn = Nothing
    xlBook.SaveAs Dlg.FileName
    xlBook.Close (True)
    xlApp.Quit
    Set xlApp = Nothing
    
    MsgBox "�տ����һ��������ɣ�" & Chr(13) & "���浽" & Dlg.FileName, vbInformation, "�����տ����һ����"
    
    Exit Sub

errmsg:
    If Err.number = 32755 Then Exit Sub     '32755���û����ȡ����ť
    MsgBox Err.Description, vbInformation, "�����տ����һ����"

End Sub
Private Sub mnuExItem_Click()
    On Error GoTo errmsg
    Dim xlApp As excel.Application
    Dim xlBook As excel.Workbook
    Dim xlSheet As excel.Worksheet
    Dim xlRange As excel.Range
    Dim rs, rsBorrow As ADODB.Recordset
    Dim strSQL As String
    Dim i, row, startRow, n As Integer
    Dim strFormat As String
    Dim strHTBH, strXMBH As String '��ͬ���,��Ŀ���
    Dim dblBalace As Double    '��֧���
    
    startRow = 3  '�ӵ�3�п�ʼ���
    
    Set rs = New ADODB.Recordset
    Set rsBorrow = New ADODB.Recordset
    DBConnect
    
    If DirExists(GetApp & "Doc") = 0 Then
        MkDir GetApp & "Doc"
    End If
    
    strSQL = "select  sub.yjs,sub.xmbh,main.wtdw,main.wtdwlxr,main.wtdwlxdh,sub.xmmc,sub.clr," & _
                  "sub.jcrq,sub.tcrq,sub.ysjzje,sub.jsj,sub.jsrq,sub.id" & " " & _
             "from main,sub" & " " & _
             "where main.id=sub.zhtid" & " " & _
             "order by main.lrrq desc,sub.xmbh"
    

    rs.Open strSQL, Conn, 1, 1
    If rs.EOF Then
        MsgBox "δ�ҵ���ؼ�¼��������ֹ��", vbExclamation, "������Ŀ����"
        rs.Close
        Conn.Close
        Exit Sub
    End If
    
    Dlg.Filter = "MS Excel�ļ�(*.xls)|*.xls"
    Dlg.FileName = "��Ŀ����(" & Format(Now(), "yyyy-mm-dd") & ")"
    Dlg.DialogTitle = "������Ŀ����"
    Dlg.InitDir = GetApp & "Doc"
    Dlg.CancelError = True
    Dlg.ShowSave
    
    strFormat = ";;;;;;;yyyy��mm��dd��;yyyy��mm��dd��;##,##0.00;yyyy��mm��dd��;##,##0.00;##,##0.00;##,##0.00;yyyy��mm��dd��"
    arrayFormat = Split(strFormat, ";")
    
    Set xlApp = CreateObject("Excel.Application")
    Set xlBook = xlApp.Workbooks.Open(GetApp & "templets\��Ŀ����.xls")
    xlApp.Visible = False
    Set xlSheet = xlBook.Worksheets("Sheet1")
    
    strXMBH = ""    '��Ŀ���
    strHTBH = ""   '��ͬ���
    n = 0
    row = 1
    
    
    
    Do While Not rs.EOF
        n = n + 1
        
        xlSheet.Cells(startRow + row, 1) = Trim(CStr(n))  '��4�У�1��
        xlSheet.Cells(startRow + row, 2) = IIf(rs("yjs"), "��", "��") '��4�У�2��
        If rs("yjs") Then xlSheet.Cells(startRow + row, 2).Font.ColorIndex = 3
        
        If IsNull(rs("ysjzje")) Then      'Ԥ���֧���
            dblBalace = 0
        Else
            dblBalace = CDbl(rs("ysjzje"))
        End If
            
        For i = 1 To 9 '1-��Ŀ���,....9-Ԥ���֧���
            If Not IsNull(rs.Fields(i).value) Then
                xlSheet.Cells(startRow + row, 2 + i) = IIf(arrayFormat(i) <> "", Format(CStr(rs.Fields(i).value), arrayFormat(i)), rs.Fields(i).value)
                    
            End If
        Next
        
        For i = 10 To 11   '10-�����,11-��������
            If Not IsNull(rs.Fields(i).value) Then
                xlSheet.Cells(startRow + row, 5 + i) = IIf(arrayFormat(3 + i) <> "", Format(CStr(rs.Fields(i).value), arrayFormat(3 + i)), rs.Fields(i).value)
            End If
        Next
        
            
        strSQL = "select jzrq,jzje from borrow where zhtid=" & rs("id") & " order by jzrq"
        rsBorrow.Open strSQL, Conn, 1, 1
            
    
        If rsBorrow.RecordCount < 1 Then
            row = row + 1
        Else
        
            
            Do While Not rsBorrow.EOF
            
                For i = 0 To 1    '��֧���
                    If Not IsNull(rsBorrow.Fields(i).value) Then
                        xlSheet.Cells(startRow + row, 12 + i) = IIf(arrayFormat(10 + i) <> "", Format(CStr(rsBorrow.Fields(i).value), arrayFormat(10 + i)), rsBorrow.Fields(i).value)
                    End If
                
                Next
            
                If Not IsNull(rsBorrow("jzje")) Then    '�����֧���
                    dblBalace = dblBalace - CDbl(rsBorrow("jzje"))
                End If
                xlSheet.Cells(startRow + row, 14) = IIf(arrayFormat(12) <> "", Format(CStr(dblBalace), arrayFormat(12)), CStr(dblBalace))
                
                rsBorrow.MoveNext
                row = row + 1
            Loop
            
            If rsBorrow.RecordCount > 1 Then
                For i = 1 To 11
                    xlSheet.Range(xlSheet.Cells(startRow + row - 1, i), xlSheet.Cells(startRow + row - rsBorrow.RecordCount, i)).Merge
                Next
                For i = 15 To 16
                    xlSheet.Range(xlSheet.Cells(startRow + row - 1, i), xlSheet.Cells(startRow + row - rsBorrow.RecordCount, i)).Merge
                Next
            
            End If
        
        End If
        
        rsBorrow.Close
        
        rs.MoveNext
    Loop
    
    Set xlRange = xlSheet.Range(xlSheet.Cells(startRow, 1), xlSheet.Cells(startRow + row - 1, 16))
    
    With xlRange.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With xlRange.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With xlRange.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With xlRange.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With xlRange.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With xlRange.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With

    
    
    rs.Close
    Conn.Close
    Set rs = Nothing
    Set Conn = Nothing
    xlBook.SaveAs Dlg.FileName
    xlBook.Close (True)
    xlApp.Quit
    Set xlApp = Nothing
    
    MsgBox "��Ŀ���ϵ�����ɣ�" & Chr(13) & "���浽" & Dlg.FileName, vbInformation, "������Ŀ����"
    
    Exit Sub

errmsg:
    If Err.number = 32755 Then Exit Sub     '32755���û����ȡ����ť
    MsgBox Err.Description, vbInformation, "������Ŀ����"



End Sub

Private Sub mnuExit_Click(Index As Integer)
    Unload Me
End Sub

Private Sub mnuGuide_Click()
    mnuGuide.Checked = Not mnuGuide.Checked
    picLeft.Visible = mnuGuide.Checked
    SaveINI "Main", "Guide", IIf(mnuGuide.Checked = True, "", "n")
End Sub

Private Sub mnuLeft_Click(Index As Integer)
    cmdLeft_Click Index
End Sub

Private Sub mnuInv_Click()
    blStocksShow = True
    frmStocks.Show
End Sub

Private Sub mnuLogin_Click()
On Error Resume Next
    Unload Me
    frmLogin.Show
End Sub

Private Sub mnuNote_Click()
    Dim RetVal As String
    RetVal = Shell("C:\WINDOWS\NOTEPAD.EXE", 1)

End Sub

Private Sub mnuParaManage_Click()
    frmOption.Show
End Sub

Private Sub mnuReturn_Click()
    frmReturn.Show
End Sub

Private Sub mnuSale_Click()
    SetCmdState False
    blBorrow = True
    frmBorrow.Show
End Sub

Private Sub mnuStock_Click()
    SetCmdState False
    frmIn.Show
End Sub

Private Sub mnuSysUsermanage_Click()
    frmUser.Show
End Sub

Private Sub mnutc_Click()
    Unload Me
End Sub

Private Sub picSB_Resize()
On Error Resume Next
    Shb2.Width = Me.Width / 15 - IIf(Me.WindowState = 2, 210, 230)
    imgLB.Visible = (Me.WindowState <> 2)
    imgLB.Left = Me.Width / 15 - 20
End Sub

Private Sub mnuSB_Click()
    mnuSB.Checked = Not mnuSB.Checked
    picSB.Visible = mnuSB.Checked
    SaveINI "Main", "StateBar", IIf(mnuSB.Checked = True, "", "n")
End Sub

Private Sub mnuTB_Click()
    mnuTB.Checked = Not mnuTB.Checked
    picTB.Visible = mnuTB.Checked
    SaveINI "Main", "ToolBar", IIf(mnuTB.Checked = True, "", "n")
End Sub

Private Sub picLeft_Resize()
On Error Resume Next
    ShLeft.Height = picLeft.Height / 15 - 23
End Sub

