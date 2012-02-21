VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmEmployee 
   Caption         =   "人员选择&管理"
   ClientHeight    =   4245
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6630
   Icon            =   "frmEmployee.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4245
   ScaleWidth      =   6630
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command3 
      Caption         =   "删除选择的员工"
      Height          =   375
      Left            =   2670
      TabIndex        =   3
      Top             =   3690
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "增加员工"
      Height          =   390
      Left            =   4590
      TabIndex        =   2
      Top             =   3690
      Width           =   1440
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   30
      Top             =   3705
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEmployee.frx":058A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEmployee.frx":0AD9
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton Command1 
      Caption         =   "选   择"
      Height          =   390
      Left            =   765
      TabIndex        =   1
      Top             =   3690
      Width           =   1440
   End
   Begin MSComctlLib.ListView lvEmployee 
      Height          =   3225
      Left            =   210
      TabIndex        =   0
      Top             =   240
      Width           =   6195
      _ExtentX        =   10927
      _ExtentY        =   5689
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      _Version        =   393217
      Icons           =   "ImageList1"
      SmallIcons      =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
End
Attribute VB_Name = "frmEmployee"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
    If lvEmployee.ListItems.Count > 0 Then
        frmBorrow.txtBorrow.Text = lvEmployee.SelectedItem.Text
        frmBorrow.txtBorrow.Tag = GetID(lvEmployee.SelectedItem.Key)
        
    End If
    
    Unload Me
End Sub

Private Sub Command2_Click()
    frmAddEmployee.Show vbModal
End Sub

Private Sub Command3_Click()
    If MsgBox("确定删除 " & lvEmployee.SelectedItem.Text & " 码？", vbYesNo + vbExclamation, "员工管理") = vbNo Then Exit Sub
    
    DBConnect
    Conn.Execute "delete from employee where eid=" & GetID(lvEmployee.SelectedItem.Key)
    
    loadList
    Command1.SetFocus
    
    
End Sub

Private Sub Form_Load()
    
    loadList
    
End Sub
Public Sub loadList()
    Dim item As ListItem
    Dim strSQL As String
    Dim rs As ADODB.Recordset
    
    Set rs = New ADODB.Recordset
    strSQL = "select * from employee order by eName"
    DBConnect
    rs.Open strSQL, Conn, adOpenDynamic, adLockBatchOptimistic
        
    lvEmployee.ListItems.Clear
    n = 1
    Do While Not rs.EOF
        Set item = lvEmployee.ListItems.Add(, rs("eid") & "k")
        item.Text = rs("eName")
        If rs("eSex") = "男" Then
            item.Icon = 1
        Else
            item.Icon = 2
        End If
        
        n = n + 1
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing
    
    If lvEmployee.ListItems.Count < 1 Then
        Command3.Enabled = False
    Else
        Command3.Enabled = True
    End If

End Sub
