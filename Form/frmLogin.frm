VERSION 5.00
Begin VB.Form frmLogin 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "办公用品管理系统登录"
   ClientHeight    =   2325
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4485
   Icon            =   "frmLogin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   155
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   299
   StartUpPosition =   2  '屏幕中心
   Begin goods.FCombo cboUser 
      Height          =   300
      Left            =   1425
      TabIndex        =   5
      Top             =   555
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   529
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      EnabledText     =   0   'False
      ListIndex       =   -1
   End
   Begin goods.FTextBox txtPW 
      Height          =   300
      Left            =   1425
      TabIndex        =   4
      Top             =   1035
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   529
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontName        =   "宋体"
      FontSize        =   9
      PasswordChar    =   "*"
   End
   Begin goods.XPButton cmdOK 
      Height          =   375
      Left            =   1785
      TabIndex        =   3
      Top             =   1755
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      Caption         =   "登录(&L)"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      cBack           =   -2147483633
   End
   Begin goods.XPButton cmdExit 
      Height          =   375
      Left            =   3105
      TabIndex        =   2
      Top             =   1755
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      Caption         =   "退出(&Q)"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      cBack           =   -2147483633
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "密  码："
      Height          =   180
      Left            =   585
      TabIndex        =   1
      Top             =   1095
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "用户名："
      Height          =   180
      Left            =   585
      TabIndex        =   0
      Top             =   615
      Width           =   720
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FFF8F0&
      BorderColor     =   &H00C5742F&
      Height          =   1335
      Left            =   285
      Top             =   255
      Width           =   3900
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cboUser_GotFocus()
    cboUser.SelAll
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub
Private Sub cmdOK_Click()
    If cboUser.Text = "" Then
        MsgBox "请填写用户名。", vbInformation
        cboUser.SetFocus
        cboUser.SetF
        Exit Sub
    End If
    If txtPW.Text = "" Then
        MsgBox "请填写密码。", vbInformation
        txtPW.SetFocus
        Exit Sub
    End If
'On Error GoTo aaaa
    Dim rs As New ADODB.Recordset, strMD5 As String
    If Conn.state <> 0 Then Conn.Close
    DBConnect
    rs.Open "Select * From UserInfo Where uName='" & cboUser.Text & "'", Conn, 1, 1
    If Not rs.EOF Then
            If StrComp(rs("uName"), cboUser.Text, 1) = 0 And StrComp(rs("uPWD"), GetMD5(txtPW.Text), 1) = 0 Then
                curUserName = rs("uName")
                cboUser.AddItem curUserName, 0
                curUserID = rs("uid")
                SaveUserList
                frmMain.Icon = Me.Icon
                Unload Me
                frmMain.Show
                Exit Sub
            End If
    End If
    MsgBox "用户名或密码错误，登陆失败！", vbCritical
    rs.Close
    Conn.Close
Exit Sub
aaaa:
    MsgBox Err.Description, vbCritical
    If Conn.state = 1 Then Conn.Close
End Sub

Private Sub cmdServer_Click()
    With frmServer
        .txtServer.Text = strSQLServer
        .txtUser.Text = strSQLUser
        If strSQLPW <> "" Then .lbPW.Visible = True
        .txtDB.Text = IIf(strSQLDB <> "", strSQLDB, "SuperMarketdb")
        .Show 1
    End With
End Sub

Private Sub Form_Activate()
On Error Resume Next
    cboUser.SetFocus
    cboUser.SetF
    If Conn.state <> 0 Then Conn.Close
    LoadUserList
 
    If cboUser.ListCount > 0 Then cboUser.ListIndex = 0
    
    txtPW.SetFocus
End Sub

Public Sub LoadUserList()
On Error GoTo ErrProcess
    Dim rs As ADODB.Recordset
    Dim strSQL As String
    
    Set rs = New ADODB.Recordset
    
    strSQL = "select * from userInfo"
    DBConnect
    rs.Open strSQL, Conn, 1, 1
    
    If rs.EOF Then
        rs.Close
        Conn.Close
        Unload frmLogin
        frmCreateUser.Show
        
    Else
        Do Until rs.EOF
            cboUser.AddItem Trim(rs("uName"))
            rs.MoveNext
        Loop
    
    End If

    Exit Sub
    
ErrProcess:
    MsgBox Err.Description, vbInformation, "登录"
    
End Sub

Public Sub SaveUserList()
On Error GoTo aaaa
    Dim strTmp As String, i As Long, j As Long
    If cboUser.ListCount <= 0 Then Exit Sub
    For i = 0 To cboUser.ListCount - 1
        strTmp = strTmp & cboUser.List(i) & vbCrLf
        j = j + 1
        If j >= 10 Then Exit For
    Next
    Open GetApp & "Files\user.inf" For Output As #1
        Print #1, strTmp
    Close #1
aaaa:
End Sub

Private Sub txtPW_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        cmdOK_Click
    End If
End Sub
