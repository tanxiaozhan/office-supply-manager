VERSION 5.00
Begin VB.Form frmCreateUser 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "建立用户"
   ClientHeight    =   3090
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5295
   Icon            =   "frmCreatUser.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   5295
   StartUpPosition =   2  '屏幕中心
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   4080
      Picture         =   "frmCreatUser.frx":058A
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   9
      ToolTipText     =   "长度为1-16位的中文、英文或数字组合"
      Top             =   1100
      Width           =   240
   End
   Begin goods.FTextBox txtPW 
      Height          =   300
      Left            =   1740
      TabIndex        =   2
      Top             =   1500
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
   Begin goods.XPButton cmdCreate 
      Height          =   375
      Left            =   2580
      TabIndex        =   4
      Top             =   2580
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      Caption         =   "创建"
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
      Left            =   3900
      TabIndex        =   5
      Top             =   2580
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      Caption         =   "退出"
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
   Begin goods.FTextBox txtUser 
      Height          =   300
      Left            =   1740
      TabIndex        =   1
      ToolTipText     =   "由长度1-16位的中文、英文或数字组成"
      Top             =   1080
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
   End
   Begin goods.FTextBox txtConform 
      Height          =   300
      Left            =   1740
      TabIndex        =   3
      Top             =   1920
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
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "确认密码："
      Height          =   180
      Left            =   780
      TabIndex        =   8
      Top             =   1995
      Width           =   900
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FFF8F0&
      BorderColor     =   &H00C5742F&
      Height          =   1575
      Left            =   240
      Top             =   840
      Width           =   4860
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "用 户 名："
      Height          =   180
      Left            =   780
      TabIndex        =   7
      Top             =   1125
      Width           =   900
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "密    码："
      Height          =   180
      Left            =   780
      TabIndex        =   6
      Top             =   1560
      Width           =   900
   End
   Begin VB.Label Label1 
      Caption         =   "第一次使用本系统，需创建您的用户帐号。"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   0
      Top             =   240
      Width           =   5055
   End
End
Attribute VB_Name = "frmCreateUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCreate_Click()
    On Error GoTo ErrProcess
    Dim strSQL As String
    
    If Me.txtPW.Text = "" Then
        MsgBox "密码不能空，请设置密码。"
        Me.txtPW.SetFocus
        Exit Sub
        
    End If
    
    If Me.txtPW.Text <> Me.txtConform.Text Then
        MsgBox "二次输入的密码不一致，请重新输入。", vbExclamation
        Me.txtConform.SetFocus
        Exit Sub
        
    Else
        strSQL = "insert into userInfo(uName,uPWD,uDescript) " & " " & _
                               "values('" & Me.txtUser.Text & "','" & GetMD5(Me.txtPW.Text) & "','管理员')"
        If Conn.state > 0 Then
            Conn.Close
        End If
        
        DBConnect
        Conn.Execute strSQL
        If Conn.state > 0 Then
            Conn.Close
        End If
        
    End If
    frmLogin.Show
    Unload Me
    Exit Sub
ErrProcess:
    MsgBox Err.Description, , "错误提示"

End Sub

Private Sub cmdExit_Click()
    End
End Sub

Private Sub Form_Activate()
    Me.txtUser.SetFocus
End Sub

Private Sub Picture1_Click()
    MsgBox "长度为1-16位的中文、英文或数字组合。", vbInformation, "创建用户"
End Sub

Private Sub txtConform_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        Me.cmdCreate.SetFocus
    End If
End Sub
Private Sub txtPW_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        Me.txtConform.SetFocus
    End If
End Sub

Private Sub txtUser_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        Me.txtPW.SetFocus
    End If
End Sub
