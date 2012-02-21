VERSION 5.00
Begin VB.Form frmHelp 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "使用说明"
   ClientHeight    =   7845
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6855
   Icon            =   "frmHelp.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7845
   ScaleWidth      =   6855
   StartUpPosition =   1  '所有者中心
   Begin VB.Frame Frame2 
      Height          =   135
      Left            =   -30
      TabIndex        =   4
      Top             =   6840
      Width           =   7500
   End
   Begin goods.XPButton XPButton1 
      Height          =   420
      Left            =   2625
      TabIndex        =   2
      Top             =   7215
      Width           =   1620
      _ExtentX        =   2858
      _ExtentY        =   741
      Caption         =   "关  闭"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      cBack           =   -2147483633
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   1125
      Left            =   30
      TabIndex        =   3
      Top             =   6855
      Width           =   6810
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   690
      Top             =   7125
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5745
      Left            =   2250
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "frmHelp.frx":058A
      Top             =   1005
      Width           =   3015
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "简要使用说明"
      BeginProperty Font 
         Name            =   "黑体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   900
      TabIndex        =   1
      Top             =   360
      Width           =   5160
   End
End
Attribute VB_Name = "frmHelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Me.Height = 8355
    Me.Width = 6975
    Text1.Text = "1、添加物品分类；" & Chr(13) & Chr(10) & Chr(13) & Chr(10)
    Text1.Text = Text1.Text & "2、添加物品名称；" & Chr(13) & Chr(10) & Chr(13) & Chr(10)
    Text1.Text = Text1.Text & "3、物品入库；" & Chr(13) & Chr(10) & Chr(13) & Chr(10)
    Text1.Text = Text1.Text & "4、借用或领用；" & Chr(13) & Chr(10) & Chr(13) & Chr(10)
    Text1.Text = Text1.Text & "5、归还物品；" & Chr(13) & Chr(10) & Chr(13) & Chr(10)
    Text1.Text = Text1.Text & "6、库存显存。" & Chr(13) & Chr(10) & Chr(13) & Chr(10)
    Text1.Top = 6500
End Sub

Private Sub Timer1_Timer()
    If Text1.Top > 200 Then Text1.Top = Text1.Top - 20
    Me.XPButton1.SetFocus
End Sub

Private Sub XPButton1_Click()
    Unload Me
End Sub
