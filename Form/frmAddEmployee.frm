VERSION 5.00
Begin VB.Form frmAddEmployee 
   Caption         =   "����Ա��"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   Icon            =   "frmAddEmployee.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CommandButton Command2 
      Caption         =   "ȡ  ��"
      Height          =   420
      Left            =   2715
      TabIndex        =   6
      Top             =   2070
      Width           =   945
   End
   Begin VB.CommandButton Command1 
      Caption         =   "��  ��"
      Height          =   420
      Left            =   1065
      TabIndex        =   5
      Top             =   2070
      Width           =   945
   End
   Begin goods.FCombo comSex 
      Height          =   300
      Left            =   1890
      TabIndex        =   4
      Top             =   930
      Width           =   1920
      _ExtentX        =   3387
      _ExtentY        =   529
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
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
   Begin goods.FTextBox txtDescript 
      Height          =   300
      Left            =   1860
      TabIndex        =   7
      Top             =   1440
      Width           =   1965
      _ExtentX        =   3466
      _ExtentY        =   529
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontName        =   "����"
      FontSize        =   9
   End
   Begin goods.FTextBox txtName 
      Height          =   300
      Left            =   1860
      TabIndex        =   0
      Top             =   390
      Width           =   1965
      _ExtentX        =   3466
      _ExtentY        =   529
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontName        =   "����"
      FontSize        =   9
   End
   Begin VB.Label Label3 
      Caption         =   "˵ ��"
      Height          =   300
      Left            =   1065
      TabIndex        =   3
      Top             =   1500
      Width           =   555
   End
   Begin VB.Label Label2 
      Caption         =   "�� ��"
      Height          =   210
      Left            =   1065
      TabIndex        =   2
      Top             =   975
      Width           =   510
   End
   Begin VB.Label Label1 
      Caption         =   "�� ��"
      Height          =   285
      Left            =   1095
      TabIndex        =   1
      Top             =   450
      Width           =   510
   End
End
Attribute VB_Name = "frmAddEmployee"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub FCombo1_Change()

End Sub

Private Sub Command1_Click()
    If Trim(txtName.Text) = "" Then
        MsgBox "����δ��д��"
        txtName.SetFocus
        Exit Sub
    End If
    
    DBConnect
    Conn.Execute "insert into employee(eName,eSex,eDescript) values('" & _
           Trim(txtName.Text) & "','" & comSex.Text & "','" & IIf(Trim(txtDescript.Text) = "", "��", Trim(txtDescript.Text)) & "')"
    
    frmEmployee.loadList
    
    Unload Me
    
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    comSex.AddItem "��"
    comSex.AddItem "Ů"
    comSex.ListIndex = 0
End Sub
