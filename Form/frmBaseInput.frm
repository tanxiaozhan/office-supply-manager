VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmBaseInput 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "期初数据录入"
   ClientHeight    =   7665
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11055
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7665
   ScaleWidth      =   11055
   Begin goods.FTextBox txtBase 
      Height          =   300
      Index           =   3
      Left            =   5595
      TabIndex        =   20
      Top             =   5505
      Visible         =   0   'False
      Width           =   930
      _ExtentX        =   1640
      _ExtentY        =   529
      BackColor       =   16777215
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
      Locked          =   -1  'True
   End
   Begin goods.XPButton XPButton6 
      Height          =   405
      Left            =   7515
      TabIndex        =   17
      Top             =   6000
      Width           =   1710
      _ExtentX        =   3016
      _ExtentY        =   714
      Caption         =   "关  闭"
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
   Begin goods.XPButton XPButton5 
      Height          =   405
      Left            =   7515
      TabIndex        =   16
      Top             =   5505
      Width           =   1710
      _ExtentX        =   3016
      _ExtentY        =   714
      Caption         =   "保  存"
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
   Begin goods.XPButton cmdLast 
      Height          =   405
      Left            =   9510
      TabIndex        =   15
      Top             =   6930
      Width           =   390
      _ExtentX        =   688
      _ExtentY        =   714
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
      Image           =   "frmBaseInput.frx":0000
      cBack           =   -2147483633
   End
   Begin goods.XPButton cmdPre 
      Height          =   405
      Left            =   7785
      TabIndex        =   14
      Top             =   6930
      Width           =   390
      _ExtentX        =   688
      _ExtentY        =   714
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
      Image           =   "frmBaseInput.frx":059A
      cBack           =   -2147483633
   End
   Begin goods.XPButton cmdTop 
      Height          =   405
      Left            =   6930
      TabIndex        =   13
      Top             =   6930
      Width           =   390
      _ExtentX        =   688
      _ExtentY        =   714
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
      Image           =   "frmBaseInput.frx":0B34
      cBack           =   -2147483633
   End
   Begin goods.FCombo fcbChain 
      Height          =   330
      Left            =   960
      TabIndex        =   5
      Top             =   225
      Width           =   5850
      _ExtentX        =   10319
      _ExtentY        =   582
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "黑体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      EnabledText     =   0   'False
      ListIndex       =   -1
   End
   Begin goods.FTextBox txtBase 
      Height          =   300
      Index           =   0
      Left            =   1455
      TabIndex        =   1
      Top             =   5460
      Width           =   3690
      _ExtentX        =   6509
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
      Enabled         =   0   'False
   End
   Begin MSComctlLib.ListView lv1 
      Height          =   4530
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   10740
      _ExtentX        =   18944
      _ExtentY        =   7990
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   8
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "序号"
         Object.Width           =   1058
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Text            =   "品牌"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "品种"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "规格"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   5
         Text            =   "单位"
         Object.Width           =   1058
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   6
         Text            =   "库存数量"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   7
         Text            =   "库存金额"
         Object.Width           =   2540
      EndProperty
   End
   Begin goods.FTextBox txtBase 
      Height          =   300
      Index           =   1
      Left            =   1470
      TabIndex        =   2
      Top             =   5850
      Width           =   3690
      _ExtentX        =   6509
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
      Enabled         =   0   'False
   End
   Begin goods.FTextBox txtBase 
      Height          =   300
      Index           =   2
      Left            =   1470
      TabIndex        =   3
      Top             =   6240
      Width           =   3690
      _ExtentX        =   6509
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
      Enabled         =   0   'False
   End
   Begin goods.FTextBox txtBase 
      Height          =   300
      Index           =   4
      Left            =   1470
      TabIndex        =   4
      Top             =   6630
      Width           =   3690
      _ExtentX        =   6509
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
      isNumber        =   -1  'True
      MaxLength       =   8
   End
   Begin goods.FTextBox txtBase 
      Height          =   300
      Index           =   5
      Left            =   1470
      TabIndex        =   12
      Top             =   7050
      Width           =   3690
      _ExtentX        =   6509
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
      isNumber        =   -1  'True
      MaxLength       =   8
   End
   Begin goods.XPButton cmdNext 
      Height          =   405
      Left            =   8655
      TabIndex        =   18
      Top             =   6930
      Width           =   390
      _ExtentX        =   688
      _ExtentY        =   714
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
      Image           =   "frmBaseInput.frx":10CE
      cBack           =   -2147483633
   End
   Begin VB.Label lblUnit 
      Caption         =   "Unit"
      Height          =   225
      Left            =   5220
      TabIndex        =   21
      Top             =   6675
      Width           =   1005
   End
   Begin VB.Label Label7 
      Caption         =   "元"
      Height          =   270
      Left            =   5235
      TabIndex        =   19
      Top             =   7110
      Width           =   270
   End
   Begin VB.Label Label6 
      Caption         =   "库存金额"
      Height          =   375
      Left            =   615
      TabIndex        =   11
      Top             =   7110
      Width           =   765
   End
   Begin VB.Label Label5 
      Caption         =   "库存数量"
      Height          =   315
      Left            =   630
      TabIndex        =   10
      Top             =   6690
      Width           =   750
   End
   Begin VB.Label Label4 
      Caption         =   "规格"
      Height          =   225
      Left            =   975
      TabIndex        =   9
      Top             =   6300
      Width           =   420
   End
   Begin VB.Label Label3 
      Caption         =   "品种"
      Height          =   255
      Left            =   990
      TabIndex        =   8
      Top             =   5925
      Width           =   420
   End
   Begin VB.Label Label2 
      Caption         =   "名牌"
      Height          =   360
      Left            =   990
      TabIndex        =   7
      Top             =   5550
      Width           =   420
   End
   Begin VB.Label Label1 
      Caption         =   "连锁店"
      BeginProperty Font 
         Name            =   "黑体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   150
      TabIndex        =   6
      Top             =   270
      Width           =   750
   End
End
Attribute VB_Name = "frmBaseInput"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim chainID(50) As Integer

Private Sub cb1_Change()
    MsgBox "change"
End Sub

Private Sub cmdLast_Click()
    LV1.ListItems.item(LV1.ListItems.Count).Selected = True
    SetDirCmdState

End Sub

Private Sub cmdNext_Click()
    LV1.ListItems.item(LV1.SelectedItem.Index + 1).Selected = True
    SetDirCmdState

End Sub

Private Sub cmdPre_Click()
    LV1.ListItems.item(LV1.SelectedItem.Index - 1).Selected = True
    SetDirCmdState
    
End Sub

Private Sub cmdTop_Click()
    LV1.ListItems.item(1).Selected = True
    SetDirCmdState
    
End Sub

Private Sub fcbChain_Click()
    fillListView
End Sub
Private Sub Form_Load()
    Dim item As ListItem
    Dim strSQL As String
    Dim rs As ADODB.Recordset
    Dim rsBalance As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Set rsBalance = New ADODB.Recordset
    DBConnect
    strSQL = "select * from chain"
    rs.Open strSQL, Conn, 1, 1
    n = 0
    Do While Not rs.EOF
        fcbChain.AddItem rs("cName")
        chainID(n) = rs("cid")
        rs.MoveNext
        n = n + 1
    Loop
    fcbChain.ListIndex = 0
   
    'rs.Close
    'Set rs = Nothing
    'Conn.Close
    
    'fillListView
        
    Me.Height = 8175
    Me.Width = 11175
    lblUnit.caption = ""
End Sub
Private Sub fillListView()
    Dim strSQL As String
    Dim rsgoods As ADODB.Recordset
    Dim rsBalance As ADODB.Recordset

    Set rsgoods = New ADODB.Recordset
    Set rsBalance = New ADODB.Recordset
    
    DBConnect
    
    strSQL = "select * from goods"
    rsgoods.Open strSQL, Conn, 1, 1
    
    Do While Not rsgoods.EOF
        'strSQL = "If Not Exists(select * from balance where bGood=" & rs("gid") & ") insert into balance(bGoods) values(" & rs("gid") & ")"
        strSQL = "select count(bid) as recCount from balance where bGoods=" & rsgoods("gid") & "  and bChain=" & chainID(fcbChain.ListIndex)
        rsBalance.Open strSQL, Conn, 1, 1
        If rsBalance("reccount") < 1 Then
            Conn.Execute "insert into balance(bChain,bGoods) values(" & chainID(fcbChain.ListIndex) & "," & rsgoods("gid") & ")"
        End If
        rsBalance.Close
        
        rsgoods.MoveNext
    Loop
    
    strSQL = "select balance.bid,bname,gname,cname,uname,bLast,bLastPrice  from balance,goods,brand,class,unit where bChain=" & chainID(fcbChain.ListIndex) & " and  balance.bgoods=goods.gid and goods.gBrand=brand.bid and goods.gSpec=class.cid and goods.gUnit=unit.uid"
    rsBalance.Open strSQL, Conn, 1, 1
    
    LV1.ListItems.Clear
    
    n = 1
    Do While Not rsBalance.EOF
        Set item = LV1.ListItems.Add(, rsBalance("bid") & "k")
        item.SubItems(1) = n
        item.SubItems(2) = rsBalance("bName")
        item.SubItems(3) = rsBalance("gName")
        item.SubItems(4) = rsBalance("cName")
        item.SubItems(5) = rsBalance("uName")
        item.SubItems(6) = rsBalance("bLast")
        item.SubItems(7) = rsBalance("bLastPrice")
        rsBalance.MoveNext
        n = n + 1
    Loop
    
    rsgoods.Close
    Set rsgoods = Nothing
    rsBalance.Close
    Set rsBalance = Nothing
    Conn.Close
    
    SetDirCmdState
    

End Sub

Private Sub Form_Unload(Cancel As Integer)
    SetCmdState True

End Sub

Private Sub lv1_Click()
    fillData
    SetDirCmdState

End Sub

Private Sub XPButton5_Click()
    DBConnect
    Conn.Execute "update balance set bCurrent=" & txtBase(4).Text & "," & _
                                    "bCurrentPrice=" & txtBase(5).Text & "   " & _
                                    "where bid=" & GetID(LV1.SelectedItem.Key)
    Conn.Close
    Set Conn = Nothing
    LV1.SelectedItem.SubItems(6) = txtBase(4).Text
    LV1.SelectedItem.SubItems(7) = txtBase(5).Text
    
    
End Sub

Private Sub XPButton6_Click()
    Unload Me
End Sub

Private Sub SetDirCmdState()
    
    cmdTop.Enabled = True
    cmdPre.Enabled = True
    cmdNext.Enabled = True
    cmdLast.Enabled = True
    
    If LV1.ListItems.Count < 2 Then
        cmdTop.Enabled = False
        cmdPre.Enabled = False
        cmdNext.Enabled = False
        cmdLast.Enabled = False
    Else
        If LV1.SelectedItem.Index = 1 Then
            cmdTop.Enabled = False
            cmdPre.Enabled = False
        Else
            If LV1.SelectedItem.Index = LV1.ListItems.Count Then
                cmdNext.Enabled = False
                cmdLast.Enabled = False
            End If
        End If
    
    End If
    
    fillData
    
End Sub
Private Sub fillData()
    For i = 0 To 5
        txtBase(i).Text = LV1.SelectedItem.SubItems(i + 2)
    Next
    lblUnit.caption = txtBase(3).Text
End Sub
