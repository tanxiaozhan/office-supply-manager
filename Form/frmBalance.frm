VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmBalance 
   Caption         =   "盘点结存情况"
   ClientHeight    =   7860
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11160
   Icon            =   "frmBalance.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7860
   ScaleWidth      =   11160
   Begin VB.Frame Frame1 
      Height          =   2295
      Left            =   120
      TabIndex        =   5
      Top             =   5400
      Width           =   10815
      Begin goods.XPButton XPButton1 
         Height          =   405
         Left            =   9000
         TabIndex        =   25
         Top             =   1080
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   714
         Caption         =   "打 印"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "黑体"
            Size            =   10.5
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
      Begin goods.FTextBox txtBase 
         Height          =   300
         Index           =   3
         Left            =   5220
         TabIndex        =   6
         Top             =   285
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
         Left            =   7320
         TabIndex        =   7
         Top             =   1080
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   714
         Caption         =   "关  闭"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "黑体"
            Size            =   10.5
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
         Left            =   5640
         TabIndex        =   8
         Top             =   1080
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   714
         Caption         =   "保  存"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "黑体"
            Size            =   10.5
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
         Left            =   9345
         TabIndex        =   9
         Top             =   1725
         Width           =   1110
         _ExtentX        =   1958
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
         ImgAlign        =   4
         Image           =   "frmBalance.frx":08CA
         cBack           =   -2147483633
      End
      Begin goods.XPButton cmdPre 
         Height          =   405
         Left            =   6895
         TabIndex        =   10
         Top             =   1725
         Width           =   1110
         _ExtentX        =   1958
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
         ImgAlign        =   4
         Image           =   "frmBalance.frx":0E64
         cBack           =   -2147483633
      End
      Begin goods.XPButton cmdTop 
         Height          =   405
         Left            =   5670
         TabIndex        =   11
         Top             =   1725
         Width           =   1110
         _ExtentX        =   1958
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
         ImgAlign        =   4
         Image           =   "frmBalance.frx":13FE
         cBack           =   -2147483633
      End
      Begin goods.FTextBox txtBase 
         Height          =   300
         Index           =   0
         Left            =   1080
         TabIndex        =   12
         Top             =   240
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
         Index           =   1
         Left            =   1095
         TabIndex        =   13
         Top             =   630
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
         Left            =   1095
         TabIndex        =   14
         Top             =   1020
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
         Left            =   1095
         TabIndex        =   15
         Top             =   1410
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
         Left            =   1095
         TabIndex        =   16
         Top             =   1830
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
         Left            =   8120
         TabIndex        =   17
         Top             =   1725
         Width           =   1110
         _ExtentX        =   1958
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
         ImgAlign        =   4
         Image           =   "frmBalance.frx":1998
         cBack           =   -2147483633
      End
      Begin VB.Label Label2 
         Caption         =   "名牌"
         Height          =   360
         Left            =   615
         TabIndex        =   24
         Top             =   330
         Width           =   420
      End
      Begin VB.Label Label3 
         Caption         =   "品种"
         Height          =   255
         Left            =   615
         TabIndex        =   23
         Top             =   705
         Width           =   420
      End
      Begin VB.Label Label4 
         Caption         =   "规格"
         Height          =   225
         Left            =   600
         TabIndex        =   22
         Top             =   1080
         Width           =   420
      End
      Begin VB.Label Label5 
         Caption         =   "库存数量"
         Height          =   315
         Left            =   255
         TabIndex        =   21
         Top             =   1470
         Width           =   750
      End
      Begin VB.Label Label6 
         Caption         =   "库存金额"
         Height          =   375
         Left            =   240
         TabIndex        =   20
         Top             =   1890
         Width           =   765
      End
      Begin VB.Label Label7 
         Caption         =   "元"
         Height          =   270
         Left            =   4860
         TabIndex        =   19
         Top             =   1890
         Width           =   270
      End
      Begin VB.Label lblUnit 
         Caption         =   "Unit"
         Height          =   225
         Left            =   4845
         TabIndex        =   18
         Top             =   1455
         Width           =   1005
      End
   End
   Begin goods.FCombo fcbDate 
      Height          =   330
      Left            =   8340
      TabIndex        =   4
      Top             =   240
      Width           =   2475
      _ExtentX        =   4366
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
   Begin goods.FCombo fcbChain 
      Height          =   330
      Left            =   870
      TabIndex        =   1
      Top             =   240
      Width           =   5475
      _ExtentX        =   9657
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
      NumItems        =   15
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
         Text            =   "上期库存"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   7
         Text            =   "上期库存金额"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   8
         Text            =   "进货量"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   9
         Text            =   "进货金额"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   10
         Text            =   "销售量"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   11
         Text            =   "销售金额"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   12
         Text            =   "实际结存"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   13
         Text            =   "结存金额"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   14
         Text            =   "盘点盈亏"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label Label8 
      Caption         =   "盘点日期"
      BeginProperty Font 
         Name            =   "黑体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   7425
      TabIndex        =   3
      Top             =   285
      Width           =   900
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
      TabIndex        =   2
      Top             =   300
      Width           =   750
   End
End
Attribute VB_Name = "frmBalance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private chainID(50) As Integer
Private isfrmLoad As Boolean       '窗体装载完毕


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

Private Sub fcbDate_Click()
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
    strSQL = "select * from chain"       '连锁店
    rs.Open strSQL, Conn, 1, 1
    n = 0
    Do While Not rs.EOF
        fcbChain.AddItem rs("cName")
        chainID(n) = rs("cid")
        rs.MoveNext
        n = n + 1
    Loop
    fcbChain.ListIndex = 0
    
    rs.Close
    
    rs.Open "select * from balanceDate order by bDate"     '盘点日期
    Do While Not rs.EOF
        fcbDate.AddItem rs("bDate")
        rs.MoveNext
    Loop
    fcbDate.ListIndex = 0
    If fcbDate.ListCount > 0 Then fcbDate.ListIndex = fcbDate.ListCount - 1
    
   
    rs.Close
    Set rs = Nothing
    Conn.Close
        
    Me.Height = 8370
    Me.Width = 11175
    lblUnit.caption = ""
    
    isfrmLoad = True
    fillListView
    
End Sub
Private Sub fillListView()
    If Not isfrmLoad Then Exit Sub
    
    Dim strSQL As String
    Dim rsBalance As ADODB.Recordset
    Set rsBalance = New ADODB.Recordset
    
    DBConnect
    
    strSQL = "select balance.bid,bname,gname,cname,uname,bLast,bLastPrice,bStock,bStockPrice,bSale,bSalePrice,bCurrent,bCurrentPrice,bProfit  " & _
                "from balance,goods,brand,class,unit " & _
                "where bChain=" & chainID(fcbChain.ListIndex) & " and balance.bDate=#" & fcbDate.Text & "#" & _
                " and  balance.bgoods=goods.gid and goods.gBrand=brand.bid and goods.gSpec=class.cid and goods.gUnit=unit.uid"

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
        item.SubItems(7) = Format(rsBalance("bLastPrice"), "##,##0.00")
        item.SubItems(8) = rsBalance("bStock")
        item.SubItems(9) = Format(rsBalance("bStockPrice"), "##,##0.00")
        item.SubItems(10) = rsBalance("bSale")
        item.SubItems(11) = Format(rsBalance("bSalePrice"), "##,##0.00")
        item.SubItems(12) = rsBalance("bCurrent")
        item.SubItems(13) = Format(rsBalance("bCurrentPrice"), "##,##0.00")
        item.SubItems(14) = Format(rsBalance("bProfit"), "##,##0.00")
        
        rsBalance.MoveNext
        n = n + 1
    Loop
    
    rsBalance.Close
    Set rsBalance = Nothing
    Conn.Close
    
    SetDirCmdState
    

End Sub

Private Sub Form_Resize()
    If Me.Width < 11175 Then Me.Width = 11175
    If Me.Height < 8370 Then Me.Height = 8370
    
    LV1.Width = Me.Width - LV1.Left * 2 - 200
    Frame1.Width = LV1.Width
    LV1.Height = Me.Height - Frame1.Height - 1600
    Frame1.Top = LV1.Height + 800
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SetCmdState True

End Sub

Private Sub lv1_Click()
    fillData
    SetDirCmdState

End Sub

Private Sub XPButton1_Click()
    PrintListView LV1, Format(fcbDate.Text, "yyyy年mm月dd日") & "盘点结存情况"
End Sub

Private Sub XPButton5_Click()
    DBConnect
    Conn.Execute "update balance set bCurrent=" & txtBase(4).Text & "," & _
                                    "bCurrentPrice=" & txtBase(5).Text & "   " & _
                                    "where bid=" & GetID(LV1.SelectedItem.Key)
    Conn.Close
    Set Conn = Nothing
    LV1.SelectedItem.SubItems(12) = txtBase(4).Text
    LV1.SelectedItem.SubItems(13) = txtBase(5).Text
    
    
End Sub

Private Sub XPButton6_Click()
    isfrmLoad = False
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
    
    If LV1.ListItems.Count > 1 Then fillData
    
End Sub
Private Sub fillData()
    For i = 0 To 3
        txtBase(i).Text = LV1.SelectedItem.SubItems(i + 2)
    Next
    
    txtBase(4).Text = LV1.SelectedItem.SubItems(12)
    txtBase(5).Text = LV1.SelectedItem.SubItems(13)
    
    lblUnit.caption = txtBase(3).Text
End Sub
