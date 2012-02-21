VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmIn 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "入库记录"
   ClientHeight    =   8685
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   13080
   Icon            =   "frmIn.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8685
   ScaleWidth      =   13080
   Begin goods.XPButton cmdPrint 
      Height          =   375
      Left            =   11775
      TabIndex        =   25
      Top             =   2595
      Width           =   870
      _ExtentX        =   1535
      _ExtentY        =   661
      Caption         =   "打 印"
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
      Left            =   11775
      TabIndex        =   5
      Top             =   3330
      Width           =   870
      _ExtentX        =   1535
      _ExtentY        =   661
      Caption         =   "关 闭"
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
   Begin goods.XPButton cmdDelete 
      Height          =   375
      Left            =   11775
      TabIndex        =   4
      Top             =   1890
      Width           =   870
      _ExtentX        =   1535
      _ExtentY        =   661
      Caption         =   "删 除"
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
   Begin goods.XPButton cmdEdit 
      Height          =   375
      Left            =   11775
      TabIndex        =   3
      Top             =   1185
      Width           =   870
      _ExtentX        =   1535
      _ExtentY        =   661
      Caption         =   "编 辑"
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
   Begin goods.XPButton cmdAdd 
      Height          =   375
      Left            =   11775
      TabIndex        =   2
      Top             =   465
      Width           =   870
      _ExtentX        =   1535
      _ExtentY        =   661
      Caption         =   "添 加"
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
   Begin MSComctlLib.ListView LV1 
      Height          =   4530
      Left            =   75
      TabIndex        =   0
      Top             =   135
      Width           =   11250
      _ExtentX        =   19844
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
      NumItems        =   10
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "序号"
         Object.Width           =   1058
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "物品名称"
         Object.Width           =   2822
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   3
         Text            =   "规格"
         Object.Width           =   2822
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "单价"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Text            =   "数量"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   6
         Text            =   "价格"
         Object.Width           =   1676
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   7
         Text            =   "日期"
         Object.Width           =   3351
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "备注"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   9
         Text            =   "操作员"
         Object.Width           =   1587
      EndProperty
   End
   Begin VB.Frame Frame1 
      Height          =   3360
      Left            =   120
      TabIndex        =   1
      Top             =   5040
      Width           =   11220
      Begin goods.XPButton cmdSaveAdd 
         Height          =   435
         Left            =   1980
         TabIndex        =   24
         Top             =   2730
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   767
         Caption         =   "保存&&增加"
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
         Enabled         =   0   'False
         cBack           =   -2147483633
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   300
         Left            =   2760
         TabIndex        =   20
         Top             =   1830
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   72024065
         CurrentDate     =   40107
      End
      Begin goods.FTextBox txtDate 
         Height          =   300
         Left            =   720
         TabIndex        =   15
         Top             =   1830
         Width           =   2055
         _ExtentX        =   3625
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
      Begin goods.XPButton cmdCancel 
         Height          =   435
         Left            =   7020
         TabIndex        =   19
         Top             =   2730
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   767
         Caption         =   "取 消"
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
         Enabled         =   0   'False
         cBack           =   -2147483633
      End
      Begin goods.XPButton cmdSave 
         Height          =   435
         Left            =   4500
         TabIndex        =   18
         Top             =   2730
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   767
         Caption         =   "保 存"
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
         Enabled         =   0   'False
         cBack           =   -2147483633
      End
      Begin goods.FTextBox txtDescript 
         Height          =   300
         Left            =   735
         TabIndex        =   17
         Top             =   2220
         Width           =   4950
         _ExtentX        =   8731
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
      Begin goods.FTextBox txtTotal 
         Height          =   300
         Left            =   735
         TabIndex        =   14
         Top             =   1440
         Width           =   2055
         _ExtentX        =   3625
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
         isNumber        =   -1  'True
         MaxLength       =   9
      End
      Begin goods.FTextBox txtPrice 
         Height          =   300
         Left            =   720
         TabIndex        =   13
         Top             =   1065
         Width           =   2055
         _ExtentX        =   3625
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
         isNumber        =   -1  'True
         MaxLength       =   9
      End
      Begin goods.FTextBox txtNum 
         Height          =   300
         Left            =   720
         TabIndex        =   12
         Top             =   690
         Width           =   2055
         _ExtentX        =   3625
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
         isNumber        =   -1  'True
         MaxLength       =   12
         afterdecimal    =   5
      End
      Begin goods.FCombo fcbName 
         Height          =   300
         Left            =   720
         TabIndex        =   9
         Top             =   300
         Width           =   5100
         _ExtentX        =   8996
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
         ListIndex       =   -1
      End
      Begin VB.Label Label10 
         Caption         =   "元"
         Height          =   285
         Left            =   2850
         TabIndex        =   23
         Top             =   1515
         Width           =   255
      End
      Begin VB.Label Label8 
         Caption         =   "元"
         Height          =   255
         Left            =   2835
         TabIndex        =   22
         Top             =   1125
         Width           =   225
      End
      Begin VB.Label lblUnit 
         Caption         =   "unit"
         Height          =   240
         Left            =   2820
         TabIndex        =   21
         Top             =   750
         Width           =   915
      End
      Begin VB.Label Label7 
         Caption         =   "日 期"
         Height          =   255
         Left            =   180
         TabIndex        =   16
         Top             =   1890
         Width           =   555
      End
      Begin VB.Label Label5 
         Caption         =   "备 注"
         Height          =   255
         Left            =   195
         TabIndex        =   11
         Top             =   2265
         Width           =   540
      End
      Begin VB.Label Label4 
         Caption         =   "价 格"
         Height          =   255
         Left            =   195
         TabIndex        =   10
         Top             =   1533
         Width           =   540
      End
      Begin VB.Label Label3 
         Caption         =   "单 价"
         Height          =   255
         Left            =   195
         TabIndex        =   8
         Top             =   1132
         Width           =   540
      End
      Begin VB.Label Label2 
         Caption         =   "数 量"
         Height          =   255
         Left            =   195
         TabIndex        =   7
         Top             =   731
         Width           =   540
      End
      Begin VB.Label Label1 
         Caption         =   "名 称"
         Height          =   255
         Left            =   195
         TabIndex        =   6
         Top             =   330
         Width           =   540
      End
   End
End
Attribute VB_Name = "frmIn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim goodsID(200) As Integer
Dim curGoods As Integer
Dim goodsUnit(20) As String
Dim goodsPrice(200) As Single    '单价
Dim iNum As Single       '数量
Dim iPrice As Single     '单价
Dim iTotal As Single     '价格


Private Sub cmdCancel_Click()
    txtNum.Text = ""
    txtPrice.Text = ""
    txtTotal.Text = ""
    txtDate.Text = ""
    txtDescript.Text = ""

    setControl (False)
End Sub
Private Sub cmdSave_Click()
    Dim strSQL As String
    Dim rs As ADODB.Recordset
    
        
    If cmdSave.Tag = "Add" Then
        strSQL = "insert into inStock(sGoods,sNumber,sPrice,sTotal,sUnit,sDate,sDescript,sOperator) " & _
                    "values(" & goodsID(curGoods) & "," & txtNum.Text & "," & txtPrice.Text & "," & _
                    txtTotal.Text & ",'" & lblUnit.caption & "','" & txtDate.Text & "','" & _
                    txtDescript.Text & "'," & curUserID & ")"
    
    Else
        strSQL = "update inStock set  sNumber=" & txtNum.Text & "," & _
                                   "sPrice=" & txtPrice.Text & "," & _
                                   "sTotal=" & txtTotal.Text & "," & _
                                   "sdate='" & txtDate.Text & "'," & _
                                   "sDescript='" & txtDescript.Text & "'  " & _
                                   "where sid=" & GetID(LV1.SelectedItem.Key)
        
    End If
    
    DBConnect
    
    Conn.Execute strSQL
    Set rs = New ADODB.Recordset
    rs.Open "select count(1) as recCount from Stock where sGoods=" & goodsID(curGoods), Conn, adOpenDynamic, adLockBatchOptimistic
    
    If rs("reccount") < 1 Then    '库存中无该物品，则创建
        Conn.Execute "insert into Stock (sGoods,sPrice,sDescript) values(" & goodsID(curGoods) & "," & txtPrice.Text & "',-')"
    End If
    rs.Close
    Set rs = Nothing
    
    
    '更新库存
    Conn.Execute "update Stock set sNumber=sNumber+" & txtNum.Text & "-" & iNum & "," & _
                                     "sTotal=sTotal+" & txtTotal.Text & "-" & iTotal & "  " & _
                                     "where sGoods=" & goodsID(curGoods)
    '更新平均单价
    Conn.Execute "update Stock set sPrice=sTotal/sNumber where sGoods=" & goodsID(curGoods)
    
    Conn.Close
    
    setControl (False)
    
    cmdCancel_Click
    
    fillListView
    If LV1.ListItems.Count > 0 Then LV1.ListItems(LV1.ListItems.Count).Selected = True
    
End Sub

Private Sub cmdSaveAdd_Click()
    cmdSave_Click
    cmdAdd_Click
    
End Sub

Private Sub DTPicker1_CloseUp()
    txtDate.Text = DTPicker1.value
    txtDescript.SetFocus
End Sub

Private Sub fcbChain_Click()
    If Me.Tag = "Over" Then
        fillListView
    End If
End Sub

Private Sub fcbName_Click()
    lblUnit.caption = goodsUnit(fcbName.ListIndex)
    txtPrice.Text = goodsPrice(fcbName.ListIndex)
    curGoods = fcbName.ListIndex
    If txtNum.Enabled Then txtNum.SetFocus
End Sub

Private Sub fcbName_LostFocus()
    fcbName_Click
End Sub

Private Sub Form_Load()
    Dim strSQL As String
    Dim rs As ADODB.Recordset
    
    Me.Tag = "Load"
    curGoods = 0
    
    Set rs = New ADODB.Recordset
    DBConnect
    
    rs.Open "select class.cName,unit.uName,goods.gid,goods.gName,goods.gPrice " & _
                         "from class,unit,goods " & _
                         "where goods.gClass=class.cid and goods.gUnit=unit.uid", Conn, adOpenForwardOnly, adLockOptimistic
    n = 0
    Do While Not rs.EOF
        goodsID(n) = rs("gid")
        goodsUnit(n) = rs("uName")
        goodsPrice(n) = rs("gPrice")
        fcbName.AddItem rs("gName")
        rs.MoveNext
        n = n + 1
    Loop
    fcbName.ListIndex = curGoods
   
    Me.Height = 9195
    Me.Width = 13080
    lblUnit.caption = goodsUnit(fcbName.ListIndex)
    Me.Tag = "Over"
    fillListView
    
    Me.Top = 0
    Me.Left = 0
    
End Sub
Private Sub fillListView()
    Dim item As ListItem
    Dim strSQL As String
    Dim rsIn As ADODB.Recordset

    Set rsIn = New ADODB.Recordset
    
    DBConnect
    
   strSQL = "select class.cName,goods.gName,goods.gSpec,sid,sNumber,sPrice,sTotal,sUnit,sDate,sDescript,userinfo.uName as uOperator " & _
                     " from class,goods,inStock,userinfo " & _
                     " where inStock.sGoods=goods.gid and  goods.gClass=class.cid  and inStock.sOperator=userinfo.uid"
    
    rsIn.Open strSQL, Conn, 1, 1
    
    LV1.ListItems.Clear
    
    n = 1
    Do While Not rsIn.EOF
        Set item = LV1.ListItems.Add(, rsIn("sid") & "k")
        item.SubItems(1) = n
        item.SubItems(2) = rsIn("gName")
        item.SubItems(3) = rsIn("gSpec")
        item.SubItems(4) = Format(rsIn("sPrice"), "##,##0.00")
        item.SubItems(5) = rsIn("sNumber") & rsIn("sUnit")
        item.SubItems(6) = Format(rsIn("sTotal"), "##,##0.00")
        item.SubItems(7) = rsIn("sDate")
        item.SubItems(8) = rsIn("SDescript")
        item.SubItems(9) = rsIn("uOperator")
        rsIn.MoveNext
        n = n + 1
    Loop
    
    rsIn.Close
    Set rsIn = Nothing
    Conn.Close
    
    If LV1.ListItems.Count < 1 Then
        cmdDelete.Enabled = False
        cmdEdit.Enabled = False
        cmdPrint.Enabled = False
    Else
        cmdDelete.Enabled = True
        cmdEdit.Enabled = True
        cmdPrint.Enabled = True
    End If
    
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SetCmdState True
       

End Sub

Private Sub txtDate_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then txtDescript.SetFocus
    
End Sub

Private Sub txtNum_KeyUp(KeyCode As Integer, Shift As Integer)
    If txtPrice.Text <> "" Then txtTotal.Text = Val(txtNum.Text) * Val(txtPrice.Text)
    If KeyCode = 13 Then txtPrice.SetFocus
    
End Sub

Private Sub txtPrice_Keyup(KeyCode As Integer, Shift As Integer)
    If txtNum.Text <> "" Then txtTotal.Text = Val(txtNum.Text) * Val(txtPrice.Text)
    If KeyCode = 13 Then txtTotal.SetFocus
End Sub

Private Sub txtTotal_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then txtDate.SetFocus
End Sub

Private Sub cmdAdd_Click()
    cmdSave.Tag = "Add"
    
    setControl (True)
    
    txtDate.Text = Format(Now, "yyyy-mm-dd hh:mm:ss")
    
    iNum = 0
    iPrice = 0
    iTotal = 0
    

End Sub

Private Sub cmdEdit_Click()
    Dim i As Integer
    Dim strGoodsName As String
    cmdSave.Tag = "Edit"
    strGoodsName = LV1.SelectedItem.SubItems(2)
    For i = 0 To fcbName.ListCount - 1
        If fcbName.List(i) = strGoodsName Then fcbName.ListIndex = i
    Next
    
    iNum = Val(LV1.SelectedItem.SubItems(5))     '编辑的时候，先保存原来的数量
    txtNum.Text = iNum
    iPrice = LV1.SelectedItem.SubItems(4)
    txtPrice.Text = iPrice
    iTotal = LV1.SelectedItem.SubItems(6)
    txtTotal.Text = iTotal
    txtDate.Text = LV1.SelectedItem.SubItems(7)
    txtDescript.Text = LV1.SelectedItem.SubItems(8)
    setControl (True)
End Sub

Private Sub cmdDelete_Click()
    Dim rs As ADODB.Recordset
    
    If MsgBox("删除序号为[ " & LV1.SelectedItem.SubItems(1) & " ] 物品名称为[" & LV1.SelectedItem.SubItems(2) & "]的入库记录吗？", vbYesNo + vbCritical, "提示") = vbYes Then
        DBConnect
        '更新库存
        Set rs = New ADODB.Recordset
        rs.Open "select * from inStock where sid=" & GetID(LV1.SelectedItem.Key)
        If rs("sNumber") >= LV1.SelectedItem.SubItems(3) Then
            MsgBox "该物品已借出或领用，不能删除该入库记录！", , "删除入库记录"
            Exit Sub
        End If
        Conn.Execute "update Stock set sNumber=sNumber-" & rs("sNumber") & "," & _
                                       "sTotal=sTotal-" & rs("sTotal") & "  " & _
                                       "where sGoods=" & rs("sGoods")
        Conn.Execute "update Stock set sPrice=sTotal/sNumber where sGoods=" & rs("sGoods")
        rs.Close
        Set rs = Nothing
        
        Conn.Execute "delete from inStock where sid=" & GetID(LV1.SelectedItem.Key)
        
        Conn.Close
        LV1.ListItems.Remove (LV1.SelectedItem.Index)
    End If
        
End Sub

Private Sub cmdExit_Click()
    SetCmdState True
    
    Unload Me
End Sub
Private Sub setControl(isEnable As Boolean)
    fcbName.Enabled = isEnable
    txtNum.Enabled = isEnable
    txtPrice.Enabled = isEnable
    txtTotal.Enabled = isEnable
    txtDate.Enabled = isEnable
    DTPicker1.Enabled = isEnable
    txtDescript.Enabled = isEnable
    cmdSave.Enabled = isEnable
    cmdSaveAdd.Enabled = isEnable
    cmdCancel.Enabled = isEnable
        
    cmdAdd.Enabled = Not isEnable
    cmdEdit.Enabled = Not isEnable
    cmdDelete.Enabled = Not isEnable
    cmdPrint.Enabled = Not isEnable
    If (isEnable And cmdSave.Tag = "Add") Then txtPrice.Text = goodsPrice(fcbName.ListIndex)
    
End Sub

Private Sub XPButton1_Click()
    PrintListView LV1, "进货情况表(" & Format(Now, "yyyy年mm月dd日") & "打印)"

End Sub
