VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmBorrow 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "借用/领用登记"
   ClientHeight    =   9225
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   13335
   Icon            =   "frmBorrow.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9225
   ScaleWidth      =   13335
   Begin goods.XPButton cmdPrint 
      Height          =   375
      Left            =   12090
      TabIndex        =   23
      Top             =   3135
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
      Left            =   12090
      TabIndex        =   5
      Top             =   3945
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
      Left            =   12090
      TabIndex        =   4
      Top             =   2310
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
      Left            =   12090
      TabIndex        =   3
      Top             =   1530
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
      Enabled         =   0   'False
      cBack           =   -2147483633
   End
   Begin goods.XPButton cmdAdd 
      Height          =   375
      Left            =   12090
      TabIndex        =   2
      Top             =   720
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
      Enabled         =   0   'False
      cBack           =   -2147483633
   End
   Begin MSComctlLib.ListView LV1 
      Height          =   4845
      Left            =   75
      TabIndex        =   0
      Top             =   60
      Width           =   11610
      _ExtentX        =   20479
      _ExtentY        =   8546
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
      NumItems        =   12
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Text            =   "序号"
         Object.Width           =   1058
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "物品名称"
         Object.Width           =   2117
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
         Object.Width           =   1058
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Text            =   "数量"
         Object.Width           =   1058
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   6
         Text            =   "价格"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   7
         Text            =   "借/领用人"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   8
         Text            =   "借/领用日期"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "说明"
         Object.Width           =   970
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   10
         Text            =   "操作员"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   11
         Text            =   "备注"
         Object.Width           =   1764
      EndProperty
   End
   Begin VB.Frame Frame1 
      Height          =   3945
      Left            =   90
      TabIndex        =   1
      Top             =   5040
      Width           =   11550
      Begin goods.FTextBox txtName 
         Height          =   300
         Left            =   960
         TabIndex        =   30
         Top             =   315
         Width           =   2835
         _ExtentX        =   5001
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
         Locked          =   -1  'True
      End
      Begin VB.CommandButton Command2 
         Caption         =   "..."
         Height          =   285
         Left            =   3945
         TabIndex        =   29
         Top             =   2295
         Width           =   420
      End
      Begin goods.FTextBox txtBorrow 
         Height          =   300
         Left            =   960
         TabIndex        =   28
         Top             =   2250
         Width           =   2835
         _ExtentX        =   5001
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
         Locked          =   -1  'True
      End
      Begin VB.CommandButton Command1 
         Caption         =   "..."
         Height          =   285
         Left            =   3945
         TabIndex        =   26
         Top             =   345
         Width           =   420
      End
      Begin VB.OptionButton optUse 
         Caption         =   "领用"
         Height          =   345
         Left            =   7830
         TabIndex        =   25
         Top             =   1230
         Width           =   795
      End
      Begin VB.OptionButton optBorrow 
         Caption         =   "借用"
         Height          =   300
         Left            =   6510
         TabIndex        =   24
         Top             =   1275
         Value           =   -1  'True
         Width           =   825
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   300
         Left            =   3945
         TabIndex        =   19
         Top             =   1815
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   529
         _Version        =   393216
         Format          =   71368705
         CurrentDate     =   40107
      End
      Begin goods.FTextBox txtDate 
         Height          =   300
         Left            =   945
         TabIndex        =   14
         Top             =   1830
         Width           =   2835
         _ExtentX        =   5001
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
      Begin goods.XPButton cmdCancel 
         Height          =   435
         Left            =   6360
         TabIndex        =   18
         Top             =   3270
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   767
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
      Begin goods.XPButton cmdSave 
         Height          =   435
         Left            =   2430
         TabIndex        =   17
         Top             =   3270
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   767
         Caption         =   "借  用"
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
      Begin goods.FTextBox txtDescript 
         Height          =   300
         Left            =   960
         TabIndex        =   16
         Top             =   2730
         Width           =   8820
         _ExtentX        =   15558
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
      Begin goods.FTextBox txtTotal 
         Height          =   300
         Left            =   960
         TabIndex        =   13
         Top             =   1440
         Width           =   2835
         _ExtentX        =   5001
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
         MaxLength       =   9
      End
      Begin goods.FTextBox txtPrice 
         Height          =   300
         Left            =   945
         TabIndex        =   12
         Top             =   1065
         Width           =   2835
         _ExtentX        =   5001
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
         MaxLength       =   9
      End
      Begin goods.FTextBox txtNum 
         Height          =   300
         Left            =   945
         TabIndex        =   11
         Top             =   690
         Width           =   2835
         _ExtentX        =   5001
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
         MaxLength       =   12
         afterdecimal    =   5
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H000000FF&
         Height          =   1275
         Left            =   6045
         Top             =   780
         Width           =   2940
      End
      Begin VB.Label Label6 
         Caption         =   "借用人"
         Height          =   255
         Left            =   195
         TabIndex        =   27
         Top             =   2280
         Width           =   795
      End
      Begin VB.Label Label10 
         Caption         =   "元"
         Height          =   285
         Left            =   3945
         TabIndex        =   22
         Top             =   1500
         Width           =   255
      End
      Begin VB.Label Label8 
         Caption         =   "元"
         Height          =   255
         Left            =   3945
         TabIndex        =   21
         Top             =   1095
         Width           =   225
      End
      Begin VB.Label lblUnit 
         Caption         =   "unit"
         ForeColor       =   &H000000FF&
         Height          =   240
         Left            =   3945
         TabIndex        =   20
         Top             =   765
         Width           =   1605
      End
      Begin VB.Label Label7 
         Caption         =   "日   期"
         Height          =   255
         Left            =   195
         TabIndex        =   15
         Top             =   1890
         Width           =   795
      End
      Begin VB.Label Label5 
         Caption         =   "备   注"
         Height          =   255
         Left            =   195
         TabIndex        =   10
         Top             =   2775
         Width           =   795
      End
      Begin VB.Label Label4 
         Caption         =   "价   格"
         Height          =   255
         Left            =   195
         TabIndex        =   9
         Top             =   1533
         Width           =   795
      End
      Begin VB.Label Label3 
         Caption         =   "单   价"
         Height          =   255
         Left            =   195
         TabIndex        =   8
         Top             =   1125
         Width           =   795
      End
      Begin VB.Label Label2 
         Caption         =   "数   量"
         Height          =   255
         Left            =   195
         TabIndex        =   7
         Top             =   735
         Width           =   795
      End
      Begin VB.Label Label1 
         Caption         =   "物品名称"
         Height          =   255
         Left            =   195
         TabIndex        =   6
         Top             =   330
         Width           =   795
      End
   End
End
Attribute VB_Name = "frmBorrow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim chainID(50) As Integer
Dim goodsID(200) As Integer
Dim curGoods As Integer
Dim goodsUnit(20) As String

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdPrint_Click()
    PrintListView LV1, "进货情况表(" & Format(Now, "yyyy年mm月dd日") & "打印)"

End Sub

Private Sub cmdSave_Click()
    Dim strSQL As String
    Dim item As ListItem
    
    If txtName.Text = "" Then
        MsgBox "请选择物品！"
        txtName.SetFocus
        Exit Sub
    End If
    If txtBorrow.Text = "" Then
        MsgBox "请选择借/领用人！"
        Command2.SetFocus
        Exit Sub
    End If
    
'    If cmdSave.Tag = "Add" Then
        strSQL = "insert into borrow(bGoods,bNumber,bPrice,bTotal,bDate,bBorrow,bOperator,bDescript,bFlag,bReturn) " & _
                    "values(" & txtName.Tag & "," & txtNum.Text & "," & txtPrice.Text & "," & txtTotal.Text & ",'" & _
                     txtDate.Text & " " & Format(Now, "hh:mm:ss") & "'," & txtBorrow.Tag & "," & curUserID & ",'" & _
                     IIf(txtDescript.Text = "", "无", txtDescript.Text) & "','" & IIf(optBorrow.value, "借用", "领用") & "',false)"
                        
'    Else
'        strSQL = "update sale  set  sNumber=" & txtNum.Text & "," & _
'                                   "sPrice=" & txtPrice.Text & "," & _
'                                   "sTotal=" & txtTotal.Text & "," & _
'                                   "sdate='" & txtDate.Text & " " & Format(Now, "hh:mm:ss") & "'," & _
'                                   "sDescript='" & txtDescript.Text & "'  " & _
'                                   "where sid=" & GetID(LV1.SelectedItem.Key)
'    End If
   
    DBConnect
    
    Conn.Execute strSQL
    
    '更新库存
    Conn.Execute "update Stock set sNumber=sNumber-" & txtNum.Text & "," & _
                                  "sTotal=sTotal-" & txtTotal.Text & " " & _
                                  "where sGoods=" & txtName.Tag
    
    Conn.Close
    
    lblUnit.caption = ""
    txtNum.Text = ""
    txtTotal.Text = ""
    txtBorrow.Text = ""
    setControl (False)
    
    
    fillListView
    If LV1.ListItems.Count > 0 Then LV1.ListItems(LV1.ListItems.Count).Selected = True
    
    If optBorrow.value Then
        MsgBox "借用数据处理完成！"
    Else
        MsgBox "领用数据处理完成！"
        
    End If
    
    
End Sub

Private Sub cmdSaveAdd_Click()
    cmdSave_Click
    cmdAdd_Click
    
End Sub

Private Sub Command1_Click()
    blStocksShow = False
    frmStocks.Show vbModal
End Sub

Private Sub Command2_Click()
    frmEmployee.Show vbModal
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
    curGoods = fcbName.ListIndex
    If txtNum.Enabled Then txtNum.SetFocus
End Sub

Private Sub Form_Load()
    Me.Tag = "Load"
    
    Me.Height = 9705
    Me.Width = 13425
    Me.Tag = "Over"
    fillListView
    
    Me.Top = 0
    Me.Left = 0
    lblUnit.caption = ""
    txtDate.Text = Format(Now, "yyyy-mm-dd")
    
    '设置借用/领用单选按钮
    optBorrow.value = blBorrow
    optUse.value = Not blBorrow
    If optBorrow.value Then
        cmdSave.caption = "借  用"
    Else
        cmdSave.caption = "领  用"
    End If
    
    
End Sub
Private Sub fillListView()
    Dim item As ListItem
    Dim strSQL As String
    Dim rsBorrow As ADODB.Recordset

    Set rsBorrow = New ADODB.Recordset
    
    DBConnect
    
    strSQL = "select goods.gName,goods.gSpec,bid,bNumber,bPrice,bTotal,bDate,bDescript,bFlag,bOperator,eName,uName " & _
                     " from goods,borrow,employee,userinfo " & _
                     " where  borrow.bGoods=goods.gid and borrow.bborrow=employee.eid and userinfo.uid=borrow.bOperator and bReturn<>true"
    
    rsBorrow.Open strSQL, Conn, 1, 1
    
    LV1.ListItems.Clear
    
    n = 1
    Do While Not rsBorrow.EOF
        Set item = LV1.ListItems.Add(, rsBorrow("bid") & "k")
        item.SubItems(1) = n
        item.SubItems(2) = rsBorrow("gName")
        item.SubItems(3) = rsBorrow("gSpec")
        item.SubItems(4) = Format(rsBorrow("bPrice"), "##,##0.00")
        item.SubItems(5) = rsBorrow("bNumber")
        item.SubItems(6) = Format(rsBorrow("bTotal"), "##,##0.00")
        item.SubItems(7) = rsBorrow("eName")
        item.SubItems(8) = rsBorrow("bDate")
        item.SubItems(9) = rsBorrow("bFlag")
        item.SubItems(10) = rsBorrow("uName")
        item.SubItems(11) = rsBorrow("bDescript")
        rsBorrow.MoveNext
        n = n + 1
    Loop
    
    
    rsBorrow.Close
    Set rsBorrow = Nothing
    Conn.Close
    
    If LV1.ListItems.Count < 1 Then
        cmdDelete.Enabled = False
        cmdEdit.Enabled = False
        cmdPrint.Enabled = False
    Else
        cmdDelete.Enabled = True
        'cmdEdit.Enabled = True
        cmdPrint.Enabled = True
    End If
    
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
        SetCmdState True
End Sub

Private Sub optBorrow_Click()
    cmdSave.caption = "借  用"
End Sub

Private Sub optUse_Click()
    cmdSave.caption = "领  用"
End Sub

Private Sub txtDate_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then txtDescript.SetFocus
    
End Sub

Private Sub txtNum_KeyUp(KeyCode As Integer, Shift As Integer)
    If txtPrice.Text <> "" Then txtTotal.Text = Val(txtNum.Text) * Val(txtPrice.Text)
    If KeyCode = 13 Then txtPrice.SetFocus
    
End Sub

Private Sub txtNum_LostFocus()
    If Val(txtNum.Text) > Val(txtNum.Tag) Then
        MsgBox "借用/领用数量不能超过库存数量!"
        txtNum.SetFocus
    End If
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
    
    txtDate.Text = Format(Now, "yyyy-mm-dd")

End Sub

Private Sub cmdEdit_Click()
    cmdSave.Tag = "Edit"
    
    txtNum.Text = Val(LV1.SelectedItem.SubItems(5))
    txtPrice.Text = LV1.SelectedItem.SubItems(4)
    txtTotal.Text = LV1.SelectedItem.SubItems(6)
    txtDate.Text = LV1.SelectedItem.SubItems(8)
    txtDescript.Text = LV1.SelectedItem.SubItems(11)
    
    If LV1.SelectedItem.SubItems(9) = "借用" Then
        optBorrow.value = True
    Else
        optUse.value = True
    End If
    
    setControl (True)
End Sub

Private Sub cmdDelete_Click()
    Dim rs As ADODB.Recordset
    
    If MsgBox("删除序号为[ " & LV1.SelectedItem.SubItems(1) & " ]的记录吗？", vbYesNo + vbExclamation, "提示") = vbYes Then
        DBConnect
            
        Set rs = New ADODB.Recordset
        rs.Open "select * from borrow where bid=" & GetID(LV1.SelectedItem.Key), Conn, adOpenDynamic, adLockBatchOptimistic
        Conn.Execute "update Stock set sNumber=sNumber+" & rs("bNumber") & "," & _
                                      "sTotal=sTotal+" & rs("bTotal") & "  " & _
                                      "where sGoods=" & rs("bGoods")
        Conn.Execute "update Stock set sPrice=sNumber/sTotal where sGoods=" & rs("bGoods")
        
        rs.Close
        Set rs = Nothing
        
        Conn.Execute "delete from sale where sid=" & GetID(LV1.SelectedItem.Key)
        Conn.Close
        LV1.ListItems.Remove (LV1.SelectedItem.Index)
    End If
        
End Sub

Private Sub cmdExit_Click()
    SetCmdState True
    Unload Me
End Sub
Private Sub setControl(isEnable As Boolean)
End Sub
