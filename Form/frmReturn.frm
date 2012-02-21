VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmReturn 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "归还登记"
   ClientHeight    =   7290
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   13335
   Icon            =   "frmReturn.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7290
   ScaleWidth      =   13335
   Begin goods.XPButton cmdPrint 
      Height          =   375
      Left            =   12135
      TabIndex        =   3
      Top             =   1725
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
      Left            =   12135
      TabIndex        =   2
      Top             =   2535
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
   Begin goods.XPButton cmdReturn 
      Height          =   375
      Left            =   12135
      TabIndex        =   1
      Top             =   900
      Width           =   870
      _ExtentX        =   1535
      _ExtentY        =   661
      Caption         =   "归 还"
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
      Height          =   6240
      Left            =   150
      TabIndex        =   0
      Top             =   630
      Width           =   11610
      _ExtentX        =   20479
      _ExtentY        =   11007
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
   Begin VB.Label Label1 
      Caption         =   "物品借用列表"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   180
      TabIndex        =   4
      Top             =   240
      Width           =   1755
   End
End
Attribute VB_Name = "frmReturn"
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

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdPrint_Click()
    PrintListView LV1, "借用情况(" & Format(Now, "yyyy年mm月dd日") & "打印)"

End Sub

Private Sub cmdReturn_Click()
    If MsgBox("归还 " & LV1.SelectedItem.SubItems(2) & "?", vbYesNo, "归还物品") = vbNo Then Exit Sub
    
    DBConnect
    Conn.Execute "update Borrow set bReturn=true where bid=" & GetID(LV1.SelectedItem.Key)
    '更新库存数量和总价
    Conn.Execute "update stock  set sNumber=sNumber+" & LV1.SelectedItem.SubItems(5) & "  " & _
                            " where sGoods=" & LV1.SelectedItem.Tag
    Conn.Execute "update stock set sTotal=sNumber * sPrice where sGoods=" & LV1.SelectedItem.Tag
    
    LV1.ListItems.Remove (LV1.SelectedItem.Index)
    
End Sub

Private Sub Form_Load()
    Me.Tag = "Load"
    
    Me.Height = 9705
    Me.Width = 13425
    Me.Tag = "Over"
    fillListView
    
    Me.Top = 0
    Me.Left = 0
    
    
End Sub
Private Sub fillListView()
    Dim item As ListItem
    Dim strSQL As String
    Dim rsBorrow As ADODB.Recordset

    Set rsBorrow = New ADODB.Recordset
    
    DBConnect
    
    strSQL = "select goods.gName,goods.gSpec,bid,bgoods,bNumber,bPrice,bTotal,bDate,bDescript,bFlag,bOperator,eName,uName " & _
                     " from goods,borrow,employee,userinfo " & _
                     " where  borrow.bGoods=goods.gid and borrow.bborrow=employee.eid and userinfo.uid=borrow.bOperator and bFlag='借用' and bReturn=false"
    
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
        item.Tag = rsBorrow("bGoods")
        rsBorrow.MoveNext
        n = n + 1
    Loop
    
    
    rsBorrow.Close
    Set rsBorrow = Nothing
    Conn.Close
    
    If LV1.ListItems.Count < 1 Then
        cmdReturn.Enabled = False
        cmdPrint.Enabled = False
    Else
        cmdReturn.Enabled = True
        cmdPrint.Enabled = True
    End If
    
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
        SetCmdState True
End Sub
