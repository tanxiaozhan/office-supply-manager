VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmStocks 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "库存"
   ClientHeight    =   8895
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   9450
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8895
   ScaleWidth      =   9450
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin MSComctlLib.ListView lvStocks 
      Height          =   7575
      Left            =   165
      TabIndex        =   1
      Top             =   510
      Width           =   9120
      _ExtentX        =   16087
      _ExtentY        =   13361
      View            =   3
      LabelWrap       =   0   'False
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      MousePointer    =   1
      NumItems        =   8
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Text            =   "序号"
         Object.Width           =   970
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "物品名称"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   3
         Text            =   "规格"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "单价"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Text            =   "数量"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   6
         Text            =   "价格"
         Object.Width           =   2822
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "备注"
         Object.Width           =   3528
      EndProperty
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "确  定"
      Height          =   375
      Left            =   3915
      TabIndex        =   0
      Top             =   8385
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Height          =   855
      Left            =   -45
      TabIndex        =   3
      Top             =   8085
      Width           =   10000
   End
   Begin VB.Label Label1 
      Caption         =   "目前库存"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   165
      TabIndex        =   2
      Top             =   195
      Width           =   1500
   End
End
Attribute VB_Name = "frmStocks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub Form_Load()
    Dim n As Integer
    Dim item As ListItem
    Dim strSQL As String
    Dim rs As ADODB.Recordset
    
    Set rs = New ADODB.Recordset
    strSQL = "select * from goods,stock where goods.gid=stock.sGoods and stock.sNumber >0 order by sGoods "
    DBConnect
    rs.Open strSQL, Conn, 1, 1
        
    lvStocks.ListItems.Clear
    n = 1
    
    Do While Not rs.EOF
        Set item = lvStocks.ListItems.Add(, rs("sid") & "k")
        item.SubItems(1) = n
        item.SubItems(2) = rs("gName")
        item.SubItems(3) = rs("gSpec")
        item.SubItems(4) = rs("sPrice")
        item.SubItems(5) = rs("sNumber")
        item.SubItems(6) = rs("sTotal")
        item.SubItems(7) = rs("sDescript")
        item.Tag = rs("sGoods")
        n = n + 1
        rs.MoveNext
    Loop
    
    rs.Close
    Set rs = Nothing
    
    
End Sub

Private Sub OKButton_Click()
    If Not blStocksShow Then
    
        frmBorrow.txtName.Text = lvStocks.SelectedItem.SubItems(2) & "(" & lvStocks.SelectedItem.SubItems(3) & ")"
        frmBorrow.txtName.Tag = lvStocks.SelectedItem.Tag
    
        frmBorrow.lblUnit.caption = "库存数量：" & lvStocks.SelectedItem.SubItems(5)
        frmBorrow.txtNum.Tag = lvStocks.SelectedItem.SubItems(5)
    
        frmBorrow.txtPrice.Text = lvStocks.SelectedItem.SubItems(4)
    End If
    
    
    'strSelectGoods = lvStocks.SelectedItem.SubItems(2) & "(" & lvStocks.SelectedItem.SubItems(3) & ")"
    'iSelectGoodsNumber = lvStocks.SelectedItem.SubItems(5)
    Unload Me
End Sub
