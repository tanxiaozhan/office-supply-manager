VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmInfo 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "基本参数设置"
   ClientHeight    =   8490
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11070
   Icon            =   "frmInfo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   566
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   738
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   10740
      Top             =   30
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInfo.frx":0E42
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInfo.frx":13DC
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInfo.frx":1976
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInfo.frx":1F10
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      Height          =   7020
      Index           =   1
      Left            =   -150
      TabIndex        =   2
      Top             =   1785
      Width           =   10785
      Begin VB.Frame freItem 
         Height          =   2925
         Index           =   1
         Left            =   780
         TabIndex        =   20
         Top             =   600
         Visible         =   0   'False
         Width           =   6285
         Begin goods.FTextBox txtName 
            Height          =   300
            Index           =   1
            Left            =   2175
            TabIndex        =   26
            Top             =   1275
            Width           =   2865
            _ExtentX        =   5054
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
         Begin goods.FTextBox txtNo 
            Height          =   300
            Index           =   1
            Left            =   2175
            TabIndex        =   21
            Top             =   735
            Width           =   2880
            _ExtentX        =   5080
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
            AutoSelAll      =   -1  'True
         End
         Begin goods.XPButton cmdExit 
            Height          =   345
            Index           =   1
            Left            =   3915
            TabIndex        =   22
            Top             =   2040
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   609
            Caption         =   "取消"
            CapAlign        =   2
            BackStyle       =   2
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mode            =   0
            Value           =   0   'False
            cBack           =   -2147483633
         End
         Begin goods.XPButton cmdOK 
            Height          =   345
            Index           =   1
            Left            =   2550
            TabIndex        =   23
            Top             =   2025
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   609
            Caption         =   "添加"
            CapAlign        =   2
            BackStyle       =   2
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "新宋体"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mode            =   0
            Value           =   0   'False
            cBack           =   -2147483633
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "类别名称"
            Height          =   180
            Left            =   1290
            TabIndex        =   25
            Top             =   1350
            Width           =   720
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "类别编号"
            Height          =   180
            Left            =   1305
            TabIndex        =   24
            Top             =   810
            Width           =   720
         End
      End
      Begin MSComctlLib.ListView List1 
         Height          =   3975
         Index           =   1
         Left            =   195
         TabIndex        =   19
         Top             =   375
         Width           =   9015
         _ExtentX        =   15901
         _ExtentY        =   7011
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
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   7
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "图标"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   1
            Text            =   "序号"
            Object.Width           =   1058
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   2
            Text            =   "名称"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "详细地址"
            Object.Width           =   4411
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   4
            Text            =   "联系人"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   5
            Text            =   "联系电话"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "备注"
            Object.Width           =   2540
         EndProperty
      End
      Begin goods.XPButton cmdExitOption 
         Height          =   345
         Index           =   1
         Left            =   9450
         TabIndex        =   17
         Top             =   2730
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   609
         Caption         =   "关闭(&Q)"
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
      Begin goods.XPButton cmdDel 
         Height          =   345
         Index           =   1
         Left            =   9450
         TabIndex        =   3
         Top             =   2090
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   609
         Caption         =   "删除(&D)"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mode            =   0
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin goods.XPButton cmdEdit 
         Height          =   345
         Index           =   1
         Left            =   9450
         TabIndex        =   4
         Top             =   1450
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   609
         Caption         =   "修改(&E)"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mode            =   0
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin goods.XPButton cmdAdd 
         Height          =   345
         Index           =   1
         Left            =   9450
         TabIndex        =   5
         Top             =   810
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   609
         Caption         =   "添加(&A)"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mode            =   0
         Value           =   0   'False
         cBack           =   -2147483633
      End
   End
   Begin VB.Frame Frame1 
      Height          =   6915
      Index           =   2
      Left            =   345
      TabIndex        =   1
      Top             =   1260
      Visible         =   0   'False
      Width           =   10680
      Begin VB.Frame freItem 
         Height          =   2535
         Index           =   2
         Left            =   435
         TabIndex        =   10
         Top             =   1155
         Visible         =   0   'False
         Width           =   4380
         Begin goods.FTextBox txtName 
            Height          =   300
            Index           =   2
            Left            =   1200
            TabIndex        =   11
            Top             =   570
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
            AutoSelAll      =   -1  'True
         End
         Begin goods.FTextBox txtDesc 
            Height          =   300
            Index           =   2
            Left            =   1200
            TabIndex        =   12
            Top             =   1320
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
            AutoSelAll      =   -1  'True
         End
         Begin goods.XPButton cmdExit 
            Height          =   345
            Index           =   2
            Left            =   2940
            TabIndex        =   13
            Top             =   1950
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   609
            Caption         =   "取消"
            CapAlign        =   2
            BackStyle       =   2
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mode            =   0
            Value           =   0   'False
            cBack           =   -2147483633
         End
         Begin goods.XPButton cmdOK 
            Height          =   345
            Index           =   2
            Left            =   1740
            TabIndex        =   14
            Top             =   1950
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   609
            Caption         =   "添加"
            CapAlign        =   2
            BackStyle       =   2
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "新宋体"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mode            =   0
            Value           =   0   'False
            cBack           =   -2147483633
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "规格名称"
            Height          =   180
            Left            =   360
            TabIndex        =   16
            Top             =   645
            Width           =   720
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "规格说明"
            Height          =   180
            Left            =   360
            TabIndex        =   15
            Top             =   1395
            Width           =   720
         End
      End
      Begin MSComctlLib.ListView List1 
         Height          =   3930
         Index           =   2
         Left            =   195
         TabIndex        =   6
         Top             =   840
         Width           =   8940
         _ExtentX        =   15769
         _ExtentY        =   6932
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   0   'False
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   7
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
            Alignment       =   2
            SubItemIndex    =   2
            Text            =   "名称"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   3
            Text            =   "详细地址"
            Object.Width           =   4410
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "联系人"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "联系电话"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "备注"
            Object.Width           =   2540
         EndProperty
      End
      Begin goods.XPButton cmdDel 
         Height          =   345
         Index           =   2
         Left            =   9345
         TabIndex        =   7
         Top             =   2595
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   609
         Caption         =   "删除(&D)"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mode            =   0
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin goods.XPButton cmdEdit 
         Height          =   345
         Index           =   2
         Left            =   9345
         TabIndex        =   8
         Top             =   1920
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   609
         Caption         =   "修改(&E)"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mode            =   0
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin goods.XPButton cmdAdd 
         Height          =   345
         Index           =   2
         Left            =   9345
         TabIndex        =   9
         Top             =   1245
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   609
         Caption         =   "添加(&A)"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mode            =   0
         Value           =   0   'False
         Enabled         =   0   'False
         cBack           =   -2147483633
      End
      Begin goods.XPButton cmdExitOption 
         Height          =   345
         Index           =   2
         Left            =   9345
         TabIndex        =   18
         Top             =   3270
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   609
         Caption         =   "关闭(&Q)"
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
      Begin VB.Label Label2 
         Caption         =   "此模块尚未编程"
         Height          =   375
         Left            =   9240
         TabIndex        =   27
         Top             =   840
         Width           =   1575
      End
   End
   Begin MSComctlLib.TabStrip tabOption 
      Height          =   735
      Left            =   6945
      TabIndex        =   0
      Top             =   60
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   1296
      MultiRow        =   -1  'True
      TabStyle        =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "类别"
            ImageVarType    =   2
            ImageIndex      =   1
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "物品"
            ImageVarType    =   2
            ImageIndex      =   2
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private intCurFrame As Integer     '当前显示的frame
Private curBrandIndex As Byte       '当前品牌
Private curSpecIndex As Byte       '当前规格
Dim brandid(100) As Integer
Dim specid(200) As Integer

Private Sub CboDec_Click()
    lblInfo.Visible = False
End Sub

Private Sub cmdSave_Click()
End Sub

Private Sub cmdAdd_Click(Index As Integer)
    Dim rs As ADODB.Recordset
    
    cmdOK(Index).caption = "添加"
    List1(Index).Visible = False
    freItem(Index).Visible = True
    Select Case Index
        Case 1, 2
            txtNo(Index).Text = ""
            txtName(Index).Text = ""
            txtName(Index).SetFocus
        
        Case 3
            Set rs = New ADODB.Recordset
            DBConnect
            rs.Open "select * from brand", Conn, 1, 1
            FCbBrand.Clear
            n = 0
            Do While Not rs.EOF
                FCbBrand.AddItem rs("bName")
                brandid(n) = rs("bid")
                n = n + 1
                rs.MoveNext
            
            Loop
            rs.Close
                        
            If n < 1 Then
                MsgBox "设置品牌后才能添加奶粉品种。", vbCritical, "未设品牌"
                List1(Index).Visible = True
                freItem(Index).Visible = False
                Exit Sub
            End If
                        
                        
                        
            rs.Open "select * from class", Conn, 1, 1
            FCbSpec.Clear
            n = 0
            Do While Not rs.EOF
                FCbSpec.AddItem rs("cName")
                specid(n) = rs("cid")
                n = n + 1
                rs.MoveNext
            
            Loop
            rs.Close
            
            If n < 1 Then
                MsgBox "设置奶粉规格后才能添加奶粉品种。", vbCritical, "未设规格"
                List1(Index).Visible = True
                freItem(Index).Visible = False
                Exit Sub
            End If
            
            FCbBrand.ListIndex = curBrandIndex
            FCbSpec.ListIndex = curSpecIndex
            
            txtName(Index).Text = ""
            txtDesc(Index).Text = ""
        
    End Select
    
    setOpCmd Index, False
    
End Sub

Private Sub cmdDel_Click(Index As Integer)
    Dim rs As ADODB.Recordset
    
    Set rs = New ADODB.Recordset
    DBConnect
    
    Select Case Index
        Case 1
            rs.Open "select top 1 sid from stock where sChain=" & GetID(List1(Index).SelectedItem.Key), Conn, 1, 1
        
            If Not rs.EOF Then
                MsgBox "已经使用了的连锁店不能删除！", vbExclamation, "参数设置"
                rs.Close
                Exit Sub
            End If
            rs.Close
    
            If MsgBox("确实删除类型名称为 [" & List1(Index).SelectedItem.SubItems(2) & "] 的连锁店吗？", vbExclamation + vbYesNo, "选项设置") = vbNo Then Exit Sub
    
            Conn.Execute "delete from chain where cid=" & GetID(List1(Index).SelectedItem.Key)
        
        Case 2
            rs.Open "select top 1 gid from goods where gClass=" & GetID(List1(Index).SelectedItem.Key), Conn, 1, 1
        
            If Not rs.EOF Then
                MsgBox "已经使用了的分类不能删除！", vbExclamation, "参数设置"
                rs.Close
                Exit Sub
            End If
            rs.Close
    
            If MsgBox("确实删除类型名称为 [" & List1(Index).SelectedItem.SubItems(2) & "] 的分类吗？", vbExclamation + vbYesNo, "选项设置") = vbNo Then Exit Sub
    
            Conn.Execute "delete from class where cid=" & GetID(List1(Index).SelectedItem.Key)
            
        Case 3
            rs.Open "select top 1 sid from stock where sGoods=" & GetID(List1(Index).SelectedItem.Key), Conn, 1, 1
        
            If Not rs.EOF Then
                MsgBox "已经使用了的品种不能删除！", vbExclamation, "参数设置"
                rs.Close
                Exit Sub
            End If
            rs.Close
    
            If MsgBox("确实删除类型名称为 [" & List1(Index).SelectedItem.SubItems(3) & "] 的品种吗？", vbExclamation + vbYesNo, "选项设置") = vbNo Then Exit Sub
    
            Conn.Execute "delete from goods where gid=" & GetID(List1(Index).SelectedItem.Key)
            
            
    End Select
    
    
    
    
    Conn.Close
    
    loadItemData Index
    
End Sub

Private Sub cmdEdit_Click(Index As Integer)
    cmdOK(Index).caption = "修改"
    
    Select Case Index
        Case 1, 2
            txtName(Index).Text = List1(Index).SelectedItem.SubItems(2)
            txtDesc(Index).Text = List1(Index).SelectedItem.SubItems(3)
    
            
        
        Case 3
            Dim rs As ADODB.Recordset
            Set rs = New ADODB.Recordset
            DBConnect
            
            strSQL = "select * from goods,brand,class where gid=" & GetID(List1(Index).SelectedItem.Key)
            rs.Open strSQL, Conn, 1, 1
            txtName(Index).Text = rs("gName")
            txtDesc(Index).Text = rs("gDescript")
            txtPrice.Text = Format(rs("gPrice"), ".00")
            gBrand = rs("gBrand")
            gSpec = rs("gSpec")
            rs.Close
            
            rs.Open "select * from brand", Conn, 1, 1
            FCbBrand.Clear
            n = 0
            Do While Not rs.EOF
                FCbBrand.AddItem rs("bName")
                brandid(n) = rs("bid")
                If gBrand = rs("bid") Then curBrandIndex = n
                n = n + 1
                rs.MoveNext
            
            Loop
            rs.Close
                        
            rs.Open "select * from class", Conn, 1, 1
            FCbSpec.Clear
            n = 0
            Do While Not rs.EOF
                FCbSpec.AddItem rs("cName")
                specid(n) = rs("cid")
                If gSpec = rs("cid") Then curSpecIndex = n
                n = n + 1
                rs.MoveNext
            
            Loop
            
            rs.Close
            Set rs = Nothing
            Conn.Close
            Set Conn = Nothing
            
            FCbBrand.ListIndex = curBrandIndex
            FCbSpec.ListIndex = curSpecIndex
            
    End Select
    
    List1(Index).Visible = False
    freItem(Index).Visible = True
      
    txtName(Index).SetFocus
    
setOpCmd Index, False
    
End Sub

Private Sub cmdExit_Click(Index As Integer)
    freItem(Index).Visible = False
    List1(Index).Visible = True
    setOpCmd Index, True
End Sub

Private Sub cmdExitOption_Click(Index As Integer)
    Unload Me
End Sub

Private Sub cmdOK_Click(Index As Integer)
    'On Error GoTo errmsg
    Dim ctlObject As Object
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    DBConnect
    
    
    For Each ctlObject In Me
        If TypeName(ctlObject) = "FTextBox" Then
            If ctlObject.Text = "" And (Not ctlObject.isNumber) Then ctlObject.Text = "-"
        End If
                            
    Next
    
    
    Select Case Index
        Case 1
            If Trim(txtName(Index).Text) = "" Then
                MsgBox "店名未填写！", vbExclamation, "参数设置"
                txtName(Index).SetFocus
                Exit Sub
            End If
    
    
            If cmdOK(Index).caption = "添加" Then
                strSQL = "select * from chain where cName='" & Trim(txtName(Index).Text) & "'"
                rs.Open strSQL, Conn, 1, 1
                recNum = rs.RecordCount
                rs.Close
                
                If recNum > 0 Then
                    MsgBox "该店名已存在！", vbCritical, "添加连锁店错误"
                    Exit Sub
                Else
                    
                    Conn.Execute "insert into chain(cName,cAddr,cContact,cTel,cDescript) values('" & _
                           Trim(txtName(Index).Text) & "','" & Trim(txtAddr(Index).Text) & "','" & Trim(txtContact(Index).Text) & "','" & _
                           Trim(txtTel(Index).Text) & "','" & IIf(Trim(txtDesc(Index).Text) <> "-", Trim(txtDesc(Index).Text), "无") & "')"
                End If
        
            Else
                Conn.Execute "update chain set cName='" & Trim(txtName(Index).Text) & "'," & _
                                  "cAddr='" & Trim(txtAddr(Index).Text) & "'," & _
                                  "cContact='" & Trim(txtContact(Index).Text) & "'," & _
                                  "cTel='" & Trim(txtTel(Index).Text) & "'," & _
                                  "cDescript='" & IIf(Trim(txtDesc(Index).Text) <> "-", Trim(txtDesc(Index).Text), "无") & "'  " & _
                                  "where cid=" & GetID(List1(Index).SelectedItem.Key)
    
            End If
    
            Conn.Close
    
            txtName(Index).Text = ""
            txtDesc(Index).Text = ""
        
        Case 2
            If Trim(txtName(Index).Text) = "" Then
                MsgBox "分类名称未填写！", vbExclamation, "参数设置"
                txtName(Index).SetFocus
                Exit Sub
            End If
    
    
            If cmdOK(Index).caption = "添加" Then
                strSQL = "select * from class where cName='" & Trim(txtName(Index).Text) & "'"
                rs.Open strSQL, Conn, 1, 1
                recNum = rs.RecordCount
                rs.Close
                
                If recNum > 0 Then
                    MsgBox "该分类名称已存在！", vbCritical, "添加分类错误"
                    Exit Sub
                Else
                    Conn.Execute "insert into class(cName,cDescript) values('" & _
                           Trim(txtName(Index).Text) & "','" & IIf(Trim(txtDesc(Index).Text) <> "", Trim(txtDesc(Index).Text), "无") & "')"
                End If
        
            Else
                Conn.Execute "update class set cName='" & Trim(txtName(Index).Text) & "'," & _
                                  "cDescript=" & IIf(Trim(txtDesc(Index).Text) <> "", "'" & Trim(txtDesc(Index).Text) & "'", "'无'") & " " & _
                                  "where cid=" & GetID(List1(Index).SelectedItem.Key)
    
            End If
    
            Conn.Close
    
        
        Case 3
            If Trim(txtName(Index).Text) = "" Then
                MsgBox "品种名称未填写！", vbExclamation, "参数设置"
                txtName(Index).SetFocus
                Exit Sub
            End If
    
            curBrandIndex = FCbBrand.ListIndex
            curSpecIndex = FCbSpec.ListIndex
    
            If cmdOK(Index).caption = "添加" Then
                Conn.Execute "insert into goods(gBrand,gName,gSpec,gPrice,gDescript) values(" & brandid(curBrandIndex) & ",'" & _
                       Trim(txtName(Index).Text) & "'," & specid(curSpecIndex) & "," & txtPrice.Text & ",'" & IIf(Trim(txtDesc(Index).Text) <> "", Trim(txtDesc(Index).Text), "无") & "')"
        
            Else
                Conn.Execute "update goods set gName='" & Trim(txtName(Index).Text) & "'," & _
                                   "gBrand=" & brandid(curBrandIndex) & "," & _
                                   "gSpec=" & specid(curSpecIndex) & "," & _
                                   "gPrice=" & txtPrice.Text & "," & _
                                  "gDescript=" & IIf(Trim(txtDesc(Index).Text) <> "", "'" & Trim(txtDesc(Index).Text) & "'", "'无'") & " " & _
                                  "where gid=" & GetID(List1(Index).SelectedItem.Key)
    
            End If
    
            Conn.Close
              
    End Select
    
    txtName(Index).Text = ""
    txtAddr(Index).Text = ""
    txtContact(Index).Text = ""
    txtTel(Index).Text = ""
    txtDesc(Index).Text = ""
    
    
    loadItemData Index
    
    List1(Index).Visible = True
    freItem(Index).Visible = False
    
    setOpCmd Index, True

    Exit Sub
errmsg:
    MsgBox Err.Description, vbCritical, "参数设置"
    
End Sub

Private Sub cmdSaveCon_Click()
    On Error GoTo errmsg
    Dim rs As ADODB.Recordset
    Dim strSQL As String
    
    Set rs = New ADODB.Recordset
    DBConnect
    
    strSQL = "select * from ItemInfo where ItemType=3"
    rs.Open strSQL, Conn, 1, 1
    If rs.EOF Then
        strSQL = "insert into ItemInfo(ItemType,ItemValue) values(3," & CboDec.ListIndex & ")"
    Else
        strSQL = "update ItemInfo set ItemValue=" & CboDec.ListIndex & " where ItemType=3"
    End If
    
    rs.Close
    Conn.Execute strSQL
    Conn.Close
    
    lblInfo.Visible = True
    bytAfterDec = CboDec.ListIndex
    Exit Sub
    
errmsg:
    MsgBox Err.Description, vbCritical, "选项设置"

End Sub

Private Sub cmdSet_Click(Index As Integer)
    On Error GoTo errmsg
    Label9.Visible = False
    
    ComDlg.CancelError = True
    ComDlg.ShowColor
    
    
    Exit Sub
    
errmsg:
    
End Sub

Private Sub Form_Load()
    intCurFrame = 1
    loadItemData (1)
    
    
End Sub

Private Sub Form_Resize()
On Error Resume Next
    'PicTop.Width = Width / 15 - 16
    Cls
    Line (2, 2)-(Width / 15 - 14, Height / 15 - 29), 10921638, B
    tabOption.Width = Width
    tabOption.Top = 0
    tabOption.Left = 0
    tabOption.Height = Height - 1500
    
    For i = 1 To 2
        Frame1(i).Top = tabOption.ClientTop
        Frame1(i).Left = tabOption.Left
        Frame1(i).Height = tabOption.Height
        Frame1(i).Width = tabOption.Width
    Next
    List1(1).Height = Frame1(1).Height
    List1(2).Width = List1(1).Width
    List1(2).Height = List1(1).Height
    List1(2).Top = List1(1).Top
    List1(2).Left = List1(1).Left
    'List1(1).ColumnHeaders.Item(4).Width = List1(1).Width - List1(1).ColumnHeaders.Item(1).Width - List1(1).ColumnHeaders.Item(2).Width - List1(1).ColumnHeaders.Item(3).Width - 90
    'List1(2).ColumnHeaders.Item(4).Width = List1(2).Width - List1(2).ColumnHeaders.Item(1).Width - List1(2).ColumnHeaders.Item(2).Width - List1(1).ColumnHeaders.Item(3).Width - 90
    

End Sub

Private Sub tabOption_Click()
    If tabOption.SelectedItem.Index = intCurFrame Then Exit Sub
    Frame1(tabOption.SelectedItem.Index).Visible = True
    Frame1(intCurFrame).Visible = False
    intCurFrame = tabOption.SelectedItem.Index
    loadItemData tabOption.SelectedItem.Index
End Sub
Sub loadItemData(intTabIndex As Integer)
    Dim strSQL As String
    Dim rs As ADODB.Recordset
    Dim item As ListItem
    Dim AfterDec As Integer
    
    Set rs = New ADODB.Recordset
    DBConnect
    
    Select Case intTabIndex
        Case 1        '1-连锁店
            strSQL = "select * from chain order by cName"
            rs.Open strSQL, Conn, 1, 1
            List1(intTabIndex).ListItems.Clear
            
            Do While Not rs.EOF
                iNo = iNo + 1
                Set item = List1(intTabIndex).ListItems.Add(, rs("cid") & "k")
                item.SubItems(1) = iNo
                item.SubItems(2) = rs("cName")
                item.SubItems(3) = rs("cAddr")
                item.SubItems(4) = rs("cContact")
                item.SubItems(5) = rs("cTel")
                item.SubItems(6) = rs("cDescript")
                rs.MoveNext
            Loop
            
            rs.Close
            
            If List1(intTabIndex).ListItems.Count > 0 Then
                cmdEdit(intTabIndex).Enabled = True
                cmdDel(intTabIndex).Enabled = True
            Else
                cmdEdit(intTabIndex).Enabled = False
                cmdDel(intTabIndex).Enabled = False
            
            End If
            
            
        Case 2   '分类
            strSQL = "select * from supplier"
            rs.Open strSQL, Conn, 1, 1
            List1(intTabIndex).ListItems.Clear
            
            Do While Not rs.EOF
                iNo = iNo + 1
                Set item = List1(intTabIndex).ListItems.Add(, rs("sid") & "k")
                item.SubItems(1) = iNo
                item.SubItems(2) = rs("sName")
                item.SubItems(3) = rs("sAddr")
                item.SubItems(3) = rs("sContact")
                item.SubItems(3) = rs("sTel")
                item.SubItems(3) = rs("sDescript")
                rs.MoveNext
            Loop
            
            rs.Close
            
            If List1(intTabIndex).ListItems.Count > 0 Then
                cmdEdit(intTabIndex).Enabled = True
                cmdDel(intTabIndex).Enabled = True
            Else
                cmdEdit(intTabIndex).Enabled = False
                cmdDel(intTabIndex).Enabled = False
            
            End If
        
        Case 3
            strSQL = "select * from brand,class,goods where goods.gBrand=brand.bid and goods.gSpec=class.cid"
            rs.Open strSQL, Conn, 1, 1
            List1(intTabIndex).ListItems.Clear
            
            Do While Not rs.EOF
                iNo = iNo + 1
                Set item = List1(intTabIndex).ListItems.Add(, rs("gid") & "k")
                item.SubItems(1) = iNo
                item.SubItems(2) = rs("bName")
                item.SubItems(3) = rs("gName")
                item.SubItems(4) = rs("cName")
                item.SubItems(5) = Format(rs("gPrice"), ".00")
                item.SubItems(6) = rs("gDescript")
                rs.MoveNext
            Loop
            
            
            rs.Close
            
            If List1(intTabIndex).ListItems.Count > 0 Then
                cmdEdit(intTabIndex).Enabled = True
                cmdDel(intTabIndex).Enabled = True
            Else
                cmdEdit(intTabIndex).Enabled = False
                cmdDel(intTabIndex).Enabled = False
            
            End If
        

        End Select
            
    If rs.state <> 0 Then rs.Close
    Conn.Close


End Sub

Private Sub txtColor_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Label9.Visible = False
End Sub

Private Sub txtDesc_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
         cmdAdd(Index).SetFocus
    End If

End Sub

Private Sub txtID_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        txtName(Index).SetFocus
    End If
End Sub

Private Sub txtName_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        txtDesc(Index).SetFocus
    End If
    
End Sub


Private Sub setOpCmd(Index As Integer, state As Boolean)
    
    cmdAdd(Index).Enabled = state
    cmdEdit(Index).Enabled = state
    cmdDel(Index).Enabled = state
End Sub

