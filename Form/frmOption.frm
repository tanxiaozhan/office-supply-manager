VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmOption 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "办公用品基本参数设置"
   ClientHeight    =   10710
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12570
   Icon            =   "frmOption.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   714
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   838
   Begin MSComDlg.CommonDialog ComDlg 
      Left            =   600
      Top             =   6930
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   8310
      Top             =   3795
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
            Picture         =   "frmOption.frx":0E42
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOption.frx":13DC
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOption.frx":1976
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOption.frx":1F10
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      Height          =   4575
      Index           =   1
      Left            =   3240
      TabIndex        =   2
      Top             =   2535
      Width           =   8445
      Begin VB.Frame freItem 
         Height          =   2205
         Index           =   1
         Left            =   540
         TabIndex        =   21
         Top             =   690
         Visible         =   0   'False
         Width           =   4380
         Begin goods.FTextBox txtName 
            Height          =   300
            Index           =   1
            Left            =   1200
            TabIndex        =   6
            Top             =   480
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
            Index           =   1
            Left            =   2940
            TabIndex        =   22
            Top             =   1515
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
            Left            =   1740
            TabIndex        =   16
            Top             =   1500
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
         Begin goods.FTextBox txtDesc 
            Height          =   300
            Index           =   1
            Left            =   1200
            TabIndex        =   11
            Top             =   930
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
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "备注说明"
            Height          =   180
            Left            =   360
            TabIndex        =   34
            Top             =   1005
            Width           =   720
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "类别名称"
            Height          =   165
            Left            =   360
            TabIndex        =   23
            Top             =   555
            Width           =   720
         End
      End
      Begin MSComctlLib.ListView List1 
         Height          =   3975
         Index           =   1
         Left            =   195
         TabIndex        =   20
         Top             =   375
         Width           =   6510
         _ExtentX        =   11483
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
         NumItems        =   4
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
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   3
            Text            =   "说明"
            Object.Width           =   5292
         EndProperty
      End
      Begin goods.XPButton cmdExitOption 
         Height          =   345
         Index           =   1
         Left            =   7080
         TabIndex        =   17
         Top             =   2670
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
         Left            =   7065
         TabIndex        =   3
         Top             =   2100
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
         Left            =   7050
         TabIndex        =   4
         Top             =   1530
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
         Left            =   7035
         TabIndex        =   7
         Top             =   990
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
      Height          =   4575
      Index           =   2
      Left            =   4155
      TabIndex        =   1
      Top             =   5685
      Visible         =   0   'False
      Width           =   6615
      Begin VB.Frame freItem 
         Height          =   4035
         Index           =   2
         Left            =   -75
         TabIndex        =   12
         Top             =   150
         Visible         =   0   'False
         Width           =   5145
         Begin goods.FCombo fcmbUnit 
            Height          =   300
            Left            =   1200
            TabIndex        =   44
            Top             =   1965
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
            ListIndex       =   -1
         End
         Begin goods.FCombo fcmbClass 
            Height          =   300
            Left            =   1200
            TabIndex        =   40
            Top             =   375
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
            ListIndex       =   -1
         End
         Begin goods.FTextBox txtName 
            Height          =   300
            Index           =   2
            Left            =   1215
            TabIndex        =   41
            Top             =   780
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
            TabIndex        =   45
            Top             =   2325
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
            TabIndex        =   47
            Top             =   2805
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
            TabIndex        =   46
            Top             =   2835
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
         Begin goods.FTextBox txtModel 
            Height          =   300
            Index           =   2
            Left            =   1200
            TabIndex        =   42
            Top             =   1185
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
         Begin goods.FTextBox txtPrice 
            Height          =   300
            Index           =   2
            Left            =   1200
            TabIndex        =   43
            Top             =   1590
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
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "计量单位"
            Height          =   180
            Left            =   330
            TabIndex        =   38
            Top             =   2040
            Width           =   720
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "类    别"
            Height          =   180
            Left            =   405
            TabIndex        =   37
            Top             =   465
            Width           =   720
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "单　　价"
            Height          =   180
            Left            =   360
            TabIndex        =   36
            Top             =   1665
            Width           =   720
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "规格型号"
            Height          =   180
            Left            =   360
            TabIndex        =   35
            Top             =   1260
            Width           =   720
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "物品名称"
            Height          =   180
            Left            =   360
            TabIndex        =   15
            Top             =   855
            Width           =   720
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "备注说明"
            Height          =   180
            Left            =   360
            TabIndex        =   14
            Top             =   2400
            Width           =   720
         End
      End
      Begin MSComctlLib.ListView List1 
         Height          =   3615
         Index           =   2
         Left            =   855
         TabIndex        =   8
         Top             =   555
         Width           =   4725
         _ExtentX        =   8334
         _ExtentY        =   6376
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
         Appearance      =   1
         NumItems        =   8
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
            Text            =   "类别"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "物品名称"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   4
            Text            =   "规格型号"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   5
            Text            =   "单价(元)"
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   6
            Text            =   "单位"
            Object.Width           =   882
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "备注说明"
            Object.Width           =   2540
         EndProperty
      End
      Begin goods.XPButton cmdDel 
         Height          =   345
         Index           =   2
         Left            =   5205
         TabIndex        =   29
         Top             =   2565
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
         Left            =   5205
         TabIndex        =   19
         Top             =   1890
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
         Left            =   5205
         TabIndex        =   9
         Top             =   1215
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
      Begin goods.XPButton cmdExitOption 
         Height          =   345
         Index           =   2
         Left            =   5205
         TabIndex        =   39
         Top             =   3240
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
   End
   Begin VB.Frame Frame1 
      Height          =   3405
      Index           =   3
      Left            =   2445
      TabIndex        =   25
      Top             =   1155
      Width           =   6090
      Begin VB.Frame freItem 
         Height          =   2145
         Index           =   3
         Left            =   420
         TabIndex        =   28
         Top             =   720
         Visible         =   0   'False
         Width           =   4140
         Begin goods.FTextBox txtDesc 
            Height          =   300
            Index           =   3
            Left            =   1155
            TabIndex        =   10
            Top             =   855
            Width           =   2640
            _ExtentX        =   4657
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
         Begin goods.XPButton cmdExit 
            Height          =   345
            Index           =   3
            Left            =   1950
            TabIndex        =   18
            Top             =   1500
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   609
            Caption         =   "取消"
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
         Begin goods.XPButton cmdOK 
            Height          =   345
            Index           =   3
            Left            =   555
            TabIndex        =   13
            Top             =   1500
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   609
            Caption         =   "添加"
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
         Begin goods.FTextBox txtName 
            Height          =   300
            Index           =   3
            Left            =   1155
            TabIndex        =   5
            Top             =   420
            Width           =   1155
            _ExtentX        =   2037
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
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "说　　明"
            Height          =   180
            Left            =   300
            TabIndex        =   33
            Top             =   930
            Width           =   720
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "计量单位"
            Height          =   180
            Left            =   300
            TabIndex        =   30
            Top             =   480
            Width           =   720
         End
      End
      Begin MSComctlLib.ListView List1 
         Height          =   2880
         Index           =   3
         Left            =   180
         TabIndex        =   26
         Top             =   300
         Width           =   4530
         _ExtentX        =   7990
         _ExtentY        =   5080
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FlatScrollBar   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   4
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
            SubItemIndex    =   3
            Text            =   "说明"
            Object.Width           =   4410
         EndProperty
      End
      Begin goods.XPButton cmdExitOption 
         Height          =   345
         Index           =   3
         Left            =   4905
         TabIndex        =   24
         Top             =   2235
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
         Index           =   3
         Left            =   4890
         TabIndex        =   27
         Top             =   1665
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
         Index           =   3
         Left            =   4875
         TabIndex        =   31
         Top             =   1095
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
         Index           =   3
         Left            =   4860
         TabIndex        =   32
         Top             =   555
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
   Begin MSComctlLib.TabStrip tabOption 
      Height          =   735
      Left            =   4215
      TabIndex        =   0
      Top             =   225
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   1296
      MultiRow        =   -1  'True
      TabStyle        =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   3
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
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "计量单位"
            ImageVarType    =   2
            ImageIndex      =   3
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmOption"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private intCurFrame As Integer     '当前显示的frame
Private curBrandIndex As Byte       '当前品牌
Private curSpecIndex As Byte       '当前规格
Private curUnitIndex As Byte
Dim iClass(100) As Integer      '类别
Dim strClass(100) As String
Dim iUnit(30) As Integer      '物品单位
Dim strUnit(30) As String

Private Sub CboDec_Click()
    lblInfo.Visible = False
End Sub

Private Sub cmdSave_Click()
End Sub

Private Sub cmdAdd_Click(Index As Integer)
    
    cmdOK(Index).caption = "添加"
    List1(Index).Visible = False
    freItem(Index).Visible = True
    Select Case Index
        Case 2
            txtModel(Index).Text = ""
            txtPrice(Index).Text = ""
    End Select
    
    
    
    txtDesc(Index).Text = ""
    txtName(Index).Text = ""
    txtName(Index).SetFocus
    
    setOpCmd Index, False
End Sub

Private Sub cmdDel_Click(Index As Integer)
    Dim rs As ADODB.Recordset
    
    Set rs = New ADODB.Recordset
    DBConnect
    
    Select Case Index
        Case 1
            rs.Open "select top 1 gid from goods where gClass=" & GetID(List1(Index).SelectedItem.Key), Conn, 1, 1
        
            If Not rs.EOF Then
                MsgBox "已经使用了的类别不能删除！", vbExclamation, "参数设置"
                rs.Close
                Exit Sub
            End If
            rs.Close
    
            If MsgBox("确实删除名称为 [" & List1(Index).SelectedItem.SubItems(2) & "] 的类别吗？", vbExclamation + vbYesNo, "参数设置") = vbNo Then Exit Sub
    
            Conn.Execute "delete from class where cid=" & GetID(List1(Index).SelectedItem.Key)
        
        Case 2
            'rs.Open "select top 1 gid from goods where gClass=" & GetID(List1(Index).SelectedItem.Key), Conn, 1, 1
        
            'If Not rs.EOF Then
            '    MsgBox "已经使用了的物品不能删除！", vbExclamation, "参数设置"
            '    rs.Close
            '    Exit Sub
            'End If
            'rs.Close
    
            If MsgBox("确实删除名称为 [" & List1(Index).SelectedItem.SubItems(3) & "] 的物品吗？", vbExclamation + vbYesNo, "选项设置") = vbNo Then Exit Sub
    
            Conn.Execute "delete from goods where gid=" & GetID(List1(Index).SelectedItem.Key)
            
        Case 3
            rs.Open "select top 1 gid from goods where gUnit=" & GetID(List1(Index).SelectedItem.Key), Conn, 1, 1
        
            If Not rs.EOF Then
                MsgBox "已经使用了的数量单位不能删除！", vbExclamation, "参数设置"
                rs.Close
                Exit Sub
            End If
            rs.Close
    
            If MsgBox("确实删除名称为 [" & List1(Index).SelectedItem.SubItems(2) & "] 的数量单位吗？", vbExclamation + vbYesNo, "选项设置") = vbNo Then Exit Sub
    
            Conn.Execute "delete from unit where uid=" & GetID(List1(Index).SelectedItem.Key)
            
            
    End Select
    
    
    
    
    Conn.Close
    
    loadItemData Index
    
End Sub

Private Sub cmdEdit_Click(Index As Integer)
    cmdOK(Index).caption = "修改"
    
    Select Case Index
        Case 1, 3
            txtName(Index).Text = List1(Index).SelectedItem.SubItems(2)
            txtDesc(Index).Text = List1(Index).SelectedItem.SubItems(3)
        
        Case 2
            fcmbClass.Text = List1(Index).SelectedItem.SubItems(2)
            txtName(Index).Text = List1(Index).SelectedItem.SubItems(3)
            txtModel(Index).Text = List1(Index).SelectedItem.SubItems(4)
            txtPrice(Index).Text = List1(Index).SelectedItem.SubItems(5)
            fcmbUnit.Text = List1(Index).SelectedItem.SubItems(6)
            txtDesc(Index).Text = List1(Index).SelectedItem.SubItems(7)
            
        
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
    Dim rs As ADODB.Recordset
    Dim curClass As Integer, curUnit As Integer
    Dim strClassName As String, strUnitName As String
    
    Set rs = New ADODB.Recordset
    DBConnect
    
    Select Case Index
        Case 1
            If Trim(txtName(Index).Text) = "" Then
                MsgBox "类别名称未填写！", vbExclamation, "参数设置"
                txtName(Index).SetFocus
                Exit Sub
            End If
    
    
            If cmdOK(Index).caption = "添加" Then
                strSQL = "select * from class where cName='" & Trim(txtName(Index).Text) & "'"
                rs.Open strSQL, Conn, 1, 1
                recNum = rs.RecordCount
                rs.Close
                
                If recNum > 0 Then
                    MsgBox "该类别已存在！", vbCritical, "添加品类错误"
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
            
            getComcoBoxData
    
            txtName(Index).Text = ""
            txtDesc(Index).Text = ""
        
        Case 2
            
            
            If Trim(txtName(Index).Text) = "" Then
                MsgBox "用品名称未填写！", vbExclamation, "参数设置"
                txtName(Index).SetFocus
                Exit Sub
            End If
            
            '类别ID、单位ID
            If fcmbClass.ListIndex <> 1000 Then
                strClassName = fcmbClass.Text
                For i = 0 To fcmbClass.ListCount - 1
                    If strClassName = fcmbClass.List(i) Then
                        fcmbClass.ListIndex = i
                        Exit For
                    End If
                Next
                If i >= fcmbClass.ListCount Then
                    addNewClassName strClassName, fcmbClass.ListCount     '如果不存在类别，增加新的类别
                    fcmbClass.AddItem strClassName
                    fcmbClass.ListIndex = fcmbClass.ListCount - 1
                End If
            End If
            
            If fcmbUnit.ListIndex <> 1000 Then
                strUnitName = fcmbUnit.Text
                For i = 0 To fcmbUnit.ListCount - 1
                    If strUnitName = fcmbUnit.List(i) Then
                        fcmbUnit.ListIndex = i
                        Exit For
                    End If
                Next
                If i >= fcmbUnit.ListCount Then
                    addNewUnitName strUnitName, fcmbUnit.ListCount     '如果不存在的度量单位，增加新的度量单位
                    fcmbUnit.AddItem strUnitName
                    fcmbUnit.ListIndex = fcmbUnit.ListCount - 1
                End If
            End If
            
            curClass = fcmbClass.ListIndex
            curUnit = fcmbUnit.ListIndex
            
            
            
    
            If cmdOK(Index).caption = "添加" Then
                strSQL = "select * from goods where gName='" & Trim(txtName(Index).Text) & "'"
                rs.Open strSQL, Conn, 1, 1
                recNum = rs.RecordCount
                rs.Close
                
                If recNum > 0 Then
                    MsgBox "该用品名称已存在！", vbCritical, "添加用品错误"
                    Exit Sub
                Else
                    Conn.Execute "insert into goods(gClass,gName,gSpec,gPrice,gUnit,gDescript) values(" & _
                           iClass(curClass) & ",'" & Trim(txtName(Index).Text) & "','" & _
                           txtModel(Index).Text & "'," & Format(txtPrice(Index).Text, "####0.00") & "," & _
                           iUnit(curUnit) & ",'" & IIf(Trim(txtDesc(Index).Text) <> "", Trim(txtDesc(Index).Text), "无") & "')"
                           
                End If
        
            Else
                Conn.Execute "update goods set gName='" & Trim(txtName(Index).Text) & "'," & _
                                  "gClass=" & iClass(curClass) & "," & _
                                  "gSpec='" & Trim(txtModel(Index).Text) & "'," & _
                                  "gPrice=" & Format(txtPrice(Index).Text, "####0.00") & "," & _
                                  "gUnit=" & iUnit(curUnit) & "," & _
                                  "gDescript=" & IIf(Trim(txtDesc(Index).Text) <> "", "'" & Trim(txtDesc(Index).Text) & "'", "'无'") & " " & _
                                  "where gid=" & GetID(List1(Index).SelectedItem.Key)
    
            End If
    
            Conn.Close
    
            txtName(Index).Text = ""
            txtDesc(Index).Text = ""
        
              
        Case 3
            If Trim(txtName(Index).Text) = "" Then
                MsgBox "数量单位名称未填写！", vbExclamation, "参数设置"
                txtName(Index).SetFocus
                Exit Sub
            End If
    
    
            If cmdOK(Index).caption = "添加" Then
                strSQL = "select * from unit where uName='" & Trim(txtName(Index).Text) & "'"
                rs.Open strSQL, Conn, 1, 1
                recNum = rs.RecordCount
                rs.Close
                
                If recNum > 0 Then
                    MsgBox "该数量单位名称已存在！", vbCritical, "添加数量单位错误"
                    Exit Sub
                Else
                    Conn.Execute "insert into unit(uName,uDescript) values('" & _
                           Trim(txtName(Index).Text) & "','" & IIf(Trim(txtDesc(Index).Text) <> "", Trim(txtDesc(Index).Text), "无") & "')"
                End If
        
            Else
                Conn.Execute "update unit set uName='" & Trim(txtName(Index).Text) & "'," & _
                                  "uDescript=" & IIf(Trim(txtDesc(Index).Text) <> "", "'" & Trim(txtDesc(Index).Text) & "'", "'无'") & " " & _
                                  "where uid=" & GetID(List1(Index).SelectedItem.Key)
    
            End If
    
            Conn.Close
    
            txtName(Index).Text = ""
            txtDesc(Index).Text = ""
    End Select
    
    
    
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

Private Sub Form_Activate()
    '填充增加物品窗口中的类别、单位组合框数据
    getComcoBoxData

End Sub

Private Sub Form_Load()
    Width = 8735
    Height = 5600

    intCurFrame = 1
    loadItemData (1)
    
    
End Sub

Private Sub Form_Resize()
On Error Resume Next
    'PicTop.Width = Width / 15 - 16
    Cls
    Line (2, 2)-(Width / 15 - 14, Height / 15 - 29), 10921638, B
    tabOption.Width = Width / 15
    tabOption.Top = 0
    tabOption.Left = 0
    tabOption.Height = Height / 15
    
    For i = 1 To 3
        Frame1(i).Top = tabOption.ClientTop
        Frame1(i).Left = tabOption.Left
        Frame1(i).Height = tabOption.Height
        Frame1(i).Width = tabOption.Width
    Next
    For i = 2 To 3
        List1(i).Top = List1(1).Top
        List1(i).Left = List1(1).Left
        List1(i).Height = List1(1).Height
        List1(i).Width = List1(1).Width
        
        cmdAdd(i).Top = cmdAdd(1).Top
        cmdEdit(i).Top = cmdEdit(1).Top
        cmdDel(i).Top = cmdDel(1).Top
        cmdExitOption(i).Top = cmdExitOption(1).Top
        cmdAdd(i).Left = cmdAdd(1).Left
        cmdEdit(i).Left = cmdEdit(1).Left
        cmdDel(i).Left = cmdDel(1).Left
        cmdExitOption(i).Left = cmdExitOption(1).Left
        
    Next
    
    List1(1).ColumnHeaders.item(4).Width = List1(1).Width - List1(1).ColumnHeaders.item(1).Width - List1(1).ColumnHeaders.item(2).Width - List1(1).ColumnHeaders.item(3).Width - 90
    List1(3).ColumnHeaders.item(4).Width = List1(2).Width - List1(2).ColumnHeaders.item(1).Width - List1(2).ColumnHeaders.item(2).Width - List1(1).ColumnHeaders.item(3).Width - 90
    

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
        Case 1        '1-类别
            strSQL = "select * from class order by cName"
            rs.Open strSQL, Conn, 1, 1
            List1(intTabIndex).ListItems.Clear
            
            Do While Not rs.EOF
                iNo = iNo + 1
                Set item = List1(intTabIndex).ListItems.Add(, rs("cid") & "k")
                item.SubItems(1) = iNo
                item.SubItems(2) = rs("cName").value
                item.SubItems(3) = rs("cDescript").value
                rs.MoveNext
            Loop
            
            
            If List1(intTabIndex).ListItems.Count > 0 Then
                cmdEdit(intTabIndex).Enabled = True
                cmdDel(intTabIndex).Enabled = True
            Else
                cmdEdit(intTabIndex).Enabled = False
                cmdDel(intTabIndex).Enabled = False
            
            End If
            
            
        Case 2   '物品名称
            strSQL = "select * from goods,class,unit where gClass=cid and gUnit=uid order by gClass"
            rs.Open strSQL, Conn, 1, 1
            List1(intTabIndex).ListItems.Clear
            
            Do While Not rs.EOF
                iNo = iNo + 1
                Set item = List1(intTabIndex).ListItems.Add(, rs("gid") & "k")
                item.SubItems(1) = iNo
                item.SubItems(2) = rs("cName").value
                item.SubItems(3) = rs("gName").value
                item.SubItems(4) = rs("gSpec").value
                item.SubItems(5) = Format(rs("gPrice").value, "###0.00")
                item.SubItems(6) = rs("uName").value
                item.SubItems(7) = rs("gDescript").value
                rs.MoveNext
            Loop
            
            If List1(intTabIndex).ListItems.Count > 0 Then
                cmdEdit(intTabIndex).Enabled = True
                cmdDel(intTabIndex).Enabled = True
            Else
                cmdEdit(intTabIndex).Enabled = False
                cmdDel(intTabIndex).Enabled = False
            
            End If
        
        
        Case 3        '单位
            strSQL = "select * from unit"
            rs.Open strSQL, Conn, 1, 1
            List1(intTabIndex).ListItems.Clear
            
            Do While Not rs.EOF
                iNo = iNo + 1
                Set item = List1(intTabIndex).ListItems.Add(, rs("uid") & "k")
                item.SubItems(1) = iNo
                item.SubItems(2) = rs("uName")
                item.SubItems(3) = rs("uDescript")
                rs.MoveNext
            Loop
            
            
            If List1(intTabIndex).ListItems.Count > 0 Then
                cmdEdit(intTabIndex).Enabled = True
                cmdDel(intTabIndex).Enabled = True
            Else
                cmdEdit(intTabIndex).Enabled = False
                cmdDel(intTabIndex).Enabled = False
            
            End If

        End Select
            
    If rs.state <> 0 Then rs.Close
    Set rs = Nothing
    
    If Conn.state <> 0 Then Conn.Close
    Set Conn = Nothing


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

Private Sub getComcoBoxData()
    Dim rs As ADODB.Recordset
    Dim n As Integer
    
    fcmbClass.Clear
    fcmbUnit.Clear
    n = 0
    
    DBConnect
    
    Set rs = New ADODB.Recordset
    rs.Open "select * from class order by cName", Conn, 1, 1
    Do While Not rs.EOF
        fcmbClass.AddItem rs("cName")
        iClass(n) = rs("cid")
        n = n + 1
        rs.MoveNext
    Loop
    rs.Close
    
    If fcmbClass.ListCount > 0 Then fcmbClass.ListIndex = 0
    
    n = 0
    rs.Open "select * from unit order by uName", Conn, 1, 1
    Do While Not rs.EOF
        fcmbUnit.AddItem rs("uName")
        iUnit(n) = rs("uid")
        n = n + 1
        rs.MoveNext
    Loop
    rs.Close
    
    If fcmbUnit.ListCount > 0 Then fcmbUnit.ListIndex = 0
    
    Set rs = Nothing
End Sub
Private Sub addNewClassName(strClassName As String, iClassCount As Integer)
    Dim rs As ADODB.Recordset
    
    DBConnect
    Conn.Execute "insert into class(cName,cDescript) values('" & strClassName & "','-')"
    Set rs = New ADODB.Recordset
    rs.Open "select top 1 cid from class  order by cid desc", Conn, 1, 1
    iClass(iClassCount) = rs("cid")
    rs.Close
    Set rs = Nothing

End Sub

Private Sub addNewUnitName(strUnitName As String, iUnitCount As Integer)
    Dim rs As ADODB.Recordset
    
    DBConnect
    Conn.Execute "insert into unit(uName,uDescript) values('" & strUnitName & "','-')"
    Set rs = New ADODB.Recordset
    rs.Open "select top 1 uid from unit  order by uid desc", Conn, 1, 1
    iUnit(iUnitCount) = rs("uid")
    rs.Close
    Set rs = Nothing

End Sub
